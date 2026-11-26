using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using Xlide.Vbe.Core;
using Xlide.Vbe.Core.Editor;
using Xlide.Vbe.Core.Updates;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;
using Xlide.Vbe.Shim.Updates;

namespace Xlide.Vbe.Shim.AddIn;

/// <summary>
/// Self-update: the session's half. One state, shown to the page's dialog and answered by the
/// api's `update` route; the network and file work runs off the host thread in
/// <see cref="UpdateService"/> and every outcome crosses back here before anything is shown.
///
/// The installer is the same executable a first install uses, started by hand or by this. It
/// finds the Office host open, says so in its own window, and installs the moment the host
/// closes - so the session's last act is to tell the developer that, and nothing here ever
/// replaces a file the running library is loaded from.
/// </summary>
internal sealed partial class AddInSession
{
    private sealed class UpdateProgress
    {
        public string Phase = "idle";
        public bool Manual;
        public Release? Release;
        public string? CheckedAt;
        public long Received;
        public string? Path;
        public int Pid;
        public int? ExitCode;
        public string? Error;
    }

    /// <summary>The ui-state key holding the version the developer declined.</summary>
    private const string SkippedUpdateKey = "updateSkipped";

    /// <summary>How long after the surface is up the automatic check runs: out of the way of boot.</summary>
    private static readonly TimeSpan UpdateCheckDelay = TimeSpan.FromSeconds(8);

    private UpdateService? _updateService;
    private UpdateProgress _update = new();
    private CancellationTokenSource? _updateWork;
    private int _updateRound;
    private bool _updateCheckScheduled;
    private UpdateSource _updateSource = UpdateSource.GitHub;

    private UpdateService Updater => _updateService ??= new UpdateService();

    /// <summary>Downloads live beside the logs, under the same root the preferences use.</summary>
    private static string UpdatesFolder() =>
        Path.Combine(PreferencesDataRoot(), ProductIdentity.DataFolderName, "updates");

    /// <summary>
    /// The installer puts the product under Programs\xlide; a library loaded from anywhere else
    /// was published by hand, and installing would register the installed copy over it. The
    /// About dialog applies the same rule to the same path.
    /// </summary>
    private static bool IsInstalledBuild(string? directory) =>
        directory is not null
        && directory.TrimEnd('\\').EndsWith(@"\Programs\xlide", StringComparison.OrdinalIgnoreCase);

    /// <summary>The Office application this library is loaded into, by its process name.</summary>
    private static string HostDisplayName() =>
        Path.GetFileNameWithoutExtension(Environment.ProcessPath)?.ToUpperInvariant() switch
        {
            "EXCEL" => "Excel",
            "WINWORD" => "Word",
            "POWERPNT" => "PowerPoint",
            "MSACCESS" => "Access",
            { Length: > 0 } other => other,
            _ => "the Office application",
        };

    /// <summary>
    /// Arms the once-per-session check, a few seconds after the first ready. A reload's ready
    /// does not arm a second one, and the setting can keep it from arming at all.
    /// </summary>
    private void ScheduleUpdateCheck()
    {
        if (_updateCheckScheduled)
        {
            return;
        }

        _updateCheckScheduled = true;

        if (!_settings.UpdateCheckOnStart)
        {
            Log.Info("update: the check on start is off; Check for updates on the wrench menu still asks");
            return;
        }

        var surface = _editorSurface;
        _ = Task.Run(async () =>
        {
            await Task.Delay(UpdateCheckDelay).ConfigureAwait(false);
            surface?.RunOnHostThread(() =>
            {
                if (!_stopped)
                {
                    BeginUpdateCheck(manual: false);
                }
            });
        });
    }

    /// <summary>
    /// A choice from the page's dialog or the api: check, install, skip, notes, cancel, or the
    /// harness's reset. Answers a refusal in words, or null when the action was taken up.
    /// </summary>
    private string? OnUpdateAction(string action)
    {
        switch (action)
        {
            case "check":
                BeginUpdateCheck(manual: true);
                return null;

            case "install":
                return BeginUpdateInstall();

            case "skip":
                return SkipOfferedUpdate();

            case "notes":
                return OpenReleasePage();

            case "cancel":
                CancelUpdateWork();
                return null;

            case "dismiss":
                // The dialog closing is the page's business; the state stands for the About
                // dialog and the api to read.
                return null;

#if DEBUG
            case "reset":
                CancelUpdateWork();
                _update = new UpdateProgress();
                ShowUpdate(open: false);
                return null;
#endif

            default:
                return $"'{action}' is not an update action; try check, install, skip, notes or cancel";
        }
    }

    private void BeginUpdateCheck(bool manual)
    {
        if (_update.Phase is "checking" or "downloading")
        {
            // One thing at a time. A check asked for by hand while one runs just shows it.
            if (manual)
            {
                ShowUpdate(open: true);
            }

            return;
        }

        var round = ++_updateRound;
        var work = ReplaceUpdateWork();
        _update = new UpdateProgress { Phase = "checking", Manual = manual };
        ShowUpdate(open: manual);
        Log.Info($"update: asking {_updateSource.FeedUrl} for the newest release ({(manual ? "by hand" : "on start")})");

        var source = _updateSource;
        var updater = Updater;
        var surface = _editorSurface;
        _ = Task.Run(async () =>
        {
            string? json = null;
            string? failure = null;
            try
            {
                json = await updater.ReadFeedAsync(source, work.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                return;
            }
            catch (Exception ex)
            {
                failure = ex.Message.Trim();
            }

            surface?.RunOnHostThread(() => CompleteUpdateCheck(round, manual, json, failure));
        });
    }

    private void CompleteUpdateCheck(int round, bool manual, string? json, string? failure)
    {
        if (round != _updateRound || _stopped)
        {
            return;
        }

        var now = DateTimeOffset.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);

        if (failure is null)
        {
            var release = ReleaseFeed.Parse(json ?? string.Empty, _updateSource, out var refusal);
            if (release is not null)
            {
                var verdict = ReleaseFeed.Judge(release, ProductVersion.Current, SkippedUpdateVersion());
                _update = new UpdateProgress
                {
                    Phase = verdict switch
                    {
                        UpdateVerdict.Available => "available",
                        UpdateVerdict.Skipped => "skipped",
                        _ => "upToDate",
                    },
                    Manual = manual,
                    Release = release,
                    CheckedAt = now,
                };

                Log.Info($"update: the newest release is {release.Version} ({release.PublishedAt}); "
                    + $"this build is {ProductVersion.Current}, so {_update.Phase}");

                // An automatic check speaks only when it has news; one asked for always answers.
                ShowUpdate(open: manual || verdict == UpdateVerdict.Available);
                return;
            }

            failure = refusal ?? "the release feed could not be read";
        }

        _update = new UpdateProgress { Phase = "failed", Manual = manual, CheckedAt = now, Error = failure };
        Log.Warn($"update: the check failed: {failure}");

        // A failed check on start is a log line, not a dialog: the developer did not ask, and a
        // machine with no route to GitHub would otherwise be told so every session.
        ShowUpdate(open: manual);
    }

    private string? BeginUpdateInstall()
    {
        if (_update.Release is not { } release)
        {
            return "no release has been offered; check for updates first";
        }

        if (_update.Phase is "checking" or "downloading" or "verified" or "launched")
        {
            return $"an update is already {_update.Phase}";
        }

        if (_update.Phase == "upToDate")
        {
            // The dialog offers no Install here, so neither does the route: downloading the
            // version already running and starting its installer serves nobody.
            return $"this build, {ProductVersion.Current}, is as new as the newest release; nothing to install";
        }

        var round = ++_updateRound;
        var work = ReplaceUpdateWork();
        _update = new UpdateProgress
        {
            Phase = "downloading",
            Manual = _update.Manual,
            Release = release,
            CheckedAt = _update.CheckedAt,
        };
        ShowUpdate(open: true);
        Log.Info($"update: downloading {release.Installer.DownloadUrl} ({release.Installer.Size:N0} bytes)");

        var folder = UpdatesFolder();
        var updater = Updater;
        var surface = _editorSurface;
        var lastShown = Stopwatch.StartNew();
        _ = Task.Run(async () =>
        {
            string path;
            try
            {
                path = await updater.DownloadAsync(release, folder, received =>
                {
                    // Progress at most five times a second: the host thread has an editor to run.
                    if (received < release.Installer.Size && lastShown.ElapsedMilliseconds < 200)
                    {
                        return;
                    }

                    lastShown.Restart();
                    surface?.RunOnHostThread(() =>
                    {
                        if (round == _updateRound && _update.Phase == "downloading")
                        {
                            _update.Received = received;
                            ShowUpdate(open: false);
                        }
                    });
                }, work.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                surface?.RunOnHostThread(() =>
                {
                    if (round == _updateRound)
                    {
                        Log.Info("update: the download was cancelled");
                        _update.Phase = "available";
                        _update.Received = 0;
                        ShowUpdate(open: false);
                    }
                });
                return;
            }
            catch (Exception ex)
            {
                var reason = ex.Message.Trim();
                surface?.RunOnHostThread(() => FailUpdate(round, reason));
                return;
            }

            surface?.RunOnHostThread(() => LaunchDownloadedInstaller(round, release, path));
        });

        return null;
    }

    private void LaunchDownloadedInstaller(int round, Release release, string path)
    {
        if (round != _updateRound || _stopped)
        {
            return;
        }

        _update.Phase = "verified";
        _update.Received = release.Installer.Size;
        _update.Path = path;
        ShowUpdate(open: false);
        Log.Info($"update: {path} verifies against the release's digest");

        Process installer;
        try
        {
            installer = UpdateService.Launch(path);
        }
        catch (Exception ex) when (ex is UpdateException or System.ComponentModel.Win32Exception or InvalidOperationException)
        {
            FailUpdate(round, $"the installer could not be started: {ex.Message.Trim()}");
            return;
        }

        _update.Phase = "launched";
        _update.Pid = installer.Id;
        Log.Info($"update: started the installer as process {installer.Id}; it installs when {HostDisplayName()} closes");
        ShowUpdate(open: true);
        _editorSurface?.Hold($"xlide {release.Version} is ready to install. Save your work and close "
            + $"{HostDisplayName()}; the installer finishes on its own.");

        var surface = _editorSurface;
        _ = Task.Run(async () =>
        {
            int code;
            try
            {
                await installer.WaitForExitAsync().ConfigureAwait(false);
                code = installer.ExitCode;
            }
            catch (Exception ex) when (ex is InvalidOperationException or System.ComponentModel.Win32Exception)
            {
                return;
            }
            finally
            {
                installer.Dispose();
            }

            // The installer ending while this session still runs means it did NOT install: an
            // install ends with the host closed and this code never reached. So say so.
            surface?.RunOnHostThread(() =>
            {
                if (round != _updateRound || _stopped)
                {
                    return;
                }

                _update.Phase = "exited";
                _update.ExitCode = code;
                Log.Warn($"update: the installer ended with code {code} before {HostDisplayName()} closed, so nothing was installed");
                _editorSurface?.Hold(string.Empty);
                ShowUpdate(open: true);
            });
        });
    }

    private void FailUpdate(int round, string reason)
    {
        if (round != _updateRound || _stopped)
        {
            return;
        }

        _update.Phase = "failed";
        _update.Error = reason;
        _update.Received = 0;
        Log.Warn($"update: {reason}");
        ShowUpdate(open: true);
    }

    private string? SkipOfferedUpdate()
    {
        if (_update.Release is not { } release || _update.Phase is not ("available" or "skipped"))
        {
            return "no update is being offered";
        }

        try
        {
            // The version is digits and dots, checked on the way in, so quoting it is the whole
            // of its JSON; the reflection serializer is not for a native build.
            UserPreferences.WriteValue(UserStateStore.UiFile, SkippedUpdateKey,
                string.Create(CultureInfo.InvariantCulture, $"\"{release.Version}\""));
        }
        catch (Exception ex)
        {
            Log.Error("update: the skipped version could not be saved", ex);
            return "the choice could not be saved; it holds for this session only";
        }
        finally
        {
            _update.Phase = "skipped";
            ShowUpdate(open: false);
        }

        Log.Info($"update: {release.Version} will not be offered again");
        _editorSurface?.Notify($"xlide {release.Version} will not be offered again. Check for updates is on the wrench menu.");
        return null;
    }

    /// <summary>The version the developer declined, read from ui-state.json, or null.</summary>
    private static string? SkippedUpdateVersion()
    {
        try
        {
            var raw = UserPreferences.ReadValue(UserStateStore.UiFile, SkippedUpdateKey);
            if (raw is null)
            {
                return null;
            }

            using var document = JsonDocument.Parse(raw);
            return document.RootElement.ValueKind == JsonValueKind.String ? document.RootElement.GetString() : null;
        }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
        {
            return null;
        }
    }

    private string? OpenReleasePage()
    {
        if (_update.Release is not { } release)
        {
            return "no release has been offered; check for updates first";
        }

        if (_updateSource != UpdateSource.GitHub)
        {
            // A harness feed names pages on its own loopback origin; a browser is not where a
            // suite's output belongs.
            return "release notes are opened only for the GitHub feed";
        }

        var address = release.PageUrl.ToString();
        if (!_updateSource.AllowsPage(release.PageUrl))
        {
            return "the release page is not on the product's repository";
        }

        var result = Win32.ShellExecute(0, "open", address, null, null, Win32.ShowNormal);
        if ((long)result <= Win32.ShellExecuteFailure)
        {
            Log.Error($"update: the shell would not open the release page (result {(long)result})");
            return "the release page could not be opened in a browser";
        }

        Log.Info($"update: opened {address}");
        return null;
    }

    private void CancelUpdateWork()
    {
        if (_updateWork is { } work)
        {
            work.Cancel();
        }
    }

    private CancellationTokenSource ReplaceUpdateWork()
    {
        var previous = _updateWork;
        _updateWork = new CancellationTokenSource();
        previous?.Cancel();
        previous?.Dispose();
        return _updateWork;
    }

    /// <summary>Stops any download and releases the client, as the session goes down.</summary>
    private void StopUpdates()
    {
        _updateRound++;
        _updateWork?.Cancel();
        _updateWork?.Dispose();
        _updateWork = null;
        _updateService?.Dispose();
        _updateService = null;
    }

    private UpdateStateReply UpdateState()
    {
        var release = _update.Release;
        var loadedFrom = ShimModule.Directory;
        return new UpdateStateReply(
            Phase: _update.Phase,
            Installed: ProductVersion.Current,
            DevelopmentBuild: !IsInstalledBuild(loadedFrom),
            LoadedFrom: loadedFrom,
            Host: HostDisplayName(),
            Feed: _updateSource.FeedUrl.ToString(),
            CheckedAt: _update.CheckedAt,
            Manual: _update.Manual,
            Version: release?.Version,
            Tag: release?.Tag,
            PublishedAt: release?.PublishedAt,
            Notes: release?.Notes,
            Page: release?.PageUrl.ToString(),
            Size: release?.Installer.Size ?? 0,
            Received: _update.Received,
            Path: _update.Path,
            Pid: _update.Pid,
            ExitCode: _update.ExitCode,
            Failure: _update.Error,
            SkippedVersion: SkippedUpdateVersion());
    }

    private void ShowUpdate(bool open) => _editorSurface?.ShowUpdate(UpdateState(), open);
}
