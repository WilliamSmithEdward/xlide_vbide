using System.Text.Json;
using System.Text.Json.Serialization;
using Xlide.Vbe.Core;
using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.AddIn;

/// <summary>
/// The api door's switch, and the address it hands out.
///
/// THE SWITCH IS NOT BEHIND THE DOOR. Everything here answers on the page's own channel to the
/// host - the one settings and the change log already use - because a switch reachable only
/// through the api it controls could never turn the api on.
///
/// WHAT TURNING IT ON MEANS, said plainly because the card says it to the developer too: a
/// loopback listener on a random port, and a token in a file under this user's profile. Any
/// program running as this user can read that file, and a program that has read it can do
/// everything this editor can do to the open projects - read code, write code, run code. That is
/// the capability, and it is why the door ships shut.
/// </summary>
internal sealed partial class AddInSession
{
    /// <summary>What the agent card reads, and what it gets back after moving the switch.</summary>
    internal sealed record ApiStateReply(
        [property: JsonPropertyName("on")] bool On,
        [property: JsonPropertyName("leansOpen")] bool LeansOpen,
        [property: JsonPropertyName("remembered")] bool Remembered,
        [property: JsonPropertyName("host")] string Host,
        [property: JsonPropertyName("pid")] int Pid,
        [property: JsonPropertyName("port")] int Port,
        [property: JsonPropertyName("token")] string Token,
        [property: JsonPropertyName("baseUrl")] string BaseUrl,
        [property: JsonPropertyName("agentUrl")] string AgentUrl,
        [property: JsonPropertyName("progId")] string ProgId,
        [property: JsonPropertyName("discovery")] string Discovery,
        [property: JsonPropertyName("project")] string Project,
        [property: JsonPropertyName("error")] string? Error = null);

    /// <summary>
    /// Reads the door's state, or moves it.
    ///
    /// `on` and `off` PERSIST. A developer who opened the door meant to open it, and a switch
    /// that forgot at the next launch would be one they had to find again every morning - so the
    /// answer is written to the settings file, and `remembered` tells the card that it was. That
    /// is also why the card says the door stays open until it is shut: a capability that outlives
    /// the session has to say so where it is turned on.
    /// </summary>
    private void OnApiRequested(int requestId, IReadOnlyDictionary<string, string> arguments)
    {
        string answer;
        try
        {
            arguments.TryGetValue("action", out var action);

            if (action is "on" or "off")
            {
                var wanted = action == "on";
                _settings = _settings with { ApiEnabled = wanted };
                SaveSettings();
                OpenOrCloseApi(wanted);
            }

            answer = JsonSerializer.Serialize(ApiState(), ApiSwitchJsonContext.Default.ApiStateReply);
        }
        catch (Exception ex)
        {
            Log.Error("xlide api: the agent card's question could not be answered", ex);
            answer = JsonSerializer.Serialize(
                ApiState() with { Error = ex.Message.Trim() },
                ApiSwitchJsonContext.Default.ApiStateReply);
        }

        _editorSurface?.ShowApiResult(requestId, answer);
    }

    /// <summary>
    /// The door as it stands. The token and the port come from the LIVE server rather than from
    /// anything remembered, so a card showing an address is showing one that answers.
    /// </summary>
    private ApiStateReply ApiState()
    {
        var server = _apiServer;
        var open = server is not null;

        return new ApiStateReply(
            On: open,
            LeansOpen: ApiOpenUnlessTold,
            Remembered: _settings.ApiEnabled is not null,
            Host: Engine.HostApp.Name,
            Pid: Environment.ProcessId,
            Port: server?.Port ?? 0,
            Token: server?.Token ?? string.Empty,
            BaseUrl: server?.BaseUrl ?? string.Empty,
            AgentUrl: open ? $"{server!.BaseUrl}/agent" : string.Empty,
            ProgId: ProductIdentity.ApiProgId,
            Discovery: server?.DiscoveryPath ?? string.Empty,
            Project: _shownProject ?? string.Empty);
    }

    /// <summary>
    /// Writes the settings file. Split out of OnSettingsChanged so the api switch persists
    /// through the same path rather than growing a second one that could disagree with it.
    /// </summary>
    private void SaveSettings()
    {
        try
        {
            var path = SettingsPath;
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, _settings.ToJson());
        }
        catch (Exception ex)
        {
            Log.Error("settings: could not be written; the choice holds for this session only", ex);
        }
    }
}

[JsonSourceGenerationOptions(WriteIndented = false)]
[JsonSerializable(typeof(AddInSession.ApiStateReply))]
internal sealed partial class ApiSwitchJsonContext : JsonSerializerContext;
