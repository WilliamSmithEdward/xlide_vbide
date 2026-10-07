using Xlide.Vbe.Core.Editor;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Editor;

namespace Xlide.Vbe.Shim.AddIn;

internal sealed partial class AddInSession
{
    private WindowPreferences? _framePreferences;
    private bool _uiStatePresented;
    private static string PreferencesDataRoot()
    {
#if DEBUG
        // Live suites share one isolated preference root across their Office processes.
        // Release builds always use the current user's LocalAppData convention.
        if (Environment.GetEnvironmentVariable("XLIDE_TEST_PREFERENCES_ROOT") is { Length: > 0 } root)
            return root;
#endif
        return Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
    }

    private static void ResetUiLayout()
    {
        foreach (var key in new[] { "paneLayout", "sourceControl", "editorGroups" })
            UserPreferences.RemoveValue(UserStateStore.UiFile, key);
    }
    private static Dictionary<string, string> LoadUiState()
    {
        var values = new Dictionary<string, string>();
        foreach (var key in new[] { "paneLayout", "sourceControl", "editorGroups" })
        {
            try
            {
                if (UserPreferences.ReadValue(UserStateStore.UiFile, key) is { } value)
                    values[key] = value;
            }
            catch (Exception ex) { Log.Error($"preferences: {key} could not be read", ex); }
        }
        return values;
    }

    private void SaveUiState(string key, string json)
    {
        try { UserPreferences.WriteValue(UserStateStore.UiFile, key, json); }
        catch (Exception ex)
        {
            Log.Error($"preferences: {key} could not be saved", ex);
            _editorSurface?.Notify("The layout could not be saved. Your arrangement holds for this session only.");
        }
    }
}
