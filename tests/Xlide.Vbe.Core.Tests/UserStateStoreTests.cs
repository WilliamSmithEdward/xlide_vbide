using System.Text.Json;
using Xlide.Vbe.Core.Editor;
using Xlide.Vbe.Core.Hosting;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

public sealed class UserStateStoreTests : IDisposable
{
    private readonly string _temporary = Path.Combine(Path.GetTempPath(), "xlide-preferences-" + Guid.NewGuid().ToString("N"));

    [Fact]
    public void ExistingSettingsKeepTheirPathAndUnknownFields()
    {
        var store = new UserStateStore(_temporary);
        Directory.CreateDirectory(store.Root);
        File.WriteAllText(store.PathFor(UserStateStore.SettingsFile), "{\"format.indentSize\":2,\"future.choice\":true}");
        var before = ProductSettings.Parse(store.Read(UserStateStore.SettingsFile));
        var after = before with { ContinueCommentOnNewline = false };
        var merged = store.ApplyChanges(UserStateStore.SettingsFile, before.ToJson(), after.ToJson());
        using var document = JsonDocument.Parse(merged);
        Assert.True(document.RootElement.GetProperty("future.choice").GetBoolean());
        Assert.Equal(2, ProductSettings.Parse(merged).FormatIndentSize);
        Assert.EndsWith(Path.Combine("xlide_vbide", "settings.json"), store.PathFor(UserStateStore.SettingsFile));
    }

    [Fact]
    public void AStaleHostChangesOnlyItsOwnSettings()
    {
        var excel = new UserStateStore(_temporary);
        var word = new UserStateStore(_temporary);
        var original = ProductSettings.Default;
        excel.ApplyChanges(UserStateStore.SettingsFile, original.ToJson(), (original with { FormatIndentSize = 2 }).ToJson());
        word.ApplyChanges(UserStateStore.SettingsFile, original.ToJson(), (original with { DesignerGridSize = 12 }).ToJson());
        var settings = ProductSettings.Parse(excel.Read(UserStateStore.SettingsFile));
        Assert.Equal(2, settings.FormatIndentSize);
        Assert.Equal(12, settings.DesignerGridSize);
    }

    [Fact]
    public void RemovingARuleOverrideDoesNotKeepTheOldOverride()
    {
        var store = new UserStateStore(_temporary);
        const string original = "{\"overrides\":{\"one\":\"off\",\"two\":\"warning\"}}";
        store.ApplyChanges(UserStateStore.SettingsFile, "{}", original);
        store.ApplyChanges(UserStateStore.SettingsFile, original, "{\"overrides\":{\"two\":\"warning\"}}");
        using var document = JsonDocument.Parse(store.Read(UserStateStore.SettingsFile));
        Assert.False(document.RootElement.GetProperty("overrides").TryGetProperty("one", out _));
    }

    [Fact]
    public async Task ConcurrentHostsKeepEachOthersUiValues()
    {
        await Task.WhenAll(Enumerable.Range(0, 20).Select(index => Task.Run(() =>
            new UserStateStore(_temporary).WriteValue(UserStateStore.UiFile, "pane" + index, "{\"width\":200}"))));
        using var document = JsonDocument.Parse(new UserStateStore(_temporary).Read(UserStateStore.UiFile));
        Assert.Equal(20, document.RootElement.EnumerateObject().Count());
        Assert.Empty(Directory.GetFiles(new UserStateStore(_temporary).Root, "*.tmp"));
    }

    [Fact]
    public void MalformedExistingDataIsReportedAndPreserved()
    {
        var store = new UserStateStore(_temporary);
        Directory.CreateDirectory(store.Root);
        File.WriteAllText(store.PathFor(UserStateStore.UiFile), "broken");
        Assert.ThrowsAny<JsonException>(() => store.WriteValue(UserStateStore.UiFile, "paneLayout", "{}"));
        Assert.Equal("broken", store.Read(UserStateStore.UiFile));
    }

    [Fact]
    public void WindowBoundsFitTheAvailableMonitorWithoutLosingMaximizedState()
    {
        var geometry = new WindowGeometry(-2000, -1000, 3000, 2000, true);
        Assert.Equal(new WindowGeometry(0, 0, 1280, 720, true), geometry.Fit(new PixelRect(0, 0, 1280, 720)));
        Assert.Equal(new WindowGeometry(-1280, 0, 1280, 720, true), geometry.Fit(new PixelRect(-1280, 0, 0, 720)));
    }

    public void Dispose()
    {
        if (Directory.Exists(_temporary)) Directory.Delete(_temporary, recursive: true);
    }
}
