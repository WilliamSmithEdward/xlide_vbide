using System.Text;
using System.Text.Json;

namespace Xlide.Vbe.Core.Editor;

/// <summary>One per-user persistence procedure for all hosts. Existing settings and sync file
/// names stay compatible; UI state is independent of WebView2's process-specific profiles.</summary>
public sealed class UserStateStore
{
    public const string SettingsFile = "settings.json";
    public const string SyncFile = "sync.json";
    public const string UiFile = "ui-state.json";

    public UserStateStore(string localApplicationData)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(localApplicationData);
        Root = Path.Combine(localApplicationData, ProductIdentity.DataFolderName);
    }

    public string Root { get; }

    public string PathFor(string document) => document is SettingsFile or SyncFile or UiFile
        ? Path.Combine(Root, document)
        : throw new ArgumentException("Unknown preference document.", nameof(document));

    public string Read(string document)
    {
        var path = PathFor(document);
        if (!File.Exists(path)) return "{}";
        using var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var reader = new StreamReader(file);
        return reader.ReadToEnd();
    }

    public string? ReadValue(string document, string key)
    {
        using var parsed = JsonDocument.Parse(Read(document));
        return parsed.RootElement.TryGetProperty(key, out var value) ? value.GetRawText() : null;
    }

    public void WriteValue(string document, string key, string json)
    {
        using var value = JsonDocument.Parse(json);
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream))
        {
            writer.WriteStartObject();
            writer.WritePropertyName(key);
            value.RootElement.WriteTo(writer);
            writer.WriteEndObject();
        }

        ApplyChanges(document, "{}", Encoding.UTF8.GetString(stream.ToArray()));
    }

    public void RemoveValue(string document, string key)
    {
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream))
        {
            writer.WriteStartObject();
            writer.WriteNull(key);
            writer.WriteEndObject();
        }
        ApplyChanges(document, Encoding.UTF8.GetString(stream.ToArray()), "{}");
    }

    /// <summary>Merge only changed fields into the latest file under a cross-process lock.
    /// A stale Excel session cannot overwrite an unrelated choice just saved in Word.
    /// Readers see a complete old or new file, including during a crash.</summary>
    public string ApplyChanges(string document, string before, string after)
    {
        var path = PathFor(document);
        Directory.CreateDirectory(Root);
        using var gate = Acquire(path + ".lock");
        using var current = JsonDocument.Parse(Read(document));
        using var previous = JsonDocument.Parse(before);
        using var updated = JsonDocument.Parse(after);
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = true }))
            Merge(writer, current.RootElement, previous.RootElement, updated.RootElement);

        var json = Encoding.UTF8.GetString(stream.ToArray());
        using var merged = JsonDocument.Parse(json);
        if (JsonElement.DeepEquals(current.RootElement, merged.RootElement)) return json;
        var temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            File.WriteAllText(temporary, json, new UTF8Encoding(false));
            File.Move(temporary, path, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
        }

        return json;
    }

    private static FileStream Acquire(string path)
    {
        var until = Environment.TickCount64 + 1000;
        while (true)
        {
            try { return new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
            catch (IOException) when (Environment.TickCount64 < until) { Thread.Sleep(10); }
        }
    }

    private static void Merge(Utf8JsonWriter writer, JsonElement current, JsonElement before, JsonElement after)
    {
        if (current.ValueKind != JsonValueKind.Object || before.ValueKind != JsonValueKind.Object
            || after.ValueKind != JsonValueKind.Object)
            throw new JsonException("A preference document must be a JSON object.");

        var values = current.EnumerateObject().ToDictionary(property => property.Name, property => property.Value);
        foreach (var property in before.EnumerateObject())
            if (!after.TryGetProperty(property.Name, out _)) values.Remove(property.Name);
        foreach (var property in after.EnumerateObject())
        {
            if (!before.TryGetProperty(property.Name, out var old) || !JsonElement.DeepEquals(old, property.Value))
                values[property.Name] = property.Value;
        }

        writer.WriteStartObject();
        foreach (var (name, value) in values)
        {
            writer.WritePropertyName(name);
            // Merge nested maps too: sync settings have one row per project.
            if (current.TryGetProperty(name, out var held) && before.TryGetProperty(name, out var old)
                && after.TryGetProperty(name, out var next) && held.ValueKind == JsonValueKind.Object
                && old.ValueKind == JsonValueKind.Object && next.ValueKind == JsonValueKind.Object)
                Merge(writer, held, old, next);
            else
                value.WriteTo(writer);
        }
        writer.WriteEndObject();
    }
}
