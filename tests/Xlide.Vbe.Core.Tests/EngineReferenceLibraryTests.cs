using System.Text.Json;
using Xlide.Vbe.Core.Engine;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

public class EngineReferenceLibraryTests
{
    [Fact]
    public void BoxedReferenceLibrariesSerializeInProjectOpenRequest()
    {
        // EngineClient sends a Dictionary<string, object>. Source-generated JSON must know the
        // boxed array's concrete type, or a NativeAOT build fails when a project is opened.
        var request = new Dictionary<string, object>
        {
            ["referenceLibraries"] = new[]
            {
                new EngineReferenceLibrary("Scripting", "{420B2830-E718-11CF-893D-00A0C9054228}",
                [new EngineReferenceType("IOMode", "enum",
                    [new EngineReferenceMember("ForAppending", "Const", "Const ForAppending = 8")])]),
            },
        };

        var json = JsonSerializer.Serialize(request, EngineJsonContext.Default.DictionaryStringObject);
        using var document = JsonDocument.Parse(json);
        var library = document.RootElement.GetProperty("referenceLibraries")[0];
        Assert.Equal("Scripting", library.GetProperty("name").GetString());
        Assert.Equal("{420B2830-E718-11CF-893D-00A0C9054228}", library.GetProperty("guid").GetString());
        Assert.Equal("ForAppending", library.GetProperty("types")[0].GetProperty("members")[0].GetProperty("name").GetString());
    }
}
