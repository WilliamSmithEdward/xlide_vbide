using Xlide.Vbe.Core.Sync;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The planner over text that is not English, which is the half of import and export that can be
/// checked without a host and therefore the half that can run in CI.
///
/// THE RISK HERE IS ARITHMETIC, not code pages. This product never decodes a workbook's bytes; the
/// companion does, and its language tests are about conversion tables for exactly that reason. What
/// runs here compares texts, counts lines and cuts headers off them, and every one of those is a
/// count of something. A count taken in bytes or in code points rather than in UTF-16 code units
/// drifts by the width of the non-ASCII text to its left, which no English fixture can show: the
/// plan still looks right, and the row it draws is about the wrong lines.
///
/// The scripts are the companion product's own language matrix (tests/vbaLanguageMatrix.test.ts),
/// so the two products are exercised over the same text.
/// </summary>
public class ModuleSyncLanguageTests
{
    public static TheoryData<string, string> Scripts() => new()
    {
        { "Western European", "déjà vu € œuvre Straße" },
        { "Central European", "Příliš žluťoučký kůň Zażółć gęślą" },
        { "Baltic", "Lietuviškas tekstas ąčęėįšųū" },
        { "Turkish", "Türkçe deneme ğüşiöç İı" },
        { "Vietnamese", "Tiếng Việt thử nghiệm" },
        { "Greek", "Δοκιμή ελληνικού κειμένου" },
        { "Cyrillic", "Проверка русского текста" },
        { "Hebrew", "בדיקת עברית" },
        { "Arabic", "اختبار العربية" },
        { "Devanagari", "नामजांच" },
        { "Thai", "ทดสอบภาษาไทย" },
        { "Japanese", "テスト用モジュール" },
        { "Chinese", "中文测试模块" },
        // Astral: two UTF-16 code units per character, which is where a code-point count and a
        // code-unit count part company for good.
        { "Emoji", "\U0001F600\U0001F4C8" },
        // The same word twice, spelled two ways. A repository authored on macOS carries the second.
        { "Latin, precomposed", "caf\u00E9 na\u00EFve" },
        { "Latin, decomposed", "cafe\u0301 nai\u0308ve" },
    };

    private static string ModuleOf(string text) => string.Join("\r\n",
        "Attribute VB_Name = \"Probe\"",
        "Option Explicit",
        string.Empty,
        $"' {text}",
        $"Public Const Note As String = \"{text}\"",
        string.Empty,
        "Public Sub Run()",
        "    Debug.Print Note",
        "End Sub",
        string.Empty);

    [Theory]
    [MemberData(nameof(Scripts))]
    public void AProjectAlreadyMatchingItsFolderIsUnchangedInEveryScript(string _, string text)
    {
        var source = ModuleOf(text);

        var plan = ModuleSync.PlanExport(
            "id", "Book.xlsm", @"C:\repo",
            [new LiveModule("Probe", "standard", source)],
            [new RepoFile("Probe.bas", source)],
            ExportMode.ExportAll);

        Assert.Equal(SyncStatus.Unchanged, plan.Items[0].Status);
        Assert.False(plan.Items[0].Checked);
    }

    [Theory]
    [MemberData(nameof(Scripts))]
    public void OneChangedCharacterIsSeenInEveryScript(string _, string text)
    {
        // The file differs from the module by ONE character, and it is at the end, so a comparison
        // that stopped short or counted the wrong units would call these the same.
        var plan = ModuleSync.PlanExport(
            "id", "Book.xlsm", @"C:\repo",
            [new LiveModule("Probe", "standard", ModuleOf(text))],
            [new RepoFile("Probe.bas", ModuleOf($"{text}."))],
            ExportMode.ExportAll);

        Assert.Equal(SyncStatus.WillWrite, plan.Items[0].Status);
        Assert.True(plan.Items[0].Checked);
    }

    [Fact]
    public void PrecomposedAndDecomposedAreDifferentText()
    {
        // They render alike and they are not the same bytes, so a plan must call them different.
        // Anything that normalised on the way past would report a file in sync that is not, and
        // silently keep two spellings of one word in two places.
        var plan = ModuleSync.PlanExport(
            "id", "Book.xlsm", @"C:\repo",
            [new LiveModule("Probe", "standard", ModuleOf("caf\u00E9"))],
            [new RepoFile("Probe.bas", ModuleOf("cafe\u0301"))],
            ExportMode.ExportAll);

        Assert.Equal(SyncStatus.WillWrite, plan.Items[0].Status);
    }

    [Theory]
    [MemberData(nameof(Scripts))]
    public void TheComparisonCountsLinesRatherThanCharacters(string _, string text)
    {
        // An unchanged row draws one line saying how many agreed. That count is where a width
        // mistake surfaces first: the module below is nine lines whatever alphabet it is in.
        var source = ModuleOf(text);

        var plan = ModuleSync.PlanExport(
            "id", "Book.xlsm", @"C:\repo",
            [new LiveModule("Probe", "standard", source)],
            [new RepoFile("Probe.bas", source)],
            ExportMode.ExportAll);

        var gap = Assert.Single(plan.Items[0].DiffWithHeaders);
        Assert.Equal(DiffKind.Gap, gap.Kind);
        Assert.Equal("10 identical lines", gap.Left);
    }

    [Theory]
    [MemberData(nameof(Scripts))]
    public void TheHeaderIsCutAtTheRightPlaceWhateverFollowsIt(string _, string text)
    {
        var body = ModuleSync.CodeWithoutHeader(ModuleOf(text));

        Assert.StartsWith("Option Explicit", body, System.StringComparison.Ordinal);
        Assert.DoesNotContain("Attribute VB_Name", body, System.StringComparison.Ordinal);
        Assert.Contains(text, body, System.StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(Scripts))]
    public void AnImportSeesTheSameChangeFromTheOtherSide(string _, string text)
    {
        var plan = ModuleSync.PlanImport(
            "id", "Book.xlsm", @"C:\repo",
            [new LiveModule("Probe", "standard", ModuleOf(text))],
            [new RepoFile("Probe.bas", ModuleOf($"{text} changed"))],
            ImportMode.UpdateOnly);

        Assert.Equal(SyncStatus.WillUpdate, plan.Items[0].Status);
    }

    [Theory]
    [MemberData(nameof(Scripts))]
    public void ADifferenceIsDrawnAtTheLineItIsOn(string _, string text)
    {
        // The changed line is the fourth. A width mistake anywhere in the comparison moves it.
        var plan = ModuleSync.PlanExport(
            "id", "Book.xlsm", @"C:\repo",
            [new LiveModule("Probe", "standard", ModuleOf(text))],
            [new RepoFile("Probe.bas", ModuleOf($"{text}!"))],
            ExportMode.ExportAll);

        var changed = plan.Items[0].DiffWithHeaders
            .Where(line => line.Kind == DiffKind.Changed)
            .ToList();

        Assert.NotEmpty(changed);
        Assert.All(changed, line => Assert.Contains(text, line.Left, System.StringComparison.Ordinal));
    }
}
