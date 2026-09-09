using Xlide.Vbe.Core.Scm;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The commit message the pane pre-fills from the change log: an author's own labels when there
/// are any, otherwise what each author did, one line, short enough to be a subject.
/// </summary>
public sealed class CommitMessageTests
{
    private static RoundSummary Round(
        string author, string? label, params (string Kind, string Module, string? From)[] entries) =>
        new(author, label, entries);

    [Fact]
    public void NoRoundsIsNoSuggestion()
    {
        Assert.Equal(string.Empty, CommitMessage.Suggest([]));
    }

    [Fact]
    public void LabelsWinAndRepeatOnce()
    {
        var rounds = new[]
        {
            Round("claude", "Fix posting", ("written", "Ledger", null)),
            Round("developer", null, ("written", "Reports", null)),
            Round("claude", "fix posting", ("written", "Ledger", null)),
            Round("claude", "Add reports", ("added", "Reports", null)),
        };

        // Newest first as the rounds arrive; a label repeated in another case is the same label.
        Assert.Equal("Fix posting; Add reports", CommitMessage.Suggest(rounds));
    }

    [Fact]
    public void WithoutLabelsEachAuthorGetsWhatTheyDid()
    {
        var rounds = new[]
        {
            Round("claude", null, ("written", "Ledger", null), ("added", "Foo", "Bar")),
            Round("developer", null, ("written", "Reports", null)),
        };

        Assert.Equal(
            "claude: Written Ledger, Added Foo (was Bar); developer: Written Reports",
            CommitMessage.Suggest(rounds));
    }

    [Fact]
    public void AModuleTouchedInTwoRoundsIsOneItemWithTheStrongerWord()
    {
        // Newest first: written today, added yesterday. To the repository it is an addition.
        var rounds = new[]
        {
            Round("claude", null, ("written", "Foo", null)),
            Round("claude", null, ("added", "Foo", null)),
            Round("claude", null, ("written", "Foo", null)),
        };

        Assert.Equal("claude: Added Foo", CommitMessage.Suggest(rounds));
    }

    [Fact]
    public void TheKindIsCapitalisedWhateverSpellingArrives()
    {
        Assert.Equal(
            "developer: Written A, Removed B",
            CommitMessage.Suggest([Round("developer", null, ("Written", "A", null), ("REMOVED", "B", null))]));
    }

    [Fact]
    public void TheSuggestionIsOneLineOfAtMostTwoHundredCharacters()
    {
        var entries = Enumerable.Range(1, 40)
            .Select(number => ("written", $"Module{number:00}", (string?)null))
            .ToArray();

        var suggestion = CommitMessage.Suggest([Round("claude", null, entries)]);

        Assert.Equal(200, suggestion.Length);
        Assert.EndsWith("...", suggestion, StringComparison.Ordinal);
        Assert.DoesNotContain('\n', suggestion);
    }

    [Fact]
    public void ALabelWithALineBreakIsFlattened()
    {
        Assert.Equal(
            "Posting fix, second try",
            CommitMessage.Suggest([Round("claude", "Posting fix,\r\nsecond try", ("written", "Ledger", null))]));
    }

    [Fact]
    public void ARoundWithNeitherLabelNorEntriesSuggestsNothing()
    {
        Assert.Equal(string.Empty, CommitMessage.Suggest([Round("claude", "  ")]));
    }
}
