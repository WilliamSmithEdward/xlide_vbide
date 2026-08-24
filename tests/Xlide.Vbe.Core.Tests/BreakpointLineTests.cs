using Xlide.Vbe.Core.Editor;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// Which lines can carry a breakpoint.
///
/// The reason this is worth a test rather than a glance: the rule it replaces was line-local, and
/// a line cannot tell you it is inside a `Type` or an `Enum`. So every member of every Enum and
/// every field of every Type was offered as breakable, the editor refused each with a modal, and
/// this product recorded them as SET - because the toggle records what it asked for and VBIDE
/// exposes no way to ask what happened. Measured on LanguageFixture 2026-08-24.
///
/// Every case below is a real line from a real module, and the block cases are the ones that
/// were wrong.
/// </summary>
public class BreakpointLineTests
{
    private static readonly string[] Module =
    [
        /*  1 */ "Option Explicit",
        /*  2 */ "",
        /*  3 */ "' AN ENUM, whose members are not statements.",
        /*  4 */ "Public Enum Corner",
        /*  5 */ "    TopLeft",
        /*  6 */ "    TopRight",
        /*  7 */ "End Enum",
        /*  8 */ "",
        /*  9 */ "Public Type Point",
        /* 10 */ "    X As Double",
        /* 11 */ "    Y As Double",
        /* 12 */ "End Type",
        /* 13 */ "",
        /* 14 */ "Private mHeld As Long",
        /* 15 */ "",
        /* 16 */ "Public Function Corners() As Long",
        /* 17 */ "    Corners = 4",
        /* 18 */ "End Function",
    ];

    [Theory]
    [InlineData(4, "the line that opens an Enum")]
    [InlineData(5, "a member INSIDE an Enum")]
    [InlineData(6, "another member inside it")]
    [InlineData(7, "the End Enum")]
    [InlineData(9, "the line that opens a Type")]
    [InlineData(10, "a field INSIDE a Type")]
    [InlineData(11, "another field inside it")]
    [InlineData(12, "the End Type")]
    public void ADeclarationBlockCarriesNothing(int line, string what)
    {
        Assert.False(Breakpoints.CanCarry(Module, line),
            $"{what} (line {line}: \"{Module[line - 1].Trim()}\") must not be offered - "
            + "the editor answers a breakpoint there with a modal");
    }

    [Theory]
    [InlineData(1, "Option Explicit")]
    [InlineData(2, "a blank line")]
    [InlineData(3, "a comment")]
    [InlineData(14, "a module-level Private declaration")]
    public void NorDoesAnythingElseThatIsNotAStatement(int line, string what)
    {
        Assert.False(Breakpoints.CanCarry(Module, line), $"{what} (line {line})");
    }

    [Theory]
    [InlineData(16, "a procedure's opening line, modifiers and all")]
    [InlineData(17, "a statement in its body")]
    [InlineData(18, "its End Function - which IS executable, unlike End Type")]
    public void AndExecutableStatementsDo(int line, string what)
    {
        Assert.True(Breakpoints.CanCarry(Module, line), $"{what} (line {line})");
    }

    [Fact]
    public void ALineThatIsNotThereCarriesNothing()
    {
        Assert.False(Breakpoints.CanCarry(Module, 0));
        Assert.False(Breakpoints.CanCarry(Module, -1));
        Assert.False(Breakpoints.CanCarry(Module, Module.Length + 1));
        Assert.False(Breakpoints.CanCarry([], 1));
    }

    [Fact]
    public void AnEnumWithNoModifierIsStillABlock()
    {
        // `Enum` and `Type` are legal bare at module level, and the walk has to catch those too -
        // the modifier is optional, so keying on "Public Enum" alone would let a bare one through
        // and take its members with it.
        string[] bare = ["Option Explicit", "", "Enum Bare", "    Only", "End Enum",
            "", "Sub After()", "End Sub"];

        Assert.False(Breakpoints.CanCarry(bare, 3), "the opener");
        Assert.False(Breakpoints.CanCarry(bare, 4), "the member inside it");
        Assert.False(Breakpoints.CanCarry(bare, 5), "the End Enum");
        Assert.True(Breakpoints.CanCarry(bare, 7), "and the procedure after it is unaffected");
    }

    [Fact]
    public void AWordIsAWordAndNotAPrefix()
    {
        // `Typed` is not `Type`, and a variable called `Options` is not `Option`. A prefix match
        // would refuse a perfectly ordinary assignment.
        string[] lines = ["Sub S()", "    Typed = 1", "    Options = 2", "    Enumerate", "End Sub"];

        Assert.True(Breakpoints.CanCarry(lines, 2));
        Assert.True(Breakpoints.CanCarry(lines, 3));
        Assert.True(Breakpoints.CanCarry(lines, 4));
    }
}
