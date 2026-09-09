using System.Text.Json.Serialization;
using Xlide.Vbe.Shim.Sync;

namespace Xlide.Vbe.Shim.Scm;

/*
 * The Source Control pane's answers, in the shapes the pane and the xlide api both read.
 *
 * One set of records for both doors, the way Changes and Sync answer theirs: the pane receives
 * the route's own JSON verbatim inside an scmResult message, and the route returns the very same
 * string, so the two surfaces cannot drift. Every reply carries `detail` first; a refusal is
 * ScmErrorReply, which the harness client throws on. Commit and import EMBED the status under
 * `status` rather than flattening it, because the pane redraws from `reply.status ?? reply` after
 * every action and a flattened copy would be a second shape to keep in step.
 */

/// <summary>One row of the pane: a module against the branch head, or against the folder.</summary>
public sealed record ScmRowReply(
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("file")] string File,
    /// <summary>modified | added | deleted | renamed, or folderNewer | missingInFolder | missingInProject.</summary>
    [property: JsonPropertyName("status")] string Status,
    [property: JsonPropertyName("from")] string? From);

public sealed record ScmBranchRow(
    /// <summary>The branch's own name; for a remote's branch, without the remote in front.</summary>
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("current")] bool Current,
    [property: JsonPropertyName("upstream")] string? Upstream,
    /// <summary>Empty for a local branch; the remote's name for a branch only that remote has.</summary>
    [property: JsonPropertyName("remote")] string Remote);

public sealed record ScmIdentityReply(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("email")] string Email);

public sealed record ScmLastCommitReply(
    [property: JsonPropertyName("hash")] string Hash,
    [property: JsonPropertyName("short")] string ShortHash,
    [property: JsonPropertyName("author")] string Author,
    [property: JsonPropertyName("when")] string When,
    [property: JsonPropertyName("subject")] string Subject);

/// <summary>The pane's whole picture, and the answer to every action that changes it.</summary>
public sealed record ScmStatusReply(
    [property: JsonPropertyName("detail")] string Detail,
    [property: JsonPropertyName("project")] string Project,
    [property: JsonPropertyName("projectId")] string ProjectId,
    /// <summary>noGit | noProject | unsaved | noFolder | noRepository | noIdentity | conflicted | ready.</summary>
    [property: JsonPropertyName("state")] string State,
    [property: JsonPropertyName("folder")] string Folder,
    /// <summary>The repository root, or empty when the folder is under none.</summary>
    [property: JsonPropertyName("repository")] string Repository,
    [property: JsonPropertyName("gitVersion")] string GitVersion,
    [property: JsonPropertyName("branch")] string Branch,
    /// <summary>The commit the branch is at, in full; empty on an unborn branch.</summary>
    [property: JsonPropertyName("head")] string Head,
    [property: JsonPropertyName("upstream")] string Upstream,
    /// <summary>The remote pushes go to - origin, else the only one - and its URL; empty with none.</summary>
    [property: JsonPropertyName("remote")] string Remote,
    [property: JsonPropertyName("remoteUrl")] string RemoteUrl,
    [property: JsonPropertyName("ahead")] int Ahead,
    [property: JsonPropertyName("behind")] int Behind,
    /// <summary>Whether the WORKBOOK has unsaved changes; the folder is one save behind it then.</summary>
    [property: JsonPropertyName("dirty")] bool Dirty,
    [property: JsonPropertyName("identity")] ScmIdentityReply? Identity,
    [property: JsonPropertyName("rows")] ScmRowReply[] Rows,
    [property: JsonPropertyName("outside")] ScmRowReply[] Outside,
    [property: JsonPropertyName("branches")] ScmBranchRow[] Branches,
    [property: JsonPropertyName("conflicts")] string[] Conflicts,
    [property: JsonPropertyName("lastCommit")] ScmLastCommitReply? LastCommit,
    /// <summary>Empty when the branch head can be undone from here; else why it cannot.</summary>
    [property: JsonPropertyName("undoBlocked")] string UndoBlocked,
    [property: JsonPropertyName("suggestedMessage")] string SuggestedMessage,
    /// <summary>For noFolder: the sync folder the project remembers, else a folder beside the workbook.</summary>
    [property: JsonPropertyName("suggestedFolder")] string SuggestedFolder,
    [property: JsonPropertyName("covers")] string Covers);

public sealed record ScmSkippedReply(
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("why")] string Why);

public sealed record ScmCommitReply(
    [property: JsonPropertyName("detail")] string Detail,
    [property: JsonPropertyName("hash")] string Hash,
    [property: JsonPropertyName("short")] string ShortHash,
    [property: JsonPropertyName("committed")] string[] Committed,
    [property: JsonPropertyName("skipped")] ScmSkippedReply[] Skipped,
    [property: JsonPropertyName("status")] ScmStatusReply Status);

public sealed record ScmImportReply(
    [property: JsonPropertyName("detail")] string Detail,
    [property: JsonPropertyName("imported")] string[] Imported,
    [property: JsonPropertyName("skipped")] ScmSkippedReply[] Skipped,
    [property: JsonPropertyName("status")] ScmStatusReply Status);

/// <summary>The commit an undo took back, its whole message for the box, and the status after.</summary>
public sealed record ScmUndoReply(
    [property: JsonPropertyName("detail")] string Detail,
    [property: JsonPropertyName("hash")] string Hash,
    [property: JsonPropertyName("short")] string ShortHash,
    [property: JsonPropertyName("subject")] string Subject,
    /// <summary>Subject and body as they were written, so the next attempt starts from them.</summary>
    [property: JsonPropertyName("message")] string Message,
    [property: JsonPropertyName("status")] ScmStatusReply Status);

/// <summary>One file a commit touched, named by the module it holds.</summary>
public sealed record ScmFileRow(
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("file")] string File,
    /// <summary>added | modified | deleted | renamed, else git's own letter.</summary>
    [property: JsonPropertyName("status")] string Status);

public sealed record ScmCommitRow(
    [property: JsonPropertyName("hash")] string Hash,
    [property: JsonPropertyName("short")] string ShortHash,
    [property: JsonPropertyName("author")] string Author,
    [property: JsonPropertyName("email")] string Email,
    [property: JsonPropertyName("when")] string When,
    [property: JsonPropertyName("subject")] string Subject,
    [property: JsonPropertyName("body")] string Body,
    [property: JsonPropertyName("files")] ScmFileRow[] Files);

public sealed record ScmLogReply(
    [property: JsonPropertyName("detail")] string Detail,
    [property: JsonPropertyName("commits")] ScmCommitRow[] Commits);

/// <summary>A module's live text lined up against its text at a ref: left the ref, right live.</summary>
public sealed record ScmDiffReply(
    [property: JsonPropertyName("detail")] string Detail,
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("ref")] string Ref,
    [property: JsonPropertyName("rows")] SyncDiffRow[] Rows);

public sealed record ScmTextReply(
    [property: JsonPropertyName("detail")] string Detail,
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("ref")] string Ref,
    /// <summary>The module's code at the ref, attribute header off, as the code pane would show it.</summary>
    [property: JsonPropertyName("text")] string Text,
    [property: JsonPropertyName("rows")] SyncDiffRow[] Rows);

public sealed record ScmRestoreReply(
    [property: JsonPropertyName("detail")] string Detail,
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("ref")] string Ref,
    /// <summary>written | added | unchanged | skipped | failed.</summary>
    [property: JsonPropertyName("did")] string Did,
    [property: JsonPropertyName("why")] string? Why);

public sealed record ScmBlameRow(
    [property: JsonPropertyName("line")] int Line,
    [property: JsonPropertyName("hash")] string Hash,
    [property: JsonPropertyName("short")] string ShortHash,
    [property: JsonPropertyName("author")] string Author,
    [property: JsonPropertyName("when")] string When,
    [property: JsonPropertyName("summary")] string Summary);

/// <summary>Blame in the editor's numbering: committed lines, and the live lines no commit holds.</summary>
public sealed record ScmBlameReply(
    [property: JsonPropertyName("detail")] string Detail,
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("head")] string Head,
    [property: JsonPropertyName("lines")] ScmBlameRow[] Lines,
    [property: JsonPropertyName("uncommitted")] int[] Uncommitted);

/// <summary>Why the request could not be answered. The harness client throws on it.</summary>
public sealed record ScmErrorReply([property: JsonPropertyName("error")] string Error);

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[JsonSerializable(typeof(ScmRowReply))]
[JsonSerializable(typeof(ScmBranchRow))]
[JsonSerializable(typeof(ScmIdentityReply))]
[JsonSerializable(typeof(ScmLastCommitReply))]
[JsonSerializable(typeof(ScmStatusReply))]
[JsonSerializable(typeof(ScmSkippedReply))]
[JsonSerializable(typeof(ScmCommitReply))]
[JsonSerializable(typeof(ScmImportReply))]
[JsonSerializable(typeof(ScmUndoReply))]
[JsonSerializable(typeof(ScmFileRow))]
[JsonSerializable(typeof(ScmCommitRow))]
[JsonSerializable(typeof(ScmLogReply))]
[JsonSerializable(typeof(SyncDiffRow))]
[JsonSerializable(typeof(ScmDiffReply))]
[JsonSerializable(typeof(ScmTextReply))]
[JsonSerializable(typeof(ScmRestoreReply))]
[JsonSerializable(typeof(ScmBlameRow))]
[JsonSerializable(typeof(ScmBlameReply))]
[JsonSerializable(typeof(ScmErrorReply))]
[JsonSerializable(typeof(ScmRowReply[]))]
[JsonSerializable(typeof(ScmBranchRow[]))]
[JsonSerializable(typeof(ScmSkippedReply[]))]
[JsonSerializable(typeof(ScmFileRow[]))]
[JsonSerializable(typeof(ScmCommitRow[]))]
[JsonSerializable(typeof(SyncDiffRow[]))]
[JsonSerializable(typeof(ScmBlameRow[]))]
[JsonSerializable(typeof(string[]))]
[JsonSerializable(typeof(int[]))]
internal sealed partial class ScmJsonContext : JsonSerializerContext;
