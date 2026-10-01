# Security policy

## Reporting a vulnerability

Report a vulnerability privately, not in a public issue or pull request:
[open a private report](https://github.com/WilliamSmithEdward/xlide_vbide/security/advisories/new).
Only the maintainer sees it. Include the xlide version, the host (Excel,
Word, PowerPoint or Access) and the impact you observed, and the smallest
file or steps that show it, with credentials and private data removed.

A confirmed vulnerability is fixed in a release on the GitHub releases page,
and the advisory is published with it, crediting you unless you ask otherwise.

## Supported versions

Only the latest release on the GitHub releases page receives security
fixes. Older releases are not maintained separately; update when a fix
ships.

## Scope

xlide is a native add-in that the Visual Basic Editor loads inside Excel,
Word, PowerPoint and Access. It shows the VBA modules of the open documents
in WebView2 editor pages, and a separate process, `xlide-engine.exe`,
analyses them over a named pipe. Module text comes from whatever a document
carries, often somebody else's, and is treated as input xlide does not
control. The Source Control pane runs the git you already have; Git for
Windows' credential manager handles signing in, and xlide never sees a
credential.

The installer, `xlide-setup.exe`, installs to `%LOCALAPPDATA%\Programs\xlide`
for the current user, needs no administrator rights, and changes nothing
outside your own profile. It carries no code signature yet, so Windows warns
before running it.

A document, module text or a local request that makes xlide run code you did
not ask for, change a project you did not ask it to, or act for someone
other than the signed-in user is a vulnerability here.

### The local API

The local API, off by default, lets an agent read and write modules, run
tests and drive the editor. Turned on from the agent card, it listens on
127.0.0.1 at a random port, and every request needs a random token written
to a file under your profile. Anything running as you can read that file,
and anything holding the token can read, write and run code in every project
the editor can reach. Turn it on only while an agent needs it. See
[docs/xlide-api.md](docs/xlide-api.md).

## How the code is checked

Three workflows check every pull request and every push to `main`, and
their gates decide whether a change can merge: **CI passed**,
**Security passed** and **Malware scan passed**. A gate passes only when
every job before it did, and any unexpected finding fails it, whatever its
severity. Security and Malware scan also run daily at 06:23 UTC, on a
published release, and by hand. A scanner failure, a compile error, a
malformed report or a missing result fails the gate too.

- **Code:** CodeQL with the `security-extended` queries, for C#,
  JavaScript/TypeScript and GitHub Actions, and Semgrep Community Edition
  with `p/security-audit` and `p/secrets`. C# analysis builds the Release
  configuration on Windows with the repository's .NET SDK, source generators
  included; other build configurations are not analysed. Semgrep runs with
  inline `nosemgrep` comments disabled. A finding fails at every severity,
  including one the SARIF marks as suppressed, and dismissing an alert in
  code scanning does not change that. Results go to the repository's code
  scanning; a release run keeps them as workflow artifacts instead.
- **Workflows:** zizmor audits the GitHub Actions workflows; a finding fails
  Security.
- **Dependencies:** the Dependency audit job runs `npm audit` over the
  engine's and the editor page's lockfiles, development tools included,
  without running install scripts, and lists every NuGet package of the
  solution with a known vulnerability, transitive ones included. Any
  advisory fails Security, whatever its severity. The editor page overrides
  the DOMPurify that monaco-editor 0.57.0 pins exactly (3.4.15) with
  3.4.16, the release that fixes GHSA-p98j-92pf-mc4p; remove the override
  in `ui/editor/package.json` once monaco-editor ships a fixed DOMPurify
  itself.
- **Malware:** ClamAV, with signatures freshclam fetches and verifies on
  every run, and YARA-X, with the YARA Forge rules pinned to a release and
  its SHA-256, scan every Git-tracked file and, on a published release,
  every asset, which must include `xlide-setup.exe`, as data without running
  it. ClamAV uses the official Talos `main`, `daily` and `bytecode`
  databases, with up to three update attempts and no fallback to stale
  signatures, and scans with the network turned off. YARA-X uses the YARA
  Forge full rule set, with no rules removed. Both engines run positive and
  negative controls, and the report checks that both scanned the same
  complete inventory. Neither engine is sure to unpack the installer's
  compressed payload, and untracked dependencies and the bundled analyzer
  source are not scanned here. Raw evidence stays a workflow artifact, not a
  code scanning alert. The details are in
  [Malware scanning and accepted detections](docs/security-malware.md).
- **Fuzzing:** fast-check properties in `ui/editor/test/properties.ts`
  check the editor page's readers of module text. Format Module must keep
  every line, change only indentation and the case of words, indent in whole
  units of spaces, and change nothing when run again; the procedure scanner
  must tile the module; and the tokenizer must cover every line with tokens
  in order. `npm test` runs every property a hundred times. The Fuzz workflow
  runs on every change to the editor page's source and daily, twenty
  thousand times per property on a change and two hundred thousand daily. It
  is not a gate: a finding becomes a regression test with its fix.
- **OpenSSF Scorecard** rates the repository's security practices on every
  change to `main` and weekly, and the README badge shows the result.
  One of its checks does not fit this project: a single maintainer cannot
  have a second person approve every change. Signed-Releases rises as
  releases carry the provenance bundle; it counts the last five.

## Accepted findings

A finding is fixed, or accepted with a written reason in
[`.github/security/malware/policy.json`](.github/security/malware/policy.json),
which holds ClamAV and YARA-X detections only. An entry matches the scanner,
the exact rule, the exact path and the file's SHA-256, so a changed file
needs another review. Each entry also names its reviewer and an expiry date;
an expired, wildcard or unexplained entry fails the report, and so does an
entry that no longer matches a finding, until it is removed. An entry for a
release asset is judged only by a scan of a release. A ClamAV alert for
a scan limit or encrypted content cannot be accepted. CodeQL and Semgrep
have no accepted list. zizmor keeps its exceptions in `.github/zizmor.yml`
or inline beside the line they excuse, each with its reason: two rules are
off there, both for publish.yml.

## Pinning and updates

Everything the workflows run is pinned: actions to full commit SHAs,
runners to named OS releases, scanner images to digests, Python tools to
hash-locked lock files, the project's own dependencies to the npm lockfiles
of `engine` and `ui/editor`, installed with `npm ci`, and to exact NuGet
versions, the YARA-X engine to its hash-locked Python wheel, and the YARA
Forge rules to a release and its SHA-256. ClamAV's signatures change too
often to pin, so freshclam fetches and verifies them on every run. Semgrep's
registry rules are also fetched at scan time.

One input is deliberately not pinned in CI: it tests the engine against the
companion analyzer's `main` in xlide_vscode, so the two move together. A
release builds the engine from the analyzer commit named in
`.github/analyzer.json`, which moves by hand; `tools/Pin-Analyzer.ps1` makes
the same pin for a local build.

Dependabot proposes updates to GitHub Actions, npm, NuGet and the ClamAV and
Semgrep images weekly, and to the hash-locked Python tools (YARA-X and
zizmor) daily, in every ecosystem once a version is a week old; and at once
for a security advisory. The Update YARA rules workflow proposes new YARA
pins each week. A minor or patch update, and the YARA pull request, merges
itself once CI, Security and Malware scan pass; a third-party major version
waits for review.

## Releases

Pushing a `vX.Y.Z` tag runs the Publish workflow. On a Windows runner it
builds `xlide-setup.exe` from the tagged commit: the language engine from
the analyzer commit in `.github/analyzer.json`, tested against it first,
then the editor page and the native shim. It refuses a tag that is not the
`<Version>` in `Directory.Build.props`, or a version without release notes
in `docs/releases/vX.Y.Z.md`. Security runs on that commit and Malware scan
on that commit and that installer, and only when both pass does Publish sign
the installer's build provenance and create the GitHub release with:

- `xlide-setup.exe`
- `xlide-<version>.sigstore.json`, its signed build provenance
- `security-report.md` and `security-report.json`
- `malware-report.md` and `malware-report.json`, which give the
  installer's SHA-256

To check that an installer was built by this repository's Publish workflow
from a tagged commit:

```bash
gh attestation verify xlide-setup.exe --repo WilliamSmithEdward/xlide_vbide
```

The provenance is not an Authenticode signature, so Windows still shows the
installer's publisher as unknown. The live gate (`tools/verify.ps1`) needs a
real Excel and cannot run in CI; it is run on the release commit before the
tag is pushed. Started by hand, Publish is a dry run: it builds, scans and
assembles the same files as the `release-preview` artifact and releases
nothing. The raw scan results stay in the workflow runs for 30 days and the
reports for 90. Releases up to 0.20.2 were built locally and carry no
provenance, and releases before these workflows carry no reports.

The checks record what was scanned. They do not certify the binaries, audit
every dependency, or replace manual review.

### Verifying a download

GitHub records a SHA-256 digest for every release asset. Compare it with
the file you downloaded:

```powershell
gh release view <tag> --repo WilliamSmithEdward/xlide_vbide --json assets --jq '.assets[] | [.name, .digest] | @tsv'
Get-FileHash .\xlide-setup.exe -Algorithm SHA256
```

A release whose malware scan ran after publication also carries
`malware-report.md`, which lists the same digest and names the commit and
workflow run that scanned the installer. A mismatch means the file is not
the one released.

## Repository settings

<!-- repo-standards:begin security-settings. Copied from WilliamSmithEdward/repo-standards, templates/security/settings-block.md. Change it there; the weekly rescan fails a copy that differs. -->
- `main` accepts changes only through a pull request that passes
  **CI passed**, **Security passed** and **Malware scan passed**. The
  ruleset has no bypass, for the owner either, and refuses force-pushes and
  deleting the branch.
- A `v*` release tag cannot be moved or deleted once pushed, except by a
  repository admin.
- A workflow that uses an action not pinned to a full commit SHA fails to
  run. Workflow tokens are read-only unless a job is granted more for
  itself.
- Secret scanning with push protection, Dependabot alerts and security
  updates, and private vulnerability reporting are on.
<!-- repo-standards:end -->
