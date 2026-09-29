# Security

## Reporting a vulnerability privately

Use [Report a vulnerability](https://github.com/WilliamSmithEdward/xlide_vbide/security/advisories/new)
to contact the maintainers privately through GitHub. Private vulnerability reporting is enabled.
Please do not disclose vulnerabilities in public issues or pull requests before coordination.
Include the affected version, reproduction steps, expected impact, and a minimal example without
real credentials or private workbook data.

## Supported versions

Security fixes target the latest release and `main`. Older releases are not maintained separately;
upgrade to the latest version when a fix is available.

## Automated analysis

The [Security workflow](https://github.com/WilliamSmithEdward/xlide_vbide/actions/workflows/security.yml)
runs on pull requests, pushes to `main`, weekly, manually, and when a new release is published.
CodeQL scans C#, JavaScript/TypeScript, and GitHub Actions with `security-extended` queries.
C# uses build-free analysis, so generated code and build-specific configurations may not be covered.
Semgrep Community Edition scans source with `p/security-audit` and `p/secrets`; registry rules are
retrieved at scan time. The scanner version is pinned and updated by Dependabot.

The **Security gate** fails on any finding, including low severity and suppressed SARIF findings,
scanner failure, warning/error diagnostic, malformed report, or missing scanner result. There is no
baseline or allowlist, and Semgrep inline `nosemgrep` suppressions are disabled. Findings must be
investigated and fixed; any future policy exception must be an explicit, reviewed repository change.
GitHub code scanning dismissal alone does not bypass this gate.

Raw SARIF is uploaded to GitHub code scanning for branch/PR runs and retained as workflow artifacts
for 30 days. Summary reports are retained for 90 days. Dependabot checks npm, NuGet, GitHub Actions,
and the Semgrep Python requirement weekly. Dependabot alerts and security updates are enabled.

## Release security reports

Releases published after this workflow is merged and included in their tagged commit receive
`security-report.md` and `security-report.json`. Reports name the scanned commit and workflow run,
list findings by scanner, and clearly mark failed or incomplete scans. Only summaries are attached
publicly; raw finding details remain in workflow artifacts (also accessible for this public repo).
Existing releases are not backfilled. Reports arrive asynchronously after publication; a failed scan
does not unpublish a release. A workflow infrastructure failure can prevent attachment and requires
a maintainer to investigate the failed run and rerun it.

For the local installer release process:

1. Merge changes and wait for the Security workflow to pass on the exact release commit.
2. Tag that commit and create a **draft** GitHub release.
3. Run `tools/release.ps1 -Tag <tag>`. It requires a clean checkout and the latest eligible Security
   run for that exact commit to have passed, validates the downloaded report, and uploads it with
   the installer. `-SkipGate` skips only the existing local verification gate, never security.
4. Review the assets and publish the draft. Publication triggers another scan of the tagged commit.

Repository administrators can require **Security gate** in branch protection/rulesets to prevent
merging failed checks; adding this workflow does not itself change those merge rules. GitHub UI or
CLI release creation outside `tools/release.ps1` is not blocked by the local gate.

These reports analyze this repository's source. They do not certify the installer binary, audit
all dependency vulnerabilities, scan the separately bundled `xlide_vscode` analyzer, or replace
manual security review. A passing scan is not a guarantee that the product is vulnerability-free.
