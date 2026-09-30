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
runs on pull requests, pushes to `main`, daily at 06:23 UTC, manually, and when a new release is published.
CodeQL scans C#, JavaScript/TypeScript, and GitHub Actions with `security-extended` queries.
C# analysis builds the Release configuration on Windows with the repository's .NET SDK, including
source generators. Other build configurations are not analyzed by this workflow.
Semgrep Community Edition scans source with `p/security-audit` and `p/secrets`; registry rules are
retrieved at scan time. Semgrep runs from its official image, pinned by digest in `.github/security/semgrep/Dockerfile` and updated by Dependabot.
The separate **Malware scan** workflow scans every tracked file, one job each for ClamAV and YARA-X. ClamAV downloads and tests fresh
official Talos signatures on every run; update failures fail the job with no stale fallback. YARA-X scans
the same inventory with the SHA-256-pinned YARA Forge **full** public collection. Published-release
runs also download and scan release assets, including `xlide-setup.exe`, without executing them.
Versions, hashes, coverage, and the bounded exception policy are in
[Malware scanning and accepted detections](docs/security-malware.md).

The **Security passed** and **Malware scan passed** gates fail on unexpected findings, scanner failure, compile errors,
malformed reports, or missing results. CodeQL/Semgrep findings still fail at every severity, including
suppressed SARIF findings; Semgrep inline `nosemgrep` suppressions are disabled. Malware detections
may be bypassed only by an explicitly reviewed, expiring entry matching scanner, rule, exact path,
and SHA-256. The current detection exception list is empty. Upstream YARA compiler warnings are
recorded, not gated, and no rules are disabled. Operational failures and missing coverage
cannot be bypassed. Changes to the accepted evidence require a reviewed repository change.
GitHub code scanning dismissal alone does not bypass this gate.

Raw SARIF is uploaded to GitHub code scanning for branch/PR runs and retained as workflow artifacts
for 30 days. Summary reports are retained for 90 days. Dependabot checks npm, NuGet, GitHub Actions,
the YARA-X Python requirements, and the ClamAV and Semgrep images weekly. Public YARA rules are pinned
for review rather than silently following a moving download. The weekly YARA updater proposes new
pins in `.github/security/yara.json` in a PR and explicitly starts CI, Security and Malware scan on
its branch. It never merges PRs. Dependabot alerts and security updates
are enabled. Raw malware evidence is a workflow artifact, not a GitHub code-scanning alert.

## Release security reports

Releases published after this workflow is merged and included in their tagged commit receive
`security-report.md` and `security-report.json` from the Security workflow, and `malware-report.md`
and `malware-report.json` from the Malware scan workflow. Reports name the scanned commit and workflow run,
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
4. Review the assets and publish the draft. Publication triggers another scan of the tagged commit
   **and the attached release assets**. That run replaces the initial source-only reports with reports
   listing each scanned asset's SHA-256. Download or scan failures produce a failing/incomplete report.

Repository administrators can require **Security passed** and **Malware scan passed** in branch
protection/rulesets to prevent merging failed checks; adding these workflows does not itself change
those merge rules. GitHub UI or
CLI release creation outside `tools/release.ps1` is not blocked by the local gate.

The pre-publication check is source-only; asset malware scanning occurs after publication and does
not unpublish a release. Neither scanner is guaranteed to unpack the installer's custom compressed
payload. Untracked dependencies and the separately bundled `xlide_vscode` analyzer source are outside
the tracked-file scan. These checks do not certify binaries, audit every dependency vulnerability,
or replace manual review. A passing scan is not a guarantee that the product is vulnerability-free.
