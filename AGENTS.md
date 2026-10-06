# Notes for agents

<!-- repo-standards:begin. Copied from WilliamSmithEdward/repo-standards, templates/agents/AGENTS-block.md. Change it there; the weekly rescan fails a copy that differs. -->
## Releases, CI and security

These rules are the same in every WilliamSmithEdward repository.

- **How a release happens here:** pushing a `vX.Y.Z` tag runs Publish, which builds the release files in CI and creates the GitHub release with them, their signed provenance and the security reports. Any other step, such as a marketplace upload, is described elsewhere in this file.
- **Starting a workflow by hand never releases anything.** Publish and every
  release report are dry runs when started with `gh workflow run` or the Run
  workflow button. They build, scan and assemble the release files exactly
  as a release would, and upload them as the `release-preview` artifact
  instead. Run one after changing anything on the release path:
  `gh workflow run <file> --ref main`, then
  `gh run download <run-id> -n release-preview`.
- **Do not create, publish, edit or delete a release or a `v*` tag** unless
  the owner asks for it. A `v*` tag cannot be moved or deleted once pushed.
- **Every change to `main` goes through a pull request** that passes CI
  passed, Security passed and Malware scan passed. No one can push to `main`
  directly or skip the checks, admins included. Push a branch, open a pull
  request, and let it merge itself: `gh pr merge --auto --squash <number>`.
- **Pins.** Actions by full commit SHA with the version as a comment. Images
  by digest, in `.github/security/<tool>/Dockerfile`. Python tools from the
  hash-locked `.github/requirements/<purpose>.txt`, compiled from the `.in`
  beside it with
  `uv pip compile <purpose>.in --universal --generate-hashes --python-version 3.12 -o <purpose>.txt`.
  Runners are named releases, never `-latest`.
- **Updates merge themselves.** Dependabot and the Update YARA rules workflow
  open pull requests that merge once the three checks pass, except a
  third-party major version, which waits for the owner. Leave them alone
  unless asked.
- **A scanner finding is fixed or accepted with a written reason** in the
  repository's accepted list. Never silence a scanner without one.
<!-- repo-standards:end -->

## Verifying issue reports

For an issue about behavior inside Excel or the VBE, use a focused live harness when the
host is available. Reproduce the contributor's concrete case before the fix when possible,
then test the same case against the built add-in in a real host. Check the observable result
the contributor reported (for example, the actual completion menu or Problems finding),
not only an engine response or a successful build. Use a disposable workbook or module,
clean up the harness-owned state, and record both the live result and any verification
limits in the issue outcome. Keep the targeted check proportional to the report; the full
live release gate belongs to releases.

## Releasing

1. Set `<Version>` in `Directory.Build.props`, write the release notes in
   `docs/releases/vX.Y.Z.md`, and set `.github/analyzer.json` to the
   xlide_vscode release the engine ships with: `ref` its tag, `commit` the
   commit that tag names
   (`gh api repos/WilliamSmithEdward/xlide_vscode/git/ref/tags/<tag>`).
2. Run the live gate on that commit, pinned the same way:
   `$env:XLIDE_ANALYZER_ROOT = (tools\Pin-Analyzer.ps1 -Ref <commit>)`, then
   `tools\verify.ps1 -Deep`. CI cannot run it: it drives a real Excel.
3. Merge to `main`. Optionally dry-run Publish:
   `gh workflow run publish.yml --ref main`, then
   `gh run download <run-id> -n release-preview`.
4. The owner pushes the `vX.Y.Z` tag on that commit. Publish builds the
   installer from it on a Windows runner, runs Security and Malware scan,
   and only when both pass signs the installer and creates the release.
