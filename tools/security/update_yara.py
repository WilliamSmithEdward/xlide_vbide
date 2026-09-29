"""Propose pinned Forge releases without approving new compiler diagnostics."""
import argparse
import copy
from datetime import datetime
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess

from malware import CONFIG, load_yara, write_json

POLICY_PATH = ".github/security/malware/policy.json"
PREFIX = "https://github.com/YARAHQ/yara-forge/releases/download/"
ASSET = "yara-forge-rules-full.zip"


def release_pin(release):
    tag = release.get("tag_name", "")
    if release.get("draft") is not False or release.get("prerelease") is not False or not re.fullmatch(r"\d{8}", tag):
        raise ValueError("Expected a stable dated YARA Forge release")
    datetime.strptime(tag, "%Y%m%d")
    assets = [asset for asset in release["assets"] if asset["name"] == ASSET]
    if len(assets) != 1:
        raise ValueError("Expected exactly one full rule archive")
    asset = assets[0]
    url = PREFIX + tag + "/" + ASSET
    checksum = asset.get("digest", "")
    if asset["browser_download_url"] != url or not re.fullmatch(r"sha256:[a-f0-9]{64}", checksum):
        raise ValueError("Missing upstream SHA-256 or unexpected archive URL")
    if not 0 < asset["size"] <= 64 * 1024 * 1024:
        raise ValueError("Invalid archive size")
    return tag, url, checksum.removeprefix("sha256:")


def prepare(policy, release, output, base_sha, compile_rules=load_yara):
    tag, url, checksum = release_pin(release)
    current = re.fullmatch(re.escape(PREFIX) + r"(\d{8})/" + re.escape(ASSET), policy["rules"]["url"])
    if not current or not re.fullmatch(r"[a-f0-9]{40}", base_sha):
        raise ValueError("Unexpected current pin or base commit")
    if tag < current[1]:
        raise ValueError("Upstream latest release would downgrade the current pin")
    if tag == current[1]:
        if checksum != policy["rules"]["sha256"]:
            raise ValueError("An already pinned upstream release changed its checksum")
        return False
    candidate = copy.deepcopy(policy)
    candidate["rules"].update(url=url, sha256=checksum)
    evidence = {}
    output.mkdir(parents=True, exist_ok=True)
    compile_rules(output, evidence, candidate)
    if not isinstance(evidence.get("ruleCount"), int) or evidence["ruleCount"] <= 0:
        raise ValueError("Empty or invalid compiled rule collection")
    candidate["rules"]["count"] = evidence["ruleCount"]
    # Never move reviewBy, accept new diagnostics, or introduce detection exceptions.
    write_json(output / "proposal.json", {"baseSha": base_sha, "release": release, "policy": candidate})
    write_json(output / "candidate-evidence.json", evidence)
    counts = {}
    for warning in evidence["diagnostics"]:
        counts[warning[0]] = counts.get(warning[0], 0) + 1
    fingerprint = hashlib.sha256(json.dumps(evidence["diagnostics"], ensure_ascii=False, separators=(",", ":")).encode()).hexdigest()
    body = (
        f"Updates YARA Forge full to [{tag}](https://github.com/YARAHQ/yara-forge/releases/tag/{tag}).\n\n"
        f"- Archive SHA-256: `{checksum}` (verified against GitHub's release asset digest).\n"
        f"- Rules: {policy['rules']['count']} → {candidate['rules']['count']}.\n"
        f"- Compiler warning counts: `{json.dumps(counts, sort_keys=True)}`.\n"
        f"- Candidate diagnostic fingerprint: `{fingerprint}`.\n\n"
        "The updater has NOT approved compiler warnings, extended the review deadline, or changed detection exceptions. "
        "Security is explicitly dispatched on this PR branch. Changed diagnostics deliberately fail its gate until "
        "a maintainer reviews the evidence and updates the documented baseline. Do not copy the fingerprint just to pass CI.\n\n"
        "Review provider/rule changes, scan results, compiler diagnostics, and docs/security-malware.md before merging. "
        "This updater never merges or approves PRs.\n"
    )
    (output / "body.md").write_text(body, encoding="utf-8")
    return True


def api(endpoint, method="GET", payload=None):
    args = ["gh", "api", "--method", method, endpoint]
    if payload is not None:
        args += ["--input", "-"]
    result = subprocess.run(args, input=json.dumps(payload) if payload is not None else None,
                            text=True, encoding="utf-8", capture_output=True, check=True, timeout=120)
    return json.loads(result.stdout) if result.stdout.strip() else None


def publish(proposal, body, original, repository, call=api):
    tag, url, checksum = release_pin(proposal["release"])
    candidate = proposal["policy"]
    allowed = copy.deepcopy(original)
    allowed["rules"].update(url=url, sha256=checksum, count=candidate["rules"]["count"])
    if candidate != allowed or not isinstance(candidate["rules"]["count"], int) or candidate["rules"]["count"] <= 0:
        raise ValueError("Proposal changed fields outside the allowed rule pin")
    base_sha = proposal["baseSha"]
    if not re.fullmatch(r"[a-f0-9]{40}", base_sha) or not re.fullmatch(r"[\w.-]+/[\w.-]+", repository):
        raise ValueError("Invalid publishing context")
    branch = "codex/yara-forge-" + tag
    root = "repos/" + repository
    pulls = call(f"{root}/pulls?state=all&head={repository.split('/')[0]}:{branch}&base=main&per_page=100")
    if pulls:
        if len(pulls) != 1:
            raise ValueError("Ambiguous existing update PRs")
        pull = pulls[0]
        if pull["state"] != "open":
            return None  # Respect a maintainer's prior close/merge decision.
    else:
        refs = call(f"{root}/git/matching-refs/heads/{branch}")
        exact = [ref for ref in refs if ref["ref"] == "refs/heads/" + branch]
        if exact:
            # An interrupted prior publish may have created the branch. Do not overwrite it.
            raise ValueError("Update branch exists without a PR; inspect it before retrying")
        base = call(f"{root}/git/commits/{base_sha}")
        tree = call(f"{root}/git/trees", "POST", {
            "base_tree": base["tree"]["sha"], "tree": [{"path": POLICY_PATH, "mode": "100644", "type": "blob",
                                                        "content": json.dumps(candidate, indent=2) + "\n"}]})
        commit = call(f"{root}/git/commits", "POST", {"message": f"Update YARA Forge full to {tag}",
                                                     "tree": tree["sha"], "parents": [base_sha]})
        call(f"{root}/git/refs", "POST", {"ref": "refs/heads/" + branch, "sha": commit["sha"]})
        pull = call(f"{root}/pulls", "POST", {"title": f"Update YARA Forge full to {tag}", "head": branch,
                                             "base": "main", "body": body})
    # workflow_dispatch runs with GITHUB_TOKEN; ordinary bot PR events can require approval.
    for workflow in ("security.yml", "build.yml"):
        call(f"{root}/actions/workflows/{workflow}/dispatches", "POST", {"ref": branch})
    return pull["html_url"]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("mode", choices=("prepare", "publish"))
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--release", type=Path)
    args = parser.parse_args()
    original = json.loads((CONFIG / "policy.json").read_text(encoding="utf-8"))
    if args.mode == "prepare":
        changed = prepare(original, json.loads(args.release.read_text(encoding="utf-8")), args.output, os.environ["GITHUB_SHA"])
        if os.environ.get("GITHUB_OUTPUT"):
            with open(os.environ["GITHUB_OUTPUT"], "a", encoding="utf-8") as output:
                output.write(f"changed={str(changed).lower()}\n")
        print("Prepared an update proposal" if changed else "Already pinned to the latest stable YARA Forge release")
    else:
        proposal = json.loads((args.output / "proposal.json").read_text(encoding="utf-8"))
        if proposal["baseSha"] != os.environ["GITHUB_SHA"]:
            raise ValueError("Proposal belongs to a different workflow commit")
        url = publish(proposal, (args.output / "body.md").read_text(encoding="utf-8"), original, os.environ["GITHUB_REPOSITORY"])
        print(url or "The existing update PR was closed; leaving that decision unchanged")
        if url and os.environ.get("GITHUB_STEP_SUMMARY"):
            with open(os.environ["GITHUB_STEP_SUMMARY"], "a", encoding="utf-8") as summary:
                summary.write(f"Proposed YARA Forge update: {url}\n\nSecurity and build explicitly dispatched on its branch.\n")


if __name__ == "__main__":
    main()
