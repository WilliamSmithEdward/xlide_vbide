import copy
import json
from pathlib import Path
import tempfile
import unittest

from update_yara import ASSET, CONFIG, PREFIX, prepare, publish, release_pin


class UpdateTests(unittest.TestCase):
    def setUp(self):
        self.policy = json.loads((CONFIG / "policy.json").read_text())
        self.release = {"tag_name": "20990101", "draft": False, "prerelease": False, "assets": [{
            "name": ASSET, "browser_download_url": PREFIX + "20990101/" + ASSET,
            "digest": "sha256:" + "a" * 64, "size": 100}]}
        self.base = "b" * 40

    def compile(self, output, evidence, candidate):
        self.assertEqual(candidate["rules"]["sha256"], "a" * 64)
        evidence.update(ruleCount=15000, diagnostics=[["new-warning", "Needs review", []]])

    def test_proposal_pins_release_but_never_approves_diagnostics(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            self.assertTrue(prepare(self.policy, self.release, output, self.base, self.compile))
            proposal = json.loads((output / "proposal.json").read_text())
            expected = copy.deepcopy(self.policy)
            expected["rules"].update(url=PREFIX + "20990101/" + ASSET, sha256="a" * 64, count=15000)
            self.assertEqual(proposal["policy"], expected)
            self.assertIn("NOT approved", (output / "body.md").read_text())

    def test_same_pin_noop_replacement_downgrade_and_bad_metadata_fail(self):
        current = self.policy["rules"]["url"].split("/")[-2]
        release = copy.deepcopy(self.release)
        release["tag_name"] = current
        release["assets"][0].update(browser_download_url=self.policy["rules"]["url"], digest="sha256:" + self.policy["rules"]["sha256"])
        with tempfile.TemporaryDirectory() as directory:
            self.assertFalse(prepare(self.policy, release, Path(directory), self.base, self.compile))
            release["assets"][0]["digest"] = "sha256:" + "f" * 64
            with self.assertRaisesRegex(ValueError, "changed its checksum"):
                prepare(self.policy, release, Path(directory), self.base, self.compile)
            release["tag_name"] = "20000101"
            release["assets"][0]["browser_download_url"] = PREFIX + "20000101/" + ASSET
            with self.assertRaisesRegex(ValueError, "downgrade"):
                prepare(self.policy, release, Path(directory), self.base, self.compile)
        for field, value in [("digest", None), ("browser_download_url", "http://example.org/rules.zip"), ("size", 0)]:
            release = copy.deepcopy(self.release)
            release["assets"][0][field] = value
            with self.assertRaises((ValueError, TypeError)):
                release_pin(release)
        for field, value in [("tag_name", "latest"), ("tag_name", "20990230"), ("prerelease", True), ("draft", True)]:
            release = copy.deepcopy(self.release)
            release[field] = value
            with self.assertRaises(ValueError):
                release_pin(release)

    def test_publication_is_policy_only_and_always_dispatches_existing_ci(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            prepare(self.policy, self.release, output, self.base, self.compile)
            proposal = json.loads((output / "proposal.json").read_text())
        calls = []

        def api(endpoint, method="GET", payload=None):
            calls.append((endpoint, method, payload))
            if "/pulls?" in endpoint or "/matching-refs/" in endpoint:
                return []
            if method == "GET":
                return {"tree": {"sha": "c" * 40}}
            if endpoint.endswith("/pulls"):
                return {"html_url": "https://github.com/owner/repo/pull/1"}
            return {"sha": "d" * 40}

        self.assertEqual(publish(proposal, "Reviewed proposal", self.policy, "owner/repo", api), "https://github.com/owner/repo/pull/1")
        tree = next(payload for endpoint, _, payload in calls if endpoint.endswith("/git/trees"))
        self.assertEqual([entry["path"] for entry in tree["tree"]], [".github/security/malware/policy.json"])
        self.assertTrue(calls[-2][0].endswith("security.yml/dispatches"))
        self.assertTrue(calls[-1][0].endswith("build.yml/dispatches"))
        proposal["policy"]["compilerDiagnostics"]["sha256"] = "e" * 64
        with self.assertRaisesRegex(ValueError, "outside the allowed"):
            publish(proposal, "body", self.policy, "owner/repo", api)

    def test_existing_pr_is_not_overwritten_or_reopened(self):
        candidate = copy.deepcopy(self.policy)
        candidate["rules"].update(url=PREFIX + "20990101/" + ASSET, sha256="a" * 64)
        proposal = {"release": self.release, "policy": candidate, "baseSha": self.base}
        for state in ["open", "closed"]:
            calls = []

            def api(endpoint, method="GET", payload=None):
                calls.append(endpoint)
                if "/pulls?" in endpoint:
                    return [{"state": state, "html_url": "https://github.com/owner/repo/pull/1"}]
                self.assertIn("/dispatches", endpoint)

            result = publish(proposal, "body", self.policy, "owner/repo", api)
            self.assertEqual(len(calls), 3 if state == "open" else 1)
            self.assertEqual(result is None, state == "closed")


if __name__ == "__main__":
    unittest.main()
