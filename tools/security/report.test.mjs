import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, writeFileSync, readFileSync, rmSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createReport, expectedScans, inspectSarif } from './report.mjs';
import { fixtures } from './malware-fixtures.mjs';

const clean = () => ({ version: '2.1.0', runs: [{ tool: { driver: { name: 'test-scanner', version: '1' } }, results: [] }] });
const statuses = { codeql: { result: 'success' }, semgrep: { result: 'success' }, dependencies: { result: 'success' } };
const malwareStatuses = { clamav: { result: 'success' }, 'yara-x': { result: 'success' } };

test('all findings count, including notes and suppressed results', () => {
  const sarif = clean();
  sarif.runs[0].results = [{ level: 'note' }, { suppressions: [{ status: 'accepted' }] }];
  assert.equal(inspectSarif(sarif).findings, 2);
});

test('failed execution and warning diagnostics cannot pass as a clean scan', () => {
  const sarif = clean();
  sarif.runs[0].invocations = [{ executionSuccessful: false, toolExecutionNotifications: [{ level: 'warning' }] }];
  assert.equal(inspectSarif(sarif).problems.length, 2);
});

test('malformed and empty SARIF is rejected', () => {
  for (const sarif of [{}, { version: '2.1.0', runs: [] }, { version: '2.1.0', runs: [{}] }]) {
    assert.throws(() => inspectSarif(sarif));
  }
});

test('gate requires every scanner, valid artifacts and successful jobs', () => {
  const root = mkdtempSync(join(tmpdir(), 'xlide-security-'));
  const { policy, result: yara, clam } = fixtures();
  const report = (status, context = {}) => createReport(root, status, policy, context);
  const malware = (status, context = {}) => createReport(root, status, policy, context, 'malware');
  try {
    assert.equal(report(statuses).passed, false);
    assert.equal(malware(malwareStatuses).passed, false);
    for (const name of expectedScans) {
      mkdirSync(join(root, name));
      writeFileSync(join(root, name, 'result.sarif'), JSON.stringify(clean()));
    }
    assert.equal(report(statuses).passed, true);
    assert.equal(malware(malwareStatuses).passed, false);
    for (const result of [yara, clam]) {
      result.commit = process.env.GITHUB_SHA;
      mkdirSync(join(root, result.scanner));
      writeFileSync(join(root, result.scanner, 'result.json'), JSON.stringify(result));
    }
    assert.equal(malware(malwareStatuses).passed, true);
    assert.equal(malware(malwareStatuses, { GITHUB_SHA: 'different-commit' }).passed, false);
    assert.equal(malware(malwareStatuses, { GITHUB_EVENT_NAME: 'release' }).passed, false);
    // The installer is the malware report's concern, not the security report's.
    assert.equal(report(statuses, { GITHUB_EVENT_NAME: 'release' }).passed, true);
    for (const result of [yara, clam]) {
      result.inventory.push({ path: 'release/xlide-setup.exe', sha256: 'd'.repeat(64), size: 123 });
      result.scannedFiles++;
      writeFileSync(join(root, result.scanner, 'result.json'), JSON.stringify(result));
    }
    assert.equal(malware(malwareStatuses, { GITHUB_EVENT_NAME: 'release' }).passed, true);
    clam.inventory[1].sha256 = 'e'.repeat(64);
    writeFileSync(join(root, 'clamav', 'result.json'), JSON.stringify(clam));
    assert.equal(malware(malwareStatuses).passed, false);
    clam.inventory[1].sha256 = 'd'.repeat(64);
    writeFileSync(join(root, 'clamav', 'result.json'), JSON.stringify(clam));
    for (const result of ['failure', 'cancelled', 'skipped', undefined]) {
      assert.equal(report({ ...statuses, semgrep: { result } }).passed, false);
      assert.equal(malware({ ...malwareStatuses, 'yara-x': { result } }).passed, false);
    }
    assert.equal(malware(null).passed, false);
    assert.equal(report(null).passed, false);
    writeFileSync(join(root, 'semgrep', 'result.sarif'), '{broken');
    assert.equal(report(statuses).passed, false);
    writeFileSync(join(root, 'semgrep', 'result.sarif'), JSON.stringify(clean()));
    writeFileSync(join(root, 'semgrep', 'unexpected.sarif'), JSON.stringify(clean()));
    assert.equal(report(statuses).passed, false);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test('a failed run still writes both release reports', () => {
  const root = mkdtempSync(join(tmpdir(), 'xlide-security-'));
  try {
    const output = join(root, 'report');
    const result = spawnSync(process.execPath, [fileURLToPath(new URL('./report.mjs', import.meta.url)), root, output], {
      env: { ...process.env, SCANNER_STATUS: '{malformed', GITHUB_STEP_SUMMARY: '' },
      encoding: 'utf8',
    });
    assert.equal(result.status, 1, result.stderr);
    assert.equal(JSON.parse(readFileSync(join(output, 'security-report.json'), 'utf8')).passed, false);
    assert.match(readFileSync(join(output, 'security-report.md'), 'utf8'), /FAIL \/ INCOMPLETE/);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});
