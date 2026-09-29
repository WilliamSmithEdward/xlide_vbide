import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, writeFileSync, readFileSync, rmSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createReport, expectedScans, inspectSarif } from './report.mjs';

const clean = () => ({ version: '2.1.0', runs: [{ tool: { driver: { name: 'test-scanner', version: '1' } }, results: [] }] });
const statuses = { codeql: { result: 'success' }, semgrep: { result: 'success' } };

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
  try {
    assert.equal(createReport(root, statuses).passed, false);
    for (const name of expectedScans) {
      mkdirSync(join(root, name));
      writeFileSync(join(root, name, 'result.sarif'), JSON.stringify(clean()));
    }
    assert.equal(createReport(root, statuses).passed, true);
    for (const result of ['failure', 'cancelled', 'skipped', undefined]) {
      assert.equal(createReport(root, { ...statuses, semgrep: { result } }).passed, false);
    }
    assert.equal(createReport(root, null).passed, false);
    writeFileSync(join(root, 'semgrep', 'result.sarif'), '{broken');
    assert.equal(createReport(root, statuses).passed, false);
    writeFileSync(join(root, 'semgrep', 'result.sarif'), JSON.stringify(clean()));
    writeFileSync(join(root, 'semgrep', 'unexpected.sarif'), JSON.stringify(clean()));
    assert.equal(createReport(root, statuses).passed, false);
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
