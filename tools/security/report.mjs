import { readdirSync, readFileSync, mkdirSync, writeFileSync, appendFileSync } from 'node:fs';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { inspectMalware } from './malware-policy.mjs';

export const expectedScans = ['codeql-csharp', 'codeql-javascript-typescript', 'codeql-actions', 'semgrep'];
export const expectedMalwareScans = ['clamav', 'yara-x'];
const malwarePolicy = JSON.parse(readFileSync(new URL('../../.github/security/malware/policy.json', import.meta.url), 'utf8'));

export function inspectSarif(sarif) {
  if (sarif.version !== '2.1.0' || !Array.isArray(sarif.runs) || !sarif.runs.length) {
    throw new Error('Missing or unsupported SARIF runs');
  }
  let findings = 0;
  const problems = [];
  const tools = [];
  for (const run of sarif.runs) {
    if (!run.tool?.driver?.name || !Array.isArray(run.results)) {
      throw new Error('Missing tool identity or results');
    }
    tools.push({ name: run.tool.driver.name, version: run.tool.driver.semanticVersion ?? run.tool.driver.version ?? 'not supplied' });
    // No severity exemptions, baselines, or silently accepted suppressions.
    findings += run.results.length;
    for (const invocation of run.invocations ?? []) {
      if (invocation.executionSuccessful === false) problems.push('Scanner execution failed');
      for (const notification of [...(invocation.toolExecutionNotifications ?? []), ...(invocation.toolConfigurationNotifications ?? [])]) {
        if (notification.level !== 'note' && notification.level !== 'none') {
          problems.push('Scanner reported an error or warning; inspect the workflow SARIF artifact');
        }
      }
    }
  }
  return { findings, problems, tools };
}

export function createReport(input, statuses, policy = malwarePolicy, context = process.env) {
  const scans = expectedScans.map(name => {
    try {
      const files = readdirSync(join(input, name)).filter(file => file.endsWith('.sarif'));
      if (files.length !== 1) throw new Error(`Expected one SARIF file; found ${files.length}`);
      return { name, ...inspectSarif(JSON.parse(readFileSync(join(input, name, files[0]), 'utf8'))) };
    } catch (error) {
      return { name, findings: null, tools: [], problems: [error.message] };
    }
  });
  for (const name of expectedMalwareScans) {
    try {
      const result = JSON.parse(readFileSync(join(input, name, 'result.json'), 'utf8'));
      if (context.GITHUB_SHA && result.commit !== context.GITHUB_SHA) throw new Error('Malware scan describes a different commit');
      scans.push({ name, ...inspectMalware(result, name, policy) });
    } catch (error) {
      scans.push({ name, findings: null, tools: [], problems: [error.message] });
    }
  }
  const problems = [];
  const malware = scans.filter(scan => expectedMalwareScans.includes(scan.name));
  if (malware.every(scan => scan.inventorySha256) && malware[0].inventorySha256 !== malware[1].inventorySha256) {
    problems.push('ClamAV and YARA-X scanned different inventories');
  }
  if (context.GITHUB_EVENT_NAME === 'release' && malware.some(scan =>
    !scan.assets?.some(asset => asset.path === 'release/xlide-setup.exe'))) {
    problems.push('Published release scan did not include the installer');
  }
  for (const name of ['codeql', 'semgrep', 'malware']) {
    if (statuses?.[name]?.result !== 'success') problems.push(`${name} job: ${statuses?.[name]?.result ?? 'missing'}`);
  }
  return { passed: problems.length === 0 && scans.every(scan => scan.findings === 0 && scan.problems.length === 0), problems, scans };
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const [input, output] = process.argv.slice(2);
  if (!input || !output) throw new Error('Usage: report.mjs INPUT OUTPUT');
  let statuses = {};
  try { statuses = JSON.parse(process.env.SCANNER_STATUS ?? '{}'); } catch { /* Missing status fails closed. */ }
  const report = {
    schemaVersion: 1,
    repository: process.env.GITHUB_REPOSITORY,
    commit: process.env.GITHUB_SHA,
    ref: process.env.GITHUB_REF,
    event: process.env.GITHUB_EVENT_NAME,
    run: `${process.env.GITHUB_SERVER_URL}/${process.env.GITHUB_REPOSITORY}/actions/runs/${process.env.GITHUB_RUN_ID}`,
    generatedAt: new Date().toISOString(),
    ...createReport(input, statuses),
  };
  const markdown = [
    '# Security report', '',
    `Status: **${report.passed ? 'PASS' : 'FAIL / INCOMPLETE'}**`, '',
    `Commit: \`${report.commit}\``, '',
    `Workflow: ${report.run}`, '',
    `Generated: ${report.generatedAt}`, '',
    '| Scanner | Unexpected findings | Accepted findings | Diagnostics |', '| --- | ---: | ---: | --- |',
    ...report.scans.map(scan => `| ${scan.name} | ${scan.findings ?? 'unavailable'} | ${scan.accepted?.length ?? 0} | ${scan.problems.length ? 'FAIL / incomplete' : `Complete${scan.acceptedDiagnostics ? `; ${scan.acceptedDiagnostics} reviewed compiler warnings` : ''}`} |`), '',
    ...report.problems.map(problem => `- ${problem}`), '',
    'Unexpected findings and incomplete scans fail the gate. Exact, expiring malware exceptions and the pinned YARA compiler-warning baseline are documented in docs/security-malware.md. Raw scan evidence is retained for 30 days.', '',
    'Scope: tracked repository files. CodeQL uses security-extended; Semgrep uses p/security-audit and p/secrets. ClamAV uses freshly updated official Talos signatures; YARA-X uses the pinned YARA Forge full collection.', '',
    ...report.scans.filter(scan => scan.name === 'yara-x').flatMap(scan => (scan.assets ?? []).map(asset => `Release asset: \`${asset.path}\`; SHA-256 \`${asset.sha256}\`.`)), '',
    'Release assets are scanned only when listed above. YARA-X scans raw bytes; neither scanner guarantees unpacking of custom installer payloads. Source scans do not include untracked dependencies or the separately bundled analyzer. This is not a guarantee that a release is malware-free.', '',
  ].join('\n');
  mkdirSync(output, { recursive: true });
  writeFileSync(join(output, 'security-report.json'), JSON.stringify(report, null, 2) + '\n');
  writeFileSync(join(output, 'security-report.md'), markdown);
  if (process.env.GITHUB_STEP_SUMMARY) appendFileSync(process.env.GITHUB_STEP_SUMMARY, markdown);
  if (!report.passed) process.exitCode = 1;
}
