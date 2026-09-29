import { readdirSync, readFileSync, mkdirSync, writeFileSync, appendFileSync } from 'node:fs';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';

export const expectedScans = ['codeql-csharp', 'codeql-javascript-typescript', 'codeql-actions', 'semgrep'];

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

export function createReport(input, statuses) {
  const scans = expectedScans.map(name => {
    try {
      const files = readdirSync(join(input, name)).filter(file => file.endsWith('.sarif'));
      if (files.length !== 1) throw new Error(`Expected one SARIF file; found ${files.length}`);
      return { name, ...inspectSarif(JSON.parse(readFileSync(join(input, name, files[0]), 'utf8'))) };
    } catch (error) {
      return { name, findings: null, tools: [], problems: [error.message] };
    }
  });
  const problems = [];
  for (const name of ['codeql', 'semgrep']) {
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
    '| Scanner | Findings | Diagnostics |', '| --- | ---: | --- |',
    ...report.scans.map(scan => `| ${scan.name} | ${scan.findings ?? 'unavailable'} | ${scan.problems.length ? 'FAIL / incomplete' : 'Complete'} |`), '',
    ...report.problems.map(problem => `- ${problem}`), '',
    'All findings and incomplete scans fail the gate. Raw SARIF and diagnostics are retained in workflow artifacts for 30 days.', '',
    'Scope: this repository at the commit above. CodeQL uses security-extended for C#, JavaScript/TypeScript, and Actions; Semgrep uses p/security-audit and p/secrets.', '',
    'This is source analysis, not an audit of the installer binary, runtime dependencies, or the separately bundled xlide_vscode analyzer. A passing scan does not guarantee absence of vulnerabilities.', '',
  ].join('\n');
  mkdirSync(output, { recursive: true });
  writeFileSync(join(output, 'security-report.json'), JSON.stringify(report, null, 2) + '\n');
  writeFileSync(join(output, 'security-report.md'), markdown);
  if (process.env.GITHUB_STEP_SUMMARY) appendFileSync(process.env.GITHUB_STEP_SUMMARY, markdown);
  if (!report.passed) process.exitCode = 1;
}
