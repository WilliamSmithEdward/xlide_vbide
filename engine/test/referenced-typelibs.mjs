/*
 * Issue #56: the VBE's referenced Scripting and RegExp type libraries must reach the same model
 * used by completion and diagnostics. The payload is deliberately the small, exact shape the
 * host's COM catalog sends; the engine test does not require Office or a registered DLL.
 */
import assert from 'node:assert/strict';
import { startEngine } from './harness.mjs';

const SCRIPTING = '{420B2830-E718-11CF-893D-00A0C9054228}';
const REGEXP = '{3F4DACA7-160D-11D2-A8E9-00104B365C9F}';
const libraries = [
  {
    name: 'Scripting', guid: SCRIPTING, types: [
      { name: 'FileSystemObject', kind: 'class', members: [
        { name: 'FileExists', kind: 'Function', signature: 'Function FileExists(FileSpec As String) As Boolean' },
        { name: 'GetFile', kind: 'Function', signature: 'Function GetFile(FilePath As String) As File' },
      ] },
      { name: 'File', kind: 'class', members: [
        { name: 'OpenAsTextStream', kind: 'Function', signature: 'Function OpenAsTextStream([IOMode As IOMode]) As TextStream' },
      ] },
      { name: 'TextStream', kind: 'class', members: [
        { name: 'WriteLine', kind: 'Sub', signature: 'Sub WriteLine([Text As String])' },
      ] },
      { name: 'IOMode', kind: 'enum', members: [
        { name: 'ForAppending', kind: 'Const', signature: 'Const ForAppending = 8' },
      ] },
    ],
  },
  {
    name: 'VBScript_RegExp_55', guid: REGEXP, types: [
      { name: 'RegExp', kind: 'class', members: [
        { name: 'Pattern', kind: 'Property', signature: 'Property Get Pattern() As String' },
        { name: 'Test', kind: 'Function', signature: 'Function Test(SourceString As String) As Boolean' },
        { name: 'Execute', kind: 'Function', signature: 'Function Execute(SourceString As String) As MatchCollection' },
      ] },
      { name: 'MatchCollection', kind: 'class', members: [] },
    ],
  },
];

const source = [
  'Option Explicit',
  'Public Sub TestIntellisense()',
  '    Dim x As RegExp',
  '    Set x = New RegExp',
  '    x.Pattern = "A"',
  '    Dim y As Scripting.FileSystemObject',
  '    Set y = New Scripting.FileSystemObject',
  '    Dim stream As Scripting.TextStream',
  '    If y.FileExists("x") Then',
  '        Set stream = y.GetFile("x").OpenAsTextStream(ForAppending)',
  '    End If',
  'End Sub',
].join('\r\n');

const { call, stop } = await startEngine('referenced-typelibs');
try {
  await call('initialize', {});
  const open = (projectId, refs, supplied = libraries) => call('project/open', {
    projectId, generation: 1, modules: [{ moduleName: 'Module1', source, type: 'standard' }],
    referenceGuids: refs, referenceLibraries: supplied,
  });
  const complete = async (projectId, text, marker) => {
    const offset = text.indexOf(marker) + marker.length;
    assert.ok(offset >= marker.length, `missing marker ${marker}`);
    const answer = await call('textDocument/completion', {
      projectId, moduleName: 'Module1', source: text, offset, moduleType: 'standard',
    });
    return answer.items.map((item) => item.label);
  };
  const diagnostics = async (projectId) => {
    const answer = await call('textDocument/diagnostics', {
      projectId, generation: 1, documentKey: `${projectId}/Module1`,
      moduleName: 'Module1', source, moduleType: 'standard',
    });
    return answer.diagnostics ?? [];
  };

  await open('with-libraries', [SCRIPTING, REGEXP]);
  const regexp = await complete('with-libraries', source.replace('x.Pattern = "A"', 'x.'), 'x.');
  assert.ok(regexp.includes('Pattern') && regexp.includes('Test') && regexp.includes('Execute'));
  assert.ok(!regexp.includes('FileExists') && !regexp.includes('Workbook'));
  const scripting = await complete('with-libraries', source.replace('y.FileExists("x")', 'y.'), 'y.');
  assert.ok(scripting.includes('FileExists') && scripting.includes('GetFile'));
  assert.ok(!scripting.includes('Pattern') && !scripting.includes('Workbook'));
  const withFindings = await diagnostics('with-libraries');
  assert.ok(!withFindings.some((item) => item.code === 'undeclared-variable' && /ForAppending/.test(item.message)));

  await open('without-libraries', []);
  const missing = await diagnostics('without-libraries');
  assert.ok(missing.some((item) => item.code === 'undeclared-variable' && /ForAppending/.test(item.message)));
  const bare = await complete('without-libraries', source.replace('x.Pattern = "A"', 'x.'), 'x.');
  assert.ok(!bare.includes('Pattern') && !bare.includes('Test'));

  // v11.0.2 ships a pinned Scripting model. A referenced known library remains usable even
  // when the host's type-library snapshot is unavailable; the GUID still gates its scope.
  await open('missing-snapshot', [SCRIPTING], [libraries[1]]);
  const fromBuiltin = await diagnostics('missing-snapshot');
  assert.ok(!fromBuiltin.some((item) => item.code === 'undeclared-variable' && /ForAppending/.test(item.message)));
  const builtinMembers = await complete('missing-snapshot', source.replace('y.FileExists("x")', 'y.'), 'y.');
  assert.ok(builtinMembers.includes('FileExists') && builtinMembers.includes('GetFile'));

  console.log('Referenced type-library completion and constants: passed');
} finally {
  stop();
}
