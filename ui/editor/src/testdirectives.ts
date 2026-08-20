/*
 * Completion for the test runner's comment directives, the sibling product's ergonomics
 * ported: typing `@` in a comment offers the three directives as snippets, and typing after
 * a directive offers its metadata keys. Registered BESIDE the host completion provider -
 * Monaco merges providers - so a comment line gains these without costing the code lines
 * anything.
 */

import * as monaco from "monaco-editor/editor/editor.api.js";

interface DirectiveTemplate {
  label: string;
  insert: string;
  detail: string;
}

const DIRECTIVES: DirectiveTemplate[] = [
  {
    label: "@xlide-test",
    insert: "@xlide-test",
    detail: "Marks the zero-argument public Sub below as a test",
  },
  {
    label: "@xlide-test-skip",
    insert: "@xlide-test-skip reason=\"$1\"",
    detail: "A test that is not run, with the reason shown in its row",
  },
  {
    label: "@xlide-test-xfail",
    insert: "@xlide-test-xfail reason=\"$1\"",
    detail: "A test expected to fail: red reads as expected, a pass reads as news",
  },
];

const METADATA: DirectiveTemplate[] = [
  { label: "tags", insert: "tags=\"$1\"", detail: "Comma-separated labels, shown as chips and usable to select runs" },
  { label: "owner", insert: "owner=$1", detail: "Who answers for this test" },
  { label: "requirement", insert: "requirement=$1", detail: "The requirement this test pins" },
  { label: "timeout", insert: "timeout=${1:30s}", detail: "Advisory only in this product: in-process VBA cannot be preempted" },
  { label: "expected-error", insert: "expected-error=${1:13}", detail: "Passes only when this VBA error number is raised" },
  { label: "expected-error (any)", insert: "expected-error", detail: "Passes when any VBA error at all is raised" },
  { label: "reason", insert: "reason=\"$1\"", detail: "Why this test is skipped or expected to fail" },
];

export function registerTestDirectiveCompletion(languageId: string): void {
  monaco.languages.registerCompletionItemProvider(languageId, {
    triggerCharacters: ["@", "-"],
    provideCompletionItems(model, position) {
      const line = model.getLineContent(position.lineNumber);
      const before = line.slice(0, position.column - 1);
      const comment = /^\s*'(.*)$/.exec(before);
      if (!comment) {
        return { suggestions: [] };
      }

      const body = comment[1] ?? "";
      const token = /@?[A-Za-z0-9-]*$/.exec(body)?.[0] ?? "";
      const start = position.column - token.length;
      const range = new monaco.Range(position.lineNumber, start, position.lineNumber, position.column);

      // At the comment's first word: the directives themselves.
      if (body.slice(0, body.length - token.length).trim().length === 0) {
        if (token.length > 0 && !"@xlide-test".startsWith(token.toLowerCase().slice(0, 11)) && !token.startsWith("@")) {
          return { suggestions: [] };
        }

        return {
          suggestions: DIRECTIVES.map((directive) => ({
            label: directive.label,
            kind: monaco.languages.CompletionItemKind.Keyword,
            detail: directive.detail,
            insertText: directive.insert,
            insertTextRules: monaco.languages.CompletionItemInsertTextRule.InsertAsSnippet,
            range,
          })),
        };
      }

      // After a directive: its metadata keys, reason only where a reason means something.
      const directive = /^\s*@xlide-test(-(skip|xfail))?\b/i.exec(body.trimStart());
      if (!directive) {
        return { suggestions: [] };
      }

      const withReason = directive[1] !== undefined;
      return {
        suggestions: METADATA
          .filter((key) => withReason || key.label !== "reason")
          .map((key) => ({
            label: key.label,
            kind: monaco.languages.CompletionItemKind.Property,
            detail: key.detail,
            insertText: key.insert,
            insertTextRules: monaco.languages.CompletionItemInsertTextRule.InsertAsSnippet,
            range,
          })),
      };
    },
  });
}
