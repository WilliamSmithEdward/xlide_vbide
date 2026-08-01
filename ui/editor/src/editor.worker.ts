// Worker entry point. Bundled separately so MonacoEnvironment.getWorker can hand Monaco a
// self-contained classic worker script that needs no module loader at runtime.
import "monaco-editor/editor/editor.worker.js";
