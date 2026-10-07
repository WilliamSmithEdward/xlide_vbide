/** UI preferences share the native per-user store. localStorage is only the current WebView
 * cache and a migration source for older builds; it is never the cross-session authority. */
type PreferenceKey = "paneLayout" | "sourceControl" | "editorGroups";
const legacyKeys: Record<PreferenceKey, string> = {
  paneLayout: "xlide.docks.v1",
  sourceControl: "xlide.scm.v1",
  editorGroups: "xlide.workspace.v1",
};
const values = new Map<PreferenceKey, string>();
const foundOnLoad = new Set<PreferenceKey>();
const listeners = new Map<PreferenceKey, (value: string) => void>();
let hydrated = false;
let save: ((key: PreferenceKey, value: string) => void) | null = null;

export function readUiPreference(key: PreferenceKey): string | null {
  if (values.has(key)) return values.get(key)!;
  try {
    const raw = localStorage.getItem(legacyKeys[key]);
    if (raw) { values.set(key, raw); foundOnLoad.add(key); }
    return raw;
  } catch { return null; }
}

export function writeUiPreference(key: PreferenceKey, value: string): void {
  values.set(key, value);
  try { localStorage.setItem(legacyKeys[key], value); }
  catch (error) { console.warn("[xlide] UI preference cache could not be written", error); }
  if (hydrated) save?.(key, value);
}

export function watchUiPreference(key: PreferenceKey, apply: (value: string) => void): void {
  listeners.set(key, apply);
}

/** Hydrate after the host's ready handshake, before enabling writes of initial defaults.
 * A page reload retains its cache; a new Office process restores the shared native values. */
export function restoreUiPreferences(
  stored: Record<string, string>,
  writer: (key: PreferenceKey, value: string) => void,
): void {
  hydrated = false;
  save = writer;
  for (const [key, apply] of listeners) {
    const raw = foundOnLoad.has(key) ? values.get(key) : stored[key] ?? values.get(key);
    if (raw) apply(raw);
  }
  hydrated = true;
  for (const [key, value] of values) save(key, value);
}
