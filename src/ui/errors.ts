import { message } from '@tauri-apps/plugin-dialog';

export function showError(e: unknown): void {
  void message(e instanceof Error ? e.message : String(e), { kind: 'error' });
}
