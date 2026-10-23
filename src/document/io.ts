import { invoke } from '@tauri-apps/api/core';

export interface TextFile {
  text: string;
  encoding: 'utf-8' | 'euc-kr';
  bom: boolean;
}

export interface DocumentIo {
  readFile(path: string): Promise<Uint8Array>;
  writeFile(path: string, bytes: Uint8Array): Promise<void>;
  readText(path: string): Promise<TextFile>;
  writeText(path: string, file: TextFile): Promise<void>;
}

export const tauriIo: DocumentIo = {
  async readFile(path) {
    const data = await invoke<ArrayBuffer>('read_file', { path });
    return new Uint8Array(data);
  },
  async writeFile(path, bytes) {
    await invoke('write_file', bytes, { headers: { 'x-path': encodeURIComponent(path) } });
  },
  readText(path) {
    return invoke<TextFile>('read_text', { path });
  },
  async writeText(path, file) {
    await invoke('write_text', {
      path,
      text: file.text,
      encoding: file.encoding,
      bom: file.bom,
    });
  },
};
