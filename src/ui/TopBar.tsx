interface Props {
  name: string;
  rows: number;
  dirty: boolean;
  saving: boolean;
  onSave(): void;
}

export function TopBar({ name, rows, dirty, saving, onSave }: Props) {
  return (
    <header className="top-bar">
      <span className="file-name">{name}</span>
      <span className="row-count">· {rows.toLocaleString()}행</span>
      {dirty && <span className="dirty">● 수정됨</span>}
      <span className="spacer" />
      <button className="save-button" onClick={onSave} disabled={saving}>
        저장 ⌘S
      </button>
    </header>
  );
}
