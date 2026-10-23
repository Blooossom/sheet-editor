interface Props {
  names: string[];
  active: number;
  onChange(index: number): void;
}

export function SheetTabs({ names, active, onChange }: Props) {
  return (
    <nav className="sheet-tabs">
      {names.map((name, i) => (
        <button key={i} className={i === active ? 'tab active' : 'tab'} onClick={() => onChange(i)}>
          {name}
        </button>
      ))}
    </nav>
  );
}
