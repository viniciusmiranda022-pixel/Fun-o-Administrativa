const colorByStatus: Record<string, string> = {
  Completed: 'bg-green-100 text-green-700',
  Running: 'bg-blue-100 text-blue-700',
  Queued: 'bg-slate-100 text-slate-700',
  Failed: 'bg-red-100 text-red-700',
  Error: 'bg-red-100 text-red-700',
  Warning: 'bg-amber-100 text-amber-700',
  Information: 'bg-slate-100 text-slate-700'
};

interface StatusBadgeProps {
  value: string;
}

export default function StatusBadge({ value }: StatusBadgeProps) {
  const cls = colorByStatus[value] ?? 'bg-slate-100 text-slate-700';
  return (
    <span className={`inline-block px-2 py-0.5 rounded text-[11px] font-medium ${cls}`}>
      {value}
    </span>
  );
}
