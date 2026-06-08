interface FilterOption {
  value?: string;
  label?: string;
}

interface FilterDropdownProps {
  label: string;
  value?: string;
  options: (FilterOption | string)[];
  onChange?: (value: string) => void;
}

export default function FilterDropdown({ label, value, options, onChange }: FilterDropdownProps) {
  return (
    <label className="flex items-center gap-2 text-xs">
      <span className="text-[#333]">{label}:</span>
      <select
        value={value ?? ''}
        onChange={(e) => onChange?.(e.target.value)}
        className="border border-border-light px-2 py-1 bg-white"
      >
        {options.map((o) => {
          const val = typeof o === 'string' ? o : (o.value ?? '');
          const lbl = typeof o === 'string' ? o : (o.label ?? o.value ?? '');
          return (
            <option key={val} value={val}>
              {lbl}
            </option>
          );
        })}
      </select>
    </label>
  );
}
