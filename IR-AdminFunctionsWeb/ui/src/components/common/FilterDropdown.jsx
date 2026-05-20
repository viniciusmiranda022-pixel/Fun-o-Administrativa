export default function FilterDropdown({ label, value, options, onChange }) {
  return (
    <label className="flex items-center gap-2 text-xs">
      <span className="text-[#333]">{label}:</span>
      <select
        value={value ?? ''}
        onChange={(e) => onChange?.(e.target.value)}
        className="border border-border-light px-2 py-1 bg-white"
      >
        {options.map((o) => (
          <option key={o.value ?? o} value={o.value ?? o}>
            {o.label ?? o}
          </option>
        ))}
      </select>
    </label>
  );
}
