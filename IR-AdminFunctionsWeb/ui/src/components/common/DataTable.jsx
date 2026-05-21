import { useState } from 'react';
import { ChevronUp, ChevronDown } from 'lucide-react';

export default function DataTable({
  columns,
  rows,
  emptyMessage = 'No items to display',
  onRowClick,
  selectedRows = [],
  showCheckboxes = true,
  totalCount,
  pageSize = 400,
}) {
  const [sortKey, setSortKey] = useState(null);
  const [sortDir, setSortDir] = useState('asc');
  const [checkedRows, setCheckedRows] = useState(new Set());
  const [allChecked, setAllChecked] = useState(false);

  function handleSort(col) {
    if (!col.sortable) return;
    if (sortKey === col.key) {
      setSortDir((d) => (d === 'asc' ? 'desc' : 'asc'));
    } else {
      setSortKey(col.key);
      setSortDir('asc');
    }
  }

  function toggleAll() {
    if (allChecked) {
      setCheckedRows(new Set());
      setAllChecked(false);
    } else {
      setCheckedRows(new Set(rows.map((r, i) => r.id ?? i)));
      setAllChecked(true);
    }
  }

  function toggleRow(id) {
    setCheckedRows((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  let sorted = [...(rows ?? [])];
  if (sortKey) {
    sorted.sort((a, b) => {
      const av = a[sortKey] ?? '';
      const bv = b[sortKey] ?? '';
      const cmp = String(av).localeCompare(String(bv));
      return sortDir === 'asc' ? cmp : -cmp;
    });
  }

  const total = totalCount ?? rows?.length ?? 0;
  const shown = sorted.length;

  if (!rows || rows.length === 0) {
    return (
      <div className="bg-white border border-[#DEE2E6] px-3 py-8 text-center text-[#888] text-xs">
        {emptyMessage}
      </div>
    );
  }

  return (
    <div className="flex flex-col">
      <div className="bg-white border border-[#DEE2E6] overflow-auto">
        <table className="w-full text-xs">
          <thead className="bg-[#F2F2F2] text-[#333] font-semibold">
            <tr>
              {showCheckboxes && (
                <th className="w-8 px-3 py-2 border-b border-[#DEE2E6] text-center">
                  <input
                    type="checkbox"
                    checked={allChecked}
                    onChange={toggleAll}
                    className="cursor-pointer"
                  />
                </th>
              )}
              {columns.map((c) => (
                <th
                  key={c.key}
                  className={`text-left px-3 py-2 border-b border-[#DEE2E6] whitespace-nowrap ${c.sortable ? 'cursor-pointer hover:bg-[#E8E8E8] select-none' : ''}`}
                  style={{ width: c.width }}
                  onClick={() => handleSort(c)}
                >
                  <span className="inline-flex items-center gap-1">
                    {c.header}
                    {c.sortable && sortKey === c.key && (
                      sortDir === 'asc'
                        ? <ChevronUp size={10} className="text-[#0078A8]" />
                        : <ChevronDown size={10} className="text-[#0078A8]" />
                    )}
                    {c.sortable && sortKey !== c.key && (
                      <ChevronDown size={10} className="text-[#ccc]" />
                    )}
                  </span>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {sorted.map((row, idx) => {
              const rowId = row.id ?? idx;
              const isChecked = checkedRows.has(rowId);
              const isSelected = selectedRows.includes(rowId);
              return (
                <tr
                  key={rowId}
                  onClick={() => onRowClick && onRowClick(row)}
                  className={`border-b border-[#DEE2E6] ${onRowClick ? 'cursor-pointer' : ''} ${
                    isSelected
                      ? 'bg-[#E6F4FA]'
                      : isChecked
                      ? 'bg-[#EDF7FF]'
                      : 'bg-white hover:bg-[#F9FAFB]'
                  }`}
                >
                  {showCheckboxes && (
                    <td className="px-3 py-2 text-center" onClick={(e) => e.stopPropagation()}>
                      <input
                        type="checkbox"
                        checked={isChecked}
                        onChange={() => toggleRow(rowId)}
                        className="cursor-pointer"
                      />
                    </td>
                  )}
                  {columns.map((c) => (
                    <td key={c.key} className="px-3 py-2 whitespace-nowrap">
                      {c.render ? c.render(row) : row[c.key] ?? '—'}
                    </td>
                  ))}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* Pagination footer */}
      <div className="bg-white border border-t-0 border-[#DEE2E6] px-4 py-2 flex items-center gap-3 text-xs text-[#555]">
        <span>
          1-{shown} of {total} Results
        </span>
        <span className="text-[#ccc]">|</span>
        <span className="flex items-center gap-1">
          Show:
          <select className="border border-[#DEE2E6] px-1 py-0.5 text-xs bg-white ml-1">
            <option>{pageSize}</option>
            <option>100</option>
            <option>200</option>
          </select>
        </span>
        <span className="text-[#ccc]">|</span>
        <button className="px-1 hover:text-[#0078A8]">&lt;</button>
        <span>1</span>
        <button className="px-1 hover:text-[#0078A8]">&gt;</button>
      </div>
    </div>
  );
}
