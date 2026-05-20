export default function DataTable({ columns, rows, emptyMessage = 'No items to display' }) {
  if (!rows || rows.length === 0) {
    return (
      <div className="bg-white border border-border-light px-3 py-8 text-center text-text-secondary text-xs">
        {emptyMessage}
      </div>
    );
  }

  return (
    <div className="bg-white border border-border-light overflow-auto">
      <table className="w-full text-xs">
        <thead className="bg-[#F2F2F2] text-[#333] font-semibold">
          <tr>
            {columns.map((c) => (
              <th key={c.key} className="text-left px-3 py-2 border-b border-border-light" style={{ width: c.width }}>
                {c.header}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, idx) => (
            <tr key={row.id ?? idx} className="border-b border-border-light hover:bg-[#F9FAFB]">
              {columns.map((c) => (
                <td key={c.key} className="px-3 py-2">
                  {c.render ? c.render(row) : row[c.key]}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
