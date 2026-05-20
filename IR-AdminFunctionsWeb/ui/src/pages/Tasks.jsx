import { useEffect, useState } from 'react';
import DataTable from '../components/common/DataTable.jsx';
import StatusBadge from '../components/common/StatusBadge.jsx';
import { api } from '../api/client.js';

const columns = [
  { key: 'time', header: 'Time', width: '18%', render: (r) => new Date(r.time).toLocaleString() },
  { key: 'name', header: 'Name', width: '24%' },
  { key: 'status', header: 'Status', width: '12%', render: (r) => <StatusBadge value={r.status} /> },
  { key: 'type', header: 'Type', width: '14%' },
  { key: 'operation', header: 'Operation', width: '32%' }
];

export default function Tasks() {
  const [rows, setRows] = useState([]);

  useEffect(() => {
    const tick = () => api.tasks().then((r) => setRows(r?.data?.items ?? [])).catch(() => {});
    tick();
    const id = setInterval(tick, 5000);
    return () => clearInterval(id);
  }, []);

  return (
    <div className="p-4 space-y-3">
      <div className="bg-white border border-border-light px-4 py-3 flex items-center">
        <h2 className="text-sm font-semibold">Tasks</h2>
        <span className="text-xs text-text-secondary ml-auto">{rows.length} items</span>
      </div>
      <DataTable columns={columns} rows={rows} emptyMessage="No tasks." />
    </div>
  );
}
