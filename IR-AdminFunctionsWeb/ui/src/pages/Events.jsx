import { useEffect, useState } from 'react';
import DataTable from '../components/common/DataTable.jsx';
import StatusBadge from '../components/common/StatusBadge.jsx';
import FilterDropdown from '../components/common/FilterDropdown.jsx';
import { api } from '../api/client.js';

const columns = [
  { key: 'time', header: 'Time', width: '20%', render: (r) => new Date(r.time).toLocaleString() },
  { key: 'severity', header: 'Severity', width: '12%', render: (r) => <StatusBadge value={r.severity} /> },
  { key: 'source', header: 'Source', width: '28%' },
  { key: 'message', header: 'Message', width: '40%' }
];

export default function Events() {
  const [rows, setRows] = useState([]);
  const [severity, setSeverity] = useState('');

  useEffect(() => {
    api.events(severity || undefined)
      .then((r) => setRows(r?.data?.items ?? []))
      .catch(() => setRows([]));
  }, [severity]);

  return (
    <div className="p-4 space-y-3">
      <div className="bg-white border border-border-light px-4 py-3 flex items-center gap-4">
        <FilterDropdown
          label="Severity"
          value={severity}
          onChange={setSeverity}
          options={[
            { value: '', label: 'Any' },
            { value: 'Information', label: 'Information' },
            { value: 'Warning', label: 'Warning' },
            { value: 'Error', label: 'Error' }
          ]}
        />
        <span className="text-xs text-text-secondary ml-auto">{rows.length} events</span>
      </div>
      <DataTable columns={columns} rows={rows} emptyMessage="No events." />
    </div>
  );
}
