import { useEffect, useState } from 'react';
import DataTable from '../components/common/DataTable.jsx';
import { api } from '../api/client.js';

const columns = [
  { key: 'displayName', header: 'Name', width: '30%', render: (r) => r.displayName ?? r.DisplayName },
  { key: 'isBuiltIn', header: 'Type', width: '15%', render: (r) => (r.isBuiltIn || r.IsBuiltIn ? 'Built-in' : 'Custom') },
  { key: 'isEnabled', header: 'Enabled', width: '15%', render: (r) => String(r.isEnabled ?? r.IsEnabled ?? '') },
  { key: 'description', header: 'Description', width: '40%', render: (r) => r.description ?? r.Description }
];

export default function UnpackedObjects() {
  const [backups, setBackups] = useState([]);
  const [selected, setSelected] = useState('');
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  useEffect(() => {
    api.backups().then((r) => {
      const items = r?.data ?? [];
      setBackups(items);
      if (items.length > 0) setSelected(items[0].id);
    });
  }, []);

  useEffect(() => {
    if (!selected) return;
    setLoading(true);
    setError(null);
    api.unpacked(selected)
      .then((r) => {
        const defs = r?.data?.roleDefinitions ?? [];
        setRows(Array.isArray(defs) ? defs : []);
      })
      .catch((e) => setError(e.message))
      .finally(() => setLoading(false));
  }, [selected]);

  return (
    <div className="p-4 space-y-3">
      <div className="bg-white border border-border-light px-4 py-3 flex items-center gap-3">
        <span className="text-xs">Backup:</span>
        <select
          value={selected}
          onChange={(e) => setSelected(e.target.value)}
          className="border border-border-light px-2 py-1 text-xs bg-white"
        >
          {backups.map((b) => (
            <option key={b.id} value={b.id}>
              {b.id} {b.collectedAt ? `(${b.collectedAt})` : ''}
            </option>
          ))}
        </select>
        <span className="text-xs text-text-secondary ml-auto">{rows.length} role definitions</span>
      </div>
      {error && (
        <div className="bg-red-50 border border-red-200 text-red-700 px-4 py-2 text-xs">{error}</div>
      )}
      {loading ? (
        <div className="bg-white border border-border-light px-4 py-8 text-center text-xs text-text-secondary">
          Carregando objetos…
        </div>
      ) : (
        <DataTable columns={columns} rows={rows} emptyMessage="Nenhum objeto." />
      )}
    </div>
  );
}
