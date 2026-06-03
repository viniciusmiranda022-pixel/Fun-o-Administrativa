import { useState, useEffect } from 'react';
import { RotateCcw, Download, Columns, Search, FileText } from 'lucide-react';
import DataTable from '../components/common/DataTable.jsx';
import { api } from '../api/client.js';

const columns = [
  {
    key: 'name', header: 'Name', sortable: true, width: '22%',
    render: (r) => (
      <div className="flex items-center gap-1.5">
        <FileText size={12} className="text-[#0096D6] flex-shrink-0" />
        <span className="text-[#0078A8] truncate">{r.name}</span>
      </div>
    )
  },
  { key: 'type', header: 'Object Type', sortable: true, width: '14%' },
  { key: 'roleId', header: 'Role ID', width: '16%', render: (r) => <span className="font-mono text-[10px] text-[#666] truncate block max-w-[140px]">{r.roleId}</span> },
  { key: 'description', header: 'Description', width: '20%', render: (r) => <span className="truncate block max-w-[180px]">{r.description || '—'}</span> },
  { key: 'permissions', header: 'Permissions', width: '8%' },
  {
    key: 'status', header: 'Status', width: '8%',
    render: (r) => <span className="text-[#28A745] font-semibold text-xs">{r.status}</span>
  },
];

function mapDefinition(d) {
  const perms = (d.rolePermissions || []).reduce((sum, rp) => sum + (rp.allowedResourceActions?.length || 0), 0);
  return {
    id: d.id || d.Id,
    name: d.displayName || d.DisplayName || '—',
    type: d.isBuiltIn || d.IsBuiltIn ? 'Built-in Role' : 'Custom Role',
    roleId: (d.id || d.Id || '').slice(0, 18) + '…',
    description: d.description || d.Description || '',
    permissions: perms,
    status: (d.isEnabled || d.IsEnabled) !== false ? 'Active' : 'Disabled',
  };
}

function mapAssignment(a) {
  return {
    id: a.id || a.Id,
    name: `${(a.principalId || a.PrincipalId || '—').slice(0, 8)}… → ${(a.roleDefinitionId || a.RoleDefinitionId || '—').slice(0, 8)}…`,
    type: 'Role Assignment',
    roleId: (a.roleDefinitionId || a.RoleDefinitionId || '').slice(0, 18) + '…',
    description: a.directoryScopeId || a.DirectoryScopeId || '/',
    permissions: '—',
    status: 'Active',
  };
}

export default function UnpackedObjects() {
  const [view, setView] = useState('list');
  const [search, setSearch] = useState('');
  const [filters, setFilters] = useState({ type: '' });
  const [backups, setBackups] = useState([]);
  const [selectedBackupId, setSelectedBackupId] = useState(null);
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  useEffect(() => {
    api.backups()
      .then((r) => {
        const list = r?.data ?? [];
        setBackups(list);
        if (list.length > 0) setSelectedBackupId(list[0].id);
        else setLoading(false);
      })
      .catch(() => setLoading(false));
  }, []);

  useEffect(() => {
    if (!selectedBackupId) return;
    setLoading(true);
    setError(null);
    api.unpacked(selectedBackupId)
      .then((r) => {
        const data = r?.data;
        if (!data) { setRows([]); return; }
        const defs = Array.isArray(data.roleDefinitions) ? data.roleDefinitions : [];
        const assigns = Array.isArray(data.roleAssignments) ? data.roleAssignments : [];
        setRows([
          ...defs.map(mapDefinition),
          ...assigns.map(mapAssignment),
        ]);
      })
      .catch((err) => setError(err.message))
      .finally(() => setLoading(false));
  }, [selectedBackupId]);

  const filtered = rows.filter((r) => {
    if (filters.type && r.type !== filters.type) return false;
    if (search && !r.name.toLowerCase().includes(search.toLowerCase())) return false;
    return true;
  });

  const typeOptions = [...new Set(rows.map(r => r.type))];

  return (
    <div className="flex flex-col h-full">
      {/* Toggle + Filter bar */}
      <div className="bg-white border-b border-[#DEE2E6] px-4 py-2 flex items-center gap-3 flex-shrink-0">
        <div className="flex border border-[#DEE2E6] overflow-hidden text-xs">
          {['list', 'objects'].map((v) => (
            <button
              key={v}
              onClick={() => setView(v)}
              className={`px-3 py-1 capitalize ${view === v ? 'bg-[#0078A8] text-white' : 'bg-white text-[#555] hover:bg-[#F2F2F2]'}`}
            >
              {v === 'list' ? 'List View' : 'Objects'}
            </button>
          ))}
        </div>

        {backups.length > 1 && (
          <div className="flex items-center gap-1 text-xs">
            <span className="text-[#555]">Backup:</span>
            <select
              value={selectedBackupId || ''}
              onChange={(e) => setSelectedBackupId(e.target.value)}
              className="border border-[#DEE2E6] px-1.5 py-0.5 text-xs bg-white max-w-[200px]"
            >
              {backups.map((b) => {
                const d = new Date(b.createdAt || b.collectedAt || '');
                return (
                  <option key={b.id} value={b.id}>
                    {isNaN(d) ? b.id : d.toLocaleString()}
                  </option>
                );
              })}
            </select>
          </div>
        )}

        <div className="flex items-center gap-1 text-xs">
          <span className="text-[#555]">Type:</span>
          <select
            value={filters.type}
            onChange={(e) => setFilters((f) => ({ ...f, type: e.target.value }))}
            className="border border-[#DEE2E6] px-1.5 py-0.5 text-xs bg-white"
          >
            <option value="">Any</option>
            {typeOptions.map((t) => <option key={t} value={t}>{t}</option>)}
          </select>
        </div>
      </div>

      <div className="p-4 space-y-3 flex-1 overflow-auto">
        {/* Action toolbar */}
        <div className="bg-white border border-[#DEE2E6] px-3 py-2 flex items-center gap-2">
          <button className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] opacity-40 cursor-not-allowed">
            <RotateCcw size={13} />
            RESTORE
          </button>
          <button className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2]">
            <Download size={13} />
            EXPORT
          </button>
          <button className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2]">
            <Columns size={13} />
            EDIT COLUMNS
          </button>
          <div className="ml-auto flex items-center gap-2">
            <span className="text-xs text-[#666]">
              {loading ? 'Loading…' : `${filtered.length} object${filtered.length !== 1 ? 's' : ''}`}
            </span>
            <div className="flex items-center border border-[#DEE2E6] overflow-hidden">
              <input
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Search..."
                className="px-2 py-1 text-xs outline-none w-40"
              />
              <button className="px-2 py-1 bg-[#F2F2F2] border-l border-[#DEE2E6] text-[#0078A8]">
                <Search size={13} />
              </button>
            </div>
          </div>
        </div>

        {loading ? (
          <div className="text-xs text-[#888] p-4">Loading objects…</div>
        ) : error ? (
          <div className="bg-white border border-[#DEE2E6] p-8 text-center text-xs text-[#DC3545]">
            {error}
          </div>
        ) : backups.length === 0 ? (
          <div className="bg-white border border-[#DEE2E6] p-8 text-center text-xs text-[#888]">
            No backups found. Create a backup first, then unpack it to view objects here.
          </div>
        ) : filtered.length === 0 ? (
          <div className="bg-white border border-[#DEE2E6] p-8 text-center text-xs text-[#888]">
            No objects found for this backup. Unpack a backup to populate this view.
          </div>
        ) : (
          <DataTable columns={columns} rows={filtered} totalCount={filtered.length} />
        )}
      </div>
    </div>
  );
}
