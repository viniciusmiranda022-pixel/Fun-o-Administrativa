import { useState, useEffect } from 'react';
import { Package, Columns, Search } from 'lucide-react';
import DataTable from '../components/common/DataTable.jsx';
import BackupUnpackingModal from '../components/BackupUnpackingModal.jsx';
import { api } from '../api/client.js';

const columns = [
  { key: 'time', header: 'Time', sortable: true, width: '18%' },
  { key: 'tenant', header: 'Tenant', sortable: true, width: '10%' },
  { key: 'customRoles', header: 'Custom Roles', width: '10%' },
  { key: 'roleAssignments', header: 'Role Assignments', width: '12%' },
  { key: 'size', header: 'Backup Size', width: '10%' },
  {
    key: 'status', header: 'Status', width: '10%',
    render: (r) => (
      <span className={`inline-flex items-center gap-1 text-xs font-semibold ${r.status === 'OK' ? 'text-[#28A745]' : 'text-[#DC3545]'}`}>
        {r.status}
      </span>
    )
  },
];

function buildBars(backups, days = 35) {
  if (!backups || backups.length === 0) {
    return Array.from({ length: days }, () => ({ h: 2 }));
  }
  const now = new Date();
  const buckets = Array.from({ length: days }, (_, i) => {
    const d = new Date(now);
    d.setDate(d.getDate() - (days - 1 - i));
    return { count: 0, date: d.toDateString() };
  });
  backups.forEach((b) => {
    const bDate = new Date(b.createdAt || b.collectedAt || '').toDateString();
    const bucket = buckets.find((bk) => bk.date === bDate);
    if (bucket) bucket.count++;
  });
  const max = Math.max(...buckets.map((b) => b.count), 1);
  return buckets.map((b) => ({ h: Math.round((b.count / max) * 100) || 2 }));
}

function MiniBarChart({ bars }) {
  return (
    <div className="bg-white border border-[#DEE2E6] px-4 pt-3 pb-2">
      <div className="flex items-end gap-0.5" style={{ height: 60 }}>
        {bars.map((b, i) => (
          <div key={i} className="flex-1 flex flex-col justify-end">
            <div className="w-full rounded-sm" style={{ height: `${b.h}%`, backgroundColor: '#6CB33F', minHeight: 2 }} />
          </div>
        ))}
      </div>
    </div>
  );
}

function mapRow(b) {
  const d = new Date(b.createdAt || b.collectedAt || '');
  return {
    id: b.id,
    time: isNaN(d) ? b.id : d.toLocaleString(),
    tenant: b.tenantId ? b.tenantId.slice(0, 8) + '…' : '—',
    customRoles: b.roleDefinitionsCount ?? '—',
    roleAssignments: b.roleAssignmentsCount ?? '—',
    size: b.sizeDisplay ?? '—',
    status: b.manifestError ? 'Error' : 'OK',
  };
}

export default function Backups() {
  const [search, setSearch] = useState('');
  const [dateFilter, setDateFilter] = useState('30days');
  const [backups, setBackups] = useState([]);
  const [loading, setLoading] = useState(true);
  const [selectedId, setSelectedId] = useState(null);
  const [showUnpack, setShowUnpack] = useState(false);

  useEffect(() => {
    api.backups()
      .then((r) => setBackups(r?.data ?? []))
      .catch(() => setBackups([]))
      .finally(() => setLoading(false));
  }, []);

  const bars = buildBars(backups, 35);

  const rows = backups
    .map(mapRow)
    .filter((r) => !search || r.time.toLowerCase().includes(search.toLowerCase()));

  return (
    <div className="flex flex-col h-full">
      {/* Filter bar */}
      <div className="bg-white border-b border-[#DEE2E6] px-4 py-2 flex items-center gap-3 flex-shrink-0">
        <span className="text-xs text-[#555]">Status:</span>
        <select className="border border-[#DEE2E6] px-2 py-1 text-xs bg-white">
          <option>Any</option>
          <option>OK</option>
          <option>Error</option>
        </select>
        <div className="flex items-center gap-0.5 ml-3">
          {[
            { id: 'today', label: 'Today' },
            { id: '7days', label: '7 days' },
            { id: '30days', label: '30 days' },
          ].map((opt) => (
            <button
              key={opt.id}
              onClick={() => setDateFilter(opt.id)}
              className={`px-3 py-1 text-xs border border-[#DEE2E6] ${dateFilter === opt.id ? 'bg-[#E6F4FA] text-[#0078A8] border-[#0096D6]' : 'bg-white text-[#555] hover:bg-[#F2F2F2]'}`}
            >
              {opt.label}
            </button>
          ))}
        </div>
      </div>

      <div className="p-4 space-y-3 flex-1 overflow-auto">
        <MiniBarChart bars={bars} />

        {/* Action toolbar */}
        <div className="bg-white border border-[#DEE2E6] px-3 py-2 flex items-center gap-2">
          <button
            disabled={!selectedId}
            onClick={() => setShowUnpack(true)}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2] disabled:opacity-40 disabled:cursor-not-allowed"
          >
            <Package size={13} />
            UNPACK
          </button>
          <button className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2]">
            <Columns size={13} />
            EDIT COLUMNS
          </button>
          <div className="ml-auto flex items-center gap-2">
            <span className="text-xs text-[#666]">
              {loading ? 'Loading…' : `${rows.length} backup${rows.length !== 1 ? 's' : ''}`}
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
          <div className="text-xs text-[#888] p-4">Loading backups…</div>
        ) : rows.length === 0 ? (
          <div className="bg-white border border-[#DEE2E6] p-8 text-center text-xs text-[#888]">
            No backups found. Use "CREATE BACKUP" on the Dashboard to create one.
          </div>
        ) : (
          <DataTable
            columns={columns}
            rows={rows}
            totalCount={rows.length}
            selectedRows={selectedId ? [selectedId] : []}
            onRowClick={(row) => setSelectedId(row.id === selectedId ? null : row.id)}
          />
        )}
      </div>

      {showUnpack && (
        <BackupUnpackingModal
          backups={backups}
          preSelectedId={selectedId}
          onClose={() => setShowUnpack(false)}
        />
      )}
    </div>
  );
}
