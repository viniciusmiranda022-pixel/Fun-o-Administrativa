import { useState, useEffect } from 'react';
import { Download, CheckSquare, Columns, Search, Info, AlertTriangle, XCircle, Trash2 } from 'lucide-react';
import DataTable from '../components/common/DataTable';
import { api } from '../api/client';

interface MappedEvent {
  id: string;
  time: string;
  description: string;
  source: string;
  severity: string;
}

interface RawEvent {
  time?: string;
  Time?: string;
  message?: string;
  Message?: string;
  description?: string;
  source?: string;
  Source?: string;
  severity?: string;
  Severity?: string;
}

function SeverityIcon({ severity }: { severity: string }) {
  const s = (severity || '').toLowerCase();
  if (s === 'error') return <XCircle size={13} className="text-[#DC3545]" />;
  if (s === 'warning') return <AlertTriangle size={13} className="text-[#FFC107]" />;
  return <Info size={13} className="text-[#17A2B8]" />;
}

import type { Column } from '../components/common/DataTable';

const columns: Column<MappedEvent>[] = [
  {
    key: 'severity', header: '', width: '3%',
    render: (r) => <SeverityIcon severity={r.severity} />
  },
  { key: 'time', header: 'Time ▼', sortable: true, width: '16%' },
  {
    key: 'description', header: 'Description', width: '45%',
    render: (r) => <span className="line-clamp-2 text-xs">{r.description}</span>
  },
  { key: 'source', header: 'Source', width: '18%' },
  {
    key: 'severity2', header: 'Severity', width: '10%',
    render: (r) => (
      <span className={`text-xs font-semibold ${(r.severity || '').toLowerCase() === 'error' ? 'text-[#DC3545]' : (r.severity || '').toLowerCase() === 'warning' ? 'text-[#FFC107]' : 'text-[#17A2B8]'}`}>
        {r.severity}
      </span>
    )
  },
];

function mapEvent(e: RawEvent): MappedEvent {
  const d = new Date(e.time || e.Time || '');
  return {
    id: `${e.time}-${Math.random()}`,
    time: isNaN(d.getTime()) ? (e.time || '—') : d.toLocaleString(),
    description: e.message || e.Message || e.description || '—',
    source: e.source || e.Source || '—',
    severity: e.severity || e.Severity || 'Information',
  };
}

function sinceFromFilter(filter: string): string | undefined {
  const now = new Date();
  if (filter === 'today') {
    const d = new Date(now); d.setHours(0, 0, 0, 0); return d.toISOString();
  }
  if (filter === '7days') return new Date(now.getTime() - 7 * 86400000).toISOString();
  if (filter === '30days') return new Date(now.getTime() - 30 * 86400000).toISOString();
  return undefined;
}

export default function Events() {
  const [severity, setSeverity] = useState('');
  const [dateFilter, setDateFilter] = useState('30days');
  const [search, setSearch] = useState('');
  const [allEvents, setAllEvents] = useState<MappedEvent[]>([]);
  const [loading, setLoading] = useState(true);

  function loadEvents() {
    setLoading(true);
    api.events(severity || undefined, sinceFromFilter(dateFilter))
      .then((r) => setAllEvents((r?.data?.items ?? []).map(mapEvent)))
      .catch(() => setAllEvents([]))
      .finally(() => setLoading(false));
  }

  function clearEvents() {
    api.clearEvents().then(() => loadEvents()).catch(() => {});
  }

  useEffect(() => { loadEvents(); }, [severity, dateFilter]);

  const rows = allEvents.filter((r) => {
    if (search && !r.description.toLowerCase().includes(search.toLowerCase()) && !r.source.toLowerCase().includes(search.toLowerCase())) return false;
    return true;
  });

  return (
    <div className="flex flex-col h-full">
      {/* Filter bar */}
      <div className="bg-white border-b border-[#DEE2E6] px-4 py-2 flex items-center gap-3 flex-shrink-0 flex-wrap">
        <div className="flex items-center gap-1 text-xs">
          <span className="text-[#555]">Severity:</span>
          <select
            value={severity}
            onChange={(e) => setSeverity(e.target.value)}
            className="border border-[#DEE2E6] px-1.5 py-0.5 text-xs bg-white"
          >
            <option value="">Any</option>
            <option value="Information">Information</option>
            <option value="Warning">Warning</option>
            <option value="Error">Error</option>
          </select>
        </div>
        <div className="flex items-center gap-0.5">
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
        {/* Action toolbar */}
        <div className="bg-white border border-[#DEE2E6] px-3 py-2 flex items-center gap-2">
          <button className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2]">
            <Download size={13} />
            EXPORT
          </button>
          <button className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2]">
            <CheckSquare size={13} />
            ACKNOWLEDGE
          </button>
          <button className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2]">
            <Columns size={13} />
            EDIT COLUMNS
          </button>
          <button
            onClick={clearEvents}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#DC3545] hover:bg-[#FFF5F5]"
          >
            <Trash2 size={13} />
            CLEAR
          </button>
          <div className="ml-auto flex items-center gap-2">
            <span className="text-xs text-[#666]">
              {loading ? 'Loading…' : `${rows.length} event${rows.length !== 1 ? 's' : ''}`}
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
          <div className="text-xs text-[#888] p-4">Loading events…</div>
        ) : rows.length === 0 ? (
          <div className="bg-white border border-[#DEE2E6] p-8 text-center text-xs text-[#888]">
            No events found. Events are generated as backup, unpack, compare, and restore operations run.
          </div>
        ) : (
          <DataTable columns={columns} rows={rows} totalCount={rows.length} showCheckboxes={true} />
        )}
      </div>
    </div>
  );
}
