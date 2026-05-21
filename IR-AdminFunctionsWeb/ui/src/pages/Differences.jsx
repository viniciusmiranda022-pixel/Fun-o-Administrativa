import { useState } from 'react';
import { RotateCcw, Download, RefreshCw, Columns, Search, FileText } from 'lucide-react';
import DataTable from '../components/common/DataTable.jsx';

const BASE_NAME = 'Quest Recovery Function - Administrative Roles Ba...';

const MOCK_DIFFS = [
  {
    id: 1,
    name: BASE_NAME,
    tenant: 'Contoso',
    change: 'Object hard deleted',
    objectType: 'Service Principal',
    attribute: 'Service Principal',
    diffType: 'deleted',
    diffText: BASE_NAME,
    backup: 'a day ago',
    refresh: '9 hours ago',
  },
  {
    id: 2,
    name: BASE_NAME,
    tenant: 'Contoso',
    change: 'Link added',
    objectType: 'App Role Assignment',
    attribute: 'AppRoleAssignment',
    diffType: 'added',
    diffText: 'Microsoft Graph',
    backup: 'a day ago',
    refresh: '9 hours ago',
  },
  {
    id: 3,
    name: BASE_NAME,
    tenant: 'Contoso',
    change: 'Link added',
    objectType: 'App Role Assignment',
    attribute: 'AppRoleAssignment',
    diffType: 'added',
    diffText: 'Microsoft Graph',
    backup: 'a day ago',
    refresh: '9 hours ago',
  },
  {
    id: 4,
    name: BASE_NAME,
    tenant: 'Contoso',
    change: 'Link added',
    objectType: 'App Role Assignment',
    attribute: 'AppRoleAssignment',
    diffType: 'added',
    diffText: 'Azure Active Directory Graph',
    backup: 'a day ago',
    refresh: '9 hours ago',
  },
  {
    id: 5,
    name: BASE_NAME,
    tenant: 'Contoso',
    change: 'Link added',
    objectType: 'App Role Assignment',
    attribute: 'AppRoleAssignment',
    diffType: 'added',
    diffText: 'Microsoft Graph',
    backup: 'a day ago',
    refresh: '9 hours ago',
  },
  {
    id: 6,
    name: BASE_NAME,
    tenant: 'Contoso',
    change: 'Link added',
    objectType: 'App Role Assignment',
    attribute: 'AppRoleAssignment',
    diffType: 'added',
    diffText: 'Office 365 SharePoint Online',
    backup: 'a day ago',
    refresh: '9 hours ago',
  },
  {
    id: 7,
    name: BASE_NAME,
    tenant: 'Contoso',
    change: 'Link added',
    objectType: 'App Role Assignment',
    attribute: 'AppRoleAssignment',
    diffType: 'added',
    diffText: 'Microsoft Graph',
    backup: 'a day ago',
    refresh: '9 hours ago',
  },
  {
    id: 8,
    name: BASE_NAME,
    tenant: 'Contoso',
    change: 'New object',
    objectType: 'Service Principal',
    attribute: 'Service Principal',
    diffType: 'new',
    diffText: 'Quest Recovery Function - Administrative Roles Backup',
    backup: 'a day ago',
    refresh: '9 hours ago',
  },
];

function DiffCell({ row }) {
  if (row.diffType === 'deleted') {
    return (
      <span className="text-[#DC3545] line-through text-xs">{row.diffText}</span>
    );
  }
  if (row.diffType === 'added') {
    return (
      <span className="text-[#28A745] text-xs">→ {row.diffText}</span>
    );
  }
  if (row.diffType === 'new') {
    return (
      <span className="text-[#28A745] text-xs">→ {row.diffText}</span>
    );
  }
  return <span className="text-xs">{row.diffText}</span>;
}

const columns = [
  {
    key: 'name', header: 'Name', sortable: true, width: '18%',
    render: (r) => (
      <div className="flex items-center gap-1.5">
        <FileText size={12} className="text-[#0096D6] flex-shrink-0" />
        <span className="text-[#0078A8] hover:underline cursor-pointer truncate text-xs">{r.name}</span>
      </div>
    )
  },
  { key: 'tenant', header: 'Tenant', width: '8%' },
  {
    key: 'change', header: 'Change', width: '12%',
    render: (r) => (
      <span className={`text-xs font-medium ${r.diffType === 'deleted' ? 'text-[#DC3545]' : r.diffType === 'added' ? 'text-[#28A745]' : 'text-[#17A2B8]'}`}>
        {r.change}
      </span>
    )
  },
  { key: 'objectType', header: 'Object Type', width: '13%' },
  { key: 'attribute', header: 'Attribute', width: '13%' },
  {
    key: 'difference', header: 'Difference', width: '18%',
    render: (r) => <DiffCell row={r} />
  },
  { key: 'backup', header: 'Backup ▲', sortable: true, width: '9%' },
  { key: 'refresh', header: 'Refresh', width: '9%' },
];

export default function Differences() {
  const [view, setView] = useState('list');
  const [search, setSearch] = useState('');
  const [filters, setFilters] = useState({ tenant: '', backup: '', type: '', changeType: '', top5: '' });

  function setFilter(k, v) { setFilters((f) => ({ ...f, [k]: v })); }

  const rows = MOCK_DIFFS.filter(
    (r) => !search || r.name.toLowerCase().includes(search.toLowerCase()) || r.change.toLowerCase().includes(search.toLowerCase())
  );

  const filterDefs = [
    { key: 'tenant', label: 'Tenant' },
    { key: 'backup', label: 'Backup' },
    { key: 'type', label: 'Type' },
    { key: 'changeType', label: 'Change Type' },
    { key: 'top5', label: 'Top 5 attributes' },
  ];

  return (
    <div className="flex flex-col h-full">
      {/* Toggle + Filter bar */}
      <div className="bg-white border-b border-[#DEE2E6] px-4 py-2 flex items-center gap-3 flex-shrink-0 flex-wrap">
        <div className="flex border border-[#DEE2E6] overflow-hidden text-xs">
          {['list', 'differences'].map((v) => (
            <button
              key={v}
              onClick={() => setView(v)}
              className={`px-3 py-1 capitalize ${view === v ? 'bg-[#0078A8] text-white' : 'bg-white text-[#555] hover:bg-[#F2F2F2]'}`}
            >
              {v === 'list' ? 'List View' : 'Differences'}
            </button>
          ))}
        </div>
        <span className="text-xs text-[#888]">Last Refresh: 9 hours ago</span>
        <div className="flex items-center gap-2 flex-wrap">
          {filterDefs.map(({ key, label }) => (
            <div key={key} className="flex items-center gap-1 text-xs">
              <span className="text-[#555]">{label}:</span>
              <select
                value={filters[key]}
                onChange={(e) => setFilter(key, e.target.value)}
                className="border border-[#DEE2E6] px-1.5 py-0.5 text-xs bg-white"
              >
                <option value="">Any</option>
              </select>
            </div>
          ))}
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
            <RefreshCw size={13} />
            REFRESH
          </button>
          <button className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2]">
            <Columns size={13} />
            EDIT COLUMNS
          </button>
          <div className="ml-auto flex items-center gap-2">
            <span className="text-xs text-[#666]">8 objects</span>
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

        <DataTable columns={columns} rows={rows} totalCount={8} />
      </div>
    </div>
  );
}
