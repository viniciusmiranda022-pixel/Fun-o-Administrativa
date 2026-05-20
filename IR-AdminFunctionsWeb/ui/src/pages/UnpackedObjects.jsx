import { useEffect, useMemo, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  Search, RefreshCw, X, ShieldCheck, Key, Users, Globe, RotateCcw, ChevronRight, ChevronDown
} from 'lucide-react';
import { api } from '../api/client.js';
import { useTenant } from '../context/TenantContext.jsx';

const TABS = [
  { id: 'roleDefinitions',  label: 'Funções customizadas', Icon: ShieldCheck },
  { id: 'rolePermissions',  label: 'Permissões',           Icon: Key },
  { id: 'roleAssignments',  label: 'Atribuições',          Icon: Users },
  { id: 'assignmentScopes', label: 'Escopos',              Icon: Globe },
];

function ObjectPanel({ obj, tab, onClose, onRestore }) {
  if (!obj) return null;
  const name = obj.displayName ?? obj.DisplayName ?? obj.roleName ?? obj.RoleName ?? obj.name ?? obj.Name ?? '—';
  const id = obj.id ?? obj.Id ?? obj.roleDefinitionId ?? obj.RoleDefinitionId ?? '—';

  return (
    <div className="fixed inset-y-0 right-0 w-[420px] bg-white border-l border-border-light shadow-xl z-40 overflow-y-auto">
      <div className="flex items-center justify-between px-5 py-4 border-b border-border-light">
        <div>
          <h3 className="text-sm font-semibold text-[#222]">{name}</h3>
          <div className="text-[11px] text-text-secondary font-mono mt-0.5">{id}</div>
        </div>
        <button onClick={onClose} className="text-gray-400 hover:text-gray-700"><X size={18} /></button>
      </div>

      <div className="p-5 space-y-4 text-sm">
        {tab === 'roleDefinitions' && (
          <>
            <div className="grid grid-cols-2 gap-3 text-xs">
              <div><span className="text-text-secondary">Tipo:</span> <span>{obj.isBuiltIn || obj.IsBuiltIn ? 'Built-in' : 'Custom'}</span></div>
              <div><span className="text-text-secondary">Habilitada:</span> <span>{String(obj.isEnabled ?? obj.IsEnabled ?? '—')}</span></div>
            </div>
            {(obj.description || obj.Description) && (
              <div>
                <div className="text-xs font-semibold text-text-secondary uppercase tracking-wide mb-1">Descrição</div>
                <div className="text-xs text-[#333]">{obj.description ?? obj.Description}</div>
              </div>
            )}
            {(obj.allowedResourceActions || obj.AllowedResourceActions || obj.permissions || obj.Permissions) && (
              <div>
                <div className="text-xs font-semibold text-text-secondary uppercase tracking-wide mb-1">Permissões permitidas</div>
                <pre className="text-[10px] font-mono bg-gray-50 border border-border-light rounded p-2 overflow-auto max-h-40">
                  {JSON.stringify(obj.allowedResourceActions ?? obj.AllowedResourceActions ?? obj.permissions ?? obj.Permissions, null, 2)}
                </pre>
              </div>
            )}
          </>
        )}

        {tab === 'roleAssignments' && (
          <div className="grid grid-cols-1 gap-2 text-xs">
            {[
              ['Principal ID',   obj.principalId ?? obj.PrincipalId],
              ['Role Def ID',    obj.roleDefinitionId ?? obj.RoleDefinitionId],
              ['Scope',          obj.directoryScopeId ?? obj.DirectoryScopeId ?? obj.scope ?? obj.Scope],
              ['Tenant',         obj.appScopeId ?? obj.tenantId],
            ].map(([label, val]) => val ? (
              <div key={label} className="flex gap-2">
                <span className="text-text-secondary w-24 flex-shrink-0">{label}:</span>
                <span className="font-mono text-[11px] break-all">{val}</span>
              </div>
            ) : null)}
          </div>
        )}

        {(tab === 'rolePermissions' || tab === 'assignmentScopes') && (
          <pre className="text-[10px] font-mono bg-gray-50 border border-border-light rounded p-3 overflow-auto max-h-96">
            {JSON.stringify(obj, null, 2)}
          </pre>
        )}

        <div className="pt-2">
          <button onClick={() => onRestore(obj)}
            className="flex items-center gap-2 text-xs px-3 py-2 bg-[#0078A8] text-white rounded hover:bg-[#006090]">
            <RotateCcw size={12} /> Restaurar este objeto
          </button>
        </div>
      </div>
    </div>
  );
}

export default function UnpackedObjects() {
  const navigate = useNavigate();
  const { selectedTenantId } = useTenant();
  const [backups, setBackups] = useState([]);
  const [selectedBackupId, setSelectedBackupId] = useState('');
  const [data, setData] = useState({});
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [activeTab, setActiveTab] = useState('roleDefinitions');
  const [query, setQuery] = useState('');
  const [panelObj, setPanelObj] = useState(null);
  const [expandedRows, setExpandedRows] = useState(new Set());

  useEffect(() => {
    api.backups().then((r) => {
      const items = (r?.data ?? []).filter((b) =>
        !selectedTenantId || !b.tenantId || b.tenantId === selectedTenantId
      );
      setBackups(items);
      if (items.length > 0) setSelectedBackupId(items[0].id);
    });
  }, [selectedTenantId]);

  useEffect(() => {
    if (!selectedBackupId) return;
    setLoading(true);
    setError(null);
    setData({});
    setExpandedRows(new Set());
    api.unpacked(selectedBackupId)
      .then((r) => setData(r?.data ?? {}))
      .catch((e) => setError(e.message))
      .finally(() => setLoading(false));
  }, [selectedBackupId]);

  const currentList = useMemo(() => {
    const raw = data[activeTab] ?? [];
    if (!Array.isArray(raw)) return [];
    if (!query.trim()) return raw;
    const q = query.trim().toLowerCase();
    return raw.filter((item) => {
      const name = (item.displayName ?? item.DisplayName ?? item.roleName ?? item.RoleName ?? item.name ?? item.id ?? '').toString().toLowerCase();
      const desc = (item.description ?? item.Description ?? '').toString().toLowerCase();
      return name.includes(q) || desc.includes(q);
    });
  }, [data, activeTab, query]);

  const tabCounts = useMemo(() => {
    const c = {};
    TABS.forEach((t) => { c[t.id] = Array.isArray(data[t.id]) ? data[t.id].length : 0; });
    return c;
  }, [data]);

  const toggleExpand = (idx) => {
    const s = new Set(expandedRows);
    s.has(idx) ? s.delete(idx) : s.add(idx);
    setExpandedRows(s);
  };

  const handleRestore = (obj) => {
    const name = obj.displayName ?? obj.DisplayName ?? obj.roleName ?? obj.RoleName ?? obj.name ?? '';
    navigate('/differences', { state: { restoreTarget: name } });
  };

  const getRowLabel = (item) =>
    item.displayName ?? item.DisplayName ?? item.roleName ?? item.RoleName ??
    item.principalId ?? item.PrincipalId ?? item.id ?? '—';

  const getRowSub = (item) => {
    if (activeTab === 'roleDefinitions') return item.isBuiltIn || item.IsBuiltIn ? 'Built-in' : 'Custom';
    if (activeTab === 'roleAssignments') return item.roleDefinitionId ?? item.RoleDefinitionId ?? '';
    if (activeTab === 'rolePermissions') return item.description ?? item.Description ?? '';
    return item.id ?? item.Id ?? '';
  };

  return (
    <div className="p-6">
      <div className="flex items-center justify-between mb-4">
        <div>
          <h2 className="text-base font-semibold text-[#222]">Unpacked Objects</h2>
          <p className="text-xs text-text-secondary mt-0.5">Objetos extraídos do backup selecionado.</p>
        </div>
        <div className="flex items-center gap-3">
          <select value={selectedBackupId} onChange={(e) => setSelectedBackupId(e.target.value)}
            className="text-xs border border-border-light rounded px-2 py-1.5 bg-white max-w-[280px]">
            {backups.length === 0 && <option value="">Nenhum backup</option>}
            {backups.map((b) => (
              <option key={b.id} value={b.id}>
                {b.id?.substring(0, 20)}… {b.collectedAt ? `(${new Date(b.collectedAt).toLocaleDateString('pt-BR')})` : ''}
              </option>
            ))}
          </select>
          <button onClick={() => { if (selectedBackupId) { setLoading(true); api.unpacked(selectedBackupId).then((r) => setData(r?.data ?? {})).finally(() => setLoading(false)); } }}
            className="text-xs flex items-center gap-1 px-3 py-1.5 border border-border-light rounded bg-white hover:bg-gray-50">
            <RefreshCw size={12} /> Recarregar
          </button>
        </div>
      </div>

      {error && <div className="bg-red-50 border border-red-200 text-red-800 text-xs px-3 py-2 rounded mb-3">{error}</div>}

      <div className="flex items-center gap-2 mb-3 flex-wrap">
        {TABS.map(({ id, label, Icon }) => (
          <button key={id} onClick={() => { setActiveTab(id); setQuery(''); setPanelObj(null); }}
            className={`flex items-center gap-1.5 px-3 py-1.5 text-xs rounded border ${activeTab === id ? 'bg-[#0078A8] text-white border-[#0078A8]' : 'bg-white border-border-light text-[#333] hover:bg-gray-50'}`}>
            <Icon size={12} /> {label}
            <span className={`ml-1 ${activeTab === id ? 'opacity-80' : 'text-text-secondary'}`}>({tabCounts[id]})</span>
          </button>
        ))}
        <div className="ml-auto relative">
          <Search size={12} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-text-secondary" />
          <input type="search" placeholder="Filtrar por nome..." value={query}
            onChange={(e) => setQuery(e.target.value)}
            className="pl-7 pr-3 py-1.5 text-xs border border-border-light rounded w-56" />
        </div>
      </div>

      <div className={`bg-white border border-border-light rounded-lg overflow-hidden ${panelObj ? 'mr-[440px]' : ''}`}>
        {loading ? (
          <div className="text-center py-12 text-text-secondary text-sm">Carregando objetos...</div>
        ) : currentList.length === 0 ? (
          <div className="text-center py-12 text-text-secondary text-sm">
            {Object.keys(data).length === 0 ? 'Selecione um backup com objetos descompactados.' : 'Nenhum objeto nesta categoria.'}
          </div>
        ) : (
          <table className="w-full text-sm">
            <thead className="bg-gray-50 text-xs text-text-secondary">
              <tr>
                <th className="w-8 px-3 py-2"></th>
                <th className="text-left px-3 py-2 font-semibold">Nome</th>
                <th className="text-left px-3 py-2 font-semibold">Detalhe</th>
                <th className="px-3 py-2 w-8"></th>
              </tr>
            </thead>
            <tbody>
              {currentList.map((item, idx) => {
                const isExpanded = expandedRows.has(idx);
                return (
                  <>
                    <tr key={idx}
                      onClick={() => { setPanelObj(item); }}
                      className={`border-t border-border-light cursor-pointer ${panelObj === item ? 'bg-blue-50' : 'hover:bg-gray-50'}`}>
                      <td className="px-3 py-2">
                        <button onClick={(e) => { e.stopPropagation(); toggleExpand(idx); }}
                          className="text-text-secondary hover:text-[#0078A8]">
                          {isExpanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
                        </button>
                      </td>
                      <td className="px-3 py-2">
                        <div className="font-medium text-[#222] text-xs">{getRowLabel(item)}</div>
                      </td>
                      <td className="px-3 py-2 text-xs text-text-secondary truncate max-w-xs">{getRowSub(item)}</td>
                      <td className="px-3 py-2 text-right">
                        <button onClick={(e) => { e.stopPropagation(); handleRestore(item); }}
                          className="text-xs text-[#0078A8] hover:underline flex items-center gap-1 justify-end whitespace-nowrap">
                          <RotateCcw size={11} /> Restaurar
                        </button>
                      </td>
                    </tr>
                    {isExpanded && (
                      <tr key={`${idx}-exp`} className="bg-gray-50 border-t border-border-light">
                        <td colSpan={4} className="px-6 py-3">
                          <pre className="text-[10px] font-mono overflow-auto max-h-32 whitespace-pre-wrap">
                            {JSON.stringify(item, null, 2)}
                          </pre>
                        </td>
                      </tr>
                    )}
                  </>
                );
              })}
            </tbody>
          </table>
        )}
      </div>

      <ObjectPanel obj={panelObj} tab={activeTab} onClose={() => setPanelObj(null)} onRestore={handleRestore} />
    </div>
  );
}
