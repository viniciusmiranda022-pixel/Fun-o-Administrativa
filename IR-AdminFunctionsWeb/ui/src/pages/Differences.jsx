import { useEffect, useMemo, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  Loader2, RefreshCw, AlertTriangle, FileMinus, FilePlus, FileEdit,
  ChevronDown, ChevronRight, RotateCcw, Eye, AlertCircle
} from 'lucide-react';
import { api, pollJob } from '../api/client.js';

function ChangeBadge({ kind }) {
  const map = {
    Deleted:    { cls: 'text-red-700 bg-red-50',       label: 'Excluída', Icon: FileMinus },
    Added:      { cls: 'text-blue-700 bg-blue-50',     label: 'Nova',      Icon: FilePlus },
    Modified:   { cls: 'text-yellow-700 bg-yellow-50', label: 'Alterada',  Icon: FileEdit },
    Unchanged:  { cls: 'text-gray-600 bg-gray-100',    label: 'Sem mudanças', Icon: null }
  };
  const c = map[kind] || map.Unchanged;
  const Icon = c.Icon;
  return (
    <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-semibold ${c.cls}`}>
      {Icon && <Icon size={11} />} {c.label}
    </span>
  );
}

function classify(row) {
  const s = (row.Status || '').toString().toLowerCase();
  if (s.includes('delet') || s.includes('miss') || s.includes('exclu')) return 'Deleted';
  if (s.includes('add') || s.includes('new') || s.includes('nov')) return 'Added';
  if (s.includes('mod') || s.includes('alter') || s.includes('chang') || s.includes('diff')) return 'Modified';
  if ((row.DifferenceCount ?? 0) > 0) return 'Modified';
  return 'Unchanged';
}

function DiffRow({ row, selected, onToggle, expanded, onExpand, onRestore }) {
  const kind = classify(row);
  return (
    <>
      <tr className={`border-t border-border-light ${selected ? 'bg-blue-50' : 'hover:bg-gray-50'}`}>
        <td className="px-3 py-2">
          <input type="checkbox" checked={selected} onChange={onToggle} disabled={kind === 'Unchanged'} />
        </td>
        <td className="px-3 py-2">
          <button onClick={onExpand} className="text-text-secondary hover:text-[#0078A8]">
            {expanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
          </button>
        </td>
        <td className="px-3 py-2">
          <div className="font-semibold text-[#222] text-sm">{row.RoleName ?? row.name ?? '—'}</div>
          <div className="text-[11px] text-text-secondary">{row.RoleType ?? row.type ?? 'Custom'}</div>
        </td>
        <td className="px-3 py-2"><ChangeBadge kind={kind} /></td>
        <td className="px-3 py-2 text-xs">{row.DefinitionChanged ?? '—'}</td>
        <td className="px-3 py-2 text-xs">{row.AssignmentsChanged ?? '—'}</td>
        <td className="px-3 py-2 text-xs text-center">{row.DifferenceCount ?? 0}</td>
        <td className="px-3 py-2 text-right">
          {kind !== 'Unchanged' && (
            <button onClick={() => onRestore(row)} className="text-xs text-[#0078A8] hover:underline flex items-center gap-1 justify-end">
              <RotateCcw size={11} /> Restore
            </button>
          )}
        </td>
      </tr>
      {expanded && (
        <tr className="bg-gray-50">
          <td colSpan={8} className="px-6 py-3 text-xs text-[#333]">
            <div className="grid grid-cols-2 gap-4">
              <div>
                <div className="font-semibold mb-1">Summary</div>
                <div className="text-text-secondary">{row.Summary || row.summary || '—'}</div>
              </div>
              <div>
                <div className="font-semibold mb-1">Raw</div>
                <pre className="text-[10px] font-mono bg-white border border-border-light rounded p-2 overflow-x-auto max-h-32">
                  {JSON.stringify(row, null, 2)}
                </pre>
              </div>
            </div>
          </td>
        </tr>
      )}
    </>
  );
}

function RestorePreviewModal({ items, onClose, onApply, applying, previewData, previewError }) {
  return (
    <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center">
      <div className="bg-white rounded-lg shadow-xl w-full max-w-3xl mx-4 max-h-[90vh] flex flex-col">
        <div className="flex items-center justify-between px-6 py-4 border-b border-border-light">
          <h2 className="text-base font-semibold text-[#222]">Restore selected administrative functions</h2>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-700 text-xl">×</button>
        </div>

        <div className="px-6 py-4 overflow-y-auto space-y-4">
          <div className="bg-yellow-50 border border-yellow-200 rounded p-3 text-xs text-yellow-900 flex gap-2">
            <AlertTriangle size={14} className="flex-shrink-0 mt-0.5" />
            <span>
              Esta operação altera configurações administrativas no Microsoft Entra ID (RBAC).
              Revise os itens selecionados antes de continuar. Requer consentimento <strong>Restore</strong>.
            </span>
          </div>

          <div>
            <h3 className="text-sm font-semibold mb-2 text-[#222]">Itens selecionados ({items.length})</h3>
            <div className="border border-border-light rounded max-h-48 overflow-y-auto">
              <table className="w-full text-xs">
                <thead className="bg-gray-50 text-text-secondary">
                  <tr>
                    <th className="text-left px-3 py-1.5 font-semibold">Role</th>
                    <th className="text-left px-3 py-1.5 font-semibold">Tipo</th>
                    <th className="text-left px-3 py-1.5 font-semibold">Mudança</th>
                  </tr>
                </thead>
                <tbody>
                  {items.map((i, idx) => (
                    <tr key={idx} className="border-t border-border-light">
                      <td className="px-3 py-1.5 font-mono">{i.RoleName ?? i.name}</td>
                      <td className="px-3 py-1.5">{i.RoleType ?? 'Custom'}</td>
                      <td className="px-3 py-1.5"><ChangeBadge kind={classify(i)} /></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {previewData && (
            <div>
              <h3 className="text-sm font-semibold mb-2 text-[#222]">Preview (What-If)</h3>
              <pre className="text-[10px] font-mono bg-gray-50 border border-border-light rounded p-3 overflow-auto max-h-40">
                {typeof previewData === 'string' ? previewData : JSON.stringify(previewData, null, 2)}
              </pre>
            </div>
          )}

          {previewError && (
            <div className="bg-red-50 border border-red-200 text-red-800 text-xs px-3 py-2 rounded">{previewError}</div>
          )}
        </div>

        <div className="flex items-center justify-end gap-2 px-6 py-3 border-t border-border-light bg-gray-50">
          <button onClick={onClose} disabled={applying}
            className="px-4 py-2 text-sm border border-border-light rounded hover:bg-white disabled:opacity-60">Cancelar</button>
          <button onClick={onApply} disabled={applying || items.length === 0}
            className="px-4 py-2 text-sm bg-red-600 text-white rounded hover:bg-red-700 disabled:opacity-60 flex items-center gap-2">
            {applying && <Loader2 size={14} className="animate-spin" />}
            Aplicar restore
          </button>
        </div>
      </div>
    </div>
  );
}

export default function Differences() {
  const navigate = useNavigate();
  const [rows, setRows] = useState([]);
  const [generatedAt, setGeneratedAt] = useState(null);
  const [running, setRunning] = useState(false);
  const [error, setError] = useState(null);
  const [filter, setFilter] = useState('All');
  const [selected, setSelected] = useState(new Set());
  const [expanded, setExpanded] = useState(new Set());
  const [showRestore, setShowRestore] = useState(false);
  const [previewData, setPreviewData] = useState(null);
  const [previewError, setPreviewError] = useState(null);
  const [applying, setApplying] = useState(false);

  const loadCached = () => {
    api.compareResults()
      .then((r) => {
        const items = Array.isArray(r?.data?.data) ? r.data.data : [];
        setRows(items);
        setGeneratedAt(r?.data?.generatedAt);
      })
      .catch((e) => setError(e.message));
  };

  useEffect(loadCached, []);

  const handleRefresh = async () => {
    setRunning(true);
    setError(null);
    try {
      const resp = await api.runCompare();
      const job = resp?.data;
      const finished = await pollJob(job.id);
      if (finished.status === 'Failed') setError(finished.error || 'Compare falhou');
      else loadCached();
    } catch (e) {
      setError(e.message);
    } finally {
      setRunning(false);
    }
  };

  const counts = useMemo(() => {
    const c = { Deleted: 0, Added: 0, Modified: 0, Unchanged: 0 };
    rows.forEach((r) => c[classify(r)]++);
    return c;
  }, [rows]);

  const filtered = useMemo(() => {
    if (filter === 'All') return rows;
    return rows.filter((r) => classify(r) === filter);
  }, [rows, filter]);

  const toggleRow = (key) => {
    const s = new Set(selected);
    s.has(key) ? s.delete(key) : s.add(key);
    setSelected(s);
  };

  const toggleExpand = (key) => {
    const s = new Set(expanded);
    s.has(key) ? s.delete(key) : s.add(key);
    setExpanded(s);
  };

  const selectedItems = useMemo(() =>
    rows.filter((r) => selected.has(r.RoleName ?? r.name)), [rows, selected]);

  const startRestoreOne = (row) => {
    setSelected(new Set([row.RoleName ?? row.name]));
    setShowRestore(true);
    setPreviewData(null);
    setPreviewError(null);
    runPreview([row]);
  };

  const startRestoreSelected = () => {
    setShowRestore(true);
    setPreviewData(null);
    setPreviewError(null);
    runPreview(selectedItems);
  };

  const runPreview = async (items) => {
    try {
      const r = await api.previewRestore({
        roleName: items[0]?.RoleName ?? items[0]?.name,
        removeExtraAssignments: false
      });
      const finished = await pollJob(r?.data?.id);
      setPreviewData(finished?.result || finished?.error || 'Sem preview disponível');
    } catch (e) {
      setPreviewError(e.message);
    }
  };

  const applyRestore = async () => {
    setApplying(true);
    try {
      for (const it of selectedItems) {
        await api.applyRestore({
          roleName: it.RoleName ?? it.name,
          removeExtraAssignments: false
        });
      }
      setShowRestore(false);
      setSelected(new Set());
      navigate('/tasks');
    } catch (e) {
      setPreviewError(e.message);
    } finally {
      setApplying(false);
    }
  };

  const FilterTab = ({ value, label, count, color }) => (
    <button onClick={() => setFilter(value)}
      className={`px-3 py-1.5 text-xs rounded border ${filter === value ? 'bg-[#0078A8] text-white border-[#0078A8]' : 'bg-white border-border-light text-[#333] hover:bg-gray-50'}`}>
      {label} {count != null && <span className={filter === value ? 'opacity-80' : color}>({count})</span>}
    </button>
  );

  return (
    <div className="p-6">
      <div className="flex items-center justify-between mb-4">
        <div>
          <h2 className="text-base font-semibold text-[#222]">Differences</h2>
          <p className="text-xs text-text-secondary mt-0.5">
            {generatedAt
              ? `Comparação gerada em ${new Date(generatedAt).toLocaleString('pt-BR')}`
              : 'Nenhuma comparação ainda. Clique em Refresh report para gerar.'}
          </p>
        </div>
        <div className="flex gap-2">
          {selected.size > 0 && (
            <button onClick={startRestoreSelected}
              className="px-3 py-1.5 text-xs bg-[#0078A8] text-white rounded hover:bg-[#006090] flex items-center gap-1.5">
              <RotateCcw size={12} /> Restore selected ({selected.size})
            </button>
          )}
          <button onClick={handleRefresh} disabled={running}
            className="px-3 py-1.5 text-xs bg-white border border-border-light rounded hover:bg-gray-50 flex items-center gap-1.5 disabled:opacity-50">
            {running ? <Loader2 size={12} className="animate-spin" /> : <RefreshCw size={12} />}
            Refresh report
          </button>
        </div>
      </div>

      {error && <div className="bg-red-50 border border-red-200 text-red-800 text-sm px-3 py-2 rounded mb-3">{error}</div>}

      <div className="flex flex-wrap gap-2 mb-4">
        <FilterTab value="All" label="Todos" count={rows.length} />
        <FilterTab value="Deleted" label="Excluídas" count={counts.Deleted} color="text-red-600" />
        <FilterTab value="Added" label="Novas" count={counts.Added} color="text-blue-600" />
        <FilterTab value="Modified" label="Alteradas" count={counts.Modified} color="text-yellow-600" />
        <FilterTab value="Unchanged" label="Sem mudanças" count={counts.Unchanged} />
      </div>

      <div className="bg-white border border-border-light rounded-lg overflow-hidden">
        {filtered.length === 0 ? (
          <div className="text-center py-12 text-text-secondary">
            <Eye size={28} className="mx-auto mb-2 opacity-30" />
            <p className="text-sm">{rows.length === 0 ? 'Sem comparação. Rode Refresh report.' : 'Nenhum item neste filtro.'}</p>
          </div>
        ) : (
          <table className="w-full text-sm">
            <thead className="bg-gray-50 text-xs text-text-secondary">
              <tr>
                <th className="px-3 py-2 w-8"></th>
                <th className="px-3 py-2 w-8"></th>
                <th className="text-left px-3 py-2 font-semibold">Role</th>
                <th className="text-left px-3 py-2 font-semibold">Mudança</th>
                <th className="text-left px-3 py-2 font-semibold">Definition</th>
                <th className="text-left px-3 py-2 font-semibold">Assignments</th>
                <th className="text-center px-3 py-2 font-semibold">Diffs</th>
                <th className="text-right px-3 py-2 font-semibold">Ações</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((row, idx) => {
                const key = row.RoleName ?? row.name ?? `row-${idx}`;
                return (
                  <DiffRow key={key} row={row}
                    selected={selected.has(key)}
                    onToggle={() => toggleRow(key)}
                    expanded={expanded.has(key)}
                    onExpand={() => toggleExpand(key)}
                    onRestore={startRestoreOne} />
                );
              })}
            </tbody>
          </table>
        )}
      </div>

      {selectedItems.length > 0 && !showRestore && (
        <div className="mt-3 bg-blue-50 border border-blue-200 rounded p-3 text-xs text-blue-900 flex items-center gap-2">
          <AlertCircle size={14} />
          {selectedItems.length} item(ns) selecionado(s). Clique em <strong>Restore selected</strong> para revisar antes de aplicar.
        </div>
      )}

      {showRestore && (
        <RestorePreviewModal items={selectedItems} onClose={() => setShowRestore(false)}
          onApply={applyRestore} applying={applying}
          previewData={previewData} previewError={previewError} />
      )}
    </div>
  );
}
