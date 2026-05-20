import { useEffect, useMemo, useState } from 'react';
import { Download, RefreshCw, AlertCircle, AlertTriangle, Info, ChevronDown, ChevronRight } from 'lucide-react';
import { api } from '../api/client.js';

function SeverityPill({ value }) {
  const map = {
    Error:       { cls: 'text-red-700 bg-red-50',     Icon: AlertCircle },
    Warning:     { cls: 'text-yellow-700 bg-yellow-50', Icon: AlertTriangle },
    Information: { cls: 'text-blue-700 bg-blue-50',    Icon: Info }
  };
  const c = map[value] || map.Information;
  const Icon = c.Icon;
  return (
    <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-semibold ${c.cls}`}>
      <Icon size={11} /> {value}
    </span>
  );
}

export default function Events() {
  const [rows, setRows] = useState([]);
  const [severity, setSeverity] = useState('');
  const [query, setQuery] = useState('');
  const [expanded, setExpanded] = useState(new Set());
  const [loading, setLoading] = useState(true);

  const load = async () => {
    setLoading(true);
    try {
      const r = await api.events(severity || undefined);
      setRows(r?.data?.items ?? []);
    } catch {
      setRows([]);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); /* eslint-disable-line */ }, [severity]);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return rows;
    return rows.filter((r) =>
      (r.message ?? '').toLowerCase().includes(q) ||
      (r.source ?? '').toLowerCase().includes(q));
  }, [rows, query]);

  const counts = useMemo(() => {
    const c = { Error: 0, Warning: 0, Information: 0 };
    rows.forEach((r) => { c[r.severity] = (c[r.severity] || 0) + 1; });
    return c;
  }, [rows]);

  const toggleExpand = (idx) => {
    const s = new Set(expanded);
    s.has(idx) ? s.delete(idx) : s.add(idx);
    setExpanded(s);
  };

  const exportCsv = async () => {
    try {
      const blob = await api.exportEvents(severity || undefined);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `events-${new Date().toISOString().replace(/[:.]/g, '-')}.csv`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (e) {
      alert(`Falha ao exportar: ${e.message}`);
    }
  };

  const SevTab = ({ value, label, count, color }) => (
    <button onClick={() => setSeverity(value)}
      className={`px-3 py-1.5 text-xs rounded border ${severity === value ? 'bg-[#0078A8] text-white border-[#0078A8]' : 'bg-white border-border-light text-[#333] hover:bg-gray-50'}`}>
      {label} {count != null && <span className={severity === value ? 'opacity-80' : color}>({count})</span>}
    </button>
  );

  return (
    <div className="p-6">
      <div className="flex items-center justify-between mb-4">
        <div>
          <h2 className="text-base font-semibold text-[#222]">Events</h2>
          <p className="text-xs text-text-secondary mt-0.5">Histórico operacional detalhado. Útil para auditoria e troubleshooting.</p>
        </div>
        <div className="flex gap-2">
          <button onClick={exportCsv} className="text-xs flex items-center gap-1 px-3 py-1.5 border border-border-light rounded bg-white hover:bg-gray-50">
            <Download size={12} /> Exportar CSV
          </button>
          <button onClick={load} className="text-xs flex items-center gap-1 px-3 py-1.5 border border-border-light rounded bg-white hover:bg-gray-50">
            <RefreshCw size={12} /> Atualizar
          </button>
        </div>
      </div>

      <div className="flex flex-wrap items-center gap-2 mb-3">
        <SevTab value="" label="Todos" count={rows.length} />
        <SevTab value="Error" label="Erros" count={counts.Error} color="text-red-600" />
        <SevTab value="Warning" label="Avisos" count={counts.Warning} color="text-yellow-600" />
        <SevTab value="Information" label="Info" count={counts.Information} color="text-blue-600" />
        <input type="search" placeholder="Filtrar por mensagem ou source..."
          value={query} onChange={(e) => setQuery(e.target.value)}
          className="ml-auto px-3 py-1.5 text-xs border border-border-light rounded w-64" />
      </div>

      <div className="bg-white border border-border-light rounded-lg overflow-hidden">
        {loading ? (
          <div className="text-center py-12 text-text-secondary text-sm">Carregando...</div>
        ) : filtered.length === 0 ? (
          <div className="text-center py-12 text-text-secondary text-sm">Nenhum evento.</div>
        ) : (
          <table className="w-full text-sm">
            <thead className="bg-gray-50 text-xs text-text-secondary">
              <tr>
                <th className="px-3 py-2 w-8"></th>
                <th className="text-left px-3 py-2 font-semibold w-44">Time</th>
                <th className="text-left px-3 py-2 font-semibold w-28">Severity</th>
                <th className="text-left px-3 py-2 font-semibold w-56">Source</th>
                <th className="text-left px-3 py-2 font-semibold">Message</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((e, idx) => {
                const isExpanded = expanded.has(idx);
                return (
                  <>
                    <tr key={idx} className="border-t border-border-light hover:bg-gray-50 cursor-pointer"
                      onClick={() => toggleExpand(idx)}>
                      <td className="px-3 py-2 text-text-secondary">
                        {isExpanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
                      </td>
                      <td className="px-3 py-2 text-xs font-mono">{new Date(e.time).toLocaleString('pt-BR')}</td>
                      <td className="px-3 py-2"><SeverityPill value={e.severity} /></td>
                      <td className="px-3 py-2 text-xs text-text-secondary truncate">{e.source}</td>
                      <td className="px-3 py-2 text-xs truncate">{e.message}</td>
                    </tr>
                    {isExpanded && (
                      <tr key={`${idx}-d`} className="bg-gray-50 border-t border-border-light">
                        <td colSpan={5} className="px-6 py-3">
                          <dl className="text-xs grid grid-cols-2 gap-2">
                            <div><dt className="text-text-secondary inline">Time:</dt> <dd className="inline font-mono">{e.time}</dd></div>
                            <div><dt className="text-text-secondary inline">Severity:</dt> <dd className="inline">{e.severity}</dd></div>
                            <div className="col-span-2"><dt className="text-text-secondary inline">Source:</dt> <dd className="inline font-mono">{e.source}</dd></div>
                            <div className="col-span-2"><dt className="text-text-secondary inline">Message:</dt>
                              <dd className="mt-1 font-mono bg-white p-2 rounded border border-border-light whitespace-pre-wrap">{e.message}</dd></div>
                          </dl>
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
    </div>
  );
}
