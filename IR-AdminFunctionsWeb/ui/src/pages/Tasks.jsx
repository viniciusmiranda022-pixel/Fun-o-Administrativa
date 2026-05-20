import { useEffect, useState } from 'react';
import {
  Loader2, RefreshCw, CheckCircle2, AlertCircle, Clock, Play,
  Database, Archive, GitCompare, RotateCcw, ClipboardList, X
} from 'lucide-react';
import { api } from '../api/client.js';

function TypeIcon({ type }) {
  const t = (type || '').toLowerCase();
  if (t.includes('backup')) return <Database size={14} className="text-blue-600" />;
  if (t.includes('unpack')) return <Archive size={14} className="text-purple-600" />;
  if (t.includes('compare') || t.includes('diff')) return <GitCompare size={14} className="text-yellow-600" />;
  if (t.includes('restore')) return <RotateCcw size={14} className="text-red-600" />;
  return <ClipboardList size={14} className="text-gray-500" />;
}

function StatusPill({ status }) {
  const map = {
    Queued:     { cls: 'text-gray-700 bg-gray-100',  Icon: Clock,         label: 'Aguardando' },
    Running:    { cls: 'text-blue-700 bg-blue-50',   Icon: Loader2,       label: 'Em execução' },
    Completed:  { cls: 'text-green-700 bg-green-50', Icon: CheckCircle2,  label: 'Concluída' },
    Failed:     { cls: 'text-red-700 bg-red-50',     Icon: AlertCircle,   label: 'Falhou' },
    Information:{ cls: 'text-gray-700 bg-gray-100',  Icon: ClipboardList, label: 'Info' }
  };
  const c = map[status] || { cls: 'text-gray-600 bg-gray-100', Icon: ClipboardList, label: status || '—' };
  const Icon = c.Icon;
  return (
    <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-semibold ${c.cls}`}>
      <Icon size={11} className={status === 'Running' ? 'animate-spin' : ''} /> {c.label}
    </span>
  );
}

function TaskDetail({ task, onClose }) {
  const [job, setJob] = useState(null);
  const [events, setEvents] = useState([]);

  useEffect(() => {
    if (task?.id) {
      api.job(task.id).then((r) => setJob(r?.data)).catch(() => setJob(null));
    }
    api.events().then((r) => {
      const all = r?.data?.items ?? [];
      setEvents(all.filter((e) => task?.name && e.source && e.source.toLowerCase().includes(task.name.toLowerCase().substring(0, 6))).slice(0, 20));
    }).catch(() => {});
  }, [task]);

  if (!task) return null;
  const fmt = (d) => d ? new Date(d).toLocaleString('pt-BR') : '—';
  const data = job ?? task;

  return (
    <div className="fixed inset-y-0 right-0 w-[460px] bg-white border-l border-border-light shadow-xl z-40 overflow-y-auto">
      <div className="flex items-center justify-between px-5 py-4 border-b border-border-light">
        <div className="flex items-center gap-2">
          <TypeIcon type={data.kind || data.type} />
          <h3 className="text-base font-semibold text-[#222]">{data.kind || data.name || 'Task'}</h3>
        </div>
        <button onClick={onClose} className="text-gray-400 hover:text-gray-700"><X size={18} /></button>
      </div>

      <div className="p-5 space-y-5 text-sm">
        <div>
          <StatusPill status={data.status} />
        </div>

        <div className="border border-border-light rounded p-3">
          <h4 className="text-xs font-semibold text-text-secondary uppercase tracking-wide mb-2">Linha do tempo</h4>
          <dl className="space-y-1.5 text-xs">
            <div className="flex justify-between"><dt className="text-text-secondary">Criada:</dt><dd>{fmt(data.createdAt || data.time)}</dd></div>
            <div className="flex justify-between"><dt className="text-text-secondary">Iniciada:</dt><dd>{fmt(data.startedAt)}</dd></div>
            <div className="flex justify-between"><dt className="text-text-secondary">Finalizada:</dt><dd>{fmt(data.finishedAt)}</dd></div>
            {data.startedAt && data.finishedAt && (
              <div className="flex justify-between"><dt className="text-text-secondary">Duração:</dt>
                <dd>{Math.round((new Date(data.finishedAt) - new Date(data.startedAt)) / 1000)}s</dd></div>
            )}
          </dl>
        </div>

        {data.input && Object.keys(data.input).length > 0 && (
          <div>
            <h4 className="text-xs font-semibold text-text-secondary uppercase tracking-wide mb-2">Parâmetros</h4>
            <pre className="text-[10px] font-mono bg-gray-50 border border-border-light rounded p-2 overflow-auto max-h-32">
              {JSON.stringify(data.input, null, 2)}
            </pre>
          </div>
        )}

        {data.error && (
          <div>
            <h4 className="text-xs font-semibold text-red-700 uppercase tracking-wide mb-2">Erro</h4>
            <div className="bg-red-50 border border-red-200 rounded p-2 text-xs text-red-800">{data.error}</div>
            {data.details && (
              <pre className="mt-1 text-[10px] font-mono bg-red-50 border border-red-200 rounded p-2 overflow-auto max-h-24 text-red-800">
                {data.details}
              </pre>
            )}
          </div>
        )}

        {data.result && (
          <div>
            <h4 className="text-xs font-semibold text-text-secondary uppercase tracking-wide mb-2">Resultado</h4>
            <pre className="text-[10px] font-mono bg-gray-50 border border-border-light rounded p-2 overflow-auto max-h-40">
              {JSON.stringify(data.result, null, 2)}
            </pre>
          </div>
        )}

        {data.operation && (
          <div>
            <h4 className="text-xs font-semibold text-text-secondary uppercase tracking-wide mb-2">Última operação</h4>
            <div className="text-xs text-[#333]">{data.operation}</div>
          </div>
        )}

        {events.length > 0 && (
          <div>
            <h4 className="text-xs font-semibold text-text-secondary uppercase tracking-wide mb-2">Eventos relacionados ({events.length})</h4>
            <div className="space-y-1 max-h-48 overflow-y-auto">
              {events.map((e, idx) => (
                <div key={idx} className="text-[11px] border-l-2 border-border-light pl-2 py-1">
                  <div className="text-text-secondary">{new Date(e.time).toLocaleTimeString('pt-BR')} · {e.severity}</div>
                  <div className="text-[#333]">{e.message}</div>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

export default function Tasks() {
  const [rows, setRows] = useState([]);
  const [jobs, setJobs] = useState([]);
  const [selected, setSelected] = useState(null);
  const [filter, setFilter] = useState('All');
  const [loading, setLoading] = useState(true);

  const load = async () => {
    setLoading(true);
    try {
      const [tasksR, jobsR] = await Promise.all([api.tasks(), api.jobs()]);
      setRows(tasksR?.data?.items ?? []);
      setJobs(jobsR?.data ?? []);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    load();
    const id = setInterval(load, 5000);
    return () => clearInterval(id);
  }, []);

  // Mescla jobs (mais ricos) com tasks dos logs
  const allTasks = [
    ...jobs.map((j) => ({
      id: j.id,
      time: j.startedAt || j.createdAt,
      createdAt: j.createdAt,
      startedAt: j.startedAt,
      finishedAt: j.finishedAt,
      name: j.kind,
      kind: j.kind,
      type: j.kind,
      status: j.status,
      operation: j.error || `Job ${j.id?.substring(0, 8)}`,
      input: j.input,
      result: j.result,
      error: j.error,
      details: j.details,
      source: 'job'
    })),
    ...rows.map((r) => ({ ...r, source: 'log' }))
  ];

  const filtered = filter === 'All' ? allTasks : allTasks.filter((t) => (t.type || '').toLowerCase().includes(filter.toLowerCase()));

  const FilterTab = ({ value, label }) => (
    <button onClick={() => setFilter(value)}
      className={`px-3 py-1 text-xs rounded ${filter === value ? 'bg-[#0078A8] text-white' : 'bg-white border border-border-light text-[#333] hover:bg-gray-50'}`}>
      {label}
    </button>
  );

  return (
    <div className="p-6">
      <div className="flex items-center justify-between mb-4">
        <div>
          <h2 className="text-base font-semibold text-[#222]">Tasks</h2>
          <p className="text-xs text-text-secondary mt-0.5">Operações iniciadas (backup, unpack, compare, restore)</p>
        </div>
        <button onClick={load} className="text-xs text-[#0078A8] hover:underline flex items-center gap-1">
          {loading ? <Loader2 className="animate-spin" size={12} /> : <RefreshCw size={12} />} Atualizar
        </button>
      </div>

      <div className="flex flex-wrap gap-2 mb-3">
        <FilterTab value="All" label="Todas" />
        <FilterTab value="backup" label="Backup" />
        <FilterTab value="unpack" label="Unpack" />
        <FilterTab value="compare" label="Compare" />
        <FilterTab value="restore" label="Restore" />
      </div>

      <div className="bg-white border border-border-light rounded-lg overflow-hidden">
        {filtered.length === 0 ? (
          <div className="text-center py-12 text-text-secondary text-sm">Nenhuma task.</div>
        ) : (
          <table className="w-full text-sm">
            <thead className="bg-gray-50 text-xs text-text-secondary">
              <tr>
                <th className="text-left px-4 py-2 font-semibold">Tipo</th>
                <th className="text-left px-4 py-2 font-semibold">Nome</th>
                <th className="text-left px-4 py-2 font-semibold">Status</th>
                <th className="text-left px-4 py-2 font-semibold">Tempo</th>
                <th className="text-left px-4 py-2 font-semibold">Última operação</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((t, idx) => (
                <tr key={t.id || idx}
                  onClick={() => setSelected(t)}
                  className="border-t border-border-light hover:bg-gray-50 cursor-pointer">
                  <td className="px-4 py-2"><TypeIcon type={t.type || t.kind} /></td>
                  <td className="px-4 py-2 text-xs font-mono">{t.name || t.kind}</td>
                  <td className="px-4 py-2"><StatusPill status={t.status} /></td>
                  <td className="px-4 py-2 text-xs">{t.time ? new Date(t.time).toLocaleString('pt-BR') : '—'}</td>
                  <td className="px-4 py-2 text-xs text-text-secondary truncate max-w-md">{t.operation}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {selected && <TaskDetail task={selected} onClose={() => setSelected(null)} />}
    </div>
  );
}
