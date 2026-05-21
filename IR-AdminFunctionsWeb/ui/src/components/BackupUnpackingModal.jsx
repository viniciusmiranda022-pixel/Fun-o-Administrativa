import { useState } from 'react';
import { CheckCircle2 } from 'lucide-react';
import Modal from './common/Modal.jsx';
import { api, pollJob } from '../api/client.js';

export default function BackupUnpackingModal({ onClose, backups, preSelectedId = null }) {
  const [opts, setOpts] = useState({ clear: true, diff: true, validate: false });
  const [selected, setSelected] = useState(preSelectedId ?? backups[0]?.id ?? null);
  const [state, setState] = useState('idle');
  const [msg, setMsg] = useState('');

  function toggle(k) { setOpts((o) => ({ ...o, [k]: !o[k] })); }

  function formatDate(snap) {
    if (!snap) return '—';
    const d = new Date(snap.createdAt || snap.collectedAt || '');
    if (isNaN(d)) return snap.id;
    return d.toLocaleString();
  }

  async function handleUnpack() {
    if (!selected) return;
    setState('running');
    setMsg('Unpacking...');
    try {
      const resp = await api.unpackBackup(selected, {
        clearPrevious: opts.clear,
        runDiff: opts.diff,
        validate: opts.validate,
      });
      const job = resp?.data;
      if (!job?.id) throw new Error('No job id returned');
      const finished = await pollJob(job.id, {
        intervalMs: 2000,
        onTick: (j) => setMsg(`Unpacking... status: ${j?.status ?? '?'}`),
      });
      if (finished?.status === 'Completed') {
        setState('done');
        setTimeout(onClose, 1500);
      } else {
        throw new Error(finished?.error || 'Unpack failed');
      }
    } catch (err) {
      setState('error');
      setMsg(err.message);
    }
  }

  return (
    <Modal title="Administrative Function Backup Unpacking" onClose={onClose} maxWidth="max-w-2xl">
      <div className="px-6 py-4 space-y-4">
        {state === 'done' ? (
          <div className="flex items-center gap-2 text-[#28A745] text-sm font-semibold">
            <CheckCircle2 size={18} /> Unpack completed!
          </div>
        ) : state === 'error' ? (
          <div className="text-[#DC3545] text-sm">{msg}</div>
        ) : state === 'running' ? (
          <div className="text-sm text-[#555]">{msg}</div>
        ) : (
          <>
            <div className="space-y-2">
              {[
                ['clear', 'Clear administrative function objects from previously unpacked backups'],
                ['diff', 'Perform administrative function differences operation during unpack'],
                ['validate', 'Validate backup integrity before unpack'],
              ].map(([k, label]) => (
                <label key={k} className="flex items-center gap-2 text-xs text-[#333] cursor-pointer">
                  <input type="checkbox" checked={opts[k]} onChange={() => toggle(k)} />
                  {label}
                </label>
              ))}
            </div>
            <div>
              <p className="text-xs font-semibold text-[#333] mb-2">Select a backup to unpack:</p>
              {backups.length === 0 ? (
                <p className="text-xs text-[#888]">No backups available. Create a backup first.</p>
              ) : (
                <table className="w-full text-xs border border-[#DEE2E6]">
                  <thead>
                    <tr className="bg-[#F2F2F2]">
                      <th className="w-8 px-2 py-1.5 border-b border-[#DEE2E6]"></th>
                      <th className="text-left px-3 py-1.5 border-b border-[#DEE2E6] font-semibold text-[#333]">Created</th>
                      <th className="text-left px-3 py-1.5 border-b border-[#DEE2E6] font-semibold text-[#333]">Roles</th>
                      <th className="text-left px-3 py-1.5 border-b border-[#DEE2E6] font-semibold text-[#333]">Size</th>
                    </tr>
                  </thead>
                  <tbody>
                    {backups.map((b) => (
                      <tr
                        key={b.id}
                        className={`border-b border-[#DEE2E6] cursor-pointer ${selected === b.id ? 'bg-[#E6F4FA]' : 'hover:bg-[#F9FAFB]'}`}
                        onClick={() => setSelected(b.id)}
                      >
                        <td className="px-2 py-1.5 text-center">
                          <input type="radio" checked={selected === b.id} onChange={() => setSelected(b.id)} />
                        </td>
                        <td className="px-3 py-1.5">{formatDate(b)}</td>
                        <td className="px-3 py-1.5">{b.roleDefinitionsCount ?? '—'}</td>
                        <td className="px-3 py-1.5">{b.sizeDisplay ?? '—'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
            <div className="flex justify-end gap-2 pt-2 border-t border-[#DEE2E6]">
              <button onClick={onClose} className="px-4 py-1.5 text-xs border border-[#DEE2E6] hover:bg-gray-50">Cancel</button>
              <button
                disabled={!selected}
                onClick={handleUnpack}
                className="px-4 py-1.5 text-xs bg-[#0078A8] text-white hover:bg-[#006090] disabled:opacity-40"
              >
                Unpack
              </button>
            </div>
          </>
        )}
      </div>
    </Modal>
  );
}
