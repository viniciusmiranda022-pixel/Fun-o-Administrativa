import { useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  ArrowLeft, RefreshCw, CheckCircle2, AlertCircle, ShieldCheck, ShieldAlert,
  ExternalLink, Eye, Loader2
} from 'lucide-react';
import { api } from '../api/client.js';

function ConsentBadge({ granted }) {
  return granted ? (
    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-semibold text-green-700 bg-green-50">
      <CheckCircle2 size={11} /> Granted
    </span>
  ) : (
    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-semibold text-yellow-700 bg-yellow-50">
      <AlertCircle size={11} /> Not Granted
    </span>
  );
}

function DetailsPanel({ tenant, onClose }) {
  const basic = tenant.consents?.basic;
  const restore = tenant.consents?.restore;
  const fmt = (d) => d ? new Date(d).toLocaleString('pt-BR') : '—';

  return (
    <div className="fixed inset-y-0 right-0 w-[420px] bg-white border-l border-border-light shadow-xl z-40 overflow-y-auto">
      <div className="flex items-center justify-between px-5 py-4 border-b border-border-light">
        <h3 className="text-base font-semibold text-[#222]">Consent Details</h3>
        <button onClick={onClose} className="text-gray-400 hover:text-gray-700 text-xl">×</button>
      </div>
      <div className="p-5 space-y-5 text-sm">
        <div>
          <div className="text-xs text-text-secondary uppercase tracking-wide">Tenant</div>
          <div className="font-semibold">{tenant.name || tenant.tenantId}</div>
          {tenant.domain && <div className="text-xs text-text-secondary">{tenant.domain}</div>}
        </div>

        <div className="border border-border-light rounded p-3">
          <div className="flex items-center justify-between mb-2">
            <div className="font-semibold text-[#222]">Basic (Backup)</div>
            <ConsentBadge granted={basic?.granted} />
          </div>
          <p className="text-xs text-text-secondary mb-2">
            Permissão de leitura. Necessária para backup, comparação e relatórios.
          </p>
          <dl className="text-xs space-y-1">
            <div className="flex justify-between"><dt className="text-text-secondary">Permission:</dt>
              <dd className="font-mono">RoleManagement.Read.Directory</dd></div>
            <div className="flex justify-between"><dt className="text-text-secondary">Concedido em:</dt>
              <dd>{fmt(basic?.grantedAt)}</dd></div>
            <div className="flex justify-between"><dt className="text-text-secondary">Por:</dt>
              <dd>{basic?.grantedByUser ?? '—'}</dd></div>
          </dl>
        </div>

        <div className="border border-orange-200 rounded p-3 bg-orange-50/40">
          <div className="flex items-center justify-between mb-2">
            <div className="font-semibold text-[#222] flex items-center gap-2">
              Restore
              <span className="text-[10px] uppercase font-semibold text-orange-700 bg-orange-100 px-2 py-0.5 rounded">Sensível</span>
            </div>
            <ConsentBadge granted={restore?.granted} />
          </div>
          <p className="text-xs text-text-secondary mb-2">
            Permissão de escrita. Necessária para restaurar funções e atribuições. Altera RBAC do Entra ID.
          </p>
          <dl className="text-xs space-y-1">
            <div className="flex justify-between"><dt className="text-text-secondary">Permission:</dt>
              <dd className="font-mono">RoleManagement.ReadWrite.Directory</dd></div>
            <div className="flex justify-between"><dt className="text-text-secondary">Concedido em:</dt>
              <dd>{fmt(restore?.grantedAt)}</dd></div>
            <div className="flex justify-between"><dt className="text-text-secondary">Por:</dt>
              <dd>{restore?.grantedByUser ?? '—'}</dd></div>
          </dl>
        </div>
      </div>
    </div>
  );
}

export default function ManageRestores() {
  const navigate = useNavigate();
  const [tenants, setTenants] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [detail, setDetail] = useState(null);
  const [granting, setGranting] = useState(null);

  const load = async () => {
    setLoading(true);
    setError(null);
    try {
      const r = await api.tenants();
      setTenants(r?.data ?? []);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    load();
    const onMsg = (ev) => { if (ev?.data?.type === 'oauth-complete') load(); };
    window.addEventListener('message', onMsg);
    return () => window.removeEventListener('message', onMsg);
  }, []);

  const startGrant = async (tenantId) => {
    setGranting(tenantId);
    try {
      const r = await api.oauthStart(tenantId);
      const w = window.open(r.data.url, 'oauth-consent', 'width=700,height=800');
      if (!w) window.location.href = r.data.url;
    } catch (e) {
      setError(e.message);
    } finally {
      setGranting(null);
    }
  };

  const overallStatus = (t) => {
    const b = t.consents?.basic?.granted;
    const r = t.consents?.restore?.granted;
    if (b && r) return { label: 'Fully authorized', color: 'text-green-700', Icon: ShieldCheck };
    if (b && !r) return { label: 'Backup only', color: 'text-yellow-700', Icon: ShieldAlert };
    return { label: 'Not authorized', color: 'text-red-700', Icon: ShieldAlert };
  };

  return (
    <div className="p-6">
      <button onClick={() => navigate('/dashboard')} className="text-xs text-[#0078A8] hover:underline flex items-center gap-1 mb-3">
        <ArrowLeft size={12} /> Voltar ao Dashboard
      </button>

      <div className="flex items-center justify-between mb-4">
        <div>
          <h2 className="text-base font-semibold text-[#222]">Manage Restores</h2>
          <p className="text-xs text-text-secondary mt-0.5">
            Status de consentimento administrativo por tenant. <strong>Basic</strong> = leitura/backup. <strong>Restore</strong> = escrita/restauração.
          </p>
        </div>
        <button onClick={load} className="text-xs text-[#0078A8] hover:underline flex items-center gap-1">
          <RefreshCw size={12} /> Atualizar
        </button>
      </div>

      <div className="bg-blue-50 border border-blue-200 rounded p-3 text-xs text-blue-900 mb-4">
        Backup usa apenas a permissão Basic (read-only). Restore exige a permissão Restore (read-write), que altera RBAC.
        Conceda apenas quando recuperação de funções administrativas for necessária.
      </div>

      {error && <div className="bg-red-50 border border-red-200 text-red-800 text-sm px-3 py-2 rounded mb-3">{error}</div>}

      {loading ? (
        <div className="text-text-secondary text-sm flex items-center gap-2"><Loader2 className="animate-spin" size={14} /> Carregando...</div>
      ) : tenants.length === 0 ? (
        <div className="text-center py-16 text-text-secondary">
          <ShieldAlert size={32} className="mx-auto mb-3 opacity-30" />
          <p className="text-sm">Nenhum tenant configurado.</p>
        </div>
      ) : (
        <div className="bg-white border border-border-light rounded-lg overflow-hidden">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 text-xs text-text-secondary">
              <tr>
                <th className="text-left px-4 py-2 font-semibold">Tenant</th>
                <th className="text-left px-4 py-2 font-semibold">Basic (Backup)</th>
                <th className="text-left px-4 py-2 font-semibold">Restore</th>
                <th className="text-left px-4 py-2 font-semibold">Overall</th>
                <th className="text-right px-4 py-2 font-semibold">Ações</th>
              </tr>
            </thead>
            <tbody>
              {tenants.map((t) => {
                const o = overallStatus(t);
                const OIcon = o.Icon;
                return (
                  <tr key={t.tenantId} className="border-t border-border-light hover:bg-gray-50">
                    <td className="px-4 py-3">
                      <div className="font-semibold text-[#222]">{t.name || t.tenantId}</div>
                      {t.domain && <div className="text-[11px] text-text-secondary">{t.domain}</div>}
                    </td>
                    <td className="px-4 py-3"><ConsentBadge granted={t.consents?.basic?.granted} /></td>
                    <td className="px-4 py-3"><ConsentBadge granted={t.consents?.restore?.granted} /></td>
                    <td className="px-4 py-3">
                      <span className={`flex items-center gap-1 text-xs font-semibold ${o.color}`}>
                        <OIcon size={12} /> {o.label}
                      </span>
                    </td>
                    <td className="px-4 py-3 text-right">
                      <div className="flex justify-end gap-3">
                        <button onClick={() => setDetail(t)}
                          className="text-xs text-[#0078A8] hover:underline flex items-center gap-1">
                          <Eye size={12} /> View Details
                        </button>
                        <button onClick={() => startGrant(t.tenantId)}
                          disabled={granting === t.tenantId}
                          className="text-xs text-[#0078A8] hover:underline flex items-center gap-1 disabled:opacity-50">
                          {granting === t.tenantId ? <Loader2 size={12} className="animate-spin" /> : <ExternalLink size={12} />}
                          {(t.consents?.basic?.granted || t.consents?.restore?.granted) ? 'Re-Grant' : 'Grant Consent'}
                        </button>
                        <button onClick={() => navigate(`/tenants/${t.tenantId}/consents`)}
                          className="text-xs text-[#0078A8] hover:underline">
                          Edit
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {detail && <DetailsPanel tenant={detail} onClose={() => setDetail(null)} />}
    </div>
  );
}
