import { useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { ArrowLeft, Archive, RefreshCw, Loader2, Package, ChevronRight } from 'lucide-react';
import { api } from '../api/client.js';
import { useTenant } from '../context/TenantContext.jsx';

export default function UnpackBackup() {
  const navigate = useNavigate();
  const { selected } = useTenant();
  const [backups, setBackups] = useState([]);
  const [selectedId, setSelectedId] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [opts, setOpts] = useState({
    clearPrevious: false,
    computeDifferences: true,
    validateIntegrity: true
  });
  const [unpacking, setUnpacking] = useState(false);
  const [result, setResult] = useState(null);

  const load = async () => {
    setLoading(true);
    try {
      const r = await api.backups();
      const list = (r?.data ?? []).filter((b) =>
        !selected?.tenantId || !b.tenantId || b.tenantId === selected.tenantId);
      setBackups(list);
      if (list.length > 0 && !selectedId) setSelectedId(list[0].id);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); /* eslint-disable-line */ }, [selected?.tenantId]);

  const handleUnpack = async () => {
    if (!selectedId) return;
    setUnpacking(true);
    setError(null);
    try {
      const r = await api.unpackBackup(selectedId, opts);
      setResult({ jobId: r?.data?.id, backupId: selectedId });
    } catch (e) {
      setError(e.message);
    } finally {
      setUnpacking(false);
    }
  };

  const selectedBackup = backups.find((b) => b.id === selectedId);

  return (
    <div className="p-6 max-w-5xl">
      <button onClick={() => navigate('/dashboard')} className="text-xs text-[#0078A8] hover:underline flex items-center gap-1 mb-3">
        <ArrowLeft size={12} /> Voltar ao Dashboard
      </button>

      <div className="flex items-center justify-between mb-4">
        <div>
          <h2 className="text-base font-semibold text-[#222]">Backup Unpacking</h2>
          <p className="text-xs text-text-secondary mt-0.5">
            Descompacte um backup para pesquisar objetos, comparar diferenças e restaurar funções administrativas.
          </p>
        </div>
        <button onClick={load} className="text-xs text-[#0078A8] hover:underline flex items-center gap-1">
          <RefreshCw size={12} /> Atualizar
        </button>
      </div>

      <div className="bg-blue-50 border border-blue-200 rounded p-3 text-xs text-blue-900 mb-5">
        O backup fica guardado compactado. Para pesquisar, comparar ou restaurar funções, primeiro descompacte para uma área de trabalho temporária.
        Depois siga para <strong>Unpacked Objects</strong> ou <strong>Differences</strong>.
      </div>

      {error && <div className="bg-red-50 border border-red-200 text-red-800 text-sm px-3 py-2 rounded mb-3">{error}</div>}

      {loading ? (
        <div className="text-text-secondary text-sm flex items-center gap-2"><Loader2 className="animate-spin" size={14} /> Carregando...</div>
      ) : backups.length === 0 ? (
        <div className="text-center py-16 text-text-secondary">
          <Archive size={32} className="mx-auto mb-3 opacity-30" />
          <p className="text-sm">Nenhum backup disponível.</p>
          <p className="text-xs mt-1">Crie um backup pelo Dashboard antes de descompactar.</p>
        </div>
      ) : result ? (
        <div className="bg-green-50 border border-green-200 rounded-lg p-5">
          <div className="flex items-start gap-3">
            <Package className="text-green-700 mt-0.5" size={20} />
            <div className="flex-1">
              <h3 className="font-semibold text-green-900">Unpack task criada</h3>
              <p className="text-sm text-green-800 mt-1">
                Backup <code className="font-mono text-xs">{result.backupId}</code> está sendo descompactado.
              </p>
              <div className="mt-3 flex gap-2">
                <button onClick={() => navigate('/tasks')} className="px-3 py-1.5 text-xs bg-green-700 text-white rounded hover:bg-green-800">
                  Ver tarefa
                </button>
                <button onClick={() => navigate(`/unpacked?backupId=${result.backupId}`)}
                  className="px-3 py-1.5 text-xs border border-green-700 text-green-800 rounded hover:bg-green-100 flex items-center gap-1">
                  Ir para Unpacked Objects <ChevronRight size={12} />
                </button>
                <button onClick={() => { setResult(null); load(); }} className="px-3 py-1.5 text-xs text-green-800 hover:underline">
                  Descompactar outro
                </button>
              </div>
            </div>
          </div>
        </div>
      ) : (
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
          <div className="lg:col-span-2 bg-white border border-border-light rounded-lg overflow-hidden">
            <div className="px-4 py-2 border-b border-border-light bg-gray-50 text-xs font-semibold text-text-secondary">
              Selecione o backup
            </div>
            <div className="max-h-[480px] overflow-y-auto">
              {backups.map((b) => {
                const sel = b.id === selectedId;
                return (
                  <button key={b.id} onClick={() => setSelectedId(b.id)}
                    className={`w-full text-left px-4 py-3 border-b border-border-light flex items-center justify-between hover:bg-gray-50 ${sel ? 'bg-blue-50' : ''}`}>
                    <div>
                      <div className="font-mono text-xs text-[#222]">{b.id}</div>
                      <div className="text-[11px] text-text-secondary mt-0.5">
                        {b.modifiedAt ? new Date(b.modifiedAt).toLocaleString('pt-BR') : '—'}
                        {b.roleDefinitionsCount != null && ` · ${b.roleDefinitionsCount} roles`}
                        {b.roleAssignmentsCount != null && ` · ${b.roleAssignmentsCount} assignments`}
                      </div>
                    </div>
                    <div className="text-xs text-text-secondary">{b.sizeDisplay}</div>
                  </button>
                );
              })}
            </div>
          </div>

          <div className="bg-white border border-border-light rounded-lg p-4">
            <h3 className="text-sm font-semibold text-[#222] mb-3">Opções de unpacking</h3>
            <div className="space-y-3 mb-4">
              <label className="flex items-start gap-2 text-sm cursor-pointer">
                <input type="checkbox" checked={opts.clearPrevious}
                  onChange={(e) => setOpts({ ...opts, clearPrevious: e.target.checked })} className="mt-0.5" />
                <div>
                  <div>Limpar objetos descompactados anteriormente</div>
                  <div className="text-xs text-text-secondary">Remove a área de trabalho atual antes de extrair.</div>
                </div>
              </label>
              <label className="flex items-start gap-2 text-sm cursor-pointer">
                <input type="checkbox" checked={opts.computeDifferences}
                  onChange={(e) => setOpts({ ...opts, computeDifferences: e.target.checked })} className="mt-0.5" />
                <div>
                  <div>Calcular diferenças durante unpack</div>
                  <div className="text-xs text-text-secondary">Compara já com o tenant atual e popula a aba Differences.</div>
                </div>
              </label>
              <label className="flex items-start gap-2 text-sm cursor-pointer">
                <input type="checkbox" checked={opts.validateIntegrity}
                  onChange={(e) => setOpts({ ...opts, validateIntegrity: e.target.checked })} className="mt-0.5" />
                <div>
                  <div>Validar integridade (SHA256)</div>
                  <div className="text-xs text-text-secondary">Verifica hash do backup antes de descompactar.</div>
                </div>
              </label>
            </div>

            {selectedBackup && (
              <div className="text-xs bg-gray-50 rounded p-2 mb-3 space-y-0.5">
                <div className="flex justify-between"><span className="text-text-secondary">ID:</span>
                  <span className="font-mono">{selectedBackup.id}</span></div>
                <div className="flex justify-between"><span className="text-text-secondary">Tamanho:</span>
                  <span>{selectedBackup.sizeDisplay}</span></div>
                <div className="flex justify-between"><span className="text-text-secondary">Roles:</span>
                  <span>{selectedBackup.roleDefinitionsCount ?? '—'}</span></div>
              </div>
            )}

            <button onClick={handleUnpack} disabled={unpacking || !selectedId}
              className="w-full px-3 py-2 bg-[#0078A8] text-white rounded text-sm hover:bg-[#006090] disabled:opacity-50 flex items-center justify-center gap-2">
              {unpacking ? <Loader2 size={14} className="animate-spin" /> : <Archive size={14} />}
              Unpack
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
