import { useEffect, useRef, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { Loader2, CheckCircle2, AlertCircle, ExternalLink, Copy, ShieldCheck } from 'lucide-react';
import { api } from '../api/client.js';

export default function Setup() {
  const navigate = useNavigate();
  const [state, setState]     = useState(null);
  const [session, setSession] = useState(null);
  const [error, setError]     = useState(null);
  const [starting, setStarting] = useState(false);
  const [copied, setCopied]   = useState(false);
  const popupRef              = useRef(null);
  const popupOpenedRef        = useRef(false);

  useEffect(() => {
    api.setupState().then((r) => setState(r.data)).catch((e) => setError(e.message));
  }, []);

  /* Poll session status */
  useEffect(() => {
    if (!session?.sessionId) return;
    const timer = setInterval(async () => {
      try {
        const r = await api.setupStatus(session.sessionId);
        const s = r.data;
        setSession((prev) => ({ ...prev, ...s }));
        if (s.status === 'Completed') {
          clearInterval(timer);
          popupRef.current?.close();
          setTimeout(() => navigate('/tenants'), 2000);
        }
        if (s.status === 'Failed') {
          clearInterval(timer);
          popupRef.current?.close();
        }
      } catch { /* keep polling */ }
    }, 2000);
    return () => clearInterval(timer);
  }, [session?.sessionId, navigate]);

  /* Auto-open popup as soon as user code is available */
  useEffect(() => {
    if (!session?.userCode || popupOpenedRef.current) return;
    popupOpenedRef.current = true;

    /* Copy code to clipboard automatically */
    navigator.clipboard.writeText(session.userCode).catch(() => {});

    /* Build URL with OTC pre-filled so user only needs to log in */
    const otc = session.userCode.replace(/-/g, '');
    const url = `https://microsoft.com/devicelogin?otc=${otc}`;
    const w = 520, h = 680;
    const left = Math.round(window.screenX + (window.outerWidth - w) / 2);
    const top  = Math.round(window.screenY + (window.outerHeight - h) / 2);
    const popup = window.open(url, 'ms-device-login',
      `width=${w},height=${h},left=${left},top=${top},toolbar=no,menubar=no,scrollbars=yes`);
    popupRef.current = popup;
  }, [session?.userCode]);

  const startSetup = async () => {
    setStarting(true);
    setError(null);
    popupOpenedRef.current = false;
    try {
      const r = await api.setupStart();
      setSession({ sessionId: r.data.sessionId, ...r.data });
    } catch (e) {
      setError(e.message);
    } finally {
      setStarting(false);
    }
  };

  const copyCode = () => {
    if (session?.userCode) {
      navigator.clipboard.writeText(session.userCode);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    }
  };

  const reopenPopup = () => {
    if (!session?.userCode) return;
    const otc = session.userCode.replace(/-/g, '');
    const url = `https://microsoft.com/devicelogin?otc=${otc}`;
    const w = 520, h = 680;
    const left = Math.round(window.screenX + (window.outerWidth - w) / 2);
    const top  = Math.round(window.screenY + (window.outerHeight - h) / 2);
    const popup = window.open(url, 'ms-device-login',
      `width=${w},height=${h},left=${left},top=${top},toolbar=no,menubar=no,scrollbars=yes`);
    popupRef.current = popup;
  };

  if (!state) {
    return (
      <div className="p-8 flex items-center gap-2 text-text-secondary text-sm">
        <Loader2 className="animate-spin" size={16} /> Carregando…
      </div>
    );
  }

  /* Already configured */
  if (state.isConfigured && !session) {
    return (
      <div className="p-8 max-w-xl mx-auto">
        <div className="bg-white border border-border-light rounded p-6">
          <div className="flex items-center gap-2 mb-4">
            <ShieldCheck size={20} className="text-green-600" />
            <h2 className="text-sm font-semibold text-[#222]">Aplicação já configurada</h2>
          </div>
          <dl className="text-xs space-y-2 mb-5">
            <div className="flex gap-2"><dt className="text-text-secondary w-36">Nome:</dt><dd className="text-[#333] font-medium">{state.displayName}</dd></div>
            <div className="flex gap-2"><dt className="text-text-secondary w-36">Client ID:</dt><dd className="font-mono">{state.clientId}</dd></div>
            <div className="flex gap-2"><dt className="text-text-secondary w-36">Tenant bootstrap:</dt><dd className="font-mono">{state.bootstrapTenantId}</dd></div>
            <div className="flex gap-2"><dt className="text-text-secondary w-36">Criado em:</dt><dd>{state.createdAt ? new Date(state.createdAt).toLocaleString('pt-BR') : '—'}</dd></div>
          </dl>
          <button onClick={() => navigate('/tenants')}
            className="px-4 py-1.5 bg-[#0078A8] text-white text-sm rounded-sm hover:bg-[#006090]">
            Ir para Tenants
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="p-8 max-w-xl mx-auto">
      <div className="bg-white border border-border-light rounded p-6">
        <h2 className="text-sm font-semibold text-[#222] mb-1">Initial Setup</h2>
        <p className="text-xs text-text-secondary mb-5">
          Antes de adicionar tenants, é necessário registrar esta aplicação como App Multi-Tenant
          no Microsoft Entra ID. O processo é feito uma única vez por um Global Administrator.
        </p>

        {/* Not started */}
        {!session && (
          <>
            {error && (
              <div className="bg-red-50 border border-red-200 text-red-800 text-xs px-3 py-2 rounded mb-4">{error}</div>
            )}
            <button onClick={startSetup} disabled={starting}
              className="px-4 py-2 bg-[#0078A8] text-white text-sm rounded-sm hover:bg-[#006090] disabled:opacity-60 flex items-center gap-2">
              {starting && <Loader2 size={14} className="animate-spin" />}
              Iniciar Setup
            </button>
          </>
        )}

        {/* Waiting for code */}
        {session && !session.userCode && session.status !== 'Failed' && session.status !== 'Completed' && (
          <div className="flex items-center gap-2 text-xs text-text-secondary py-4">
            <Loader2 size={13} className="animate-spin" /> Solicitando autenticação à Microsoft…
          </div>
        )}

        {/* Code received — popup opened automatically */}
        {session?.userCode && session.status !== 'Completed' && session.status !== 'Failed' && (
          <div className="space-y-4">
            <div className="bg-blue-50 border border-blue-200 rounded p-4">
              <div className="flex items-center gap-2 text-blue-800 font-semibold text-sm mb-2">
                <Loader2 size={14} className="animate-spin text-blue-600" />
                Aguardando aprovação no popup da Microsoft…
              </div>
              <p className="text-xs text-blue-700">
                Uma janela de login da Microsoft foi aberta automaticamente. Faça login com uma conta
                <strong> Global Administrator</strong> e aprove as permissões.
              </p>
            </div>

            {/* Code display as fallback */}
            <div className="border border-border-light rounded p-4 bg-gray-50">
              <div className="text-[10px] text-text-secondary uppercase tracking-wide mb-1">Código de autenticação</div>
              <div className="flex items-center gap-3">
                <span className="text-2xl font-mono font-bold tracking-widest text-[#0078A8]">
                  {session.userCode}
                </span>
                <button onClick={copyCode}
                  className="text-xs text-[#0078A8] hover:underline inline-flex items-center gap-1">
                  <Copy size={11} /> {copied ? 'Copiado!' : 'Copiar'}
                </button>
              </div>
              <p className="text-[10px] text-text-secondary mt-2">
                Se o popup foi bloqueado,{' '}
                <button onClick={reopenPopup} className="text-[#0078A8] hover:underline inline-flex items-center gap-0.5">
                  clique aqui para abrir novamente <ExternalLink size={10} />
                </button>
              </p>
            </div>

            {session.status === 'Registering' && (
              <div className="flex items-center gap-2 text-xs text-green-700">
                <Loader2 size={13} className="animate-spin" />
                Autenticado. Registrando aplicação no Azure AD…
              </div>
            )}
          </div>
        )}

        {/* Success */}
        {session?.status === 'Completed' && (
          <div className="bg-green-50 border border-green-200 rounded p-4">
            <div className="flex items-center gap-2 text-green-800 font-semibold text-sm mb-1">
              <CheckCircle2 size={16} /> Setup concluído com sucesso!
            </div>
            <p className="text-xs text-green-700">Client ID: <code className="font-mono">{session.clientId}</code></p>
            <p className="text-xs text-green-600 mt-2">Redirecionando para a página de Tenants…</p>
          </div>
        )}

        {/* Failed */}
        {session?.status === 'Failed' && (
          <div className="bg-red-50 border border-red-200 rounded p-4">
            <div className="flex items-center gap-2 text-red-800 font-semibold text-sm mb-2">
              <AlertCircle size={16} /> Setup falhou
            </div>
            <p className="text-xs text-red-700">{session.message}</p>
            <button
              onClick={() => { setSession(null); setError(null); popupOpenedRef.current = false; }}
              className="mt-3 text-xs text-[#0078A8] hover:underline">
              Tentar novamente
            </button>
          </div>
        )}
      </div>
    </div>
  );
}
