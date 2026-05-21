import axios from 'axios';

const client = axios.create({
  baseURL: '/api',
  timeout: 30000,
  headers: { 'Content-Type': 'application/json' }
});

client.interceptors.response.use(
  (resp) => resp,
  (err) => {
    const message = err.response?.data?.error || err.message || 'Erro desconhecido';
    return Promise.reject(new Error(message));
  }
);

export const api = {
  health: () => client.get('/health').then((r) => r.data),
  settings: () => client.get('/settings').then((r) => r.data),
  tenants: () => client.get('/tenants').then((r) => r.data),
  addTenant: (body) => client.post('/tenants', body).then((r) => r.data),
  removeTenant: (tenantId) => client.delete(`/tenants/${tenantId}`).then((r) => r.data),
  tenantStatus: (tenantId) => client.get(`/tenants/${tenantId}/status`).then((r) => r.data),
  tenantConsents: (tenantId) => client.get(`/tenants/${tenantId}/consents`).then((r) => r.data),
  setupState: () => client.get('/setup/state').then((r) => r.data),
  setupReset: () => client.delete('/setup/reset').then((r) => r.data),
  setupStart: () => client.post('/setup/start').then((r) => r.data),
  setupStatus: (sessionId) => client.get(`/setup/status/${sessionId}`).then((r) => r.data),
  oauthStart: (tenantHint) => client.get('/oauth/start', { params: tenantHint ? { tenantHint } : {} }).then((r) => r.data),
  backups: () => client.get('/backups').then((r) => r.data),
  backup: (id) => client.get(`/backups/${id}`).then((r) => r.data),
  runBackup: () => client.post('/backups/run').then((r) => r.data),
  unpacked: (backupId) => client.get('/unpacked-objects', { params: { backupId } }).then((r) => r.data),
  runCompare: (backupId) => client.post('/compare/run', { backupId }).then((r) => r.data),
  compareResults: () => client.get('/compare/results').then((r) => r.data),
  previewRestore: (payload) => client.post('/restore/preview', payload).then((r) => r.data),
  applyRestore: (payload) => client.post('/restore/apply', payload).then((r) => r.data),
  tasks: (since) => client.get('/tasks', { params: since ? { since } : {} }).then((r) => r.data),
  clearTasks: () => client.delete('/tasks').then((r) => r.data),
  events: (severity, since) => client.get('/events', { params: { ...(severity ? { severity } : {}), ...(since ? { since } : {}) } }).then((r) => r.data),
  clearEvents: () => client.delete('/events').then((r) => r.data),
  job: (id) => client.get(`/jobs/${id}`).then((r) => r.data),
  jobs: () => client.get('/jobs').then((r) => r.data),
  unpackBackup: (backupId, opts) => client.post(`/backups/${backupId}/unpack`, opts).then((r) => r.data),
  backupSettings: (tenantId) => client.get(`/tenants/${tenantId}/backup-settings`).then((r) => r.data),
  updateBackupSettings: (tenantId, data) => client.put(`/tenants/${tenantId}/backup-settings`, data).then((r) => r.data),
};

export async function pollJob(id, { intervalMs = 2000, timeoutMs = 600000, onTick } = {}) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    const resp = await api.job(id);
    const job = resp?.data;
    if (onTick) onTick(job);
    if (job?.status === 'Completed' || job?.status === 'Failed') {
      return job;
    }
    await new Promise((r) => setTimeout(r, intervalMs));
  }
  throw new Error(`Job ${id} excedeu timeout de polling`);
}
