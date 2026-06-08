import axios from 'axios';
import type {
  Job,
  ApiResponse,
  TenantEntry,
  BackupSnapshot,
  TaskEntry,
  EventEntry,
  AppConfig,
  SetupStartResult,
  SetupStatus,
  TenantConsentsResponse,
  UnpackedObjects,
  CompareResults,
} from '../types/api';

const client = axios.create({
  baseURL: '/api',
  timeout: 30000,
  headers: { 'Content-Type': 'application/json' }
});

client.interceptors.response.use(
  (resp) => resp,
  (err: unknown) => {
    const anyErr = err as { response?: { data?: { error?: string } }; message?: string };
    const message = anyErr.response?.data?.error || anyErr.message || 'Unknown error';
    return Promise.reject(new Error(message));
  }
);

interface UnpackOptions {
  clearPrevious: boolean;
  runDiff: boolean;
  validate: boolean;
}

interface RestorePayload {
  backupId: string;
  roleName: string;
  currentRoleName?: string;
  removeExtraAssignments?: boolean;
}

type SetupStateResponse = AppConfig & { isConfigured: boolean };

export const api = {
  health: (): Promise<ApiResponse<unknown>> => client.get('/health').then((r) => r.data),
  settings: (): Promise<ApiResponse<unknown>> => client.get('/settings').then((r) => r.data),
  tenants: (): Promise<ApiResponse<TenantEntry[]>> => client.get('/tenants').then((r) => r.data),
  addTenant: (body: Record<string, unknown>): Promise<ApiResponse<TenantEntry>> => client.post('/tenants', body).then((r) => r.data),
  removeTenant: (tenantId: string): Promise<ApiResponse<unknown>> => client.delete(`/tenants/${tenantId}`).then((r) => r.data),
  tenantStatus: (tenantId: string): Promise<ApiResponse<unknown>> => client.get(`/tenants/${tenantId}/status`).then((r) => r.data),
  tenantConsents: (tenantId: string): Promise<ApiResponse<TenantConsentsResponse>> => client.get(`/tenants/${tenantId}/consents`).then((r) => r.data),
  updateTenantThumbprint: (tenantId: string, thumbprint: string): Promise<ApiResponse<unknown>> => client.put(`/tenants/${tenantId}/thumbprint`, { thumbprint }).then((r) => r.data),
  setupState: (): Promise<ApiResponse<SetupStateResponse>> => client.get('/setup/state').then((r) => r.data),
  setupReset: (): Promise<ApiResponse<unknown>> => client.delete('/setup/reset').then((r) => r.data),
  setupStart: (): Promise<ApiResponse<SetupStartResult>> => client.post('/setup/start').then((r) => r.data),
  setupStatus: (sessionId: string): Promise<ApiResponse<SetupStatus>> => client.get(`/setup/status/${sessionId}`).then((r) => r.data),
  oauthStart: (tenantHint?: string): Promise<ApiResponse<{ url: string }>> => client.get('/oauth/start', { params: tenantHint ? { tenantHint } : {} }).then((r) => r.data),
  backups: (): Promise<ApiResponse<BackupSnapshot[]>> => client.get('/backups').then((r) => r.data),
  backup: (id: string): Promise<ApiResponse<BackupSnapshot>> => client.get(`/backups/${id}`).then((r) => r.data),
  runBackup: (): Promise<ApiResponse<Job>> => client.post('/backups/run').then((r) => r.data),
  unpacked: (backupId: string): Promise<ApiResponse<UnpackedObjects>> => client.get('/unpacked-objects', { params: { backupId } }).then((r) => r.data),
  runCompare: (backupId?: string): Promise<ApiResponse<Job>> => client.post('/compare/run', { backupId }).then((r) => r.data),
  compareResults: (): Promise<ApiResponse<CompareResults>> => client.get('/compare/results').then((r) => r.data),
  previewRestore: (payload: RestorePayload): Promise<ApiResponse<unknown>> => client.post('/restore/preview', payload).then((r) => r.data),
  applyRestore: (payload: RestorePayload): Promise<ApiResponse<Job>> => client.post('/restore/apply', payload).then((r) => r.data),
  tasks: (since?: string): Promise<ApiResponse<{ items: TaskEntry[] }>> => client.get('/tasks', { params: since ? { since } : {} }).then((r) => r.data),
  clearTasks: (): Promise<ApiResponse<unknown>> => client.delete('/tasks').then((r) => r.data),
  events: (severity?: string, since?: string): Promise<ApiResponse<{ items: EventEntry[] }>> => client.get('/events', { params: { ...(severity ? { severity } : {}), ...(since ? { since } : {}) } }).then((r) => r.data),
  clearEvents: (): Promise<ApiResponse<unknown>> => client.delete('/events').then((r) => r.data),
  job: (id: string): Promise<ApiResponse<Job>> => client.get(`/jobs/${id}`).then((r) => r.data),
  jobs: (): Promise<ApiResponse<Job[]>> => client.get('/jobs').then((r) => r.data),
  unpackBackup: (backupId: string, opts: UnpackOptions): Promise<ApiResponse<Job>> => client.post(`/backups/${backupId}/unpack`, opts).then((r) => r.data),
  backupSettings: (tenantId: string): Promise<ApiResponse<unknown>> => client.get(`/tenants/${tenantId}/backup-settings`).then((r) => r.data),
  updateBackupSettings: (tenantId: string, data: Record<string, unknown>): Promise<ApiResponse<unknown>> => client.put(`/tenants/${tenantId}/backup-settings`, data).then((r) => r.data),
};

interface PollJobOptions {
  intervalMs?: number;
  timeoutMs?: number;
  onTick?: (job: Job | undefined) => void;
}

export async function pollJob(id: string, { intervalMs = 5000, timeoutMs = 600000, onTick }: PollJobOptions = {}): Promise<Job> {
  const start = Date.now();
  let currentInterval = intervalMs;
  while (Date.now() - start < timeoutMs) {
    const resp = await api.job(id);
    const job = resp?.data;
    if (onTick) onTick(job);
    if (job?.status === 'Completed' || job?.status === 'Failed') {
      return job;
    }
    await new Promise((r) => setTimeout(r, currentInterval));
    // Exponential backoff: double interval up to 30s for long-running jobs
    currentInterval = Math.min(currentInterval * 1.5, 30000);
  }
  throw new Error(`Job ${id} exceeded polling timeout`);
}
