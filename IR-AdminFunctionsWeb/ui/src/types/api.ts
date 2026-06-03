// API model types derived from the .NET Models/Models.cs

export type JobStatus = 'Queued' | 'Running' | 'Completed' | 'Failed';

export interface Job {
  id: string;
  kind: string;
  status: JobStatus;
  createdAt: string;
  startedAt?: string;
  finishedAt?: string;
  result?: unknown;
  error?: string;
  details?: string;
  input?: Record<string, unknown>;
}

export interface BackupSnapshot {
  id: string;
  path: string;
  createdAt: string;
  modifiedAt: string;
  collectedAt?: string;
  tenantId?: string;
  roleDefinitionsCount?: number;
  roleAssignmentsCount?: number;
  roleDefinitionsHash?: string;
  roleAssignmentsHash?: string;
  sizeBytes?: number;
  sizeDisplay?: string;
  manifestError?: string;
}

export interface ApiResponse<T> {
  success: boolean;
  data: T;
  error?: string;
  details?: string;
}

export interface TaskEntry {
  time: string;
  name: string;
  status: string;
  type: string;
  operation: string;
}

export interface EventEntry {
  time: string;
  severity: string;
  source: string;
  message: string;
}

export interface ConsentInfo {
  granted: boolean;
  grantedAt?: string;
  grantedByUser?: string;
}

export interface TenantConsentsState {
  basic: ConsentInfo;
  restore: ConsentInfo;
}

export interface TenantEntry {
  id: string;
  name: string;
  tenantId: string;
  clientId?: string;
  certificateThumbprint?: string;
  domain?: string;
  addedAt: string;
  addedByUser?: string;
  consents: TenantConsentsState;
  isConfigured?: boolean;
}

export interface TenantConfig {
  tenantId: string;
  displayName?: string;
  clientId?: string;
  certificateThumbprint?: string;
}

export interface PaginatedResponse<T> {
  items: T[];
  total: number;
  skip: number;
  take: number;
}

export interface AppConfig {
  clientId?: string;
  clientSecret?: string;
  certificateThumbprint?: string;
  applicationObjectId?: string;
  servicePrincipalId?: string;
  displayName?: string;
  bootstrapTenantId?: string;
  createdAt?: string;
  isConfigured?: boolean;
}

export interface SetupStartResult {
  deviceCode: string;
  userCode: string;
  verificationUrl: string;
  message: string;
  expiresInSeconds: number;
  sessionId: string;
}

export interface SetupStatus {
  status: string;
  message?: string;
  clientId?: string;
  tenantId?: string;
  userCode?: string;
  verificationUrl?: string;
  sessionId?: string;
}

export interface AddTenantRequest {
  name: string;
  tenantId: string;
  clientId: string;
  certificateThumbprint: string;
  domain?: string;
}

export interface UnpackedObjects {
  backupId: string;
  roleDefinitions: unknown[];
  roleAssignments: unknown[];
}

export interface CompareResultItem {
  RoleName?: string;
  roleName?: string;
  RoleType?: string;
  roleType?: string;
  Status?: string;
  status?: string;
  DefinitionChanged?: string;
  definitionChanged?: string;
  AssignmentsChanged?: string;
  assignmentsChanged?: string;
  DifferenceCount?: number;
  differenceCount?: number;
  Summary?: string;
  summary?: string;
}

export interface CompareResults {
  generatedAt?: string;
  backupId?: string;
  data: CompareResultItem[];
}

export interface ConsentBlock {
  key: string;
  title: string;
  granted: boolean;
  grantedAt?: string;
  description?: string;
  operations: string[];
  permission?: string;
  sensitive?: boolean;
}

export interface TenantConsentsResponse {
  basic: ConsentBlock;
  restore: ConsentBlock & { sensitive?: boolean };
}
