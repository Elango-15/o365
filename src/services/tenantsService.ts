declare global {
  interface Window {
    __API_BASE_URL__?: string;
  }
}

const API_ORIGIN =
  (typeof import.meta !== 'undefined' && (import.meta as any).env?.VITE_API_BASE_URL) ||
  (typeof window !== 'undefined' ? window.__API_BASE_URL__ : undefined) ||
  "http://127.0.0.1:5000";

const API_BASE = `${API_ORIGIN.replace(/\/$/, '')}/api/tenants`;

export interface TenantPayload {
  id?: string;
  name: string;
  tenantId: string;
  clientId: string;
  clientSecret?: string;
  isActive?: boolean;
  lastSync?: string;
  userCount?: number;
  licenseCount?: number;
}

export interface TenantRecord extends TenantPayload { id: string; }

export async function listTenants(): Promise<TenantRecord[]> {
  const res = await fetch(API_BASE);
  if (!res.ok) throw new Error(`List tenants failed: ${res.status}`);
  const data = await res.json();
  return data.tenants ?? [];
}

export async function createTenant(payload: TenantPayload): Promise<TenantRecord> {
  const res = await fetch(API_BASE, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  if (!res.ok) throw new Error(`Create tenant failed: ${res.status}`);
  return res.json();
}

export async function updateTenant(id: string, payload: TenantPayload): Promise<TenantRecord> {
  // Omit empty secret to avoid wiping existing value on server
  const toSend: any = { ...payload };
  if (!toSend.clientSecret) delete toSend.clientSecret;
  const res = await fetch(`${API_BASE}/${id}`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(toSend),
  });
  if (!res.ok) throw new Error(`Update tenant failed: ${res.status}`);
  return res.json();
}

export async function deleteTenant(id: string): Promise<void> {
  const res = await fetch(`${API_BASE}/${id}`, { method: 'DELETE' });
  if (!res.ok) throw new Error(`Delete tenant failed: ${res.status}`);
}


