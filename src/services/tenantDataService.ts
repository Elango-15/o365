declare global {
  interface Window {
    __API_BASE_URL__?: string;
  }
}

const API_ORIGIN =
  (typeof import.meta !== 'undefined' && (import.meta as any).env?.VITE_API_BASE_URL) ||
  (typeof window !== 'undefined' ? window.__API_BASE_URL__ : undefined) ||
  "http://127.0.0.1:5000";

const API_BASE = `${API_ORIGIN.replace(/\/$/, '')}/api`;

export interface TenantData {
  users: any[];
  groups: any[];
  sites: any[];
  licenses: any[];
  metrics: {
    totalUsers: number;
    activeUsers: number;
    disabledUsers: number;
    totalLicenses: number;
    usedLicenses: number;
    availableLicenses: number;
    userStatus: {
      active: number;
      disabled: number;
    };
    licenseStatus: {
      used: number;
      available: number;
    };
  };
}

export async function fetchTenantData(tenantId: string): Promise<TenantData> {
  const response = await fetch(`${API_BASE}/tenants/${tenantId}/data`);
  if (!response.ok) {
    throw new Error(`Failed to fetch tenant data: ${response.status} ${response.statusText}`);
  }
  return response.json();
}

export async function syncTenantData(tenantId: string): Promise<any> {
  const response = await fetch(`${API_BASE}/tenants/${tenantId}/sync`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    }
  });
  if (!response.ok) {
    throw new Error(`Failed to sync tenant data: ${response.status} ${response.statusText}`);
  }
  return response.json();
}

export async function getActiveTenants(): Promise<any[]> {
  const response = await fetch(`${API_BASE}/tenants`);
  if (!response.ok) {
    throw new Error(`Failed to fetch tenants: ${response.status} ${response.statusText}`);
  }
  const data = await response.json();
  return data.tenants.filter((tenant: any) => tenant.isActive);
}

export async function fetchMultiTenantData(): Promise<{
  tenants: any[];
  aggregatedData: TenantData;
}> {
  const activeTenants = await getActiveTenants();
  
  if (activeTenants.length === 0) {
    throw new Error('No active tenants configured. Please add and activate at least one tenant.');
  }

  // Fetch data from all active tenants
  const tenantDataPromises = activeTenants.map(async (tenant) => {
    try {
      const data = await fetchTenantData(tenant.id);
      return { tenant, data, success: true };
    } catch (error) {
      console.warn(`Failed to fetch data for tenant ${tenant.name}:`, error);
      return { tenant, data: null, success: false, error };
    }
  });

  const results = await Promise.all(tenantDataPromises);
  const successfulResults = results.filter(r => r.success && r.data);

  if (successfulResults.length === 0) {
    throw new Error('Failed to fetch data from any active tenant. Please check your tenant configurations.');
  }

  // Aggregate data from all successful tenants
  const aggregatedData: TenantData = {
    users: [],
    groups: [],
    sites: [],
    licenses: [],
    metrics: {
      totalUsers: 0,
      activeUsers: 0,
      disabledUsers: 0,
      totalLicenses: 0,
      usedLicenses: 0,
      availableLicenses: 0,
      userStatus: { active: 0, disabled: 0 },
      licenseStatus: { used: 0, available: 0 }
    }
  };

  // Combine data from all tenants
  successfulResults.forEach(({ data }) => {
    if (data) {
      aggregatedData.users.push(...data.users);
      aggregatedData.groups.push(...data.groups);
      aggregatedData.sites.push(...data.sites);
      aggregatedData.licenses.push(...data.licenses);
      
      // Aggregate metrics
      aggregatedData.metrics.totalUsers += data.metrics.totalUsers;
      aggregatedData.metrics.activeUsers += data.metrics.activeUsers;
      aggregatedData.metrics.disabledUsers += data.metrics.disabledUsers;
      aggregatedData.metrics.totalLicenses += data.metrics.totalLicenses;
      aggregatedData.metrics.usedLicenses += data.metrics.usedLicenses;
      aggregatedData.metrics.availableLicenses += data.metrics.availableLicenses;
      aggregatedData.metrics.userStatus.active += data.metrics.userStatus.active;
      aggregatedData.metrics.userStatus.disabled += data.metrics.userStatus.disabled;
      aggregatedData.metrics.licenseStatus.used += data.metrics.licenseStatus.used;
      aggregatedData.metrics.licenseStatus.available += data.metrics.licenseStatus.available;
    }
  });

  return {
    tenants: results,
    aggregatedData
  };
}
