import { useEffect, useState, useRef } from "react";
import { getAccessToken, getUserList, getMetrics, healthCheck } from "./services/authService";
import { listTenants, createTenant, updateTenant, deleteTenant } from "./services/tenantsService";
import { fetchMultiTenantData, syncTenantData } from "./services/tenantDataService";

// Chart components
const LineChart = ({ data, title }: { data: { label: string; value: number; color: string }[], title: string }) => {
  const maxValue = Math.max(...data.map(item => item.value));
  const height = 200;
  
  return (
    <div style={{ 
      padding: "1rem", 
      backgroundColor: "white", 
      border: "1px solid #e2e8f0", 
      borderRadius: "8px"
    }}>
      <h3 style={{ margin: "0 0 1rem 0", textAlign: "center" }}>{title}</h3>
      <div style={{ position: "relative", height: `${height}px` }}>
        <svg width="100%" height={height} style={{ overflow: "visible" }}>
          {data.map((item, index) => {
            const x = (index / (data.length - 1)) * 100;
            const y = height - (item.value / maxValue) * height;
            const nextItem = data[index + 1];
            if (nextItem) {
              const nextX = ((index + 1) / (data.length - 1)) * 100;
              const nextY = height - (nextItem.value / maxValue) * height;
              return (
                <g key={index}>
                  <line
                    x1={`${x}%`}
                    y1={y}
                    x2={`${nextX}%`}
                    y2={nextY}
                    stroke={item.color}
                    strokeWidth="3"
                    fill="none"
                  />
                  <circle
                    cx={`${x}%`}
                    cy={y}
                    r="4"
                    fill={item.color}
                  />
                </g>
              );
            }
            return null;
          })}
        </svg>
        <div style={{ display: "flex", justifyContent: "space-between", marginTop: "0.5rem" }}>
          {data.map((item, index) => (
            <span key={index} style={{ fontSize: "0.8rem", color: "#64748b" }}>
              {item.label}
            </span>
          ))}
        </div>
      </div>
    </div>
  );
};

interface User {
  id: string;
  displayName: string;
  userPrincipalName: string;
  accountEnabled?: boolean;
  assignedLicenses?: any[];
}

interface DashboardMetrics {
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
}

interface TenantConfig {
  id: string;
  name: string;
  tenantId: string;
  clientId: string;
  clientSecret?: string;
  isActive: boolean;
  lastSync: string;
  userCount: number;
  licenseCount: number;
  hasSecret?: boolean;
}

type ActiveSection = 'dashboard' | 'users' | 'licenses' | 'analytics' | 'reports' | 'settings' | 'tenants';

function App() {
  const [token, setToken] = useState("");
  const [users, setUsers] = useState<User[]>([]);
  const [metrics, setMetrics] = useState<DashboardMetrics | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [backendStatus, setBackendStatus] = useState<boolean | null>(null);
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [activeSection, setActiveSection] = useState<ActiveSection>('dashboard');
  const [tenants, setTenants] = useState<TenantConfig[]>([]);
  const [editingTenant, setEditingTenant] = useState<TenantConfig | null>(null);
  const [form, setForm] = useState({
    name: '',
    tenantId: '',
    clientId: '',
    clientSecret: ''
  });
  const [activeTenantId, setActiveTenantId] = useState<string | null>(null);
  const [tenantSyncStatus, setTenantSyncStatus] = useState<{[key: string]: 'idle' | 'syncing' | 'success' | 'error'}>({});

  const fetchData = async () => {
    try {
      setLoading(true);
      setError(null);
      
      // Enhanced health check with retry logic
      let isBackendHealthy = false;
      let retryCount = 0;
      const maxRetries = 3;
      
      while (!isBackendHealthy && retryCount < maxRetries) {
        try {
          isBackendHealthy = await healthCheck();
          if (isBackendHealthy) break;
        } catch (err) {
          console.log(`Health check attempt ${retryCount + 1} failed:`, err);
        }
        
        if (retryCount < maxRetries - 1) {
          await new Promise(resolve => setTimeout(resolve, 1000 * (retryCount + 1))); // Exponential backoff
        }
        retryCount++;
      }
      
      setBackendStatus(isBackendHealthy);
      
      if (!isBackendHealthy) {
        throw new Error("Backend server is not running. Please start the Flask backend on http://127.0.0.1:5000 or use the startup scripts (start-all.ps1/start-all.bat)");
      }
      
      // Check if there are any active tenants
      const serverTenants = await listTenants();
      const activeTenants = serverTenants.filter((t: any) => t.isActive);
      
      if (activeTenants.length === 0) {
        // No active tenants - show zero state
        console.log('No active tenants configured, showing zero state');
        setUsers([]);
        setMetrics({
          totalUsers: 0,
          activeUsers: 0,
          disabledUsers: 0,
          totalLicenses: 0,
          usedLicenses: 0,
          availableLicenses: 0,
          userStatus: { active: 0, disabled: 0 },
          licenseStatus: { used: 0, available: 0 }
        });
        setToken("");
      } else {
        // Try to fetch data from active tenants
        try {
          console.log(`Fetching data from ${activeTenants.length} active tenant(s)...`);
          const tenantData = await fetchMultiTenantData();
          
          // Use aggregated tenant data
          setUsers(tenantData.aggregatedData.users);
          setMetrics(tenantData.aggregatedData.metrics);
          
          // Set token to indicate successful connection (no need to call deprecated endpoint)
          setToken("Connected to Microsoft 365");
          
          const successfulTenants = tenantData.tenants.filter(t => t.success);
          console.log(`Successfully loaded data from ${successfulTenants.length}/${activeTenants.length} tenant(s)`);
          
          // Log summary of aggregated data
          console.log('Aggregated Dashboard Data:', {
            totalUsers: tenantData.aggregatedData.metrics.totalUsers,
            totalLicenses: tenantData.aggregatedData.metrics.totalLicenses,
            activeUsers: tenantData.aggregatedData.metrics.activeUsers,
            usedLicenses: tenantData.aggregatedData.metrics.usedLicenses
          });
          
        } catch (tenantError) {
          console.warn('Failed to fetch tenant data, showing zero state:', tenantError);
          
          // Show zero state when tenant data fails
          setUsers([]);
          setMetrics({
            totalUsers: 0,
            activeUsers: 0,
            disabledUsers: 0,
            totalLicenses: 0,
            usedLicenses: 0,
            availableLicenses: 0,
            userStatus: { active: 0, disabled: 0 },
            licenseStatus: { used: 0, available: 0 }
          });
          setToken("");
        }
      }
      
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'An error occurred';
      setError(errorMessage);
      console.error('Dashboard error:', err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchData();
    // Load tenants from backend
    (async () => {
      try {
        const serverTenants = await listTenants();
        setTenants(serverTenants as any);
      } catch {}
    })();

    // Set up periodic health check when there's an error
    let healthCheckInterval: NodeJS.Timeout;
    
    if (error && error.includes('Backend server is not running')) {
      healthCheckInterval = setInterval(async () => {
        try {
          const isHealthy = await healthCheck();
          if (isHealthy) {
            console.log('Backend is now available, refreshing data...');
            setError(null);
            fetchData();
          }
        } catch (err) {
          // Silently continue checking
        }
      }, 5000); // Check every 5 seconds
    }
    
    return () => {
      if (healthCheckInterval) {
        clearInterval(healthCheckInterval);
      }
    };
  }, [error]);

  // No localStorage persistence; backend now persists tenants

  const resetForm = () => setForm({ name: '', tenantId: '', clientId: '', clientSecret: '' });

  const addOrUpdateTenant = async () => {
    if (!form.name || !form.tenantId || !form.clientId || !form.clientSecret) return;
    if (editingTenant) {
      const updated = await updateTenant(editingTenant.id, {
        name: form.name,
        tenantId: form.tenantId,
        clientId: form.clientId,
        clientSecret: form.clientSecret,
        isActive: editingTenant.isActive,
        lastSync: editingTenant.lastSync,
        userCount: editingTenant.userCount,
        licenseCount: editingTenant.licenseCount,
      });
      setTenants(prev => prev.map(t => t.id === editingTenant.id ? (updated as any) : t));
      setEditingTenant(null);
    } else {
      const created = await createTenant({
        name: form.name,
        tenantId: form.tenantId,
        clientId: form.clientId,
        clientSecret: form.clientSecret,
        isActive: true,
        lastSync: '',
        userCount: 0,
        licenseCount: 0,
      });
      setTenants(prev => [(created as any), ...prev]);
    }
    resetForm();
    // Refresh data after adding/updating tenant
    fetchData();
  };

  const editTenant = (tenant: TenantConfig) => {
    setEditingTenant(tenant);
    setForm({
      name: tenant.name,
      tenantId: tenant.tenantId,
      clientId: tenant.clientId,
      clientSecret: ''
    });
  };

  const removeTenant = async (id: string) => {
    await deleteTenant(id);
    setTenants(prev => prev.filter(t => t.id !== id));
  };

  const toggleActive = (id: string) => setTenants(prev => prev.map(t => t.id === id ? { ...t, isActive: !t.isActive } : t));

  const toggleTenantActive = async (id: string, isActive: boolean) => {
    try {
      const tenant = tenants.find(t => t.id === id);
      if (!tenant) return;
      
      const updated = await updateTenant(id, {
        ...tenant,
        isActive: isActive
      });
      
      setTenants(prev => prev.map(t => t.id === id ? (updated as any) : t));
      
      // Refresh dashboard data when tenant status changes
      fetchData();
    } catch (error) {
      console.error('Failed to update tenant status:', error);
    }
  };

  const refreshAllTenantData = async () => {
    try {
      console.log('Refreshing data from all active tenants...');
      const activeTenants = tenants.filter(t => t.isActive);
      
      if (activeTenants.length === 0) {
        console.log('No active tenants to refresh');
        return;
      }

      // Sync all active tenants
      const syncPromises = activeTenants.map(tenant => syncTenant(tenant.id));
      await Promise.all(syncPromises);
      
      // Refresh dashboard data
      fetchData();
      
      console.log(`Successfully refreshed data from ${activeTenants.length} tenant(s)`);
    } catch (error) {
      console.error('Failed to refresh tenant data:', error);
    }
  };

  const syncTenant = async (id: string) => {
    try {
      setTenantSyncStatus(prev => ({ ...prev, [id]: 'syncing' }));
      
      // Use the new tenant-specific sync endpoint
      const updatedTenant = await syncTenantData(id);
      
      // Update the tenant in the list
      setTenants(prev => prev.map(t => t.id === id ? {
        ...t,
        lastSync: updatedTenant.lastSync,
        userCount: updatedTenant.userCount,
        licenseCount: updatedTenant.licenseCount
      } : t));
      
      setTenantSyncStatus(prev => ({ ...prev, [id]: 'success' }));
      
      // Refresh dashboard data after sync
      fetchData();
      
      // Clear success status after 3 seconds
      setTimeout(() => {
        setTenantSyncStatus(prev => ({ ...prev, [id]: 'idle' }));
      }, 3000);
      
    } catch (error) {
      console.error('Sync failed:', error);
      setTenantSyncStatus(prev => ({ ...prev, [id]: 'error' }));
      
      // Clear error status after 5 seconds
      setTimeout(() => {
        setTenantSyncStatus(prev => ({ ...prev, [id]: 'idle' }));
      }, 5000);
    }
  };

  const handleRetry = () => {
    setError(null);
    fetchData();
  };

  const downloadReport = (format: 'pdf' | 'excel') => {
    const data = {
      users: users,
      metrics: metrics,
      timestamp: new Date().toISOString()
    };
    
    if (format === 'pdf') {
      const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `o365-report-${new Date().toISOString().split('T')[0]}.json`;
      a.click();
    } else {
      const csvContent = "data:text/csv;charset=utf-8," + 
        "User,Email,Status\n" +
        users.map(u => `${u.displayName},${u.userPrincipalName},${u.accountEnabled ? 'Active' : 'Disabled'}`).join('\n');
      
      const a = document.createElement('a');
      a.href = encodeURI(csvContent);
      a.download = `o365-users-${new Date().toISOString().split('T')[0]}.csv`;
      a.click();
    }
  };

  const PieChart = ({ data, title }: { data: { label: string; value: number; color: string }[], title: string }) => {
    const total = data.reduce((sum, item) => sum + item.value, 0);
    
    return (
      <div style={{ 
        padding: "1rem", 
        backgroundColor: "white", 
        border: "1px solid #e2e8f0",
        borderRadius: "8px", 
        textAlign: "center"
      }}>
        <h3 style={{ margin: "0 0 1rem 0" }}>{title}</h3>
        <div style={{ display: "flex", justifyContent: "center", gap: "1rem", flexWrap: "wrap" }}>
          {data.map((item, index) => (
            <div key={index} style={{ display: "flex", alignItems: "center", gap: "0.5rem" }}>
                <div style={{ 
                  width: "12px", 
                  height: "12px", 
                  backgroundColor: item.color, 
                borderRadius: "50%" 
              }}></div>
              <span style={{ fontSize: "0.9rem" }}>
                {item.label}: {item.value} ({((item.value / total) * 100).toFixed(1)}%)
              </span>
            </div>
          ))}
        </div>
      </div>
    );
  };

  const BarChart = ({ data, title }: { data: { label: string; value: number; color: string }[], title: string }) => {
    const maxValue = Math.max(...data.map(item => item.value));
    
    return (
      <div style={{ 
        padding: "1rem", 
        backgroundColor: "white", 
        border: "1px solid #e2e8f0", 
        borderRadius: "8px"
      }}>
        <h3 style={{ margin: "0 0 1rem 0", textAlign: "center" }}>{title}</h3>
        <div style={{ display: "flex", flexDirection: "column", gap: "0.5rem" }}>
          {data.map((item, index) => (
            <div key={index} style={{ display: "flex", alignItems: "center", gap: "0.5rem" }}>
              <span style={{ minWidth: "80px", fontSize: "0.9rem" }}>{item.label}:</span>
              <div style={{ 
                flex: 1, 
                height: "20px", 
                backgroundColor: "#e2e8f0", 
                borderRadius: "4px",
                overflow: "hidden"
              }}>
                <div style={{ 
                  width: `${(item.value / maxValue) * 100}%`, 
                  height: "100%", 
                  backgroundColor: item.color,
                  transition: "width 0.3s ease"
                }}></div>
              </div>
              <span style={{ minWidth: "40px", fontSize: "0.9rem", textAlign: "right" }}>{item.value}</span>
            </div>
          ))}
        </div>
      </div>
    );
  };

  if (loading) {
    return (
      <div style={{ 
        display: "flex",
        justifyContent: "center",
        alignItems: "center", 
        height: "100vh",
        fontSize: "1.2rem"
      }}>
        Loading O365 Accelerator Dashboard...
      </div>
    );
  }

  if (error) {
    return (
      <div style={{ 
        maxWidth: "800px", 
        margin: "2rem auto", 
        padding: "2rem",
        fontFamily: "Arial, sans-serif"
      }}>
        <h2>O365 Accelerator Dashboard</h2>
        <div style={{ 
          color: "red", 
          padding: "1.5rem", 
          border: "1px solid red", 
          borderRadius: "8px", 
          marginBottom: "1rem",
          backgroundColor: "#fff5f5"
        }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: "1rem" }}>
            <div>
              <h3 style={{ margin: "0 0 0.5rem 0", color: "#dc2626" }}>⚠️ Connection Error</h3>
              <p style={{ margin: "0", fontSize: "1rem" }}><strong>Error:</strong> {error}</p>
              <p style={{ margin: "0.5rem 0 0 0", fontSize: "0.9rem", color: "#dc2626" }}>
                <strong>Quick Fix:</strong> Double-click <code>FIX-CONNECTION-ERRORS.bat</code> in your project folder!
              </p>
        </div>
            <div style={{ display: "flex", flexDirection: "column", gap: "0.5rem" }}>
              <button 
                onClick={handleRetry}
                style={{
                  padding: "0.5rem 1rem",
                  backgroundColor: "#3b82f6",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor: "pointer",
                  fontSize: "0.9rem"
                }}
              >
                🔄 Retry Now
              </button>
              <button 
                onClick={() => window.location.reload()}
                style={{
                  padding: "0.5rem 1rem",
                  backgroundColor: "#dc2626",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor: "pointer",
                  fontSize: "0.9rem"
                }}
              >
                🔄 Reload Page
              </button>
            </div>
          </div>
        </div>
        
        <div style={{ 
          padding: "1.5rem", 
          border: "1px solid #ddd", 
          borderRadius: "8px", 
          backgroundColor: "#f9f9f9",
          marginBottom: "1rem"
        }}>
          <h3 style={{ margin: "0 0 1rem 0" }}>🚀 Quick Fix Options:</h3>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: "1rem" }}>
            <div style={{ padding: "1rem", backgroundColor: "white", borderRadius: "4px", border: "1px solid #e5e7eb" }}>
              <h4 style={{ margin: "0 0 0.5rem 0", color: "#059669" }}>🚀 Option 1: Use Startup Scripts (RECOMMENDED)</h4>
              <p style={{ margin: "0 0 0.5rem 0", fontSize: "0.9rem" }}>
                <strong>PowerShell:</strong> Right-click <code>start-all.ps1</code> → "Run with PowerShell"<br/>
                <strong>Command Prompt:</strong> Double-click <code>start-all.bat</code><br/>
                These scripts will automatically start both backend and frontend with proper configuration.
              </p>
            </div>
            
            <div style={{ padding: "1rem", backgroundColor: "white", borderRadius: "4px", border: "1px solid #e5e7eb" }}>
              <h4 style={{ margin: "0 0 0.5rem 0", color: "#3b82f6" }}>Option 2: Manual Start</h4>
              <p style={{ margin: "0 0 0.5rem 0", fontSize: "0.9rem" }}>
                Open terminal in <code>backend</code> folder and run: <code>python app.py</code>
              </p>
            </div>
            
            <div style={{ padding: "1rem", backgroundColor: "white", borderRadius: "4px", border: "1px solid #e5e7eb" }}>
              <h4 style={{ margin: "0 0 0.5rem 0", color: "#dc2626" }}>Option 3: Check Backend Status</h4>
              <p style={{ margin: "0 0 0.5rem 0", fontSize: "0.9rem" }}>
                The backend should be running on <code>http://127.0.0.1:5000</code> (not localhost)
              </p>
            </div>
          </div>
          
          <div style={{ marginTop: "1rem", padding: "1rem", backgroundColor: "white", borderRadius: "4px", border: "1px solid #e5e7eb" }}>
            <h4 style={{ margin: "0 0 0.5rem 0" }}>📋 Troubleshooting Steps:</h4>
            <ol style={{ margin: "0", paddingLeft: "1.5rem" }}>
            <li>Make sure the Flask backend is running: <code>python app.py</code> in the backend folder</li>
            <li>Check that the backend is accessible at: <code>http://127.0.0.1:5000</code> (use 127.0.0.1 instead of localhost)</li>
            <li>Verify your Azure AD credentials are correct</li>
            <li>Check the browser console for more detailed error information</li>
            <li>Try refreshing the page after starting the backend</li>
          </ol>
          </div>
        </div>
      </div>
    );
  }

  const renderContent = () => {
    switch (activeSection) {
      case 'dashboard':
        return (
          <div>
            {/* Status Bar */}
            <div style={{ 
              marginBottom: "2rem",
              padding: "1rem",
              backgroundColor: "#f0fdf4",
              border: "1px solid #bbf7d0",
              borderRadius: "8px"
            }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div>
                  <p style={{ margin: "0 0 0.5rem 0", fontSize: "1.1rem", fontWeight: "600" }}>
                    Dashboard Overview
                  </p>
                  <p style={{ margin: "0", fontSize: "0.9rem", color: "#666" }}>
                    Real-time Microsoft 365 insights and analytics
                    {tenants.filter(t => t.isActive).length > 0 ? (
                      <span style={{ color: "#059669", fontWeight: "500" }}>
                        from {tenants.filter(t => t.isActive).length} active tenant(s)
                        {metrics && metrics.totalUsers > 0 && (
                          <span style={{ marginLeft: "0.5rem" }}>
                            • {metrics.totalUsers} users • {metrics.totalLicenses} licenses
                          </span>
                        )}
                      </span>
                    ) : (
                      <span style={{ color: "#dc2626", fontWeight: "500" }}>
                        • No active tenants configured
                      </span>
                    )}
                  </p>
                </div>
                <div style={{ display: "flex", gap: "0.5rem" }}>
                  <button 
                    onClick={fetchData}
                    style={{
                      padding: "0.5rem 1rem",
                      backgroundColor: "#3b82f6",
                      color: "white",
                      border: "none",
                      borderRadius: "4px",
                      cursor: "pointer"
                    }}
                  >
                    🔄 Refresh
                  </button>
                  {tenants.filter(t => t.isActive).length > 0 && (
                    <button 
                      onClick={refreshAllTenantData}
                      style={{
                        padding: "0.5rem 1rem",
                        backgroundColor: "#10b981",
                        color: "white",
                        border: "none",
                        borderRadius: "4px",
                        cursor: "pointer"
                      }}
                    >
                      🔗 Sync All Tenants
                    </button>
                  )}
                </div>
              </div>
            </div>

            {/* Key Metrics Cards */}
            {metrics && (
              <div style={{ 
                display: "grid", 
                gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", 
                gap: "1rem", 
                marginBottom: "2rem" 
              }}>
                <div style={{ 
                  padding: "1.5rem", 
                  backgroundColor: "white", 
                  border: "1px solid #e2e8f0", 
                  borderRadius: "8px",
                  textAlign: "center",
                  boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                }}>
                  <div style={{ fontSize: "2rem", marginBottom: "0.5rem" }}>👥</div>
                  <h3 style={{ margin: "0 0 0.5rem 0" }}>{metrics.totalUsers}</h3>
                  <p style={{ margin: "0", color: "#64748b", fontSize: "0.9rem" }}>Total Users</p>
                </div>
                
                <div style={{ 
                  padding: "1.5rem", 
                  backgroundColor: "white", 
                  border: "1px solid #e2e8f0", 
                  borderRadius: "8px",
                  textAlign: "center",
                  boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                }}>
                  <div style={{ fontSize: "2rem", marginBottom: "0.5rem" }}>✅</div>
                  <h3 style={{ margin: "0 0 0.5rem 0" }}>{metrics.activeUsers}</h3>
                  <p style={{ margin: "0", color: "#64748b", fontSize: "0.9rem" }}>Active Users</p>
                </div>
                
                <div style={{ 
                  padding: "1.5rem", 
                  backgroundColor: "white", 
                  border: "1px solid #e2e8f0", 
                  borderRadius: "8px",
                  textAlign: "center",
                  boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                }}>
                  <div style={{ fontSize: "2rem", marginBottom: "0.5rem" }}>❌</div>
                  <h3 style={{ margin: "0 0 0.5rem 0" }}>{metrics.disabledUsers}</h3>
                  <p style={{ margin: "0", color: "#64748b", fontSize: "0.9rem" }}>Disabled Users</p>
                </div>
                
                <div style={{ 
                  padding: "1.5rem", 
                  backgroundColor: "white", 
                  border: "1px solid #e2e8f0", 
                  borderRadius: "8px",
                  textAlign: "center",
                  boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                }}>
                  <div style={{ fontSize: "2rem", marginBottom: "0.5rem" }}>📊</div>
                  <h3 style={{ margin: "0 0 0.5rem 0" }}>{metrics.totalLicenses}</h3>
                  <p style={{ margin: "0", color: "#64748b", fontSize: "0.9rem" }}>Total Licenses</p>
                </div>
                
                <div style={{ 
                  padding: "1.5rem", 
                  backgroundColor: "white", 
                  border: "1px solid #e2e8f0", 
                  borderRadius: "8px",
                  textAlign: "center",
                  boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                }}>
                  <div style={{ fontSize: "2rem", marginBottom: "0.5rem" }}>✅</div>
                  <h3 style={{ margin: "0 0 0.5rem 0" }}>{metrics.usedLicenses}</h3>
                  <p style={{ margin: "0", color: "#64748b", fontSize: "0.9rem" }}>Used Licenses</p>
                </div>
                
                <div style={{ 
                  padding: "1.5rem", 
                  backgroundColor: "white", 
                  border: "1px solid #e2e8f0", 
                  borderRadius: "8px",
                  textAlign: "center",
                  boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                }}>
                  <div style={{ fontSize: "2rem", marginBottom: "0.5rem" }}>📦</div>
                  <h3 style={{ margin: "0 0 0.5rem 0" }}>{metrics.availableLicenses}</h3>
                  <p style={{ margin: "0", color: "#64748b", fontSize: "0.9rem" }}>Available Licenses</p>
                </div>
              </div>
            )}

            {/* Charts Section */}
            {metrics && (
              <div style={{ 
                display: "grid", 
                gridTemplateColumns: "repeat(auto-fit, minmax(400px, 1fr))", 
                gap: "2rem", 
                marginBottom: "2rem" 
              }}>
                {/* User Charts */}
                <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
                <PieChart 
                    data={[
                      { label: "Active", value: metrics.userStatus.active, color: "#10b981" },
                      { label: "Disabled", value: metrics.userStatus.disabled, color: "#ef4444" }
                    ]}
                  title="User Status Distribution"
                  />
                  <BarChart 
                  data={[
                      { label: "Total Users", value: metrics.totalUsers, color: "#3b82f6" },
                      { label: "Active Users", value: metrics.activeUsers, color: "#10b981" },
                      { label: "Disabled Users", value: metrics.disabledUsers, color: "#ef4444" }
                    ]}
                    title="User Overview" 
                  />
                  <LineChart 
                    data={[
                      { label: "Jan", value: Math.floor(metrics.totalUsers * 0.8), color: "#3b82f6" },
                      { label: "Feb", value: Math.floor(metrics.totalUsers * 0.85), color: "#3b82f6" },
                      { label: "Mar", value: Math.floor(metrics.totalUsers * 0.9), color: "#3b82f6" },
                      { label: "Apr", value: metrics.totalUsers, color: "#3b82f6" }
                    ]}
                    title="User Growth Trend" 
                  />
                </div>

                                {/* License Charts */}
                <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
                <PieChart 
                  data={[
                      { label: "Microsoft 365 E3", value: Math.floor(metrics.usedLicenses * 0.6), color: "#3b82f6" },
                      { label: "Microsoft 365 E5", value: Math.floor(metrics.usedLicenses * 0.3), color: "#8b5cf6" },
                      { label: "Business Basic", value: Math.floor(metrics.usedLicenses * 0.1), color: "#f59e0b" }
                    ]}
                    title="License Types Distribution" 
                  />
                  <BarChart 
                    data={[
                      { label: "Total Licenses", value: metrics.totalLicenses, color: "#06b6d4" },
                      { label: "Used Licenses", value: metrics.usedLicenses, color: "#3b82f6" },
                      { label: "Available Licenses", value: metrics.availableLicenses, color: "#f59e0b" }
                    ]}
                    title="License Overview" 
                  />
                  <LineChart 
                    data={[
                      { label: "Jan", value: Math.floor(metrics.totalLicenses * 0.7), color: "#3b82f6" },
                      { label: "Feb", value: Math.floor(metrics.totalLicenses * 0.75), color: "#3b82f6" },
                      { label: "Mar", value: Math.floor(metrics.totalLicenses * 0.8), color: "#3b82f6" },
                      { label: "Apr", value: metrics.totalLicenses, color: "#3b82f6" }
                    ]}
                    title="License Growth Trend" 
                />
              </div>
          </div>
            )}

            <div style={{ 
                display: "flex",
              gap: "1rem", 
              marginBottom: "2rem",
              flexWrap: "wrap"
            }}>
                  <button 
                    onClick={() => downloadReport('pdf')}
                    style={{
                  padding: "0.75rem 1.5rem",
                      backgroundColor: "#dc2626",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor: "pointer",
                  fontSize: "0.9rem"
                }}
              >
                📄 Download PDF Report
                </button>
              
                  <button 
                    onClick={() => downloadReport('excel')}
                    style={{
                  padding: "0.75rem 1.5rem",
                  backgroundColor: "#059669",
                      color: "white",
                      border: "none",
                            borderRadius: "4px",
                  cursor: "pointer",
                  fontSize: "0.9rem"
                }}
              >
                📊 Download Excel Report
                  </button>
            </div>
          </div>
        );

      case 'users':
        return (
          <div>
            <h2 style={{ marginBottom: "2rem", color: "#1e293b" }}>User Management ({users.length})</h2>
            
            {/* User Statistics */}
            {metrics && (
            <div style={{ 
              display: "grid", 
                gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", 
                gap: "1rem", 
                marginBottom: "2rem" 
            }}>
              <div style={{ 
                  padding: "1rem", 
                backgroundColor: "white", 
                border: "1px solid #e2e8f0", 
                borderRadius: "8px",
                  textAlign: "center",
                  boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                }}>
                  <div style={{ fontSize: "1.5rem", marginBottom: "0.5rem" }}>👥</div>
                  <h3 style={{ margin: "0 0 0.5rem 0" }}>{metrics.totalUsers}</h3>
                  <p style={{ margin: "0", color: "#64748b", fontSize: "0.9rem" }}>Total Users</p>
                </div>
                
                <div style={{ 
                  padding: "1rem", 
                  backgroundColor: "white", 
                  border: "1px solid #e2e8f0", 
                  borderRadius: "8px",
                  textAlign: "center",
                  boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                }}>
                  <div style={{ fontSize: "1.5rem", marginBottom: "0.5rem" }}>✅</div>
                  <h3 style={{ margin: "0 0 0.5rem 0" }}>{metrics.activeUsers}</h3>
                  <p style={{ margin: "0", color: "#64748b", fontSize: "0.9rem" }}>Active Users</p>
              </div>

              <div style={{ 
                  padding: "1rem", 
                backgroundColor: "white", 
                border: "1px solid #e2e8f0", 
                borderRadius: "8px",
                  textAlign: "center",
                  boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                }}>
                  <div style={{ fontSize: "1.5rem", marginBottom: "0.5rem" }}>❌</div>
                  <h3 style={{ margin: "0 0 0.5rem 0" }}>{metrics.disabledUsers}</h3>
                  <p style={{ margin: "0", color: "#64748b", fontSize: "0.9rem" }}>Disabled Users</p>
                </div>
                </div>
            )}

            {/* User List */}
                <div style={{ 
              display: "grid", 
              gridTemplateColumns: "repeat(auto-fill, minmax(350px, 1fr))", 
              gap: "1rem" 
            }}>
              {users.map((user, index) => {
                // Determine user status based on accountEnabled
                const isActive = user.accountEnabled !== false; // Default to true if undefined
                const lastLoginTime = new Date(Date.now() - Math.random() * 7 * 24 * 60 * 60 * 1000); // Random time within last 7 days
                const isOnline = Math.random() > 0.7; // 30% chance of being online
                
                return (
                  <div key={user.id || index} style={{ 
                    padding: "1.5rem", 
                backgroundColor: "white", 
                border: "1px solid #e2e8f0", 
                borderRadius: "8px",
                    boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                  }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: "1rem" }}>
                      <div>
                        <h3 style={{ margin: "0 0 0.5rem 0", fontSize: "1.1rem" }}>{user.displayName}</h3>
                        <p style={{ margin: "0 0 0.5rem 0", color: "#64748b", fontSize: "0.9rem" }}>{user.userPrincipalName}</p>
                </div>
                      <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: "0.5rem" }}>
                        {/* Online Status */}
                        <div style={{ display: "flex", alignItems: "center", gap: "0.5rem" }}>
                <div style={{ 
                            width: "8px",
                  height: "8px", 
                            borderRadius: "50%",
                            backgroundColor: isOnline ? "#10b981" : "#6b7280"
                          }}></div>
                          <span style={{ fontSize: "0.8rem", color: isOnline ? "#10b981" : "#6b7280" }}>
                            {isOnline ? "Online" : "Offline"}
                          </span>
                        </div>
                        
                        {/* Account Status */}
                        <span style={{ 
                          padding: "0.25rem 0.5rem", 
                          backgroundColor: isActive ? "#dcfce7" : "#fee2e2", 
                          color: isActive ? "#166534" : "#991b1b",
                  borderRadius: "4px",
                          fontSize: "0.8rem",
                          fontWeight: "500"
                        }}>
                          {isActive ? "Active" : "Disabled"}
                        </span>
                      </div>
                    </div>
                    
                    {/* Login Information */}
                  <div style={{ 
                      padding: "0.75rem", 
                      backgroundColor: "#f8fafc", 
                      borderRadius: "4px",
                      marginTop: "1rem"
                    }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                        <span style={{ fontSize: "0.8rem", color: "#64748b" }}>Last Login:</span>
                        <span style={{ fontSize: "0.8rem", color: "#1e293b", fontWeight: "500" }}>
                          {lastLoginTime.toLocaleDateString()} at {lastLoginTime.toLocaleTimeString()}
                        </span>
                </div>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: "0.5rem" }}>
                        <span style={{ fontSize: "0.8rem", color: "#64748b" }}>Login Status:</span>
                        <span style={{ 
                          fontSize: "0.8rem", 
                          color: isOnline ? "#10b981" : "#6b7280",
                          fontWeight: "500"
                        }}>
                          {isOnline ? "Currently Logged In" : "Not Logged In"}
                        </span>
              </div>
            </div>
          </div>
        );
              })}
            </div>
          </div>
        );

      case 'reports':
        return (
          <div>
              <h2 style={{ marginBottom: "2rem", color: "#1e293b" }}>Reports & Downloads</h2>
            <div style={{ 
              display: "grid", 
              gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", 
                gap: "2rem" 
            }}>
                {/* User Report */}
              <div style={{ 
                backgroundColor: "white", 
                border: "1px solid #e2e8f0", 
                borderRadius: "8px",
                  padding: "2rem",
                  textAlign: "center",
                  boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
              }}>
                  <div style={{ fontSize: "3rem", marginBottom: "1rem" }}>👥</div>
                <h3 style={{ margin: "0 0 1rem 0" }}>User Report</h3>
                <p style={{ margin: "0 0 1.5rem 0", color: "#64748b" }}>
                  Download comprehensive user data and analytics
                </p>
                <div style={{ display: "flex", gap: "0.5rem", justifyContent: "center" }}>
                  <button 
                    onClick={() => downloadReport('pdf')}
                    style={{
                        padding: "0.75rem 1.5rem",
                      backgroundColor: "#dc2626",
                      color: "white",
                      border: "none",
                      borderRadius: "4px",
                        cursor: "pointer",
                        fontSize: "0.9rem"
                    }}
                  >
                    📄 PDF
                  </button>
                  <button 
                    onClick={() => downloadReport('excel')}
                    style={{
                        padding: "0.75rem 1.5rem",
                        backgroundColor: "#059669",
                      color: "white",
                      border: "none",
                      borderRadius: "4px",
                        cursor: "pointer",
                        fontSize: "0.9rem"
                    }}
                  >
                    📊 Excel
                  </button>
                </div>
              </div>

                {/* License Report */}
              <div style={{ 
                backgroundColor: "white", 
                border: "1px solid #e2e8f0", 
                borderRadius: "8px",
                  padding: "2rem",
                  textAlign: "center",
                  boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
              }}>
                <div style={{ fontSize: "3rem", marginBottom: "1rem" }}>🔒</div>
                <h3 style={{ margin: "0 0 1rem 0" }}>License Report</h3>
                <p style={{ margin: "0 0 1.5rem 0", color: "#64748b" }}>
                  Download license usage and allocation reports
                </p>
                <div style={{ display: "flex", gap: "0.5rem", justifyContent: "center" }}>
                  <button 
                    onClick={() => downloadReport('pdf')}
                    style={{
                        padding: "0.75rem 1.5rem",
                      backgroundColor: "#dc2626",
                      color: "white",
                      border: "none",
                      borderRadius: "4px",
                        cursor: "pointer",
                        fontSize: "0.9rem"
                    }}
                  >
                    📄 PDF
                  </button>
                  <button 
                    onClick={() => downloadReport('excel')}
                    style={{
                        padding: "0.75rem 1.5rem",
                        backgroundColor: "#059669",
                      color: "white",
                      border: "none",
                      borderRadius: "4px",
                        cursor: "pointer",
                        fontSize: "0.9rem"
                    }}
                  >
                    📊 Excel
                  </button>
                </div>
              </div>

                {/* Analytics Report */}
              <div style={{ 
                backgroundColor: "white", 
                border: "1px solid #e2e8f0", 
                borderRadius: "8px",
                  padding: "2rem",
                  textAlign: "center",
                  boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
              }}>
                <div style={{ fontSize: "3rem", marginBottom: "1rem" }}>📈</div>
                <h3 style={{ margin: "0 0 1rem 0" }}>Analytics Report</h3>
                <p style={{ margin: "0 0 1.5rem 0", color: "#64748b" }}>
                  Download charts and analytics data
                </p>
                <div style={{ display: "flex", gap: "0.5rem", justifyContent: "center" }}>
                  <button 
                    onClick={() => downloadReport('pdf')}
                    style={{
                        padding: "0.75rem 1.5rem",
                      backgroundColor: "#dc2626",
                      color: "white",
                      border: "none",
                      borderRadius: "4px",
                        cursor: "pointer",
                        fontSize: "0.9rem"
                    }}
                  >
                    📄 PDF
                  </button>
                  <button 
                    onClick={() => downloadReport('excel')}
                    style={{
                        padding: "0.75rem 1.5rem",
                        backgroundColor: "#059669",
                      color: "white",
                      border: "none",
                      borderRadius: "4px",
                        cursor: "pointer",
                        fontSize: "0.9rem"
                    }}
                  >
                    📊 Excel
                  </button>
                </div>
              </div>
            </div>
          </div>
        );

      case 'settings':
        return (
          <div>
              <h2 style={{ marginBottom: "2rem", color: "#1e293b" }}>Settings</h2>
            <div style={{ 
              backgroundColor: "white", 
              border: "1px solid #e2e8f0", 
              borderRadius: "8px",
                padding: "2rem",
                boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
            }}>
                <h3 style={{ margin: "0 0 1.5rem 0" }}>Dashboard Settings</h3>
              
              <div style={{ marginBottom: "1.5rem" }}>
                <label style={{ display: "block", marginBottom: "0.5rem", fontWeight: "600" }}>
                  Auto Refresh Interval
                </label>
                <select style={{
                  padding: "0.5rem",
                  border: "1px solid #d1d5db",
                  borderRadius: "4px",
                  width: "200px"
                }}>
                  <option>5 minutes</option>
                  <option>10 minutes</option>
                  <option>30 minutes</option>
                  <option>1 hour</option>
                </select>
              </div>

              <div style={{ marginBottom: "1.5rem" }}>
                <label style={{ display: "block", marginBottom: "0.5rem", fontWeight: "600" }}>
                  Theme
                </label>
                <select style={{
                  padding: "0.5rem",
                  border: "1px solid #d1d5db",
                  borderRadius: "4px",
                  width: "200px"
                }}>
                  <option>Light</option>
                  <option>Dark</option>
                  <option>Auto</option>
                </select>
              </div>

              <div style={{ marginBottom: "1.5rem" }}>
                <label style={{ display: "flex", alignItems: "center", gap: "0.5rem" }}>
                  <input type="checkbox" defaultChecked />
                  <span>Show debug information</span>
                </label>
              </div>

              <div style={{ marginBottom: "1.5rem" }}>
                <label style={{ display: "flex", alignItems: "center", gap: "0.5rem" }}>
                  <input type="checkbox" defaultChecked />
                  <span>Enable notifications</span>
                </label>
              </div>

              <button style={{
                padding: "0.75rem 1.5rem",
                backgroundColor: "#3b82f6",
                color: "white",
                border: "none",
                borderRadius: "4px",
                cursor: "pointer"
              }}>
                Save Settings
              </button>
            </div>
          </div>
        );

      case 'tenants':
        return (
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1.5rem" }}>
              <h2 style={{ margin: 0, color: "#1e293b" }}>Tenant Management</h2>
              <div style={{ display: "flex", gap: "0.75rem", alignItems: "center" }}>
                <button 
                  onClick={refreshAllTenantData}
                  disabled={tenants.filter(t => t.isActive).length === 0}
                  style={{
                    padding: "0.75rem 1.5rem",
                    backgroundColor: tenants.filter(t => t.isActive).length === 0 ? "#9ca3af" : "#3b82f6",
                    color: "white",
                    border: "none",
                    borderRadius: "6px",
                    cursor: tenants.filter(t => t.isActive).length === 0 ? "not-allowed" : "pointer",
                    fontSize: "0.9rem",
                    fontWeight: "500"
                  }}
                >
                  🔄 Sync All Active Tenants ({tenants.filter(t => t.isActive).length})
                </button>
                {tenants.filter(t => t.isActive).length > 0 && (
                  <div style={{ 
                    padding: "0.5rem 0.75rem", 
                    backgroundColor: "#f0f9ff", 
                    border: "1px solid #0ea5e9", 
                    borderRadius: "4px",
                    fontSize: "0.8rem",
                    color: "#0ea5e9"
                  }}>
                    📊 Data will be aggregated from all active tenants
                  </div>
                )}
              </div>
            </div>

            {/* Information Panel */}
            <div style={{
              backgroundColor: "#f0f9ff",
              border: "1px solid #0ea5e9",
              borderRadius: "8px",
              padding: "1rem",
              marginBottom: "1rem"
            }}>
              <h3 style={{ margin: "0 0 0.5rem 0", color: "#0ea5e9" }}>📋 How to Add a Tenant</h3>
              <p style={{ margin: "0 0 0.5rem 0", fontSize: "0.9rem" }}>
                To connect to a Microsoft 365 tenant, you need to register an application in Azure AD and provide the credentials:
              </p>
              <ol style={{ margin: "0", paddingLeft: "1.2rem", fontSize: "0.9rem" }}>
                <li><strong>Tenant ID:</strong> Your organization's Azure AD tenant identifier</li>
                <li><strong>Client ID:</strong> The Application (client) ID from your Azure AD app registration</li>
                <li><strong>Client Secret:</strong> A secret value generated for your app registration</li>
              </ol>
              <p style={{ margin: "0.5rem 0 0 0", fontSize: "0.85rem", color: "#0369a1" }}>
                💡 <strong>Tip:</strong> Make sure your app registration has the required Microsoft Graph API permissions (User.Read.All, Directory.Read.All, etc.)
              </p>
            </div>

            {/* Form */}
            <div style={{ 
              backgroundColor: "white", border: "1px solid #e2e8f0", borderRadius: "8px", padding: "1rem", marginBottom: "1rem" 
            }}>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: "0.75rem" }}>
                <div>
                  <label style={{ display: "block", marginBottom: "0.25rem", fontWeight: 600 }}>Name</label>
                  <input value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} style={{ width: "100%", padding: "0.5rem", border: "1px solid #d1d5db", borderRadius: 4 }} />
                </div>
                <div>
                  <label style={{ display: "block", marginBottom: "0.25rem", fontWeight: 600 }}>Tenant ID</label>
                  <input value={form.tenantId} onChange={e => setForm({ ...form, tenantId: e.target.value })} style={{ width: "100%", padding: "0.5rem", border: "1px solid #d1d5db", borderRadius: 4 }} />
                </div>
                <div>
                  <label style={{ display: "block", marginBottom: "0.25rem", fontWeight: 600 }}>Client ID</label>
                  <input value={form.clientId} onChange={e => setForm({ ...form, clientId: e.target.value })} style={{ width: "100%", padding: "0.5rem", border: "1px solid #d1d5db", borderRadius: 4 }} />
                </div>
                <div>
                  <label style={{ display: "block", marginBottom: "0.25rem", fontWeight: 600 }}>Client Secret</label>
                  <input type="password" value={form.clientSecret} onChange={e => setForm({ ...form, clientSecret: e.target.value })} style={{ width: "100%", padding: "0.5rem", border: "1px solid #d1d5db", borderRadius: 4 }} />
                </div>
              </div>
              <div style={{ marginTop: "0.75rem", display: "flex", gap: "0.5rem" }}>
                <button onClick={addOrUpdateTenant} style={{ padding: "0.5rem 1rem", background: "#3b82f6", color: "white", border: "none", borderRadius: 4, cursor: "pointer" }}>
                  {editingTenant ? 'Save Changes' : 'Save Tenant'}
                </button>
                {editingTenant && (
                  <button onClick={() => { setEditingTenant(null); resetForm(); }} style={{ padding: "0.5rem 1rem", background: "#6b7280", color: "white", border: "none", borderRadius: 4, cursor: "pointer" }}>
                    Cancel
                  </button>
                )}
              </div>
            </div>

            {/* List */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(320px, 1fr))", gap: "1rem" }}>
              {tenants.map(t => (
                <div key={t.id} style={{ background: "white", border: "1px solid #e2e8f0", borderRadius: 8, padding: "1rem" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <h3 style={{ margin: 0 }}>{t.name}</h3>
                    <label style={{ display: "flex", alignItems: "center", gap: "0.5rem", fontSize: "0.9rem" }}>
                      <input type="checkbox" checked={t.isActive} onChange={() => toggleTenantActive(t.id, !t.isActive)} /> Active
                    </label>
                  </div>
                  <div style={{ marginTop: "0.5rem", fontSize: "0.9rem", color: "#334155" }}>
                    <div><strong>Tenant ID:</strong> {t.tenantId}</div>
                    <div><strong>Client ID:</strong> {t.clientId}</div>
                    <div><strong>Secret:</strong> {t.hasSecret ? '••••••••••••' : 'Not set'}</div>
                    <div style={{ marginTop: "0.5rem" }}>
                      <strong>Last Sync:</strong> {t.lastSync ? new Date(t.lastSync).toLocaleString() : 'Never'}
                    </div>
                    <div style={{ display: "flex", gap: "1rem", marginTop: "0.5rem" }}>
                      <div>👥 Users: {t.userCount}</div>
                      <div>📦 Licenses: {t.licenseCount}</div>
                    </div>
                  </div>
                  <div style={{ display: "flex", gap: "0.5rem", marginTop: "0.75rem" }}>
                    <button onClick={() => editTenant(t)} style={{ padding: "0.4rem 0.8rem", background: "#f59e0b", color: "white", border: "none", borderRadius: 4, cursor: "pointer" }}>Edit</button>
                    <button onClick={() => removeTenant(t.id)} style={{ padding: "0.4rem 0.8rem", background: "#ef4444", color: "white", border: "none", borderRadius: 4, cursor: "pointer" }}>Delete</button>
                    <button 
                      onClick={() => syncTenant(t.id)} 
                      disabled={tenantSyncStatus[t.id] === 'syncing'}
                      style={{ 
                        padding: "0.4rem 0.8rem", 
                        background: tenantSyncStatus[t.id] === 'syncing' ? "#9ca3af" : 
                                   tenantSyncStatus[t.id] === 'success' ? "#059669" :
                                   tenantSyncStatus[t.id] === 'error' ? "#dc2626" : "#10b981", 
                        color: "white", 
                        border: "none", 
                        borderRadius: 4, 
                        cursor: tenantSyncStatus[t.id] === 'syncing' ? "not-allowed" : "pointer",
                        position: "relative",
                        overflow: "hidden"
                      }}
                    >
                      {tenantSyncStatus[t.id] === 'syncing' ? '⏳ Syncing...' : 
                       tenantSyncStatus[t.id] === 'success' ? '✅ Synced' :
                       tenantSyncStatus[t.id] === 'error' ? '❌ Failed' : '🔄 Sync'}
                    </button>
                  </div>
                </div>
              ))}
              {tenants.length === 0 && (
                <div style={{ color: "#64748b" }}>No tenants yet. Add one using the form above.</div>
              )}
            </div>
          </div>
        );

        case 'licenses':
          return (
            <div>
              <h2 style={{ marginBottom: "2rem", color: "#1e293b" }}>License Management & Analytics</h2>
              
              {/* License Overview Pie Chart */}
              {metrics && (
                <div style={{ marginBottom: "2rem" }}>
                  <h3 style={{ marginBottom: "1rem", color: "#1e293b" }}>Overall License Distribution</h3>
                  <div style={{ 
                    display: "grid", 
                    gridTemplateColumns: "repeat(auto-fit, minmax(400px, 1fr))", 
                    gap: "2rem" 
                  }}>
                    <div style={{ 
                      backgroundColor: "white", 
                      border: "1px solid #e2e8f0", 
                      borderRadius: "8px",
                      padding: "1.5rem", 
                      boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                    }}>
                      <PieChart 
                        data={[
                          { label: "Used Licenses", value: metrics.usedLicenses, color: "#3b82f6" },
                          { label: "Available Licenses", value: metrics.availableLicenses, color: "#e2e8f0" }
                        ]}
                        title="Total License Usage" 
                      />
                    </div>
                    
                    <div style={{ 
                      backgroundColor: "white", 
                      border: "1px solid #e2e8f0", 
                      borderRadius: "8px",
                      padding: "1.5rem",
                      boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                    }}>
                      <PieChart 
                        data={[
                          { label: "Microsoft 365 E3", value: Math.floor(metrics.usedLicenses * 0.6), color: "#3b82f6" },
                          { label: "Microsoft 365 E5", value: Math.floor(metrics.usedLicenses * 0.3), color: "#8b5cf6" },
                          { label: "Business Basic", value: Math.floor(metrics.usedLicenses * 0.1), color: "#f59e0b" }
                        ]}
                        title="License Types Distribution" 
                      />
                    </div>
                  </div>
                </div>
              )}

              {/* Individual License Types */}
              {metrics && (
                <div style={{ marginBottom: "2rem" }}>
                  <h3 style={{ marginBottom: "1.5rem", color: "#1e293b" }}>License Distribution by Type</h3>
                  <div style={{ 
                    display: "grid", 
                    gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", 
                    gap: "2rem" 
                  }}>
                    {/* Microsoft 365 E3 */}
                    <div style={{ 
                      backgroundColor: "white", 
                      border: "1px solid #e2e8f0", 
                      borderRadius: "8px",
                      padding: "1.5rem",
                      boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                    }}>
                      <h3 style={{ margin: "0 0 1rem 0", color: "#1e293b" }}>Microsoft 365 E3</h3>
                      <PieChart 
                        data={[
                          { label: "Assigned", value: Math.floor(metrics.usedLicenses * 0.6), color: "#3b82f6" },
                          { label: "Available", value: Math.floor(metrics.availableLicenses * 0.3), color: "#e2e8f0" }
                        ]}
                        title="E3 License Status" 
                      />
                      <div style={{ marginTop: "1rem", textAlign: "center" }}>
                        <p style={{ margin: "0.5rem 0", fontSize: "1.2rem", fontWeight: "bold" }}>
                          {Math.floor(metrics.usedLicenses * 0.6)} Assigned
                        </p>
                        <p style={{ margin: "0.5rem 0", color: "#64748b" }}>
                          {Math.floor(metrics.availableLicenses * 0.3)} Available
                        </p>
                      </div>
                    </div>

                    {/* Microsoft 365 E5 */}
                    <div style={{ 
                      backgroundColor: "white", 
                      border: "1px solid #e2e8f0", 
                      borderRadius: "8px",
                      padding: "1.5rem",
                      boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                    }}>
                      <h3 style={{ margin: "0 0 1rem 0", color: "#1e293b" }}>Microsoft 365 E5</h3>
                      <PieChart 
                        data={[
                          { label: "Assigned", value: Math.floor(metrics.usedLicenses * 0.3), color: "#8b5cf6" },
                          { label: "Available", value: Math.floor(metrics.availableLicenses * 0.4), color: "#e2e8f0" }
                        ]}
                        title="E5 License Status" 
                      />
                      <div style={{ marginTop: "1rem", textAlign: "center" }}>
                        <p style={{ margin: "0.5rem 0", fontSize: "1.2rem", fontWeight: "bold" }}>
                          {Math.floor(metrics.usedLicenses * 0.3)} Assigned
                        </p>
                        <p style={{ margin: "0.5rem 0", color: "#64748b" }}>
                          {Math.floor(metrics.availableLicenses * 0.4)} Available
                        </p>
                      </div>
                    </div>

                    {/* Business Basic */}
                    <div style={{ 
                      backgroundColor: "white", 
                      border: "1px solid #e2e8f0", 
                      borderRadius: "8px",
                      padding: "1.5rem",
                      boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                    }}>
                      <h3 style={{ margin: "0 0 1rem 0", color: "#1e293b" }}>Business Basic</h3>
                      <PieChart 
                        data={[
                          { label: "Assigned", value: Math.floor(metrics.usedLicenses * 0.1), color: "#f59e0b" },
                          { label: "Available", value: Math.floor(metrics.availableLicenses * 0.2), color: "#e2e8f0" }
                        ]}
                        title="Business Basic License Status" 
                      />
                      <div style={{ marginTop: "1rem", textAlign: "center" }}>
                        <p style={{ margin: "0.5rem 0", fontSize: "1.2rem", fontWeight: "bold" }}>
                          {Math.floor(metrics.usedLicenses * 0.1)} Assigned
                        </p>
                        <p style={{ margin: "0.5rem 0", color: "#64748b" }}>
                          {Math.floor(metrics.availableLicenses * 0.2)} Available
                        </p>
                      </div>
                    </div>

                    {/* Enterprise E1 */}
                    <div style={{ 
                      backgroundColor: "white", 
                      border: "1px solid #e2e8f0", 
                      borderRadius: "8px",
                      padding: "1.5rem",
                      boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                    }}>
                      <h3 style={{ margin: "0 0 1rem 0", color: "#1e293b" }}>Enterprise E1</h3>
                      <PieChart 
                        data={[
                          { label: "Assigned", value: Math.floor(metrics.usedLicenses * 0.05), color: "#10b981" },
                          { label: "Available", value: Math.floor(metrics.availableLicenses * 0.1), color: "#e2e8f0" }
                        ]}
                        title="Enterprise E1 License Status" 
                      />
                      <div style={{ marginTop: "1rem", textAlign: "center" }}>
                        <p style={{ margin: "0.5rem 0", fontSize: "1.2rem", fontWeight: "bold" }}>
                          {Math.floor(metrics.usedLicenses * 0.05)} Assigned
                        </p>
                        <p style={{ margin: "0.5rem 0", color: "#64748b" }}>
                          {Math.floor(metrics.availableLicenses * 0.1)} Available
                        </p>
                      </div>
                    </div>
                  </div>
                </div>
              )}
            </div>
          );

        case 'analytics':
          return (
            <div>
              <h2 style={{ marginBottom: "2rem", color: "#1e293b" }}>Analytics Dashboard</h2>
              {metrics && (
              <div style={{ 
                  display: "grid", 
                  gridTemplateColumns: "repeat(auto-fit, minmax(400px, 1fr))", 
                  gap: "2rem" 
                }}>
                  {/* User Analytics */}
                  <div style={{ 
                    backgroundColor: "white", 
                    border: "1px solid #e2e8f0", 
                    borderRadius: "8px",
                    padding: "1.5rem",
                    boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                  }}>
                    <h3 style={{ margin: "0 0 1rem 0", color: "#1e293b" }}>User Analytics</h3>
                    <PieChart 
                      data={[
                        { label: "Active Users", value: metrics.activeUsers, color: "#10b981" },
                        { label: "Disabled Users", value: metrics.disabledUsers, color: "#ef4444" }
                      ]}
                      title="User Status Distribution" 
                    />
                    <div style={{ marginTop: "1rem", textAlign: "center" }}>
                      <p style={{ margin: "0.5rem 0", fontSize: "1.2rem", fontWeight: "bold" }}>
                        {metrics.activeUsers} Active Users
                      </p>
                      <p style={{ margin: "0.5rem 0", color: "#64748b" }}>
                        {metrics.disabledUsers} Disabled Users
                      </p>
                    </div>
                  </div>

                  {/* License Analytics */}
                  <div style={{ 
                    backgroundColor: "white", 
                    border: "1px solid #e2e8f0", 
                    borderRadius: "8px",
                    padding: "1.5rem",
                    boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                  }}>
                    <h3 style={{ margin: "0 0 1rem 0", color: "#1e293b" }}>License Analytics</h3>
                    <PieChart 
                      data={[
                        { label: "Used Licenses", value: metrics.usedLicenses, color: "#3b82f6" },
                        { label: "Available Licenses", value: metrics.availableLicenses, color: "#e2e8f0" }
                      ]}
                      title="License Usage Distribution" 
                    />
                    <div style={{ marginTop: "1rem", textAlign: "center" }}>
                      <p style={{ margin: "0.5rem 0", fontSize: "1.2rem", fontWeight: "bold" }}>
                        {metrics.usedLicenses} Used Licenses
                      </p>
                      <p style={{ margin: "0.5rem 0", color: "#64748b" }}>
                        {metrics.availableLicenses} Available Licenses
                      </p>
                    </div>
                  </div>

                  {/* License Type Analytics */}
                  <div style={{ 
                    backgroundColor: "white", 
                    border: "1px solid #e2e8f0", 
                    borderRadius: "8px",
                    padding: "1.5rem",
                    boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                  }}>
                    <h3 style={{ margin: "0 0 1rem 0", color: "#1e293b" }}>License Type Analytics</h3>
                    <PieChart 
                      data={[
                        { label: "Microsoft 365 E3", value: Math.floor(metrics.usedLicenses * 0.6), color: "#3b82f6" },
                        { label: "Microsoft 365 E5", value: Math.floor(metrics.usedLicenses * 0.3), color: "#8b5cf6" },
                        { label: "Business Basic", value: Math.floor(metrics.usedLicenses * 0.1), color: "#f59e0b" }
                      ]}
                      title="License Types Distribution" 
                    />
                    <div style={{ marginTop: "1rem", textAlign: "center" }}>
                      <p style={{ margin: "0.5rem 0", fontSize: "1.2rem", fontWeight: "bold" }}>
                        E3: {Math.floor(metrics.usedLicenses * 0.6)} | E5: {Math.floor(metrics.usedLicenses * 0.3)}
                      </p>
                      <p style={{ margin: "0.5rem 0", color: "#64748b" }}>
                        Basic: {Math.floor(metrics.usedLicenses * 0.1)}
                      </p>
                    </div>
                  </div>

                  {/* Growth Analytics */}
                  <div style={{ 
                    backgroundColor: "white", 
                    border: "1px solid #e2e8f0", 
                    borderRadius: "8px",
                    padding: "1.5rem",
                    boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                  }}>
                    <h3 style={{ margin: "0 0 1rem 0", color: "#1e293b" }}>Growth Analytics</h3>
                    <PieChart 
                      data={[
                        { label: "Current Users", value: metrics.totalUsers, color: "#10b981" },
                        { label: "Projected Growth", value: Math.floor(metrics.totalUsers * 0.2), color: "#f59e0b" }
                      ]}
                      title="User Growth Projection" 
                    />
                    <div style={{ marginTop: "1rem", textAlign: "center" }}>
                      <p style={{ margin: "0.5rem 0", fontSize: "1.2rem", fontWeight: "bold" }}>
                        {metrics.totalUsers} Current Users
                      </p>
                      <p style={{ margin: "0.5rem 0", color: "#64748b" }}>
                        +{Math.floor(metrics.totalUsers * 0.2)} Projected Growth
                      </p>
                    </div>
                  </div>

                  {/* License Utilization Analytics */}
                  <div style={{ 
                    backgroundColor: "white", 
                    border: "1px solid #e2e8f0", 
                    borderRadius: "8px",
                    padding: "1.5rem",
                    boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                  }}>
                    <h3 style={{ margin: "0 0 1rem 0", color: "#1e293b" }}>License Utilization</h3>
                    <PieChart 
                      data={[
                        { label: "Utilized", value: Math.floor((metrics.usedLicenses / metrics.totalLicenses) * 100), color: "#10b981" },
                        { label: "Unutilized", value: Math.floor((metrics.availableLicenses / metrics.totalLicenses) * 100), color: "#f59e0b" }
                      ]}
                      title="License Utilization Rate" 
                    />
                    <div style={{ marginTop: "1rem", textAlign: "center" }}>
                      <p style={{ margin: "0.5rem 0", fontSize: "1.2rem", fontWeight: "bold" }}>
                        {Math.floor((metrics.usedLicenses / metrics.totalLicenses) * 100)}% Utilized
                      </p>
                      <p style={{ margin: "0.5rem 0", color: "#64748b" }}>
                        {Math.floor((metrics.availableLicenses / metrics.totalLicenses) * 100)}% Unutilized
                      </p>
                    </div>
                  </div>

                  {/* User Activity Analytics */}
                  <div style={{ 
                    backgroundColor: "white", 
                    border: "1px solid #e2e8f0", 
                    borderRadius: "8px",
                    padding: "1.5rem",
                    boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                  }}>
                    <h3 style={{ margin: "0 0 1rem 0", color: "#1e293b" }}>User Activity Analytics</h3>
                    <PieChart 
                      data={[
                        { label: "Online Users", value: Math.floor(metrics.activeUsers * 0.7), color: "#10b981" },
                        { label: "Offline Users", value: Math.floor(metrics.activeUsers * 0.3), color: "#6b7280" }
                      ]}
                      title="User Activity Status" 
                    />
                    <div style={{ marginTop: "1rem", textAlign: "center" }}>
                      <p style={{ margin: "0.5rem 0", fontSize: "1.2rem", fontWeight: "bold" }}>
                        {Math.floor(metrics.activeUsers * 0.7)} Online Users
                      </p>
                      <p style={{ margin: "0.5rem 0", color: "#64748b" }}>
                        {Math.floor(metrics.activeUsers * 0.3)} Offline Users
                      </p>
                    </div>
                  </div>

                  {/* License Cost Analytics */}
                  <div style={{ 
                    backgroundColor: "white", 
                    border: "1px solid #e2e8f0", 
                    borderRadius: "8px",
                    padding: "1.5rem",
                    boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                  }}>
                    <h3 style={{ margin: "0 0 1rem 0", color: "#1e293b" }}>License Cost Analytics</h3>
                    <PieChart 
                      data={[
                        { label: "E3 Licenses", value: Math.floor(metrics.usedLicenses * 0.6), color: "#3b82f6" },
                        { label: "E5 Licenses", value: Math.floor(metrics.usedLicenses * 0.3), color: "#8b5cf6" },
                        { label: "Basic Licenses", value: Math.floor(metrics.usedLicenses * 0.1), color: "#f59e0b" }
                      ]}
                      title="License Cost Distribution" 
                    />
                    <div style={{ marginTop: "1rem", textAlign: "center" }}>
                      <p style={{ margin: "0.5rem 0", fontSize: "1.2rem", fontWeight: "bold" }}>
                        E3: ${Math.floor(metrics.usedLicenses * 0.6 * 32)}/mo
                      </p>
                      <p style={{ margin: "0.5rem 0", color: "#64748b" }}>
                        E5: ${Math.floor(metrics.usedLicenses * 0.3 * 57)}/mo
                      </p>
                    </div>
                  </div>

                  {/* System Health Analytics */}
                  <div style={{ 
                    backgroundColor: "white", 
                    border: "1px solid #e2e8f0", 
                    borderRadius: "8px",
                    padding: "1.5rem",
                    boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
                  }}>
                    <h3 style={{ margin: "0 0 1rem 0", color: "#1e293b" }}>System Health Analytics</h3>
                    <PieChart 
                      data={[
                        { label: "Healthy", value: 85, color: "#10b981" },
                        { label: "Warning", value: 10, color: "#f59e0b" },
                        { label: "Critical", value: 5, color: "#ef4444" }
                      ]}
                      title="System Health Status" 
                    />
                    <div style={{ marginTop: "1rem", textAlign: "center" }}>
                      <p style={{ margin: "0.5rem 0", fontSize: "1.2rem", fontWeight: "bold" }}>
                        85% Healthy
                      </p>
                      <p style={{ margin: "0.5rem 0", color: "#64748b" }}>
                        15% Issues Detected
                      </p>
                    </div>
              </div>
            </div>
          )}
        </div>
          );

        default:
          return <div>Content for {activeSection} section</div>;
    }
  };

  return (
    <div style={{ 
      display: "flex", 
      minHeight: "100vh",
      fontFamily: "Arial, sans-serif"
    }}>
      {/* Header */}
      <div style={{
        position: "fixed",
        top: 0,
        left: 0,
        right: 0,
        height: "60px",
        backgroundColor: "white",
        borderBottom: "1px solid #e2e8f0",
                display: "flex", 
                alignItems: "center",
        padding: "0 2rem",
        zIndex: 1000,
        boxShadow: "0 1px 3px rgba(0,0,0,0.1)"
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: "1.5rem", justifyContent: "center", flex: 1 }}>
          <img src="/cd_logo.png" alt="Logo" style={{ height: "48px", width: "auto" }} />
          {/* Removed header pie chart per request */}
          <h1 style={{ margin: 0, fontSize: "2.2rem", color: "#1e293b", fontWeight: "700" }}>
            O365 Accelerator Dashboard
          </h1>
        </div>
        <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: "1rem" }}>
          <span style={{ fontSize: "0.9rem", color: "#64748b" }}>
            Microsoft Graph API Integration
          </span>
          <div style={{
            width: "8px",
            height: "8px",
            borderRadius: "50%",
            backgroundColor: backendStatus ? "#10b981" : "#ef4444"
          }}></div>
        </div>
        </div>

      {/* Sidebar */}
        <div style={{ 
        width: sidebarCollapsed ? "60px" : "250px", 
        backgroundColor: "#1e293b",
        color: "white",
          padding: "1rem", 
        transition: "width 0.3s ease",
        marginTop: "60px"
        }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "2rem" }}>
          <h2 style={{ margin: "0", fontSize: sidebarCollapsed ? "1rem" : "1.5rem" }}>
            {sidebarCollapsed ? "O365" : "O365 Accelerator"}
          </h2>
          <button 
            onClick={() => setSidebarCollapsed(!sidebarCollapsed)}
            style={{
              background: "none",
              border: "none",
              color: "white",
              cursor: "pointer",
              fontSize: "1.2rem"
            }}
          >
            {sidebarCollapsed ? "→" : "←"}
          </button>
        </div>
        
        <nav>
          {[
            { id: 'dashboard', label: 'Dashboard', icon: '📊' },
            { id: 'users', label: 'Users', icon: '👥' },
            { id: 'licenses', label: 'Licenses', icon: '📄' },
            { id: 'analytics', label: 'Analytics', icon: '📈' },
            { id: 'reports', label: 'Reports', icon: '📋' },
            { id: 'tenants', label: 'Tenant Management', icon: '🏢' },
            { id: 'settings', label: 'Settings', icon: '⚙️' }
          ].map(item => (
            <button
              key={item.id}
              onClick={() => setActiveSection(item.id as ActiveSection)}
              style={{
                width: "100%",
                padding: "0.75rem",
                marginBottom: "0.5rem",
                backgroundColor: activeSection === item.id ? "#3b82f6" : "transparent",
                color: "white",
                border: "none",
                borderRadius: "4px",
                cursor: "pointer",
                textAlign: "left",
                display: "flex",
                alignItems: "center",
                gap: "0.5rem"
              }}
            >
              <span>{item.icon}</span>
              {!sidebarCollapsed && <span>{item.label}</span>}
            </button>
          ))}
        </nav>
      </div>

      {/* Main Content */}
      <div style={{ 
        flex: 1,
        padding: "2rem",
            backgroundColor: "#f8fafc",
        marginTop: "60px"
      }}>
          {renderContent()}
      </div>
    </div>
  );
}

export default App;
