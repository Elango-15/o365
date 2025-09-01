// Allow configuring API base via Vite env or a global injected var, with a safe fallback
declare global {
  interface Window {
    __API_BASE_URL__?: string;
  }
}

// Use localhost/127.0.0.1 for development to avoid network issues
const API_ORIGIN =
  (typeof import.meta !== 'undefined' && (import.meta as any).env?.VITE_API_BASE_URL) ||
  (typeof window !== 'undefined' ? window.__API_BASE_URL__ : undefined) ||
  "http://127.0.0.1:5000"; // Always use localhost for development

const API_BASE = `${API_ORIGIN.replace(/\/$/, '')}/api`;

// Retry configuration
const RETRY_ATTEMPTS = 3;
const RETRY_DELAY = 2000; // 2 seconds

const delay = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

const fetchWithRetry = async (url: string, options: RequestInit = {}, retries = RETRY_ATTEMPTS): Promise<Response> => {
  try {
    const response = await fetch(url, {
      ...options,
      headers: {
        'Content-Type': 'application/json',
        ...options.headers,
      },
    });
    
    if (response.ok) {
      return response;
    }
    
    throw new Error(`HTTP ${response.status}: ${response.statusText}`);
  } catch (error) {
    if (retries > 0) {
      console.log(`Request failed, retrying... (${retries} attempts left)`);
      await delay(RETRY_DELAY);
      return fetchWithRetry(url, options, retries - 1);
    }
    throw error;
  }
};

export const getAccessToken = async (): Promise<string> => {
  try {
    const res = await fetchWithRetry(`${API_BASE}/token`);
    const data = await res.json();
    
    if (data.error) {
      throw new Error(`Token API Error: ${data.error}`);
    }
    
    if (!data.access_token) {
      throw new Error('No access token received from server');
    }
    
    return data.access_token;
  } catch (error) {
    console.error('Error fetching token:', error);
    
    if (error instanceof TypeError && error.message.includes('fetch')) {
      throw new Error('Cannot connect to backend server. Please ensure the Flask backend is running on http://127.0.0.1:5000');
    }
    
    throw new Error(`Failed to fetch token: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
};

export const getUserList = async (): Promise<any[]> => {
  try {
    const res = await fetchWithRetry(`${API_BASE}/users`);
    const data = await res.json();
    
    if (data.error) {
      throw new Error(`Users API Error: ${data.error}`);
    }
    
    return data.value || [];
  } catch (error) {
    console.error('Error fetching users:', error);
    
    if (error instanceof TypeError && error.message.includes('fetch')) {
      throw new Error('Cannot connect to backend server. Please ensure the Flask backend is running on http://127.0.0.1:5000');
    }
    
    throw new Error(`Failed to fetch users: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
};

export const getMetrics = async (): Promise<any> => {
  try {
    const res = await fetchWithRetry(`${API_BASE}/metrics`);
    const data = await res.json();
    
    if (data.error) {
      throw new Error(`Metrics API Error: ${data.error}`);
    }
    
    return data;
  } catch (error) {
    console.error('Error fetching metrics:', error);
    throw new Error(`Failed to fetch metrics: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
};

export const healthCheck = async (): Promise<boolean> => {
  try {
    const res = await fetchWithRetry(`${API_BASE}/health`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
      },
    });
    
    if (res.ok) {
      const data = await res.json();
      return data.status === 'healthy';
    }
    
    return false;
  } catch (error) {
    console.error('Health check failed:', error);
    return false;
  }
};

// Enhanced health check with detailed status
export const getBackendStatus = async (): Promise<{
  isHealthy: boolean;
  message: string;
  details?: any;
}> => {
  try {
    const res = await fetch(`${API_BASE}/health`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
      },
    });
    
    if (res.ok) {
      const data = await res.json();
      return {
        isHealthy: data.status === 'healthy',
        message: data.message || 'Backend is running',
        details: data
      };
    }
    
    return {
      isHealthy: false,
      message: `Backend responded with status ${res.status}`,
      details: { status: res.status, statusText: res.statusText }
    };
  } catch (error) {
    return {
      isHealthy: false,
      message: 'Cannot connect to backend server',
      details: { error: error instanceof Error ? error.message : 'Unknown error' }
    };
  }
}; 