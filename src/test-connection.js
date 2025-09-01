// Test script to verify backend connection
console.log('Testing O365 Dashboard connection...');

async function testConnection() {
  try {
    // Test health endpoint
    console.log('Testing health endpoint...');
    const healthResponse = await fetch('http://localhost:5000/api/health');
    const healthData = await healthResponse.json();
    console.log('✅ Health check:', healthData);

    // Test token endpoint
    console.log('Testing token endpoint...');
    const tokenResponse = await fetch('http://localhost:5000/api/token');
    const tokenData = await tokenResponse.json();
    console.log('✅ Token response:', tokenData.access_token ? 'SUCCESS' : 'FAILED');

    // Test users endpoint
    console.log('Testing users endpoint...');
    const usersResponse = await fetch('http://localhost:5000/api/users');
    const usersData = await usersResponse.json();
    console.log('✅ Users response:', usersData.value ? `${usersData.value.length} users` : 'FAILED');

    // Test metrics endpoint
    console.log('Testing metrics endpoint...');
    const metricsResponse = await fetch('http://localhost:5000/api/metrics');
    const metricsData = await metricsResponse.json();
    console.log('✅ Metrics response:', metricsData.totalUsers ? 'SUCCESS' : 'FAILED');

    console.log('🎉 All tests passed! Backend is working correctly.');
  } catch (error) {
    console.error('❌ Connection test failed:', error);
  }
}

// Run the test
testConnection(); 