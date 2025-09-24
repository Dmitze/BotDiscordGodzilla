// Improved health check script for Docker
const http = require('http');

// Check if the application is responding
const options = {
  hostname: 'localhost',
  port: 3000,
  path: '/health',
  method: 'GET',
  timeout: 5000
};

const request = http.request(options, (res) => {
  if (res.statusCode === 200) {
    console.log('Health check passed');
    process.exit(0);
  } else {
    console.log('Health check failed with status:', res.statusCode);
    process.exit(1);
  }
});

request.on('error', (err) => {
  console.log('Health check failed with error:', err.message);
  process.exit(1);
});

request.on('timeout', () => {
  console.log('Health check timed out');
  request.destroy();
  process.exit(1);
});

request.end();