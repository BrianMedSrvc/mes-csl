// proxy-server/api/index.js

const express = require('express');
const fetch = require('node-fetch');
const app = express();

app.use(express.json());

// ✅ Allow GPT to call us
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-Requested-With');
  next();
});

app.post('/', async (req, res) => {
  const { command } = req.body;

  if (!command) {
    return res.status(400).json({ error: 'Missing command' });
  }

  try {
    const response = await fetch("https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec", {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ command })
    });

    const result = await response.json();
    res.json({ proxy_status: 'ok', google_response: result });
  } catch (err) {
    console.error('Proxy error:', err);
    res.status(500).json({ error: 'Proxy failed', detail: err.message });
  }
});

module.exports = app;
