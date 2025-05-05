// /api/triggerMESCommand/index.js

export default async function handler(req, res) {
  // Only allow POST requests
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ error: 'Method Not Allowed' });
  }

  // Handle CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-Requested-With');

  // Handle preflight requests
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  const { command } = req.body;

  if (!command) {
    return res.status(400).json({ error: 'Missing command' });
  }

  try {
    const response = await fetch(
      'https://script.google.com/macros/s/AKfycbyHm-EjXQWoXSExo7_PZDPBT0XBgSprwg1v9sW4NmHrWmSCLEpepf7WlfuenOPE2NPQ/exec',
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ command })
      }
    );

    const result = await response.json();
    return res.status(200).json({
      proxy_status: 'ok',
      google_response: result
    });
  } catch (err) {
    console.error('Proxy error:', err);
    return res.status(500).json({
      error: 'Proxy failed',
      detail: err.message
    });
  }
}
