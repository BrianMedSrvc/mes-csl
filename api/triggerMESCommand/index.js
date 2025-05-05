export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method Not Allowed' });
  }

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
    return res.status(200).json({ proxy_status: 'ok', google_response: result });

  } catch (err) {
    console.error('Proxy error:', err);
    return res.status(500).json({ error: 'Proxy failed', detail: err.message });
  }
}
