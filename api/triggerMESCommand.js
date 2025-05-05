// Vercel Serverless Function Format (NO express)

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method Not Allowed' });
  }

  const { command } = req.body;

  if (!command) {
    return res.status(400).json({ error: 'Missing command' });
  }

  try {
    const response = await fetch("https://script.google.com/macros/s/AKfycbyHm-EjXQWoXSExo7_PZDPBT0XBgSprwg1v9sW4NmHrWmSCLEpepf7WlfuenOPE2NPQ/exec", {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ command }),
    });

    const result = await response.json();
    res.status(200).json({ proxy_status: 'ok', google_response: result });

  } catch (err) {
    console.error('Proxy error:', err);
    res.status(500).json({ error: 'Proxy failed', detail: err.message });
  }
}
