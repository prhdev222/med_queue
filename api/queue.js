/**
 * Vercel Serverless Proxy — ไม่ส่ง Apps Script URL ไปที่ browser
 * ใส่ URL ใน Vercel Env: QUEUE_APPSCRIPT_URL
 */
const APPSCRIPT_URL = process.env.QUEUE_APPSCRIPT_URL;

export default async function handler(req, res) {
  if (!APPSCRIPT_URL) {
    res.status(500).json({ error: 'QUEUE_APPSCRIPT_URL not configured' });
    return;
  }

  // CORS สำหรับเรียกจากหน้าเว็บ
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.method !== 'GET') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  try {
    const url = new URL(APPSCRIPT_URL);
    Object.entries(req.query || {}).forEach(([k, v]) => {
      if (v !== undefined && v !== '') url.searchParams.set(k, v);
    });
    const response = await fetch(url.toString(), { headers: { Accept: 'application/json' } });
    const data = await response.json();
    res.status(200).json(data);
  } catch (err) {
    console.error('Queue proxy error:', err);
    res.status(502).json({ error: 'Proxy error', message: err.message });
  }
}
