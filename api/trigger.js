// Vercel Serverless Function: POST /api/trigger
export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ ok: false, error: "Use POST" });
    }

    const n8nUrl = process.env.N8N_WEBHOOK_URL;
    if (!n8nUrl) throw new Error("Missing N8N_WEBHOOK_URL");

    const body = req.body || {};
    if (!body.date || body.ready !== true) {
      return res.status(400).json({ ok: false, error: "Not ready / invalid payload" });
    }

    const r = await fetch(n8nUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });

    const text = await r.text();
    if (!r.ok) {
      return res.status(502).json({ ok: false, error: `n8n ${r.status}: ${text}` });
    }

    return res.status(200).json({ ok: true, n8n_response: text });
  } catch (err) {
    return res.status(500).json({ ok: false, error: String(err?.message || err) });
  }
}
