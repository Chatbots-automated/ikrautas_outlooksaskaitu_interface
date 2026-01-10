import express from "express";
import { google } from "googleapis";

/**
 * ENV REQUIRED:
 *  MS_TENANT_ID
 *  MS_CLIENT_ID
 *  MS_CLIENT_SECRET
 *  MAILBOX                     (default: invoices@ikrautas.lt)
 *
 *  SHEET_ID
 *  GOOGLE_CLIENT_EMAIL
 *  GOOGLE_PRIVATE_KEY          (keep \n as literal, we convert)
 *  SHEET_RECOGNIZED_TAB        (default: Atpažintos saskaitos)
 *  SHEET_UNRECOGNIZED_TAB      (default: Neatpažintos saskaitos)
 *
 *  N8N_WEBHOOK_URL             (for /api/trigger)
 */

const app = express();
app.use(express.json({ limit: "2mb" }));

// ---------------------------- UI (single HTML) ----------------------------
const HTML = `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Ikrautas – Batch Verifier</title>
  <style>
    body { font-family: ui-sans-serif, system-ui, Arial; margin: 24px; }
    .row { display:flex; gap:12px; align-items:center; flex-wrap:wrap; }
    input, button { padding:10px 12px; font-size:14px; }
    button { cursor:pointer; }
    .card { border:1px solid #ddd; border-radius:12px; padding:16px; margin-top:16px; }
    .ok { color: #0a7b2f; font-weight: 800; }
    .bad { color: #b00020; font-weight: 800; }
    .muted { color: #666; }
    table { width: 100%; border-collapse: collapse; margin-top: 12px; }
    th, td { border-bottom:1px solid #eee; padding:8px; text-align:left; font-size: 13px; }
    .pill { display:inline-block; padding:4px 10px; border-radius:999px; background:#f3f3f3; }
    pre { background:#0b0b0b; color:#e7e7e7; padding:12px; border-radius:10px; overflow:auto; }
    .btn-primary { background:#111; color:#fff; border: none; border-radius:10px; }
    .btn-secondary { background:#f3f3f3; border:1px solid #ddd; border-radius:10px; }
    .btn-primary:disabled { opacity:0.5; cursor:not-allowed; }
    .small { font-size:12px; }
  </style>
</head>

<body>
  <h1>Invoice Batch Verifier</h1>
  <p class="muted">Compares Outlook vs Sheets by <b>messageId + attachmentId</b>. GREEN only when <b>Missing=0</b> and <b>Extra=0</b>.</p>

  <div class="row">
    <label>Date (UTC):</label>
    <input id="date" type="date" />
    <button class="btn-secondary" id="btnCheck">Check batch</button>
    <button class="btn-primary" id="btnSend" disabled>Send to n8n</button>
    <span id="statusPill" class="pill muted">Not checked</span>
  </div>

  <div class="card" id="summaryCard" style="display:none;">
    <h3>Summary</h3>
    <div id="summary"></div>
    <h3>Payload</h3>
    <pre id="payload"></pre>
    <div class="muted small">This payload is what gets POSTed to your n8n webhook when you click “Send to n8n”.</div>
  </div>

  <div class="card" id="detailsCard" style="display:none;">
    <h3>Details</h3>
    <div id="details"></div>
  </div>

  <script>
    const dateEl = document.getElementById('date');
    const btnCheck = document.getElementById('btnCheck');
    const btnSend = document.getElementById('btnSend');
    const statusPill = document.getElementById('statusPill');
    const summaryCard = document.getElementById('summaryCard');
    const detailsCard = document.getElementById('detailsCard');
    const summaryEl = document.getElementById('summary');
    const payloadEl = document.getElementById('payload');
    const detailsEl = document.getElementById('details');

    const now = new Date();
    const yyyy = now.getUTCFullYear();
    const mm = String(now.getUTCMonth() + 1).padStart(2, '0');
    const dd = String(now.getUTCDate()).padStart(2, '0');
    dateEl.value = \`\${yyyy}-\${mm}-\${dd}\`;

    let lastCheck = null;

    function setStatus(text, ok=null) {
      statusPill.textContent = text;
      statusPill.className = 'pill ' + (ok === true ? 'ok' : ok === false ? 'bad' : 'muted');
    }

    btnCheck.onclick = async () => {
      btnSend.disabled = true;
      setStatus('Checking...', null);
      summaryCard.style.display = 'none';
      detailsCard.style.display = 'none';

      const date = dateEl.value;
      const res = await fetch(\`/api/check?date=\${encodeURIComponent(date)}\`);
      const data = await res.json();
      lastCheck = data;

      summaryCard.style.display = 'block';
      detailsCard.style.display = 'block';

      if (data.error) {
        setStatus('ERROR ❌', false);
        summaryEl.innerHTML = \`<div class="bad">\${escapeHtml(data.error)}</div>\`;
        payloadEl.textContent = '';
        detailsEl.innerHTML = '';
        return;
      }

      summaryEl.innerHTML = \`
        <div class="row">
          <div><b>Date:</b> \${data.date}</div>
          <div><b>Expected (Outlook):</b> \${data.expected_total}</div>
          <div><b>Sheets unique:</b> \${data.sheets_unique_total}</div>
          <div><b>Missing:</b> \${data.missing_count}</div>
          <div><b>Extra:</b> \${data.extra_count}</div>
        </div>
        <div class="row" style="margin-top:8px;">
          <div><b>Recognized unique:</b> \${data.recognized_unique}</div>
          <div><b>Unrecognized unique:</b> \${data.unrecognized_unique}</div>
          <div><b>Outlook messages:</b> \${data.messages_count}</div>
        </div>
      \`;

      payloadEl.textContent = JSON.stringify(data.payload, null, 2);

      detailsEl.innerHTML = \`
        <div class="row">
          <div><b>Status:</b> <span class="\${data.ready ? 'ok' : 'bad'}">\${data.ready ? 'READY ✅' : 'NOT READY ❌'}</span></div>
          <div class="muted">READY only when Missing=0 and Extra=0</div>
        </div>

        <h4>Missing (Outlook expected but not found in Sheets):</h4>
        \${renderKeyTable(data.missing)}

        <h4>Extra (in Sheets but not in Outlook expected set):</h4>
        \${renderKeyTable(data.extra)}
      \`;

      setStatus(data.ready ? 'READY ✅' : 'NOT READY ❌', data.ready);
      btnSend.disabled = !data.ready;
    };

    btnSend.onclick = async () => {
      if (!lastCheck?.ready) return;

      btnSend.disabled = true;
      setStatus('Sending...', null);

      const res = await fetch('/api/trigger', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(lastCheck.payload),
      });

      const out = await res.json().catch(() => ({}));
      if (out.ok) {
        setStatus('Sent ✅', true);
      } else {
        setStatus('Send failed ❌', false);
        alert(out.error || 'Send failed');
        btnSend.disabled = false;
      }
    };

    function renderKeyTable(list) {
      if (!list || list.length === 0) return '<div class="muted">None</div>';
      const rows = list.slice(0, 200).map(x => \`
        <tr>
          <td>\${escapeHtml(x.messageId || '')}</td>
          <td>\${escapeHtml(x.attachmentId || '')}</td>
          <td>\${escapeHtml(x.name || '')}</td>
          <td>\${escapeHtml(x.contentType || '')}</td>
        </tr>\`).join('');
      return \`
        <table>
          <thead><tr><th>messageId</th><th>attachmentId</th><th>name</th><th>contentType</th></tr></thead>
          <tbody>\${rows}</tbody>
        </table>
        \${list.length > 200 ? \`<div class="muted">Showing first 200 of \${list.length}</div>\` : ''}\`;
    }

    function escapeHtml(s) {
      return String(s).replace(/[&<>"']/g, c => ({
        '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
      }[c]));
    }
  </script>
</body>
</html>`;

app.get("/", (_, res) => res.set("Content-Type", "text/html").send(HTML));

// ---------------------------- Outlook (Graph) helpers ----------------------------
const IMAGE_EXT = /\.(png|jpe?g|gif|bmp|tiff?|webp|heic|heif|svg)$/i;

function isPdf(att) {
  const ct = String(att.contentType || "").toLowerCase();
  const name = String(att.name || "");
  return ct.includes("application/pdf") || /\.pdf$/i.test(name);
}

function isImage(att) {
  const ct = String(att.contentType || "").toLowerCase();
  const name = String(att.name || "").toLowerCase();
  if (ct.startsWith("image/")) return true;
  return IMAGE_EXT.test(name);
}

function utcRangeForDate(dateYYYYMMDD) {
  return {
    start: `${dateYYYYMMDD}T00:00:00Z`,
    end: `${dateYYYYMMDD}T23:59:59Z`,
  };
}

async function getGraphToken() {
  const tenant = process.env.MS_TENANT_ID;
  const clientId = process.env.MS_CLIENT_ID;
  const clientSecret = process.env.MS_CLIENT_SECRET;

  if (!tenant || !clientId || !clientSecret) {
    throw new Error("Missing MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET");
  }

  const url = `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/token`;
  const body = new URLSearchParams();
  body.set("grant_type", "client_credentials");
  body.set("client_id", clientId);
  body.set("client_secret", clientSecret);
  body.set("scope", "https://graph.microsoft.com/.default");

  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  const json = await resp.json();
  if (!resp.ok) {
    throw new Error(`Token error ${resp.status}: ${JSON.stringify(json)}`);
  }
  return json.access_token;
}

async function graphGet(token, url) {
  const resp = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  });

  const json = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    throw new Error(`Graph GET ${resp.status}: ${JSON.stringify(json)}`);
  }
  return json;
}

async function fetchAllPaged(token, firstUrl) {
  const out = [];
  let url = firstUrl;
  while (url) {
    const json = await graphGet(token, url);
    if (Array.isArray(json.value)) out.push(...json.value);
    url = json["@odata.nextLink"] || null;
  }
  return out;
}

async function getExpectedOutlookKeys({ mailbox, date }) {
  const token = await getGraphToken();
  const { start, end } = utcRangeForDate(date);

  const messagesUrl =
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}/messages` +
    `?$select=id,receivedDateTime,hasAttachments` +
    `&$filter=hasAttachments eq true and receivedDateTime ge ${start} and receivedDateTime le ${end}` +
    `&$top=999`;

  const messages = await fetchAllPaged(token, messagesUrl);

  const expected = [];
  for (const m of messages) {
    const messageId = m.id;

    const attsUrl =
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}/messages/${encodeURIComponent(messageId)}/attachments` +
      `?$top=999`;

    const attachments = await fetchAllPaged(token, attsUrl);

    for (const a of attachments) {
      const otype = a["@odata.type"] || "#microsoft.graph.fileAttachment";
      if (!String(otype).toLowerCase().endsWith("fileattachment")) continue;

      // RULES (match your n8n):
      // - Include PDFs always
      // - Include non-PDF unless (inline-ish AND image)
      const pdf = isPdf(a);
      const inlineish = a.isInline === true || !!a.contentId;

      if (!pdf && inlineish && isImage(a)) continue;

      expected.push({
        messageId,
        attachmentId: a.id,
        name: a.name || "",
        contentType: a.contentType || "",
      });
    }
  }

  // unique by messageId::attachmentId
  const seen = new Set();
  const uniq = [];
  for (const e of expected) {
    const k = `${e.messageId}::${e.attachmentId}`;
    if (seen.has(k)) continue;
    seen.add(k);
    uniq.push(e);
  }

  return { messages_count: messages.length, expected: uniq };
}

// ---------------------------- Google Sheets helpers ----------------------------
function getGoogleAuth() {
  const clientEmail = process.env.GOOGLE_CLIENT_EMAIL;
  const privateKey = (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n");
  if (!clientEmail || !privateKey) throw new Error("Missing GOOGLE_CLIENT_EMAIL / GOOGLE_PRIVATE_KEY");

  return new google.auth.JWT({
    email: clientEmail,
    key: privateKey,
    scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
  });
}

function indexByHeader(headers) {
  const map = {};
  headers.forEach((h, i) => {
    map[String(h || "").trim().toLowerCase()] = i;
  });
  return map;
}

function pickIdx(map, names) {
  for (const n of names) {
    const idx = map[n.toLowerCase()];
    if (idx !== undefined) return idx;
  }
  return -1;
}

function datePrefix(ts) {
  return String(ts || "").slice(0, 10);
}

async function getSheetKeysForDate({ spreadsheetId, recognizedTab, unrecognizedTab, date }) {
  const auth = getGoogleAuth();
  const sheets = google.sheets({ version: "v4", auth });

  async function readTab(tabName) {
    const range = `${tabName}!A:Z`;
    const resp = await sheets.spreadsheets.values.get({ spreadsheetId, range });
    const values = resp.data.values || [];
    if (values.length < 2) return { headers: [], rows: [] };
    return { headers: values[0], rows: values.slice(1) };
  }

  const rec = await readTab(recognizedTab);
  const unr = await readTab(unrecognizedTab);

  function extract(tab) {
    const map = indexByHeader(tab.headers);

    const idxDate = pickIdx(map, ["data", "date", "timestamp"]);
    const idxMsg  = pickIdx(map, ["messageid", "message id", "message_id"]);
    const idxAtt  = pickIdx(map, ["attachmentid", "attachment id", "attachment_id"]);
    const idxName = pickIdx(map, [
      "originalus failo pavadinimas",
      "original filename",
      "filename",
      "file name",
      "pavadinimas",
    ]);

    if (idxDate < 0 || idxMsg < 0 || idxAtt < 0) {
      throw new Error("Sheets must contain columns: Data, MessageId, AttachmentId");
    }

    const out = [];
    for (const r of tab.rows) {
      const ts = r[idxDate];
      if (datePrefix(ts) !== date) continue;

      const messageId = r[idxMsg] || "";
      const attachmentId = r[idxAtt] || "";
      if (!messageId || !attachmentId) continue;

      out.push({
        messageId,
        attachmentId,
        name: idxName >= 0 ? String(r[idxName] || "") : "",
      });
    }

    // unique
    const seen = new Set();
    const uniq = [];
    for (const x of out) {
      const k = `${x.messageId}::${x.attachmentId}`;
      if (seen.has(k)) continue;
      seen.add(k);
      uniq.push(x);
    }

    return uniq;
  }

  const recUniq = extract(rec);
  const unrUniq = extract(unr);

  return {
    recognized_unique: recUniq,
    unrecognized_unique: unrUniq,
  };
}

// ---------------------------- API: check + trigger ----------------------------
app.get("/api/check", async (req, res) => {
  try {
    const date = String(req.query.date || "").trim(); // YYYY-MM-DD
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      return res.status(400).json({ error: "Invalid date. Use YYYY-MM-DD" });
    }

    const mailbox = process.env.MAILBOX || "invoices@ikrautas.lt";

    const spreadsheetId = process.env.SHEET_ID;
    const recognizedTab = process.env.SHEET_RECOGNIZED_TAB || "Atpažintos saskaitos";
    const unrecognizedTab = process.env.SHEET_UNRECOGNIZED_TAB || "Neatpažintos saskaitos";
    if (!spreadsheetId) throw new Error("Missing SHEET_ID");

    const outlook = await getExpectedOutlookKeys({ mailbox, date });
    const sheets = await getSheetKeysForDate({
      spreadsheetId,
      recognizedTab,
      unrecognizedTab,
      date,
    });

    const expectedSet = new Map(); // key -> metadata
    for (const e of outlook.expected) expectedSet.set(`${e.messageId}::${e.attachmentId}`, e);

    const sheetsAll = [
      ...sheets.recognized_unique.map((x) => ({ ...x, source: "recognized" })),
      ...sheets.unrecognized_unique.map((x) => ({ ...x, source: "unrecognized" })),
    ];

    const sheetsSet = new Map();
    for (const s of sheetsAll) sheetsSet.set(`${s.messageId}::${s.attachmentId}`, s);

    const missing = [];
    for (const [k, e] of expectedSet.entries()) if (!sheetsSet.has(k)) missing.push(e);

    const extra = [];
    for (const [k, s] of sheetsSet.entries()) if (!expectedSet.has(k)) extra.push(s);

    const ready = missing.length === 0 && extra.length === 0;

    const payload = {
      date,
      mailbox,
      expected_total: outlook.expected.length,
      sheets_unique_total: sheetsAll.length,
      recognized_unique: sheets.recognized_unique.length,
      unrecognized_unique: sheets.unrecognized_unique.length,
      missing_count: missing.length,
      extra_count: extra.length,
      missing: missing.map((x) => ({
        messageId: x.messageId,
        attachmentId: x.attachmentId,
        name: x.name,
        contentType: x.contentType,
      })),
      extra: extra.map((x) => ({
        messageId: x.messageId,
        attachmentId: x.attachmentId,
        name: x.name,
      })),
      ready,
      generated_at: new Date().toISOString(),
    };

    res.json({
      date,
      ready,
      messages_count: outlook.messages_count,
      expected_total: outlook.expected.length,
      sheets_unique_total: sheetsAll.length,
      recognized_unique: sheets.recognized_unique.length,
      unrecognized_unique: sheets.unrecognized_unique.length,
      missing_count: missing.length,
      extra_count: extra.length,
      missing,
      extra,
      payload,
    });
  } catch (err) {
    res.status(500).json({ error: String(err?.message || err) });
  }
});

app.post("/api/trigger", async (req, res) => {
  try {
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
    if (!r.ok) return res.status(502).json({ ok: false, error: `n8n ${r.status}: ${text}` });

    res.json({ ok: true, n8n_response: text });
  } catch (err) {
    res.status(500).json({ ok: false, error: String(err?.message || err) });
  }
});

// ---------------------------- start ----------------------------
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`UI running on http://localhost:${PORT}`));
