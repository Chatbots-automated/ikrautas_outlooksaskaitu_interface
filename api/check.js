// Vercel Serverless Function: GET /api/check?date=YYYY-MM-DD
// No external deps. Uses:
// - Microsoft Graph (client_credentials)
// - Google Sheets API (OAuth refresh_token flow)

// ---------- helpers ----------
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

function utcStartEndExclusive(dateYYYYMMDD) {
  // start inclusive, end exclusive (next day 00:00)
  const [y, m, d] = dateYYYYMMDD.split("-").map(Number);
  const start = new Date(Date.UTC(y, m - 1, d, 0, 0, 0));
  const end = new Date(Date.UTC(y, m - 1, d + 1, 0, 0, 0));
  return { start: start.toISOString(), end: end.toISOString() };
}

async function fetchJson(url, opts = {}) {
  const r = await fetch(url, opts);
  const text = await r.text();
  let json;
  try { json = text ? JSON.parse(text) : {}; } catch { json = { raw: text }; }
  if (!r.ok) throw new Error(`${r.status} ${r.statusText}: ${text}`);
  return json;
}

async function fetchAllPagedGraph(token, firstUrl) {
  const out = [];
  let url = firstUrl;
  while (url) {
    const json = await fetchJson(url, {
      headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
    });
    if (Array.isArray(json.value)) out.push(...json.value);
    url = json["@odata.nextLink"] || null;
  }
  return out;
}

// ---------- Microsoft Graph ----------
async function getGraphToken() {
  const tenant = process.env.MS_TENANT_ID;
  const clientId = process.env.MS_CLIENT_ID;
  const clientSecret = process.env.MS_CLIENT_SECRET;
  if (!tenant || !clientId || !clientSecret) {
    throw new Error("Missing MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET");
  }

  const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/token`;
  const body = new URLSearchParams();
  body.set("grant_type", "client_credentials");
  body.set("client_id", clientId);
  body.set("client_secret", clientSecret);
  body.set("scope", "https://graph.microsoft.com/.default");

  const r = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  const text = await r.text();
  if (!r.ok) throw new Error(`Token error ${r.status}: ${text}`);
  const json = JSON.parse(text);
  return json.access_token;
}

async function getExpectedOutlookKeys({ mailbox, date }) {
  const token = await getGraphToken();
  const { start, end } = utcStartEndExclusive(date);

  // Filter: hasAttachments and received in [start, end)
  const messagesUrl =
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}/messages` +
    `?$select=id,receivedDateTime,hasAttachments` +
    `&$filter=hasAttachments eq true and receivedDateTime ge ${start} and receivedDateTime lt ${end}` +
    `&$top=999`;

  const messages = await fetchAllPagedGraph(token, messagesUrl);

  const expected = [];

  for (const m of messages) {
    const messageId = m.id;

    const attsUrl =
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}/messages/${encodeURIComponent(messageId)}/attachments` +
      `?$top=999`;

    const attachments = await fetchAllPagedGraph(token, attsUrl);

    for (const a of attachments) {
      const otype = a["@odata.type"] || "#microsoft.graph.fileAttachment";
      if (!String(otype).toLowerCase().endsWith("fileattachment")) continue;

      // RULES (match your n8n logic):
      // - include PDFs
      // - include non-PDF unless (inline-ish AND image)
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

// ---------- Google Sheets via OAuth refresh_token ----------
async function getGoogleAccessToken() {
  const clientId = process.env.GOOGLE_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
  const refreshToken = process.env.GOOGLE_REFRESH_TOKEN;

  if (!clientId || !clientSecret || !refreshToken) {
    throw new Error("Missing GOOGLE_CLIENT_ID / GOOGLE_CLIENT_SECRET / GOOGLE_REFRESH_TOKEN");
  }

  const tokenUrl = "https://oauth2.googleapis.com/token";
  const body = new URLSearchParams();
  body.set("client_id", clientId);
  body.set("client_secret", clientSecret);
  body.set("refresh_token", refreshToken);
  body.set("grant_type", "refresh_token");

  const r = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  const text = await r.text();
  if (!r.ok) throw new Error(`Google token error ${r.status}: ${text}`);
  const json = JSON.parse(text);
  return json.access_token;
}

async function sheetsValuesGet(accessToken, spreadsheetId, rangeA1) {
  const url =
    `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(rangeA1)}` +
    `?majorDimension=ROWS`;

  return fetchJson(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
}

function indexByHeader(headers) {
  const map = {};
  headers.forEach((h, i) => (map[String(h || "").trim().toLowerCase()] = i));
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
  const accessToken = await getGoogleAccessToken();

  async function readTab(tabName) {
    const range = `${tabName}!A:Z`;
    const resp = await sheetsValuesGet(accessToken, spreadsheetId, range);
    const values = resp.values || [];
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

  const recognized_unique = extract(rec);
  const unrecognized_unique = extract(unr);

  return { recognized_unique, unrecognized_unique };
}

// ---------- handler ----------
export default async function handler(req, res) {
  try {
    const date = String(req.query.date || "").trim();
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

    const expectedSet = new Map();
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

    return res.status(200).json({
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
    return res.status(500).json({ error: String(err?.message || err) });
  }
}
