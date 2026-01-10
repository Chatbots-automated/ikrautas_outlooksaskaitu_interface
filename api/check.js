// Vercel Serverless Function: GET /api/check?date=YYYY-MM-DD&debug=1
// No external deps. Uses:
// - Microsoft Graph (client_credentials)
// - Google Sheets API (OAuth refresh_token flow)  <-- keep as-is for now

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
  const [y, m, d] = dateYYYYMMDD.split("-").map(Number);
  const start = new Date(Date.UTC(y, m - 1, d, 0, 0, 0));
  const end = new Date(Date.UTC(y, m - 1, d + 1, 0, 0, 0));
  return { start: start.toISOString(), end: end.toISOString() };
}

function safeTruncate(s, max = 800) {
  const str = String(s ?? "");
  return str.length > max ? str.slice(0, max) + "…(truncated)" : str;
}

function nowIso() {
  return new Date().toISOString();
}

function decodeJwtClaims(accessToken) {
  // ONLY for debugging roles/scp/tid/appid - no validation
  try {
    const parts = String(accessToken || "").split(".");
    if (parts.length < 2) return null;
    const payload = parts[1].replace(/-/g, "+").replace(/_/g, "/");
    const json = Buffer.from(payload, "base64").toString("utf8");
    const claims = JSON.parse(json);
    // return only relevant fields
    return {
      aud: claims.aud,
      iss: claims.iss,
      tid: claims.tid,
      appid: claims.appid,
      roles: claims.roles || null,
      scp: claims.scp || null,
      exp: claims.exp || null,
    };
  } catch {
    return null;
  }
}

function randId() {
  // Vercel/node supports crypto.randomUUID in modern runtimes
  try {
    return crypto.randomUUID();
  } catch {
    return "req_" + Math.random().toString(16).slice(2);
  }
}

async function fetchWithMeta(url, opts = {}) {
  const r = await fetch(url, opts);
  const text = await r.text();
  const headers = {};
  r.headers.forEach((v, k) => (headers[k.toLowerCase()] = v));
  let json = null;
  try {
    json = text ? JSON.parse(text) : {};
  } catch {
    json = null;
  }
  return {
    ok: r.ok,
    status: r.status,
    statusText: r.statusText,
    url,
    headers,
    text,
    json,
  };
}

async function fetchJsonOrThrow(url, opts, logs, stepName) {
  const started = Date.now();
  const meta = await fetchWithMeta(url, opts);

  logs.push({
    t: nowIso(),
    step: stepName,
    url,
    status: meta.status,
    ok: meta.ok,
    request_id: meta.headers["request-id"] || null,
    client_request_id: opts?.headers?.["client-request-id"] || null,
    duration_ms: Date.now() - started,
  });

  if (!meta.ok) {
    // include body snippet for debugging
    const body = meta.json ?? meta.text;
    throw new Error(
      `${meta.status} ${meta.statusText}: ${safeTruncate(
        typeof body === "string" ? body : JSON.stringify(body)
      )}`
    );
  }

  // prefer json if parse ok, else throw because graph should be json for our calls
  const out = meta.json ?? (() => {
    throw new Error(`Non-JSON response: ${safeTruncate(meta.text)}`);
  })();

  return out;
}

async function fetchAllPagedGraph(token, firstUrl, logs, stepPrefix) {
  const out = [];
  let url = firstUrl;
  let page = 0;

  while (url) {
    page += 1;
    const clientRequestId = randId();

    const json = await fetchJsonOrThrow(
      url,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json",
          "client-request-id": clientRequestId,
        },
      },
      logs,
      `${stepPrefix}.page_${page}`
    );

    if (Array.isArray(json.value)) out.push(...json.value);
    url = json["@odata.nextLink"] || null;

    logs.push({
      t: nowIso(),
      step: `${stepPrefix}.page_${page}.count`,
      items_in_page: Array.isArray(json.value) ? json.value.length : 0,
      total_so_far: out.length,
    });
  }

  return out;
}

// ---------- Microsoft Graph ----------
async function getGraphToken(logs) {
  const tenant = process.env.MS_TENANT_ID;
  const clientId = process.env.MS_CLIENT_ID;
  const clientSecret = process.env.MS_CLIENT_SECRET;

  if (!tenant || !clientId || !clientSecret) {
    throw new Error("Missing MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET");
  }

  const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(
    tenant
  )}/oauth2/v2.0/token`;

  const body = new URLSearchParams();
  body.set("grant_type", "client_credentials");
  body.set("client_id", clientId);
  body.set("client_secret", clientSecret);
  body.set("scope", "https://graph.microsoft.com/.default");

  const started = Date.now();
  const meta = await fetchWithMeta(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  logs.push({
    t: nowIso(),
    step: "graph.token",
    url: tokenUrl,
    status: meta.status,
    ok: meta.ok,
    duration_ms: Date.now() - started,
  });

  if (!meta.ok) {
    throw new Error(`Token error ${meta.status}: ${safeTruncate(meta.text)}`);
  }

  const json = meta.json || JSON.parse(meta.text || "{}");
  const token = json.access_token;

  const claims = decodeJwtClaims(token);
  logs.push({
    t: nowIso(),
    step: "graph.token.claims",
    claims,
    note:
      "For app-only access, you must see roles like Mail.Read (Application) in claims.roles. If roles is null -> permissions not granted.",
  });

  return token;
}

async function getExpectedOutlookKeys({ mailbox, date, logs }) {
  logs.push({ t: nowIso(), step: "outlook.start", mailbox, date });

  const token = await getGraphToken(logs);
  const { start, end } = utcStartEndExclusive(date);

  logs.push({ t: nowIso(), step: "outlook.date_window_utc", start, end });

  const messagesUrl =
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}/messages` +
    `?$select=id,receivedDateTime,hasAttachments` +
    `&$filter=hasAttachments eq true and receivedDateTime ge ${start} and receivedDateTime lt ${end}` +
    `&$top=999`;

  logs.push({ t: nowIso(), step: "outlook.messages.url", messagesUrl });

  const messages = await fetchAllPagedGraph(token, messagesUrl, logs, "outlook.messages");
  logs.push({ t: nowIso(), step: "outlook.messages.total", count: messages.length });

  const expected = [];

  for (let i = 0; i < messages.length; i++) {
    const m = messages[i];
    const messageId = m.id;

    const attsUrl =
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}/messages/${encodeURIComponent(messageId)}/attachments` +
      `?$top=999`;

    logs.push({
      t: nowIso(),
      step: "outlook.attachments.url",
      index: i,
      messageId,
      attsUrl,
    });

    const attachments = await fetchAllPagedGraph(token, attsUrl, logs, `outlook.attachments.msg_${i + 1}`);

    // count raw attachments
    logs.push({
      t: nowIso(),
      step: "outlook.attachments.total_for_message",
      index: i,
      messageId,
      attachments_count: attachments.length,
    });

    for (const a of attachments) {
      const otype = a["@odata.type"] || "#microsoft.graph.fileAttachment";
      if (!String(otype).toLowerCase().endsWith("fileattachment")) continue;

      const pdf = isPdf(a);
      const inlineish = a.isInline === true || !!a.contentId;
      if (!pdf && inlineish && isImage(a)) continue;

      expected.push({
        messageId,
        attachmentId: a.id,
        name: a.name || "",
        contentType: a.contentType || "",
        isPdf: pdf,
        isInline: a.isInline === true,
      });
    }
  }

  // unique
  const seen = new Set();
  const uniq = [];
  for (const e of expected) {
    const k = `${e.messageId}::${e.attachmentId}`;
    if (seen.has(k)) continue;
    seen.add(k);
    uniq.push(e);
  }

  logs.push({
    t: nowIso(),
    step: "outlook.expected.unique_total",
    expected_raw: expected.length,
    expected_unique: uniq.length,
  });

  return { messages_count: messages.length, expected: uniq };
}

// ---------- Google Sheets via OAuth refresh_token ----------
async function getGoogleAccessToken(logs) {
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

  const started = Date.now();
  const meta = await fetchWithMeta(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  logs.push({
    t: nowIso(),
    step: "google.token",
    status: meta.status,
    ok: meta.ok,
    duration_ms: Date.now() - started,
  });

  if (!meta.ok) {
    throw new Error(`Google token error ${meta.status}: ${safeTruncate(meta.text)}`);
  }

  const json = meta.json || JSON.parse(meta.text || "{}");
  return json.access_token;
}

async function sheetsValuesGet(accessToken, spreadsheetId, rangeA1, logs) {
  const url =
    `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(rangeA1)}` +
    `?majorDimension=ROWS`;

  logs.push({ t: nowIso(), step: "sheets.values.url", url });

  return fetchJsonOrThrow(
    url,
    { headers: { Authorization: `Bearer ${accessToken}` } },
    logs,
    "sheets.values.get"
  );
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

async function getSheetKeysForDate({ spreadsheetId, recognizedTab, unrecognizedTab, date, logs }) {
  logs.push({ t: nowIso(), step: "sheets.start", spreadsheetId, recognizedTab, unrecognizedTab, date });

  const accessToken = await getGoogleAccessToken(logs);

  async function readTab(tabName) {
    const range = `${tabName}!A:Z`;
    const resp = await sheetsValuesGet(accessToken, spreadsheetId, range, logs);
    const values = resp.values || [];
    logs.push({ t: nowIso(), step: "sheets.tab.read", tabName, rows: values.length });
    if (values.length < 2) return { headers: [], rows: [] };
    return { headers: values[0], rows: values.slice(1) };
  }

  const rec = await readTab(recognizedTab);
  const unr = await readTab(unrecognizedTab);

  function extract(tab, tabLabel) {
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

    logs.push({
      t: nowIso(),
      step: "sheets.tab.columns",
      tab: tabLabel,
      idxDate,
      idxMsg,
      idxAtt,
      idxName,
      headers_preview: tab.headers.slice(0, 12),
    });

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

    logs.push({
      t: nowIso(),
      step: "sheets.tab.filtered",
      tab: tabLabel,
      filtered_raw: out.length,
      filtered_unique: uniq.length,
    });

    return uniq;
  }

  const recognized_unique = extract(rec, "recognized");
  const unrecognized_unique = extract(unr, "unrecognized");

  return { recognized_unique, unrecognized_unique };
}

// ---------- handler ----------
export default async function handler(req, res) {
  const debug = String(req.query.debug || "") === "1";
  const logs = [];

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

    logs.push({
      t: nowIso(),
      step: "env.summary",
      mailbox,
      sheet_id_present: !!process.env.SHEET_ID,
      ms_env_present: !!process.env.MS_TENANT_ID && !!process.env.MS_CLIENT_ID && !!process.env.MS_CLIENT_SECRET,
      google_env_present: !!process.env.GOOGLE_CLIENT_ID && !!process.env.GOOGLE_CLIENT_SECRET && !!process.env.GOOGLE_REFRESH_TOKEN,
    });

    // OUTLOOK
    const outlook = await getExpectedOutlookKeys({ mailbox, date, logs });

    // SHEETS
    const sheets = await getSheetKeysForDate({
      spreadsheetId,
      recognizedTab,
      unrecognizedTab,
      date,
      logs,
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
      ready,
      generated_at: new Date().toISOString(),
    };

    return res.status(200).json({
      ok: true,
      date,
      ready,
      messages_count: outlook.messages_count,
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
        isPdf: x.isPdf,
        isInline: x.isInline,
      })),
      extra,
      payload,
      ...(debug ? { logs } : {}),
    });
  } catch (err) {
    const errorMsg = String(err?.message || err);
    logs.push({ t: nowIso(), step: "error", message: errorMsg });

    return res.status(500).json({
      ok: false,
      error: errorMsg,
      ...(debug ? { logs } : {}),
    });
  }
}
