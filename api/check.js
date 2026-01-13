// Vercel Serverless Function: GET /api/check?date=YYYY-MM-DD&debug=1
// Uses:
// - Microsoft Graph: prefers DELEGATED token from Authorization header,
//   falls back to client_credentials if no Authorization header.
// - Google Sheets API: Service Account JWT (GOOGLE_SA_CLIENT_EMAIL + GOOGLE_SA_PRIVATE_KEY)

import crypto from "crypto";

const IMAGE_EXT = /\.(png|jpe?g|gif|bmp|tiff?|webp|heic|heif|svg)$/i;

// ---------- helpers ----------
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

function safeTruncate(s, max = 1200) {
  const str = String(s ?? "");
  return str.length > max ? str.slice(0, max) + "…(truncated)" : str;
}

function nowIso() {
  return new Date().toISOString();
}

function decodeJwtClaims(accessToken) {
  try {
    const parts = String(accessToken || "").split(".");
    if (parts.length < 2) return null;
    const payload = parts[1].replace(/-/g, "+").replace(/_/g, "/");
    const json = Buffer.from(payload, "base64").toString("utf8");
    const claims = JSON.parse(json);
    return {
      aud: claims.aud,
      iss: claims.iss,
      tid: claims.tid,
      appid: claims.appid,
      upn: claims.upn || null,
      preferred_username: claims.preferred_username || null,
      roles: claims.roles || null,
      scp: claims.scp || null,
      exp: claims.exp || null,
    };
  } catch {
    return null;
  }
}

function randId() {
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
  return { ok: r.ok, status: r.status, statusText: r.statusText, url, headers, text, json };
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
    "x-ms-ags-diagnostic": meta.headers["x-ms-ags-diagnostic"]
      ? safeTruncate(meta.headers["x-ms-ags-diagnostic"], 600)
      : null,
  });

  if (!meta.ok) {
    const body = meta.json ?? meta.text;
    const err = new Error(
      `${meta.status} ${meta.statusText}: ${safeTruncate(
        typeof body === "string" ? body : JSON.stringify(body)
      )}`
    );
    err.__httpStatus = meta.status;
    throw err;
  }

  const out =
    meta.json ??
    (() => {
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
async function getGraphTokenAppOnly(logs) {
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

  const started = Date.now();
  const meta = await fetchWithMeta(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  logs.push({
    t: nowIso(),
    step: "graph.token.app_only",
    url: tokenUrl,
    status: meta.status,
    ok: meta.ok,
    duration_ms: Date.now() - started,
  });

  if (!meta.ok) throw new Error(`Token error ${meta.status}: ${safeTruncate(meta.text)}`);

  const json = meta.json || JSON.parse(meta.text || "{}");
  const token = json.access_token;

  logs.push({ t: nowIso(), step: "graph.token.app_only.claims", claims: decodeJwtClaims(token) });

  return token;
}

function getBearerFromReq(req) {
  const auth = req.headers?.authorization || req.headers?.Authorization || "";
  const m = String(auth).match(/^Bearer\s+(.+)$/i);
  return m ? m[1].trim() : null;
}

async function getExpectedOutlookKeys({ mailbox, date, logs, token, mode }) {
  logs.push({ t: nowIso(), step: "outlook.start", mailbox, date, mode });

  const { start, end } = utcStartEndExclusive(date);
  logs.push({ t: nowIso(), step: "outlook.date_window_utc", start, end });

  // ✅ HARD RULE: ALWAYS query the target mailbox via /users/{mailbox}
  // No fallback to /me because it produces wrong data.
  const usedBase = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}`;

  const messagesUrl =
    `${usedBase}/messages?$select=id,receivedDateTime,hasAttachments&` +
    `$filter=hasAttachments eq true and receivedDateTime ge ${start} and receivedDateTime lt ${end}&$top=999`;

  logs.push({ t: nowIso(), step: "outlook.messages.url", messagesUrl, usedBase });

  let messages;
  try {
    messages = await fetchAllPagedGraph(token, messagesUrl, logs, "outlook.messages");
  } catch (e) {
    // ✅ If delegated token can’t read the shared mailbox -> fail loud with clear error
    const status = e?.__httpStatus;
    if (mode === "delegated" && (status === 401 || status === 403)) {
      const msg =
        `Delegated token cannot read mailbox "${mailbox}". ` +
        `You must login with a user that has access to this shared mailbox (Full Access), ` +
        `and consent scopes like Mail.Read.Shared. Raw: ${String(e.message || e)}`;
      const err = new Error(msg);
      err.__httpStatus = status;
      throw err;
    }
    throw e;
  }

  logs.push({ t: nowIso(), step: "outlook.messages.total", count: messages.length });

  const expected = [];

  for (let i = 0; i < messages.length; i++) {
    const m = messages[i];
    const messageId = m.id;

    const attsUrl = `${usedBase}/messages/${encodeURIComponent(messageId)}/attachments?$top=999`;
    logs.push({ t: nowIso(), step: "outlook.attachments.url", index: i, messageId, attsUrl });

    const attachments = await fetchAllPagedGraph(token, attsUrl, logs, `outlook.attachments.msg_${i + 1}`);

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

  // Outlook de-dupe by key
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
    usedBase,
  });

  return { messages_count: messages.length, expected_unique: uniq, expected_raw_total: expected.length, usedBase };
}

// ---------- Google Sheets via Service Account JWT ----------
function base64UrlEncode(bufOrStr) {
  const b = Buffer.isBuffer(bufOrStr) ? bufOrStr : Buffer.from(String(bufOrStr));
  return b
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}

function signJwtRS256({ header, payload, privateKeyPem }) {
  const encodedHeader = base64UrlEncode(JSON.stringify(header));
  const encodedPayload = base64UrlEncode(JSON.stringify(payload));
  const data = `${encodedHeader}.${encodedPayload}`;

  const signer = crypto.createSign("RSA-SHA256");
  signer.update(data);
  signer.end();

  const signature = signer.sign(privateKeyPem);
  const encodedSig = base64UrlEncode(signature);
  return `${data}.${encodedSig}`;
}

async function getGoogleAccessTokenServiceAccount(logs) {
  const clientEmail = process.env.GOOGLE_SA_CLIENT_EMAIL;
  let privateKey = process.env.GOOGLE_SA_PRIVATE_KEY;

  if (!clientEmail || !privateKey) {
    throw new Error("Missing GOOGLE_SA_CLIENT_EMAIL / GOOGLE_SA_PRIVATE_KEY");
  }

  privateKey = privateKey.replace(/\\n/g, "\n");

  const now = Math.floor(Date.now() / 1000);
  const tokenUrl = "https://oauth2.googleapis.com/token";

  const header = { alg: "RS256", typ: "JWT" };
  const payload = {
    iss: clientEmail,
    scope: "https://www.googleapis.com/auth/spreadsheets.readonly",
    aud: tokenUrl,
    iat: now,
    exp: now + 3600,
  };

  const assertion = signJwtRS256({ header, payload, privateKeyPem: privateKey });

  const body = new URLSearchParams();
  body.set("grant_type", "urn:ietf:params:oauth:grant-type:jwt-bearer");
  body.set("assertion", assertion);

  const started = Date.now();
  const meta = await fetchWithMeta(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  logs.push({
    t: nowIso(),
    step: "google.sa.token",
    status: meta.status,
    ok: meta.ok,
    duration_ms: Date.now() - started,
  });

  if (!meta.ok) {
    throw new Error(`Google SA token error ${meta.status}: ${safeTruncate(meta.text)}`);
  }

  const json = meta.json || JSON.parse(meta.text || "{}");
  if (!json.access_token) throw new Error("Google SA token response missing access_token");
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

function normalizeToYMD(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  if (!s) return null;

  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);

  if (/^\d+(\.\d+)?$/.test(s)) {
    const n = Number(s);
    if (Number.isFinite(n) && n > 1000) {
      const ms = Math.round((n - 25569) * 86400 * 1000);
      const d = new Date(ms);
      if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
    }
  }

  let m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);
  if (m) {
    const dd = String(m[1]).padStart(2, "0");
    const mm = String(m[2]).padStart(2, "0");
    const yy = m[3];
    return `${yy}-${mm}-${dd}`;
  }

  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) {
    const dd = String(m[1]).padStart(2, "0");
    const mm = String(m[2]).padStart(2, "0");
    const yy = m[3];
    return `${yy}-${mm}-${dd}`;
  }

  const t = Date.parse(s);
  if (!Number.isNaN(t)) return new Date(t).toISOString().slice(0, 10);

  return null;
}

function rowBelongsToDate({ row, idxDate, date, scanCellsCount = 12 }) {
  const v = idxDate >= 0 ? row[idxDate] : null;
  const ymd = normalizeToYMD(v);
  if (ymd === date) return true;

  const lim = Math.min(row.length, scanCellsCount);
  for (let i = 0; i < lim; i++) {
    const cell = String(row[i] ?? "");
    if (!cell) continue;
    if (cell.includes(date)) return true;
  }
  return false;
}

function isMarkedDuplicate(cell) {
  return String(cell || "").trim().toUpperCase() === "DUPLICATE";
}

async function getSheetAttachmentsForDate({ spreadsheetId, recognizedTab, unrecognizedTab, date, logs }) {
  logs.push({ t: nowIso(), step: "sheets.start", spreadsheetId, recognizedTab, unrecognizedTab, date });

  const accessToken = await getGoogleAccessTokenServiceAccount(logs);

  async function readTab(tabName) {
    // ✅ IMPORTANT: read wide range, because MessageId/AttachmentId might be past Z
    const range = `${tabName}!A:ZZ`;
    const resp = await sheetsValuesGet(accessToken, spreadsheetId, range, logs);
    const values = resp.values || [];
    logs.push({ t: nowIso(), step: "sheets.tab.read", tabName, rows: values.length, range });
    if (values.length < 2) return { headers: [], rows: [] };
    return { headers: values[0], rows: values.slice(1) };
  }

  const rec = await readTab(recognizedTab);
  const unr = await readTab(unrecognizedTab);

  function extract(tab, tabLabel) {
    const map = indexByHeader(tab.headers);

    const idxDate = pickIdx(map, ["data", "date", "timestamp"]);
    const idxMsg = pickIdx(map, ["messageid", "message id", "message_id"]);
    const idxAtt = pickIdx(map, ["attachmentid", "attachment id", "attachment_id"]);
    const idxDup = pickIdx(map, ["duplicate", "duplicates", "dup", "status"]);

    const idxName = pickIdx(map, [
      "originalus failo pavadinimas",
      "original filename",
      "filename",
      "file name",
      "pavadinimas",
      "failas",
    ]);

    logs.push({
      t: nowIso(),
      step: "sheets.tab.columns",
      tab: tabLabel,
      idxDate,
      idxMsg,
      idxAtt,
      idxDup,
      idxName,
      headers_count: tab.headers.length,
      headers_preview: tab.headers.slice(0, 30),
      headers_tail_preview: tab.headers.slice(Math.max(0, tab.headers.length - 10)),
    });

    if (idxDate < 0) {
      throw new Error("Sheets must contain a date column named Data/Date/Timestamp");
    }

    // We will count total rows by date EVEN if ids are missing
    let totalRowsByDate = 0;
    let rowsMissingIds = 0;

    const allWithIds = [];
    const markedDuplicates = [];

    for (const r of tab.rows) {
      if (!rowBelongsToDate({ row: r, idxDate, date })) continue;

      totalRowsByDate += 1;

      const messageId = idxMsg >= 0 ? (r[idxMsg] || "") : "";
      const attachmentId = idxAtt >= 0 ? (r[idxAtt] || "") : "";

      const name = idxName >= 0 ? String(r[idxName] || "") : "";

      const dupMarked = idxDup >= 0 ? isMarkedDuplicate(r[idxDup]) : false;
      if (dupMarked && messageId && attachmentId) {
        markedDuplicates.push({ messageId, attachmentId, name });
      }

      if (!messageId || !attachmentId) {
        rowsMissingIds += 1;
        continue;
      }

      allWithIds.push({ messageId, attachmentId, name });
    }

    // unique keys for matching
    const seen = new Set();
    const unique = [];
    for (const x of allWithIds) {
      const k = `${x.messageId}::${x.attachmentId}`;
      if (seen.has(k)) continue;
      seen.add(k);
      unique.push(x);
    }

    const collisions = allWithIds.length - unique.length;

    logs.push({
      t: nowIso(),
      step: "sheets.tab.filtered",
      tab: tabLabel,
      total_rows_by_date: totalRowsByDate,
      rows_missing_ids: rowsMissingIds,
      rows_with_ids: allWithIds.length,
      unique_keys: unique.length,
      duplicate_key_collisions: collisions,
      duplicate_marked_rows: markedDuplicates.length,
    });

    return {
      totalRowsByDate,
      rowsMissingIds,
      allWithIds,
      unique,
      collisions,
      markedDuplicatesCount: markedDuplicates.length,
    };
  }

  const recognized = extract(rec, "recognized");
  const unrecognized = extract(unr, "unrecognized");

  const allUniquePool = [...recognized.unique, ...unrecognized.unique];
  const seen = new Set();
  const unique = [];
  for (const x of allUniquePool) {
    const k = `${x.messageId}::${x.attachmentId}`;
    if (seen.has(k)) continue;
    seen.add(k);
    unique.push(x);
  }

  const sheets_total_rows_by_date = recognized.totalRowsByDate + unrecognized.totalRowsByDate;
  const sheets_rows_missing_ids = recognized.rowsMissingIds + unrecognized.rowsMissingIds;

  return {
    recognized,
    unrecognized,
    unique,
    sheets_total_rows_by_date,     // ✅ THIS matches your “total attachments” logic
    sheets_rows_missing_ids,       // diagnostics
    unique_keys_total: unique.length,
    collisions_total: (recognized.collisions + unrecognized.collisions),
  };
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

    const bearer = getBearerFromReq(req);
    const mode = bearer ? "delegated" : "app_only";

    logs.push({
      t: nowIso(),
      step: "env.summary",
      mailbox,
      mode,
      sheet_id_present: !!process.env.SHEET_ID,
      ms_env_present: !!process.env.MS_TENANT_ID && !!process.env.MS_CLIENT_ID && !!process.env.MS_CLIENT_SECRET,
      google_env_present: !!process.env.GOOGLE_SA_CLIENT_EMAIL && !!process.env.GOOGLE_SA_PRIVATE_KEY,
      auth_header_present: !!bearer,
    });

    let token;
    if (bearer) {
      token = bearer;
      logs.push({ t: nowIso(), step: "graph.token.delegated.claims", claims: decodeJwtClaims(token) });
    } else {
      token = await getGraphTokenAppOnly(logs);
    }

    // OUTLOOK (always /users/{mailbox})
    const outlook = await getExpectedOutlookKeys({ mailbox, date, logs, token, mode });

    // SHEETS
    const sheets = await getSheetAttachmentsForDate({ spreadsheetId, recognizedTab, unrecognizedTab, date, logs });

    // Matching on UNIQUE keys only
    const expectedSet = new Map();
    for (const e of outlook.expected_unique) expectedSet.set(`${e.messageId}::${e.attachmentId}`, e);

    const sheetsSet = new Map();
    for (const s of sheets.unique) sheetsSet.set(`${s.messageId}::${s.attachmentId}`, s);

    const missing = [];
    for (const [k, e] of expectedSet.entries()) if (!sheetsSet.has(k)) missing.push(e);

    const extra = [];
    for (const [k, s] of sheetsSet.entries()) if (!expectedSet.has(k)) extra.push(s);

    const ready = missing.length === 0 && extra.length === 0 && sheets.sheets_rows_missing_ids === 0;

    const payload = {
      date,
      mailbox,

      // totals shown in UI
      outlook_attachments_total: outlook.expected_unique.length,
      sheets_attachments_total: sheets.sheets_total_rows_by_date,
      recognized_attachments: sheets.recognized.totalRowsByDate,
      unrecognized_attachments: sheets.unrecognized.totalRowsByDate,

      // matching
      missing_count: missing.length,
      extra_count: extra.length,
      ready,
      generated_at: new Date().toISOString(),

      // diagnostics
      sheets_rows_missing_ids: sheets.sheets_rows_missing_ids,
      matching_keys_sheets_unique_total: sheets.unique_keys_total,
      matching_keys_duplicate_collisions_in_sheets: sheets.collisions_total,
      sheets_marked_duplicates_recognized: sheets.recognized.markedDuplicatesCount,
      sheets_marked_duplicates_unrecognized: sheets.unrecognized.markedDuplicatesCount,
      usedBase: outlook.usedBase,
    };

    return res.status(200).json({
      ok: true,
      date,
      ready,
      mode,
      mailbox,
      usedBase: outlook.usedBase,
      messages_count: outlook.messages_count,

      outlook_attachments_total: outlook.expected_unique.length,
      sheets_attachments_total: sheets.sheets_total_rows_by_date,
      recognized_attachments: sheets.recognized.totalRowsByDate,
      unrecognized_attachments: sheets.unrecognized.totalRowsByDate,

      sheets_rows_missing_ids: sheets.sheets_rows_missing_ids,

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
    const status = err?.__httpStatus || 500;
    logs.push({ t: nowIso(), step: "error", message: errorMsg, status });

    return res.status(status).json({
      ok: false,
      error: errorMsg,
      ...(debug ? { logs } : {}),
    });
  }
}
