// Vercel Serverless Function: GET /api/check?date=YYYY-MM-DD&debug=1
// Uses:
// - Microsoft Graph: prefers DELEGATED token from Authorization header,
//   falls back to client_credentials if no Authorization header.
// - Google Sheets API: Service Account JWT (GOOGLE_SA_CLIENT_EMAIL + GOOGLE_SA_PRIVATE_KEY)

import crypto from "crypto";

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

function safeTruncate(s, max = 1200) {
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
    throw new Error(
      `${meta.status} ${meta.statusText}: ${safeTruncate(
        typeof body === "string" ? body : JSON.stringify(body)
      )}`
    );
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

  const baseUsers = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}`;
  const baseMe = `https://graph.microsoft.com/v1.0/me`;

  const buildMessagesUrl = (base) =>
    `${base}/messages?$select=id,receivedDateTime,hasAttachments&$filter=hasAttachments eq true and receivedDateTime ge ${start} and receivedDateTime lt ${end}&$top=999`;

  let usedBase = baseUsers;
  let messagesUrl = buildMessagesUrl(usedBase);
  logs.push({ t: nowIso(), step: "outlook.messages.url", messagesUrl, usedBase });

  let messages;
  try {
    messages = await fetchAllPagedGraph(token, messagesUrl, logs, "outlook.messages");
  } catch (e) {
    const msg = String(e?.message || e);
    if (mode === "delegated" && msg.includes("403")) {
      usedBase = baseMe;
      messagesUrl = buildMessagesUrl(usedBase);
      logs.push({ t: nowIso(), step: "outlook.messages.fallback_to_me", reason: safeTruncate(msg), messagesUrl });
      messages = await fetchAllPagedGraph(token, messagesUrl, logs, "outlook.messages");
    } else {
      throw e;
    }
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

  return { messages_count: messages.length, expected: uniq, usedBase };
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

function datePrefix(ts) {
  return String(ts || "").slice(0, 10);
}

function isDuplicateFlagCell(v) {
  const s = String(v ?? "").trim().toLowerCase();
  if (!s) return false;
  if (s === "duplicate") return true;
  if (s === "dup") return true;
  if (s === "dublikatas") return true;
  if (s === "dublis") return true;
  if (s === "1" || s === "true" || s === "yes" || s === "y") return true;
  if (s.includes("duplicate")) return true;
  if (s.includes("dublik")) return true;
  if (s.includes("dubl")) return true;
  return false;
}

async function getSheetKeysForDate({ spreadsheetId, recognizedTab, unrecognizedTab, date, logs }) {
  logs.push({ t: nowIso(), step: "sheets.start", spreadsheetId, recognizedTab, unrecognizedTab, date });

  const accessToken = await getGoogleAccessTokenServiceAccount(logs);

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
    const idxMsg = pickIdx(map, ["messageid", "message id", "message_id"]);
    const idxAtt = pickIdx(map, ["attachmentid", "attachment id", "attachment_id"]);
    const idxName = pickIdx(map, [
      "originalus failo pavadinimas",
      "original filename",
      "filename",
      "file name",
      "pavadinimas",
    ]);

    // IMPORTANT: detect the "duplicate" marker column used by your workflow
    const idxDup = pickIdx(map, [
      "duplicate",
      "duplicates",
      "duplicate flag",
      "duplicate_flag",
      "dup",
      "dublikatas",
      "dublis",
      "status",
      "flag",
      "žyma",
      "zyme",
    ]);

    logs.push({
      t: nowIso(),
      step: "sheets.tab.columns",
      tab: tabLabel,
      idxDate,
      idxMsg,
      idxAtt,
      idxName,
      idxDup,
      headers_preview: tab.headers.slice(0, 12),
    });

    if (idxDate < 0 || idxMsg < 0 || idxAtt < 0) {
      throw new Error("Sheets must contain columns: Data, MessageId, AttachmentId");
    }

    // We track:
    // - raw rows for date (including duplicates flagged)
    // - duplicates flagged by the DUPLICATE column
    // - non-duplicate rows (raw - flagged)
    // - keys used for matching (MessageId+AttachmentId unique), INCLUDING duplicates flagged
    // - keys non-duplicate only (excluding flagged duplicates) - this matches your "total=18" sheet summary
    let rawRowsForDate = 0;
    let duplicatesFlaggedRows = 0;

    const keyCountsAll = new Map();      // key -> sample
    const keyCountsNonDup = new Map();   // key -> sample (only rows NOT flagged duplicate)

    for (const r of tab.rows) {
      const ts = r[idxDate];
      if (datePrefix(ts) !== date) continue;

      rawRowsForDate += 1;

      const dupCell = idxDup >= 0 ? r[idxDup] : "";
      const isDupFlagged = idxDup >= 0 ? isDuplicateFlagCell(dupCell) : false;
      if (isDupFlagged) duplicatesFlaggedRows += 1;

      const messageId = r[idxMsg] || "";
      const attachmentId = r[idxAtt] || "";
      if (!messageId || !attachmentId) continue;

      const sample = {
        messageId,
        attachmentId,
        name: idxName >= 0 ? String(r[idxName] || "") : "",
      };

      const key = `${messageId}::${attachmentId}`;

      // all rows (including duplicates flagged)
      if (!keyCountsAll.has(key)) keyCountsAll.set(key, sample);

      // only non-duplicate-flagged rows
      if (!isDupFlagged) {
        if (!keyCountsNonDup.has(key)) keyCountsNonDup.set(key, sample);
      }
    }

    const unique_all = Array.from(keyCountsAll.values());
    const unique_nondup = Array.from(keyCountsNonDup.values());

    const nonDupRowsForDate = rawRowsForDate - duplicatesFlaggedRows;

    logs.push({
      t: nowIso(),
      step: "sheets.tab.filtered",
      tab: tabLabel,
      raw_rows_for_date: rawRowsForDate,
      duplicates_flagged_rows: duplicatesFlaggedRows,
      nondup_rows_for_date: nonDupRowsForDate,
      unique_all_keys: unique_all.length,
      unique_nondup_keys: unique_nondup.length,
    });

    return {
      unique_all,
      unique_nondup,
      raw_rows_for_date: rawRowsForDate,
      duplicates_flagged_rows: duplicatesFlaggedRows,
      nondup_rows_for_date: nonDupRowsForDate,
      has_dup_column: idxDup >= 0,
    };
  }

  const recognized = extract(rec, "recognized");
  const unrecognized = extract(unr, "unrecognized");

  return {
    // keys used for matching (includes duplicates flagged)
    recognized_unique_all: recognized.unique_all,
    unrecognized_unique_all: unrecognized.unique_all,

    // keys excluding duplicates flagged (matches your "total=18" sheet summary)
    recognized_unique_nondup: recognized.unique_nondup,
    unrecognized_unique_nondup: unrecognized.unique_nondup,

    // row counts
    recognized_total_rows: recognized.raw_rows_for_date,
    unrecognized_total_rows: unrecognized.raw_rows_for_date,
    recognized_duplicates_flagged: recognized.duplicates_flagged_rows,
    unrecognized_duplicates_flagged: unrecognized.duplicates_flagged_rows,
    recognized_nondup_rows: recognized.nondup_rows_for_date,
    unrecognized_nondup_rows: unrecognized.nondup_rows_for_date,

    sheets_total_rows: recognized.raw_rows_for_date + unrecognized.raw_rows_for_date,
    sheets_duplicates_flagged_total: recognized.duplicates_flagged_rows + unrecognized.duplicates_flagged_rows,
    sheets_nondup_rows_total: recognized.nondup_rows_for_date + unrecognized.nondup_rows_for_date,

    has_dup_column: recognized.has_dup_column || unrecognized.has_dup_column,
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

    // OUTLOOK expected keys (unique by messageId+attachmentId)
    const outlook = await getExpectedOutlookKeys({ mailbox, date, logs, token, mode });

    // SHEETS
    const sheets = await getSheetKeysForDate({ spreadsheetId, recognizedTab, unrecognizedTab, date, logs });

    // Use SHEETS UNIQUE ALL KEYS for matching (includes duplicate-flagged rows)
    const sheetsAll = [
      ...sheets.recognized_unique_all.map((x) => ({ ...x, source: "recognized" })),
      ...sheets.unrecognized_unique_all.map((x) => ({ ...x, source: "unrecognized" })),
    ];

    const expectedSet = new Map();
    for (const e of outlook.expected) expectedSet.set(`${e.messageId}::${e.attachmentId}`, e);

    const sheetsSet = new Map();
    for (const s of sheetsAll) sheetsSet.set(`${s.messageId}::${s.attachmentId}`, s);

    const missing = [];
    for (const [k, e] of expectedSet.entries()) if (!sheetsSet.has(k)) missing.push(e);

    const extra = [];
    for (const [k, s] of sheetsSet.entries()) if (!expectedSet.has(k)) extra.push(s);

    const ready = missing.length === 0 && extra.length === 0;

    // Nice human explanation
    const warnings = [];
    if (sheets.has_dup_column && (sheets.sheets_duplicates_flagged_total || 0) > 0) {
      warnings.push(
        `Sheets has ${sheets.sheets_duplicates_flagged_total} rows marked as DUPLICATE (recognized ${sheets.recognized_duplicates_flagged}, unrecognized ${sheets.unrecognized_duplicates_flagged}). ` +
        `These are business duplicates (flag column), not MessageId+AttachmentId collisions. Matching uses keys and still includes them.`
      );
    }

    const payload = {
      date,
      mailbox,

      // Matching counts (keys)
      expected_total: outlook.expected.length,
      sheets_keys_total: sheetsAll.length, // renamed for clarity
      recognized_keys: sheets.recognized_unique_all.length,
      unrecognized_keys: sheets.unrecognized_unique_all.length,

      // Sheet “business” totals (exclude DUPLICATE-flagged rows)
      sheets_total_rows: sheets.sheets_total_rows,
      sheets_duplicates_flagged_total: sheets.sheets_duplicates_flagged_total,
      sheets_nondup_rows_total: sheets.sheets_nondup_rows_total,
      recognized_nondup_rows: sheets.recognized_nondup_rows,
      unrecognized_nondup_rows: sheets.unrecognized_nondup_rows,
      recognized_duplicates_flagged: sheets.recognized_duplicates_flagged,
      unrecognized_duplicates_flagged: sheets.unrecognized_duplicates_flagged,

      // These match your summary tab meaning:
      recognized_total_nondup_keys: sheets.recognized_unique_nondup.length,
      unrecognized_total_nondup_keys: sheets.unrecognized_unique_nondup.length,

      // Ready logic
      missing_count: missing.length,
      extra_count: extra.length,
      ready,
      warnings,
      generated_at: new Date().toISOString(),
    };

    return res.status(200).json({
      ok: true,
      date,
      ready,
      mode,
      mailbox,
      usedBase: outlook.usedBase,
      messages_count: outlook.messages_count,

      // Matching counts (keys)
      expected_total: outlook.expected.length,
      sheets_keys_total: sheetsAll.length,
      recognized_keys: sheets.recognized_unique_all.length,
      unrecognized_keys: sheets.unrecognized_unique_all.length,

      // Business sheet clarity
      sheets_total_rows: sheets.sheets_total_rows,
      sheets_duplicates_flagged_total: sheets.sheets_duplicates_flagged_total,
      sheets_nondup_rows_total: sheets.sheets_nondup_rows_total,
      recognized_nondup_rows: sheets.recognized_nondup_rows,
      unrecognized_nondup_rows: sheets.unrecognized_nondup_rows,
      recognized_duplicates_flagged: sheets.recognized_duplicates_flagged,
      unrecognized_duplicates_flagged: sheets.unrecognized_duplicates_flagged,

      recognized_total_nondup_keys: sheets.recognized_unique_nondup.length,
      unrecognized_total_nondup_keys: sheets.unrecognized_unique_nondup.length,

      missing_count: missing.length,
      extra_count: extra.length,
      warnings,

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
