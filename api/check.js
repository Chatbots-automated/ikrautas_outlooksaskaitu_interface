// Vercel Serverless Function: GET /api/check?date=YYYY-MM-DD&debug=1
// - Microsoft Graph: prefers DELEGATED token from Authorization header,
//   falls back to client_credentials if no Authorization header.
// - Google Sheets API: Service Account JWT (GOOGLE_SA_CLIENT_EMAIL + GOOGLE_SA_PRIVATE_KEY)
//
// KEY FIXES VS YOUR VERSION:
// ✅ Graph attachments query uses $select to avoid downloading contentBytes (huge base64)
// ✅ Global deadline guard to prevent "booting forever"
// ✅ Safer fetch JSON parsing + content-type checks
// ✅ Conservative concurrency + paging caps
//
// Env:
// MAILBOX (optional, default invoices@ikrautas.lt)
// SHEET_ID (required)
// SHEET_RECOGNIZED_TAB (optional)
// SHEET_UNRECOGNIZED_TAB (optional)
// MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET (required for app_only)
// GOOGLE_SA_CLIENT_EMAIL, GOOGLE_SA_PRIVATE_KEY (required)
// FETCH_TIMEOUT_MS (optional, default 15000)
// ATTACHMENTS_CONCURRENCY (optional, default 4)
// MAX_TOTAL_MS (optional, default 25000)  <-- hard stop for whole function

import crypto from "crypto";

export const config = {
  api: {
    bodyParser: false,
  },
  // If your project supports it, this lets Vercel allow longer runs.
  // Not all setups honor this, but it doesn't hurt.
  maxDuration: 60,
};

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
  // Graph is fine with ISO; keep milliseconds (works), but you can drop if you want.
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

// Per-request timeout
const FETCH_TIMEOUT_MS = Number(process.env.FETCH_TIMEOUT_MS || 15000);
// Whole function hard cap (prevents “booting forever”)
const MAX_TOTAL_MS = Number(process.env.MAX_TOTAL_MS || 25000);

function pLimit(limit) {
  let active = 0;
  const queue = [];
  const next = () => {
    active = Math.max(0, active - 1);
    if (queue.length) queue.shift()();
  };
  return (fn) =>
    new Promise((resolve, reject) => {
      const run = () => {
        active++;
        Promise.resolve()
          .then(fn)
          .then((val) => {
            resolve(val);
            next();
          })
          .catch((err) => {
            reject(err);
            next();
          });
      };
      if (active < limit) run();
      else queue.push(run);
    });
}

function getBearerFromReq(req) {
  const auth = req.headers?.authorization || req.headers?.Authorization || "";
  const m = String(auth).match(/^Bearer\s+(.+)$/i);
  return m ? m[1].trim() : null;
}

function assertTime(deadlineTs, logs, step) {
  if (Date.now() > deadlineTs) {
    logs.push({
      t: nowIso(),
      step: step || "deadline.hit",
      message: `Exceeded MAX_TOTAL_MS=${MAX_TOTAL_MS}`,
    });
    const err = new Error(`Timeout: exceeded MAX_TOTAL_MS=${MAX_TOTAL_MS}`);
    err.__httpStatus = 504;
    throw err;
  }
}

async function fetchWithMeta(url, opts = {}, deadlineTs, logs, stepName) {
  assertTime(deadlineTs, logs, stepName ? `${stepName}.pre_fetch` : "pre_fetch");

  const controller = new AbortController();
  const id = setTimeout(() => controller.abort(), FETCH_TIMEOUT_MS);

  const started = Date.now();
  let r;

  try {
    r = await fetch(url, { ...opts, signal: controller.signal });
  } catch (e) {
    clearTimeout(id);
    return {
      ok: false,
      status: 0,
      statusText: e?.name === "AbortError" ? "Fetch timeout" : "Fetch error",
      url,
      headers: {},
      text: e?.message ? String(e.message) : String(e),
      json: null,
      duration_ms: Date.now() - started,
    };
  } finally {
    clearTimeout(id);
  }

  const headers = {};
  r.headers.forEach((v, k) => (headers[k.toLowerCase()] = v));

  const ct = String(headers["content-type"] || "");
  let text = "";
  let json = null;

  try {
    // Graph sometimes returns huge JSON — but we avoid it by $select on attachments.
    // Still parse carefully.
    if (ct.includes("application/json")) {
      json = await r.json();
    } else {
      text = await r.text();
      try {
        json = text ? JSON.parse(text) : null;
      } catch {
        json = null;
      }
    }
  } catch (e) {
    // Fallback: try text
    try {
      text = await r.text();
    } catch {}
    json = null;
    return {
      ok: false,
      status: r.status,
      statusText: `Parse error: ${String(e?.message || e)}`,
      url,
      headers,
      text: safeTruncate(text, 1500),
      json: null,
      duration_ms: Date.now() - started,
    };
  }

  return {
    ok: r.ok,
    status: r.status,
    statusText: r.statusText,
    url,
    headers,
    text: safeTruncate(text, 2000),
    json,
    duration_ms: Date.now() - started,
  };
}

async function fetchJsonOrThrow(url, opts, logs, stepName, deadlineTs) {
  const meta = await fetchWithMeta(url, opts, deadlineTs, logs, stepName);

  logs.push({
    t: nowIso(),
    step: stepName,
    url,
    status: meta.status,
    ok: meta.ok,
    request_id: meta.headers["request-id"] || null,
    client_request_id: opts?.headers?.["client-request-id"] || null,
    duration_ms: meta.duration_ms,
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
    err.__httpStatus = meta.status || 500;
    throw err;
  }

  if (meta.json == null) {
    const err = new Error(`Non-JSON response: ${safeTruncate(meta.text)}`);
    err.__httpStatus = 502;
    throw err;
  }

  return meta.json;
}

async function fetchAllPagedGraph(token, firstUrl, logs, stepPrefix, deadlineTs, maxPages = 12) {
  const out = [];
  let url = firstUrl;
  let page = 0;
  const seenLinks = new Set();

  while (url) {
    assertTime(deadlineTs, logs, `${stepPrefix}.deadline_check`);

    if (seenLinks.has(url)) {
      logs.push({ t: nowIso(), step: `${stepPrefix}.paging.loop_detected`, url });
      break;
    }
    seenLinks.add(url);

    page += 1;
    if (page > maxPages) {
      logs.push({ t: nowIso(), step: `${stepPrefix}.paging.page_cap_hit`, maxPages });
      break;
    }

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
      `${stepPrefix}.page_${page}`,
      deadlineTs
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
async function getGraphTokenAppOnly(logs, deadlineTs) {
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

  const meta = await fetchWithMeta(
    tokenUrl,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
    },
    deadlineTs,
    logs,
    "graph.token.app_only"
  );

  logs.push({
    t: nowIso(),
    step: "graph.token.app_only",
    url: tokenUrl,
    status: meta.status,
    ok: meta.ok,
    duration_ms: meta.duration_ms,
  });

  if (!meta.ok) throw new Error(`Token error ${meta.status}: ${safeTruncate(meta.text)}`);

  const token = meta.json?.access_token;
  if (!token) throw new Error("Token response missing access_token");

  logs.push({ t: nowIso(), step: "graph.token.app_only.claims", claims: decodeJwtClaims(token) });
  return token;
}

async function getExpectedOutlookKeys({ mailbox, date, logs, token, mode, deadlineTs }) {
  logs.push({ t: nowIso(), step: "outlook.start", mailbox, date, mode });

  const { start, end } = utcStartEndExclusive(date);
  logs.push({ t: nowIso(), step: "outlook.date_window_utc", start, end });

  // Always query target mailbox via /users/{mailbox}
  const usedBase = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}`;

  // Messages
  const messagesUrl =
    `${usedBase}/messages?$select=id,receivedDateTime,hasAttachments&` +
    `$filter=hasAttachments eq true and receivedDateTime ge ${start} and receivedDateTime lt ${end}&$top=100`;

  logs.push({ t: nowIso(), step: "outlook.messages.url", messagesUrl, usedBase });

  let messages;
  try {
    messages = await fetchAllPagedGraph(token, messagesUrl, logs, "outlook.messages", deadlineTs, 10);
  } catch (e) {
    const status = e?.__httpStatus;
    if (mode === "delegated" && (status === 401 || status === 403)) {
      const msg =
        `Delegated token cannot read mailbox "${mailbox}". ` +
        `Login with a user that has access to this shared mailbox (Full Access), ` +
        `and consent scopes like Mail.Read.Shared. Raw: ${String(e.message || e)}`;
      const err = new Error(msg);
      err.__httpStatus = status;
      throw err;
    }
    throw e;
  }

  logs.push({ t: nowIso(), step: "outlook.messages.total", count: messages.length });

  const expected = [];
  const limit = pLimit(Number(process.env.ATTACHMENTS_CONCURRENCY || 4));

  // ✅ Critical: $select excludes contentBytes so Graph doesn’t dump base64 attachment bodies into JSON.
  const attSelect = [
    "id",
    "name",
    "contentType",
    "isInline",
    "contentId",
    "size",
    "@odata.type",
  ].join(",");

  await Promise.all(
    messages.map((m, i) =>
      limit(async () => {
        assertTime(deadlineTs, logs, `outlook.attachments.msg_${i + 1}.deadline_check`);

        const messageId = m.id;
        const attsUrl =
          `${usedBase}/messages/${encodeURIComponent(messageId)}/attachments` +
          `?$select=${encodeURIComponent(attSelect)}&$top=50`;

        logs.push({ t: nowIso(), step: "outlook.attachments.url", index: i, messageId, attsUrl });

        const attachments = await fetchAllPagedGraph(
          token,
          attsUrl,
          logs,
          `outlook.attachments.msg_${i + 1}`,
          deadlineTs,
          6
        );

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
      })
    )
  );

  // De-dupe by key
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

async function getGoogleAccessTokenServiceAccount(logs, deadlineTs) {
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

  const meta = await fetchWithMeta(
    tokenUrl,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
    },
    deadlineTs,
    logs,
    "google.sa.token"
  );

  logs.push({
    t: nowIso(),
    step: "google.sa.token",
    status: meta.status,
    ok: meta.ok,
    duration_ms: meta.duration_ms,
  });

  if (!meta.ok) {
    throw new Error(`Google SA token error ${meta.status}: ${safeTruncate(meta.text)}`);
  }

  const accessToken = meta.json?.access_token;
  if (!accessToken) throw new Error("Google SA token response missing access_token");
  return accessToken;
}

async function sheetsValuesGet(accessToken, spreadsheetId, rangeA1, logs, deadlineTs) {
  const url =
    `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values/` +
    `${encodeURIComponent(rangeA1)}?majorDimension=ROWS`;

  logs.push({ t: nowIso(), step: "sheets.values.url", url });

  return fetchJsonOrThrow(
    url,
    { headers: { Authorization: `Bearer ${accessToken}` } },
    logs,
    "sheets.values.get",
    deadlineTs
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

// Robust: turn many date formats into YYYY-MM-DD (or null)
function normalizeToYMD(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  if (!s) return null;

  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);

  // Google serial date
  if (/^\d+(\.\d+)?$/.test(s)) {
    const n = Number(s);
    if (Number.isFinite(n) && n > 1000) {
      const ms = Math.round((n - 25569) * 86400 * 1000);
      const d = new Date(ms);
      if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
    }
  }

  // dd.mm.yyyy
  let m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);
  if (m) {
    const dd = String(m[1]).padStart(2, "0");
    const mm = String(m[2]).padStart(2, "0");
    const yy = m[3];
    return `${yy}-${mm}-${dd}`;
  }

  // dd/mm/yyyy
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

/**
 * If timestamp parses & is NOT selected date -> false
 * Only fallback scan when timestamp is blank/unparseable.
 */
function rowBelongsToDate({ row, idxDate, date, scanCellsCount = 12 }) {
  const v = idxDate >= 0 ? row[idxDate] : null;
  const ymd = normalizeToYMD(v);

  if (ymd === date) return true;
  if (ymd !== null) return false;

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

async function getSheetAttachmentsForDate({
  spreadsheetId,
  recognizedTab,
  unrecognizedTab,
  date,
  logs,
  deadlineTs,
}) {
  logs.push({ t: nowIso(), step: "sheets.start", spreadsheetId, recognizedTab, unrecognizedTab, date });

  const accessToken = await getGoogleAccessTokenServiceAccount(logs, deadlineTs);

  async function readTab(tabName) {
    assertTime(deadlineTs, logs, `sheets.tab.${tabName}.deadline_check`);
    const range = `${tabName}!A:ZZ`;
    const resp = await sheetsValuesGet(accessToken, spreadsheetId, range, logs, deadlineTs);
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
    sheets_total_rows_by_date,
    sheets_rows_missing_ids,
    unique_keys_total: unique.length,
    collisions_total: recognized.collisions + unrecognized.collisions,
  };
}

// ---------- handler ----------
export default async function handler(req, res) {
  res.setHeader("Cache-Control", "no-store");

  const debug = String(req.query.debug || "") === "1";
  const logs = [];
  const deadlineTs = Date.now() + MAX_TOTAL_MS;

  try {
    const date = String(req.query.date || "").trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      return res.status(400).json({ ok: false, error: "Invalid date. Use YYYY-MM-DD" });
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
      fetch_timeout_ms: FETCH_TIMEOUT_MS,
      max_total_ms: MAX_TOTAL_MS,
      attachments_concurrency: Number(process.env.ATTACHMENTS_CONCURRENCY || 4),
    });

    let token;
    if (bearer) {
      token = bearer;
      logs.push({ t: nowIso(), step: "graph.token.delegated.claims", claims: decodeJwtClaims(token) });
    } else {
      token = await getGraphTokenAppOnly(logs, deadlineTs);
    }

    // OUTLOOK
    const outlook = await getExpectedOutlookKeys({ mailbox, date, logs, token, mode, deadlineTs });

    // SHEETS
    const sheets = await getSheetAttachmentsForDate({
      spreadsheetId,
      recognizedTab,
      unrecognizedTab,
      date,
      logs,
      deadlineTs,
    });

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
      outlook_attachments_total: outlook.expected_unique.length,
      sheets_attachments_total: sheets.sheets_total_rows_by_date,
      recognized_attachments: sheets.recognized.totalRowsByDate,
      unrecognized_attachments: sheets.unrecognized.totalRowsByDate,
      missing_count: missing.length,
      extra_count: extra.length,
      ready,
      generated_at: new Date().toISOString(),
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
