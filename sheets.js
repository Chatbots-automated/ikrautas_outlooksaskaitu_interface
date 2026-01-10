import { google } from 'googleapis';

function getGoogleAuth() {
  const clientEmail = process.env.GOOGLE_CLIENT_EMAIL;
  const privateKey = (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n');
  if (!clientEmail || !privateKey) {
    throw new Error('Missing GOOGLE_CLIENT_EMAIL / GOOGLE_PRIVATE_KEY');
  }

  return new google.auth.JWT({
    email: clientEmail,
    key: privateKey,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });
}

function indexByHeader(headers) {
  const map = {};
  headers.forEach((h, i) => {
    map[String(h || '').trim().toLowerCase()] = i;
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
  // expects "2025-12-22T..." -> "2025-12-22"
  return String(ts || '').slice(0, 10);
}

export async function getSheetKeysForDate({ spreadsheetId, recognizedTab, unrecognizedTab, date }) {
  const auth = getGoogleAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  async function readTab(tabName) {
    const range = `${tabName}!A:Z`;
    const resp = await sheets.spreadsheets.values.get({ spreadsheetId, range });
    const values = resp.data.values || [];
    if (values.length < 2) return { rows: [], headers: [] };
    const headers = values[0];
    const rows = values.slice(1);
    return { headers, rows };
  }

  const rec = await readTab(recognizedTab);
  const unr = await readTab(unrecognizedTab);

  function extract(tab) {
    const map = indexByHeader(tab.headers);

    const idxDate = pickIdx(map, ['data', 'date', 'timestamp']);
    const idxMsg  = pickIdx(map, ['messageid', 'message id', 'message_id']);
    const idxAtt  = pickIdx(map, ['attachmentid', 'attachment id', 'attachment_id']);
    const idxDup  = pickIdx(map, ['duplicate?', 'duplicate', 'dubliuotas', 'dublikatas']);
    const idxName = pickIdx(map, ['originalus failo pavadinimas', 'original filename', 'filename', 'file name', 'pavadinimas']);

    if (idxDate < 0 || idxMsg < 0 || idxAtt < 0) {
      throw new Error(`Sheet "${recognizedTab}" or "${unrecognizedTab}" missing required columns (Data, MessageId, AttachmentId).`);
    }

    const out = [];
    for (const r of tab.rows) {
      const ts = r[idxDate];
      if (datePrefix(ts) !== date) continue;

      const messageId = r[idxMsg] || '';
      const attachmentId = r[idxAtt] || '';
      if (!messageId || !attachmentId) continue;

      const dup = idxDup >= 0 ? String(r[idxDup] || '') : '';
      const name = idxName >= 0 ? String(r[idxName] || '') : '';

      out.push({
        messageId,
        attachmentId,
        duplicateFlag: dup,
        name,
      });
    }
    return out;
  }

  const recRows = extract(rec);
  const unrRows = extract(unr);

  function uniqByKey(rows) {
    const seen = new Set();
    const uniq = [];
    let dups = 0;

    for (const r of rows) {
      const k = `${r.messageId}::${r.attachmentId}`;
      if (seen.has(k)) {
        dups++;
        continue;
      }
      seen.add(k);
      uniq.push(r);
    }
    return { uniq, dups };
  }

  const recUniq = uniqByKey(recRows);
  const unrUniq = uniqByKey(unrRows);

  return {
    date,
    recognized_all: recRows,
    unrecognized_all: unrRows,
    recognized_unique: recUniq.uniq,
    unrecognized_unique: unrUniq.uniq,
    recognized_key_dups: recUniq.dups,
    unrecognized_key_dups: unrUniq.dups,
  };
}
