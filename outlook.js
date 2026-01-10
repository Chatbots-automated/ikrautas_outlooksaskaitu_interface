const IMAGE_EXT = /\.(png|jpe?g|gif|bmp|tiff?|webp|heic|heif|svg)$/i;

function isPdf(att) {
  const ct = String(att.contentType || '').toLowerCase();
  const name = String(att.name || '');
  return ct.includes('application/pdf') || /\.pdf$/i.test(name);
}

function isImage(att) {
  const ct = String(att.contentType || '').toLowerCase();
  const name = String(att.name || '').toLowerCase();
  if (ct.startsWith('image/')) return true;
  return IMAGE_EXT.test(name);
}

async function getAccessToken() {
  const tenant = process.env.MS_TENANT_ID;
  const clientId = process.env.MS_CLIENT_ID;
  const clientSecret = process.env.MS_CLIENT_SECRET;

  if (!tenant || !clientId || !clientSecret) {
    throw new Error('Missing MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET');
  }

  const url = `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/token`;
  const body = new URLSearchParams();
  body.set('grant_type', 'client_credentials');
  body.set('client_id', clientId);
  body.set('client_secret', clientSecret);
  body.set('scope', 'https://graph.microsoft.com/.default');

  const resp = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
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
      Accept: 'application/json',
    },
  });
  const json = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    throw new Error(`Graph GET ${resp.status}: ${JSON.stringify(json)}`);
  }
  return json;
}

function utcRangeForDate(dateYYYYMMDD) {
  const start = `${dateYYYYMMDD}T00:00:00Z`;
  const end = `${dateYYYYMMDD}T23:59:59Z`;
  return { start, end };
}

async function fetchAllPaged(token, firstUrl) {
  const out = [];
  let url = firstUrl;
  while (url) {
    const json = await graphGet(token, url);
    if (Array.isArray(json.value)) out.push(...json.value);
    url = json['@odata.nextLink'] || null;
  }
  return out;
}

export async function getExpectedAttachmentsForDate({ mailbox, date }) {
  const token = await getAccessToken();
  const { start, end } = utcRangeForDate(date);

  // Messages for the date
  const messagesUrl =
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}/messages` +
    `?$select=id,receivedDateTime,hasAttachments` +
    `&$filter=hasAttachments eq true and receivedDateTime ge ${start} and receivedDateTime le ${end}` +
    `&$top=999`;

  const messages = await fetchAllPaged(token, messagesUrl);

  // For each message, pull attachments and apply your exact include/exclude rules
  const expected = [];
  for (const m of messages) {
    const messageId = m.id;
    const attsUrl =
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}/messages/${encodeURIComponent(messageId)}/attachments` +
      `?$top=999`;

    const attachments = await fetchAllPaged(token, attsUrl);

    for (const a of attachments) {
      const otype = a['@odata.type'] || '#microsoft.graph.fileAttachment';
      if (!String(otype).toLowerCase().endsWith('fileattachment')) continue;

      // Rule:
      // - Include PDFs always
      // - Include non-PDF unless it's an inline-ish IMAGE (signature/logo)
      const pdf = isPdf(a);
      const inlineish = a.isInline === true || !!a.contentId;

      if (!pdf) {
        if (inlineish && isImage(a)) continue;
      }

      expected.push({
        messageId,
        attachmentId: a.id,
        name: a.name || '',
        contentType: a.contentType || '',
        isInline: a.isInline === true,
        size: a.size ?? null,
        pdf,
      });
    }
  }

  // Create unique expected set by (messageId, attachmentId)
  const seen = new Set();
  const uniq = [];
  for (const e of expected) {
    const k = `${e.messageId}::${e.attachmentId}`;
    if (seen.has(k)) continue;
    seen.add(k);
    uniq.push(e);
  }

  return {
    date,
    mailbox,
    messages_count: messages.length,
    expected: uniq,
  };
}
