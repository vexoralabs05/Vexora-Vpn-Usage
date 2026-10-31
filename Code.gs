/** Vexora Labs Customer Usage Portal — Google Apps Script backend.
 *  Deploy as Web App: Execute as Me; Who has access: Anyone.
 *  The browser receives usage fields only; the subscription URL is never returned.
 *  Phone-only lookup is intentionally simple but is NOT strong customer authentication.
 */
const CONFIG = {
  SPREADSHEET_ID: '1fzn7kAJQNFtG3TqhMfPRsQEXV8cqUP7HnfyAUzz01SY',
  HEADER_ROW: 1,
  // Restrict outbound subscription fetches to the provider host you verified.
  ALLOWED_SUBSCRIPTION_HOSTS: ['sg3.vpnlanka.store'],
  // Existing sheet columns: B=Contact, H=Subscribtion usage(link).
  FALLBACK_PHONE_COLUMN: 2,
  FALLBACK_LINK_COLUMN: 8,
  FAILURE_LIMIT: 6,
  LOCK_MINUTES: 10
};

function doGet(e) {
  const p = (e && e.parameter) || {};
  const callback = String(p.callback || '');
  const action = String(p.action || '');
  let result;
  try {
    if (action !== 'usage') throw new Error('Invalid request.');
    result = getCustomerUsage_(String(p.phone || ''));
  } catch (err) {
    result = { success: false, error: safeError_(err) };
  }
  const body = JSON.stringify(result);
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]{0,80}$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + body + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function getCustomerUsage_(phoneInput) {
  const phone = normalizePhone_(phoneInput);
  if (phone.length < 9 || phone.length > 15) {
    throw new Error('Enter a valid registered phone number.');
  }
  const cache = CacheService.getScriptCache();
  const key = 'vex_fail_' + digest_(phone);
  if (cache.get(key + '_lock')) throw new Error('Too many failed attempts. Try again in 10 minutes.');

  const row = findCustomerRow_(phone);
  if (!row) {
    const count = Number(cache.get(key) || 0) + 1;
    cache.put(key, String(count), CONFIG.LOCK_MINUTES * 60);
    if (count >= CONFIG.FAILURE_LIMIT) cache.put(key + '_lock', '1', CONFIG.LOCK_MINUTES * 60);
    throw new Error('No subscription was found for this phone number. Check the number or contact Vexora Labs.');
  }
  cache.remove(key);
  const subscriptionUrl = String(row.subscriptionUrl || '').trim();
  if (!subscriptionUrl) throw new Error('No subscription link is saved for this customer. Contact Vexora Labs.');
  const data = fetchAndParseSubscription_(subscriptionUrl, row);
  return { success: true, data: data };
}

function findCustomerRow_(phone) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheets = ss.getSheets();
  // Search all monthly tabs; skip hidden/empty tabs and locate columns by header labels.
  for (let s = sheets.length - 1; s >= 0; s--) {
    const sh = sheets[s];
    if (sh.getLastRow() <= CONFIG.HEADER_ROW) continue;
    const lastCol = Math.max(sh.getLastColumn(), CONFIG.FALLBACK_LINK_COLUMN);
    const headers = sh.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0].map(normalizeHeader_);
    const phoneCol = findHeader_(headers, ['contact','phone','phone number','mobile','mobile number']) || CONFIG.FALLBACK_PHONE_COLUMN;
    const linkCol = findHeader_(headers, ['subscribtion usage(link)','subscription usage(link)','subscription link','usage link','configuration link']) || CONFIG.FALLBACK_LINK_COLUMN;
    const nameCol = findHeader_(headers, ['name','customer name']) || 3;
    const values = sh.getRange(CONFIG.HEADER_ROW + 1, 1, sh.getLastRow() - CONFIG.HEADER_ROW, lastCol).getDisplayValues();
    for (let r = values.length - 1; r >= 0; r--) {
      if (normalizePhone_(values[r][phoneCol - 1]) !== phone) continue;
      return {
        phone: phone,
        subscriptionUrl: values[r][linkCol - 1],
        customerName: values[r][nameCol - 1] || 'Vexora Subscriber',
        sheetName: sh.getName()
      };
    }
  }
  return null;
}

function fetchAndParseSubscription_(url, row) {
  const match = url.match(/^https:\/\/([^/?#]+)(\/sub\/[A-Za-z0-9-]+)\/?$/i);
  const host = match ? match[1].split(':')[0].toLowerCase() : '';
  if (!match || CONFIG.ALLOWED_SUBSCRIPTION_HOSTS.indexOf(host) < 0) {
    throw new Error('The saved subscription link is not from an approved provider host.');
  }
  const response = UrlFetchApp.fetch(url + (url.indexOf('?') >= 0 ? '&' : '?') + '_vex=' + Date.now(), {
    method: 'get', muteHttpExceptions: true, followRedirects: true,
    headers: { 'Cache-Control': 'no-cache', 'Accept': 'text/html' }
  });
  if (response.getResponseCode() !== 200) throw new Error('Subscription server returned HTTP ' + response.getResponseCode() + '.');
  const html = response.getContentText();
  const template = html.match(/<template\b[^>]*\bid=["']pxn-data["'][^>]*>/i);
  if (!template) throw new Error('Could not read the subscription usage data. The provider page format may have changed.');
  const attrs = {};
  const attrRegex = /([\w-]+)=["']([^"']*)["']/g;
  let m;
  while ((m = attrRegex.exec(template[0])) !== null) attrs[m[1].toLowerCase()] = m[2];
  const statsMatch = html.match(/<script\b[^>]*\bid=["']pxn-stats-embed["'][^>]*>([\s\S]*?)<\/script>/i);
  let stats = {};
  if (statsMatch) { try { stats = JSON.parse(statsMatch[1].trim()); } catch (_) {} }
  const used = attrs['data-used'] || '—';
  const total = attrs['data-total'] || '—';
  const percent = percentFrom_(used, total);
  const expiry = Number(attrs['data-expire'] || 0);
  const lastOnline = Number(attrs['data-last-online'] || 0);
  const online = lastOnline > 0 && (Date.now() - lastOnline) < 120000;
  const enabled = String(attrs['data-enabled'] || '').toLowerCase() === 'true';
  const expireText = expiry ? Utilities.formatDate(new Date(expiry * 1000), Session.getScriptTimeZone(), 'MMM d, yyyy') : 'Never';
  const daysLeft = expiry ? Math.ceil((expiry * 1000 - Date.now()) / 86400000) : null;
  return {
    customerName: row.customerName || attrs['data-emails'] || 'Vexora Subscriber',
    subscriptionId: attrs['data-sid'] || '—',
    status: enabled ? 'Active' : 'Disabled', enabled: enabled,
    statusMessage: enabled ? 'Your subscription is active and working normally.' : 'This subscription is disabled. Contact support.',
    upload: attrs['data-upload'] || '—', download: attrs['data-download'] || '—', used: used,
    remaining: attrs['data-remained'] || '—', total: total, usagePercent: percent,
    expiryText: daysLeft === null ? expireText : (expireText + (daysLeft >= 0 ? ' · ' + daysLeft + ' days left' : ' · Expired')),
    lastOnlineText: relativeTime_(lastOnline), online: online,
    cpu: isFinite(Number(stats.cpu)) ? Number(stats.cpu) : null,
    ram: isFinite(Number(stats.ram)) ? Number(stats.ram) : null,
    uploadSpeed: speedText_(stats.net_out), downloadSpeed: speedText_(stats.net_in),
    provider: stats.isp || '—', region: stats.region || '—'
  };
}

function percentFrom_(used, total) {
  const u = parseUnit_(used), t = parseUnit_(total);
  return t > 0 ? Math.max(0, Math.min(100, Math.round(u / t * 1000) / 10)) : 0;
}
function parseUnit_(v) {
  const m = String(v || '').replace(/,/g, '').match(/^([\d.]+)\s*(B|KB|MB|GB|TB)?$/i);
  if (!m) return 0;
  const n = Number(m[1]), unit = (m[2] || 'B').toUpperCase();
  return n * ({B:1,KB:1024,MB:1048576,GB:1073741824,TB:1099511627776}[unit] || 1);
}
function speedText_(v) { const n = Number(v); return isFinite(n) ? (n >= 1024 ? (n/1024).toFixed(1)+' MB/s' : n.toFixed(0)+' KB/s') : '—'; }
function relativeTime_(ms) {
  if (!ms) return 'Never seen';
  const sec = Math.max(0, Math.floor((Date.now() - ms) / 1000));
  if (sec < 45) return 'Just now'; if (sec < 3600) return Math.floor(sec/60)+'m ago';
  if (sec < 86400) return Math.floor(sec/3600)+'h ago'; return Math.floor(sec/86400)+'d ago';
}
function normalizePhone_(v) { let n=String(v || '').replace(/\D/g, ''); if(n.length===11 && n.indexOf('94')===0) n='0'+n.slice(2); return n.replace(/^0+/, ''); }
function normalizeHeader_(v) { return String(v || '').toLowerCase().replace(/\s+/g, ' ').trim(); }
function findHeader_(headers, names) { for (let i=0;i<headers.length;i++) if (names.indexOf(headers[i]) >= 0) return i+1; return 0; }
function digest_(s) { return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s)).slice(0,24); }
function safeError_(err) { const msg=String(err && err.message || 'Request failed.'); return msg.length>180 ? msg.slice(0,180) : msg; }
