/**
 * Barides Receipt List — fire claim receipt tracker
 *
 * How to deploy:
 * 1) Paste this file into the Apps Script project bound to your sheet.
 * 2) Fill in and run setR2Credentials() once, then delete the values from the source.
 * 3) Run setup() once.
 * 4) Deploy > New deployment > Web app  (Execute as: Me, Access: Anyone).
 * 5) Copy the /exec URL and paste it into SCRIPT_URL at the top of index.html
 *    (the form hosted on GitHub Pages).
 *
 * The GitHub Pages form talks to this script via fetch():
 *   GET  ?action=categories  -> JSON list of categories
 *   POST (JSON body)         -> appends a row, uploads photo to R2
 */

const SHEET_NAME  = 'Sheet1';   // the tab holding the receipts
const LISTS_SHEET = 'Lists';    // hidden tab that backs the Category dropdown

const HEADERS = [
  'Unique Reciept Id',
  'Date of purchase',
  'Photo Url',
  'Description',
  'Category',
  'Purchase Amount'
];

/** Category is prefixed with its claim group so duplicates (Clothing, Furniture) stay distinct. */
const CATEGORIES = [
  'Food & Dining — Restaurants',
  'Food & Dining — Takeout',
  'Food & Dining — Grocery expenses above normal grocery spending',
  'Food & Dining — Food delivery fees',

  'Transportation — Uber, Lyft, or taxis',

  'Essentials — Clothing',
  'Essentials — Shoes',
  'Essentials — Toiletries',
  'Essentials — Medication replacements',
  'Essentials — Basic household necessities',

  'Temporary Housing — Bedding and towels',
  'Temporary Housing — Kitchen supplies',
  'Temporary Housing — Small appliances',
  'Temporary Housing — Furniture for temporary housing',

  'Cleaning & Laundry — Laundromat',
  'Cleaning & Laundry — Dry cleaning',
  'Cleaning & Laundry — Smoke or soot cleaning of clothing',
  'Cleaning & Laundry — Cleaning supplies',

  'Storage & Moving — Storage unit',
  'Storage & Moving — Moving company',
  'Storage & Moving — Boxes and packing supplies',
  'Storage & Moving — Delivery charges',

  'Emergency Services — Board-up service',
  'Emergency Services — Tarping',
  'Emergency Services — Water removal',
  'Emergency Services — Emergency electrician or plumber',
  'Emergency Services — Locksmith',
  'Emergency Services — Temporary fencing or security',

  'Remediation — Smoke remediation',
  'Remediation — Soot removal',
  'Remediation — Water-damage remediation',
  'Remediation — Mold prevention',
  'Remediation — Contents cleaning',

  'Contents Lost — Furniture',
  'Contents Lost — Electronics',
  'Contents Lost — Clothing',
  'Contents Lost — Appliances',
  'Contents Lost — Kitchenware',
  'Contents Lost — Music or photography equipment',
  'Contents Lost — Other destroyed belongings',

  'Documents & Records — Replacement identification',
  'Documents & Records — Passport or birth-certificate fees',
  'Documents & Records — Printing, postage, and shipping',
  'Documents & Records — Notary fees',
  'Documents & Records — Copies of records',

  'School — Temporary school-related expenses',

  'Work & Business — Temporary workspace',
  'Work & Business — Replacement work equipment',
  'Work & Business — Internet installation',
  'Work & Business — Equipment rental',
  'Work & Business — Lost or damaged business inventory'
];

/* ------------------------------------------------------------------ setup */

function setup() {
  const ss = SpreadsheetApp.getActive();
  const warnings = [];

  // Receipts tab
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  try {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setFontColor('#ffffff')
      .setBackground('#274e13');
  } catch (e) {
    warnings.push('headers (' + e.message + ')');
  }
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 220); // id
  sh.setColumnWidth(2, 130); // date
  sh.setColumnWidth(3, 260); // photo url
  sh.setColumnWidth(4, 300); // description
  sh.setColumnWidth(5, 320); // category
  sh.setColumnWidth(6, 130); // amount

  // If the sheet was converted to a Table (typed columns), Sheets refuses
  // setNumberFormat / setDataValidation. Typed columns already format dates
  // and currency themselves, so we skip instead of crashing.
  const last = sh.getMaxRows() - 1;
  try {
    sh.getRange(2, 2, last, 1).setNumberFormat('yyyy-mm-dd');   // Date of purchase
    sh.getRange(2, 6, last, 1).setNumberFormat('$#,##0.00');    // Purchase Amount
  } catch (e) {
    warnings.push('number formats (typed/table columns handle this themselves)');
  }

  // Hidden list tab (a range beats an inline list — no length limits, easy to edit)
  const lists = ss.getSheetByName(LISTS_SHEET) || ss.insertSheet(LISTS_SHEET);
  lists.clear();
  lists.getRange(1, 1).setValue('Category').setFontWeight('bold');
  lists.getRange(2, 1, CATEGORIES.length, 1)
       .setValues(CATEGORIES.map(c => [c]));
  lists.hideSheet();

  try {
    const source = lists.getRange(2, 1, CATEGORIES.length, 1);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(source, true)
      .setAllowInvalid(false)
      .setHelpText('Pick a category from the list.')
      .build();
    sh.getRange(2, 5, last, 1).setDataValidation(rule);
  } catch (e) {
    warnings.push('category dropdown (typed/table column blocked it)');
  }

  ss.setActiveSheet(sh);
  const note = warnings.length
    ? 'Setup done with skips: ' + warnings.join('; ') +
      '. To restore full formatting, right-click the table name on Sheet1 and choose "Convert to unformatted range", then rerun setup().'
    : 'Setup complete — ' + CATEGORIES.length + ' categories loaded.';
  SpreadsheetApp.getActive().toast(note, 'Setup', 10);
  Logger.log(note);
}

/* -------------------------------------------------------------- JSON API */

/**
 * GET endpoint used by the GitHub Pages form.
 *   ?action=categories -> { ok, categories: [...] }
 *   anything else      -> { ok, status } health check
 */
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'categories') {
    return json_({ ok: true, categories: CATEGORIES });
  }
  return json_({ ok: true, status: 'Barides receipt API is running.' });
}

/**
 * POST endpoint used by the GitHub Pages form.
 * Body is a JSON string (sent as text/plain to avoid a CORS preflight):
 *   { date, description, category, amount, photo: { base64, name, mimeType } | null }
 */
function doPost(e) {
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'Bad request: body must be JSON.' });
  }
  return json_(addReceipt(payload));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Appends one receipt row (and uploads the photo to R2 if present).
 * Returns { ok, id, photoUrl } or { ok: false, error }.
 */
function addReceipt(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const id = 'RCPT-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss')
             + '-' + Math.random().toString(36).slice(2, 6).toUpperCase();

    let photoUrl = '';
    if (payload.photo && payload.photo.base64) {
      photoUrl = uploadToR2(payload.photo, id);
    }

    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    sh.appendRow([
      id,
      payload.date ? new Date(payload.date + 'T12:00:00') : new Date(),
      photoUrl,
      payload.description || '',
      payload.category || '',
      Number(payload.amount) || 0
    ]);

    return { ok: true, id: id, photoUrl: photoUrl };
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------- Cloudflare R2 (S3) */

const R2_PUBLIC_BASE = 'https://pub-8fa925d5f99d4ebfa509fd8ed8758839.r2.dev';

function r2Config_() {
  const p = PropertiesService.getScriptProperties();
  return {
    accountId: p.getProperty('R2_ACCOUNT_ID'),
    bucket:    p.getProperty('R2_BUCKET'),
    accessKey: p.getProperty('R2_ACCESS_KEY_ID'),
    secretKey: p.getProperty('R2_SECRET_ACCESS_KEY')
  };
}

/** Uploads a base64 photo straight to R2 with SigV4 and returns its public URL. */
function uploadToR2(photo, receiptId) {
  const cfg = r2Config_();
  if (!cfg.accountId || !cfg.bucket || !cfg.accessKey || !cfg.secretKey) {
    throw new Error('R2 credentials are missing. Run setR2Credentials() first.');
  }

  const ext = (photo.name && photo.name.indexOf('.') > -1)
    ? photo.name.split('.').pop().toLowerCase().replace(/[^a-z0-9]/g, '')
    : (photo.mimeType === 'image/png' ? 'png' : 'jpg');
  const key = 'receipts/' + receiptId + '.' + ext;

  const bytes = Utilities.base64Decode(photo.base64);
  const host  = cfg.accountId + '.r2.cloudflarestorage.com';
  const path  = '/' + cfg.bucket + '/' + key;
  const contentType = photo.mimeType || 'image/jpeg';

  const now       = new Date();
  const amzDate   = Utilities.formatDate(now, 'GMT', "yyyyMMdd'T'HHmmss'Z'");
  const dateStamp = amzDate.slice(0, 8);
  const payloadHash = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes));

  const canonicalHeaders =
      'content-type:' + contentType + '\n' +
      'host:' + host + '\n' +
      'x-amz-content-sha256:' + payloadHash + '\n' +
      'x-amz-date:' + amzDate + '\n';
  const signedHeaders = 'content-type;host;x-amz-content-sha256;x-amz-date';

  const canonicalRequest = [
    'PUT',
    path.split('/').map(encodeURIComponent).join('/'),
    '',
    canonicalHeaders,
    signedHeaders,
    payloadHash
  ].join('\n');

  const scope = dateStamp + '/auto/s3/aws4_request';
  const stringToSign = [
    'AWS4-HMAC-SHA256',
    amzDate,
    scope,
    hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, canonicalRequest, Utilities.Charset.UTF_8))
  ].join('\n');

  let k = Utilities.computeHmacSha256Signature(dateStamp, 'AWS4' + cfg.secretKey);
  k = Utilities.computeHmacSha256Signature(Utilities.newBlob('auto').getBytes(), k);
  k = Utilities.computeHmacSha256Signature(Utilities.newBlob('s3').getBytes(), k);
  k = Utilities.computeHmacSha256Signature(Utilities.newBlob('aws4_request').getBytes(), k);
  const signature = hex_(Utilities.computeHmacSha256Signature(Utilities.newBlob(stringToSign).getBytes(), k));

  const authorization = 'AWS4-HMAC-SHA256 Credential=' + cfg.accessKey + '/' + scope +
    ', SignedHeaders=' + signedHeaders + ', Signature=' + signature;

  const res = UrlFetchApp.fetch('https://' + host + path, {
    method: 'put',
    contentType: contentType,
    payload: bytes,
    headers: {
      'Authorization': authorization,
      'x-amz-date': amzDate,
      'x-amz-content-sha256': payloadHash
    },
    muteHttpExceptions: true
  });

  if (res.getResponseCode() >= 300) {
    throw new Error('R2 upload failed (' + res.getResponseCode() + '): ' + res.getContentText().slice(0, 300));
  }
  return R2_PUBLIC_BASE + '/' + key;
}

function hex_(bytes) {
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

/**
 * Run once, then DELETE the values from this function so keys aren't sitting in the source.
 * Create the token in Cloudflare: R2 > API > Manage API tokens > Object Read & Write.
 */
function setR2Credentials() {
  PropertiesService.getScriptProperties().setProperties({
    R2_ACCOUNT_ID: 'your-account-id',
    R2_BUCKET: 'your-bucket-name',
    R2_ACCESS_KEY_ID: 'your-access-key-id',
    R2_SECRET_ACCESS_KEY: 'your-secret-access-key'
  });
}
