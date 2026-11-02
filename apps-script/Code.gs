// CODE.GS VERSION 8.0 - 2026-11-02 (paste this whole file, then Deploy > Manage deployments > New version; adds ?f=counts for the HSSE Tracker)
/**
 * Mentari HSSE Induction – receiver v7
 * 1) The sheet ROW is written first (small, fast, safe)
 * 2) The certificate copy is saved to Drive second (month folder), and its link is added to the row
 * 3) The certificate QR opens the induction app's own verification screen, which asks this script for the stored record as JSON (?v=..&c=..&f=json)
 * v5: the ID column keeps ONLY the last 4 characters (an old or cached page that still sends a full number is trimmed here too).
 * v6: optional MCU / health-certificate photo or PDF. Saved to a separate 'MCU' sub-folder (month folder), link in column 23 'MCU Link'. Never shown on the verification page.
 * v7: HSSE familiarisation sign-off. After the induction an HSSE officer opens the certificate QR, ticks the site-familiarisation checklist, enters the sticker number and their personal PIN.
 *     Officers and PINs live in the 'HSSE Officers' sheet tab (Name | PIN | Active). Each sign-off is written to the 'Familiarisation' tab and to columns 24-26 of the induction row, and shows on the verification screen.
 *     Run setupOfficers once (adds the HSSE team with random PINs), then emailPins. Forgotten PIN: the officer taps 'Email me my PIN' (goes to the email on file).
 * Replies {ok, rowSaved, imageSaved, link} so the app knows exactly what was saved.
 */
// Leave both EMPTY to let this script create its own Sheet and Drive folder in the account that runs it (first use), or paste the IDs of an existing Sheet / folder you own.
const ROOT_FOLDER_ID = '';
const SHEET_ID = '';
const SHEET_NAME = 'Induction Records';
const TZ = 'Asia/Jakarta';
const SCRIPT_VERSION = 'v8.0 (2026-11-02)';   // shown by the plain address, so you can see which code the live address runs
const MAX_IMAGE_CHARS = 4000000;
const MAX_MCU_CHARS = 3200000;                 // ~2.3 MB file limit for the MCU               // ~3 MB image limit

const HEADERS = ['Timestamp','Certificate No','Name','Company','ID/Passport (last 4)','Host/Contractor',
  'Host Contact Name','Host Contact Phone','Purpose','Details Certified','Health Cert Submitted',
  'Score','Total Questions','Percentage','Result','Attempts Taken','Start Time','Completion Time',
  'Duration','Language','Certificate Link','Verify Code','MCU Link','Familiarised By','Familiarised At','Sticker No'];
const COL_LINK = 21, COL_CODE = 22, COL_MCU = 23, COL_FAM_BY = 24, COL_FAM_AT = 25, COL_STICKER = 26;
const FAM_SHEET = 'Familiarisation', OFFICER_SHEET = 'HSSE Officers';
const FAM_HEADERS = ['Timestamp','Certificate No','Name','Company','ID/Passport (last 4)','HSSE Officer','Sticker No','Checklist'];
const OFFICER_SEED = [                          // the HSSE personnel who may sign off (edit the sheet tab afterwards, not this list)
  ['Ignatius Fernandes', 'Ignatius.Ambrosio@ecooils.com.my', 'Project HSSE Manager'],
  ['Putri Yuslina', 'H.PutriYuslina@ecooils.co.id', 'Operations HSSE Officer'],
  ['Hocine Boudida', 'Hocine.Boudida@ecooils.com.my', 'Project HSSE Engineer'],
  ['Ronore Lapat', 'ronore.lapat@ecooils.co.id', 'Project HSSE Engineer'],
  ['Agung Tri Muliawan', 'Agung.Muliawan@ecooils.co.id', 'Project HSSE Supervisor'],
  ['Irfan Amin', 'irfan.amin@ecooils.co.id', 'Shell Secondee (SM) - HSSE Manager'],
  ['Garnet B', 'Garnet.Balangalibun@ecooils.com.my', 'Shell Secondee (IND) - HSSE Manager']
];
const MAIL_REPLY_TO = 'Ignatius.Ambrosio@ecooils.com.my';      // replies to the PIN emails go to you
const MAIL_SENDER_NAME = 'Ignatius Fernandes - EcoOils HSSE';  // name the officer sees as the sender
const MAX_PIN_TRIES = 5, MAX_PIN_TRIES_ALL = 30, LOCK_SECS = 900;   // 5 wrong PINs lock that officer for 15 min; 30 in total pause everybody

/**
 * ONE-TIME IMPORT of the earlier induction records from the old Sheet into this one.
 * Run importOldRecords once from the editor. It matches columns by their header names, adds only certificate
 * numbers that are not already here (safe to run twice) and leaves the old Sheet untouched.
 */
const OLD_SHEET_ID = '1491t3_AvjvW_gT3mSfqkvFicKxFA8WhlD0ABtbahi7w';
function importOldRecords() {
  const old = SpreadsheetApp.openById(OLD_SHEET_ID);
  const src = old.getSheetByName(SHEET_NAME) || old.getSheets()[0];
  const sv = src.getDataRange().getValues();
  if (sv.length < 2) { Logger.log('Old sheet has no data rows.'); return; }
  const sh = sheet_();
  const dv = sh.getDataRange().getValues();
  const dh = dv[0].map(String), sh0 = sv[0].map(String);
  const dCert = dh.indexOf('Certificate No'), sCert = sh0.indexOf('Certificate No');
  if (dCert < 0 || sCert < 0) throw new Error('No "Certificate No" column found.');
  const have = {}; for (let i = 1; i < dv.length; i++) have[String(dv[i][dCert])] = true;
  const map = dh.map(function (h) { return sh0.indexOf(h); });          // new column -> old column (or -1)
  const out = [];
  for (let r = 1; r < sv.length; r++) {
    const c = String(sv[r][sCert] || '');
    if (!c || have[c]) continue;
    have[c] = true;
    out.push(map.map(function (k) { return k < 0 ? '' : sv[r][k]; }));
  }
  if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, dh.length).setValues(out);
  Logger.log('Imported ' + out.length + ' record(s); skipped ' + (sv.length - 1 - out.length) + ' already present or blank.');
}

/** Adds ONE test record (TEST VISITOR, MTR-HSSE-261026-999) so you can scan the test certificate and try the verification + sign-off. Safe to run twice. */
function addTestRecord() {
  const sh = sheet_(), last = sh.getLastRow();
  if (last > 1 && sh.getRange(2, 2, last - 1, 1).getValues().some(function (r) { return String(r[0]) === 'MTR-HSSE-261026-999'; })) { Logger.log('Test record already present.'); return; }
  const row = new Array(HEADERS.length).fill('');
  const v = ['2026-10-26 14:30:00','MTR-HSSE-261026-999','TEST VISITOR','EcoOils Test Co','1234','EcoOils','Test Host','','Test of verification','Yes','Yes',18,20,'90%','PASS',1,'','','','EN','','S3HA6HZ2U3'];
  for (let i = 0; i < v.length; i++) row[i] = v[i];
  sh.appendRow(row);
  Logger.log('Test record added.');
}

/** If setupOfficers was already run, this fills in / corrects Putri's email in the 'HSSE Officers' tab (no PIN is changed, nothing is sent). */
function setPutriEmail() {
  const sh = officersSheet_(), v = sh.getDataRange().getValues(), h = v[0].map(String);
  const ni = h.indexOf('Name'), ei = h.indexOf('Email');
  if (ni < 0 || ei < 0) throw new Error('Name/Email columns not found in HSSE Officers.');
  for (let r = 1; r < v.length; r++) if (String(v[r][ni]).indexOf('Putri') === 0) { sh.getRange(r + 1, ei + 1).setValue('H.PutriYuslina@ecooils.co.id'); Logger.log('Putri email set.'); return; }
  Logger.log('Putri not found.');
}

/**
 * Add the contractor focal points. Edit the list below (name, email, company), Save, run addContractorFocalPoints once.
 * Each gets a random PIN and can sign off ONLY workers whose Host/Contractor is their own company. Nothing is emailed:
 * they tap 'Email me my PIN' on the sign-off screen. People already listed are left alone.
 */
const CONTRACTOR_FOCAL = [
  ['KKS 1 - Edwin', 'edwin.wantah@kokohsemesta.com', 'KKS'],
  ['KKS 2 - Pesta', 'pesta.kokohsemesta1312@gmail.com', 'KKS'],
  ['GSB 1 - Hartis', 'dc02_hse@gsb.co.id', 'GSB'],
  ['GSB 2', 'hsegsb@gmail.com', 'GSB'],
  ['PBS 1', 'hse.ecooils@paramita.co.id', 'PBS'],
  ['PBS 2 - Maradu Naipospos', 'maradu.naipospos@paramita.co.id', 'PBS'],
  ['WEN 1', 'evamyudi23@gmail.com', 'WEN'],
  ['MPE 1', 'aqchal003@gmail.com', 'MPE']
];
function addContractorFocalPoints() {
  const sh = officersSheet_(), last = sh.getLastRow(), have = {}, used = {};
  if (last > 1) sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) { have[String(r[0]).trim().toLowerCase()] = true; if (String(r[1]).trim()) used[('0000' + String(r[1]).trim()).slice(-4)] = true; });
  CONTRACTOR_FOCAL.forEach(function (o) {
    if (have[o[0].toLowerCase()]) { Logger.log('Already listed: ' + o[0]); return; }
    if (['GSB', 'WEN', 'KKS', 'PBS', 'MPE'].indexOf(String(o[2]).toUpperCase()) < 0) { Logger.log('Skipped (company must be GSB/WEN/KKS/PBS/MPE): ' + o[0]); return; }
    const pin = newPin_(used), r = sh.getLastRow() + 1;
    sh.getRange(r, 2).setNumberFormat('@');
    sh.getRange(r, 1, 1, 6).setValues([[o[0], pin, 'Yes', o[1], o[2].toUpperCase() + ' HSSE focal point', o[2].toUpperCase()]]);
    Logger.log('Added ' + o[0] + ' (' + o[2].toUpperCase() + ')');
  });
}

/* ---------- v7.8: statistics view (Ctrl+Alt+M in the app) ----------
 * Same officer + PIN as the sign-off (and the same 'Email me my PIN'). ALL-scope officers see every contractor;
 * a contractor focal point sees only their own company. History from the old Microsoft Forms is read from the
 * 'Forms History' tab (import it once), the new app's records from 'Induction Records'. PASS = 16 or more out of 20.
 */
const FORMS_SHEET = 'Forms History';
const PASS_MARK = 16;
const STATS_FROM = '2026-01-01';   // statistics cover 2026 onwards; earlier records are archived and only counted
function ts_(v) { return (v instanceof Date) ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm:ss') : String(v || ''); }
function statRows_() {
  const out = [], ss = SpreadsheetApp.openById(sheetId_());
  const fh = ss.getSheetByName(FORMS_SHEET);
  if (fh && fh.getLastRow() > 1) {
    fh.getRange(2, 1, fh.getLastRow() - 1, 6).getValues().forEach(function (r) {
      if (!r[1]) return;
      const sc = Number(r[4]);
      out.push({ date: ts_(r[0]), name: String(r[1]), company: String(r[2] || 'Unassigned (2025)'), score: isNaN(sc) ? 0 : sc, pass: /^p/i.test(String(r[5])) });
    });
  }
  const sh = sheet_(), last = sh.getLastRow();
  if (last > 1) {
    sh.getRange(2, 1, last - 1, 15).getValues().forEach(function (r) {
      if (!r[2]) return;
      const sc = Number(r[11]), tot = Number(r[12]) || 20;
      out.push({ date: ts_(r[0]), name: String(r[2]), company: String(r[5] || 'Unassigned'), score: isNaN(sc) ? 0 : sc, pass: /pass/i.test(String(r[14])) || (!isNaN(sc) && sc / tot >= 0.8) });
    });
  }
  out.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
  return out;
}
function admin_(d) {
  const name = String(d.officer || '').trim(), cache = CacheService.getScriptCache();
  const key = 'pin_' + name.toLowerCase().replace(/[^a-z0-9]/g, '').slice(0, 40), keyAll = 'pin_all';
  if (Number(cache.get(key) || 0) >= MAX_PIN_TRIES || Number(cache.get(keyAll) || 0) >= MAX_PIN_TRIES_ALL) return { ok: false, error: 'LOCKED' };
  let officer = null; const list = activeOfficers_();
  for (let i = 0; i < list.length; i++) if (list[i].name === name) { officer = list[i]; break; }
  if (!officer || !pinEq_(officer.pin, d.pin)) {
    cache.put(key, String(Number(cache.get(key) || 0) + 1), LOCK_SECS);
    cache.put(keyAll, String(Number(cache.get(keyAll) || 0) + 1), LOCK_SECS);
    return { ok: false, error: 'PIN' };
  }
  cache.remove(key);
  const scope = officer.scope || 'ALL';
  let rows = statRows_();
  if (scope !== 'ALL') rows = rows.filter(function (r) { return r.company.toUpperCase() === scope; });
  const archive = { inductions: 0, passed: 0 };
  rows = rows.filter(function (r) { if (r.date < STATS_FROM) { archive.inductions++; if (r.pass) archive.passed++; return false; } return true; });
  const ppl = {}, sum = {}, mon = {}, order = [];
  rows.forEach(function (r) {
    const co = r.company, k = r.name.toLowerCase().replace(/[^a-z]/g, '') + '|' + co;
    if (!sum[co]) { sum[co] = { attempts: 0, pass: 0, fail: 0, people: 0, done: 0 }; order.push(co); }
    sum[co].attempts++; if (r.pass) sum[co].pass++; else sum[co].fail++;
    let p = ppl[k]; if (!p) { p = ppl[k] = { name: r.name, company: co, done: false }; sum[co].people++; }
    if (r.pass && !p.done) {
      p.done = true; p.date = r.date; p.score = r.score; sum[co].done++;
      const m = r.date.slice(0, 7); mon[m] = mon[m] || {}; mon[m][co] = (mon[m][co] || 0) + 1;
    }
  });
  const main = ['PBS', 'KKS', 'GSB', 'WEN', 'MPE', 'EcoOils'];
  order.sort(function (a, b) { const x = main.indexOf(a), y = main.indexOf(b); return (x < 0 ? 99 : x) - (y < 0 ? 99 : y) || (a < b ? -1 : 1); });
  const months = Object.keys(mon).sort().map(function (m) { return { m: m, c: mon[m] }; });
  const recent = Object.keys(ppl).map(function (k) { return ppl[k]; }).filter(function (p) { return p.done; })
    .sort(function (a, b) { return a.date < b.date ? 1 : -1; }).slice(0, 40).map(function (p) { return { date: p.date.slice(0, 10), name: p.name, company: p.company, score: p.score }; });
  return { ok: true, officer: officer.name, scope: scope, passMark: PASS_MARK, companies: order, summary: sum, months: months, recent: recent, total: rows.length, archive: archive, from: STATS_FROM };
}

// v8.0: counts only, for the Mentari HSSE Performance Tracker. No names, PINs or companies leave the script.
// Same rule as the Induction Stats screen: one induction per person (name + company), counted in the month of their first PASS. Attempts = every sitting.
function counts_() {
  const rows = statRows_().filter(function (r) { return r.date >= STATS_FROM; });
  const seen = {}, months = {};
  rows.forEach(function (r) {
    const m = r.date.slice(0, 7); months[m] = months[m] || { inductions: 0, attempts: 0 }; months[m].attempts++;
    const k = r.name.toLowerCase().replace(/[^a-z]/g, '') + '|' + r.company;
    if (r.pass && !seen[k]) { seen[k] = 1; months[m].inductions++; }
  });
  return { ok: true, months: months, from: STATS_FROM };
}

function doGet(e) {
  try { return doGet_(e); } catch (err) { return reply_({ ok: false, error: 'doGet: ' + String(err) }); }
}

function doGet_(e) {
  const p = (e && e.parameter) || {};
  if (String(p.f) === 'counts') return reply_(counts_());
  if (String(p.f) === 'officers') return (function () { const L = activeOfficers_(), sc = {}; L.forEach(function (o) { sc[o.name] = o.scope; }); return reply_({ ok: true, officers: L.map(function (o) { return o.name; }), scopes: sc }); })();   // names only, never PINs
  // New names (cert, code) avoid the 'c' parameter that Google's web-app front door rejects; v and c still accepted.
  const cert = String(p.cert || p.v || ''), code = String(p.code || p.c || '');
  if (!cert) return reply_({ ok: true, service: 'Mentari induction receiver', version: SCRIPT_VERSION });
  if (String(p.f) === 'json') return reply_(verifyData_(cert, code));   // used by the app's verification screen
  return verifyPage_(cert, code);                                       // plain page (backup)
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'familiarise') return reply_(familiarise_(d));   // HSSE sign-off
    if (d.action === 'admin') return reply_(admin_(d));               // Ctrl+Alt+M statistics view (officer PIN)
    if (d.action === 'emailpin') return reply_(emailPin_(d));         // 'Email me my PIN'
    if (!/^MTR-HSSE-\d{6}-\d{3}$/.test(String(d.certNumber || '')) || !d.name) {
      return reply_({ ok: false, error: 'Invalid record' });
    }
    logRecord_(d);                                   // 1) the row first
    let link = '', imageSaved = false, imageError = '';
    if (d.certImage && d.passed) {                   // 2) then the certificate copy
      try { link = saveCertificate_(d); setLink_(d, link); imageSaved = true; }
      catch (err) { imageError = String(err); }     // a Drive problem can no longer lose the record
    }
    let mcuSaved = false, mcuError = '';
    if (d.mcuImage) {                                // 3) the MCU copy (independent of pass/fail and of the certificate)
      try { const m = saveMcu_(d); setMcuLink_(d, m); mcuSaved = true; }
      catch (err) { mcuError = String(err); }
    }
    return reply_({ ok: true, rowSaved: true, imageSaved: imageSaved, link: link, imageError: imageError, mcuSaved: mcuSaved, mcuError: mcuError });
  } catch (err) {
    return reply_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function saveCertificate_(d) {
  const s = String(d.certImage);
  const comma = s.indexOf(',');
  const head = /^data:image\/(png|jpeg);base64$/.exec(s.slice(0, comma));
  if (comma < 0 || !head || s.length > MAX_IMAGE_CHARS) throw new Error('Bad or oversized image');
  const ext = head[1] === 'jpeg' ? 'jpg' : 'png';
  const folder = monthFolder_();
  const fileName = clean_(d.certNumber + ' - ' + d.name + ' - ' + d.company) + '.' + ext;
  const existing = folder.getFilesByName(fileName);
  if (existing.hasNext()) return existing.next().getUrl();   // retry-safe
  const blob = Utilities.newBlob(Utilities.base64Decode(s.slice(comma + 1)), 'image/' + head[1], fileName);
  return folder.createFile(blob).getUrl();
}

function saveMcu_(d) {
  const s = String(d.mcuImage);
  const comma = s.indexOf(',');
  const head = /^data:(image\/png|image\/jpeg|application\/pdf);base64$/.exec(s.slice(0, comma));
  if (comma < 0 || !head || s.length > MAX_MCU_CHARS) throw new Error('Bad or oversized MCU file');
  const ext = head[1] === 'application/pdf' ? 'pdf' : (head[1] === 'image/png' ? 'png' : 'jpg');
  const month = monthFolder_();
  const it = month.getFoldersByName('MCU');
  const folder = it.hasNext() ? it.next() : month.createFolder('MCU');
  const fileName = clean_(d.certNumber + ' - ' + d.name + ' - MCU') + '.' + ext;
  const existing = folder.getFilesByName(fileName);
  if (existing.hasNext()) return existing.next().getUrl();   // retry-safe
  const blob = Utilities.newBlob(Utilities.base64Decode(s.slice(comma + 1)), head[1], fileName);
  return folder.createFile(blob).getUrl();
}

function setMcuLink_(d, link) {
  const sh = sheet_();
  const r = findRow_(sh, d.certNumber, d.idnum);
  if (r && link && !sh.getRange(r, COL_MCU).getValue()) sh.getRange(r, COL_MCU).setValue(link);
}

/* ---------- v7: HSSE familiarisation sign-off ---------- */
function officersSheet_() {
  const ss = SpreadsheetApp.openById(sheetId_());
  let sh = ss.getSheetByName(OFFICER_SHEET);
  if (!sh) {
    sh = ss.insertSheet(OFFICER_SHEET);
    sh.setFrozenRows(1);
    sh.getRange(2, 2, 200, 1).setNumberFormat('@');           // PINs are text, so a leading 0 is kept
  }
  const head = ['Name', 'PIN', 'Active (Yes/No)', 'Email', 'Role', 'Scope (ALL or GSB/WEN/KKS/PBS/MPE)'];
  for (let c = 1; c <= head.length; c++) if (sh.getRange(1, c).getValue() !== head[c - 1]) sh.getRange(1, c).setValue(head[c - 1]).setFontWeight('bold');
  return sh;
}

function activeOfficers_() {
  const sh = officersSheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  const out = [], vals = sh.getRange(2, 1, last - 1, 6).getValues();
  for (let i = 0; i < vals.length; i++) {
    const name = String(vals[i][0] || '').trim();
    if (name && /^y/i.test(String(vals[i][2] || '').trim()) && String(vals[i][1] || '').trim()) {
      out.push({ name: name, pin: String(vals[i][1]).trim(), email: String(vals[i][3] || '').trim(), scope: String(vals[i][5] || 'ALL').trim().toUpperCase() || 'ALL' });
    }
  }
  return out;
}

function newPin_(used) {                                        // random 4 digits; never 0000 / 1111 / 1234 style, never a duplicate
  for (let t = 0; t < 200; t++) {
    const n = parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 8), 16) % 10000;
    const pin = ('0000' + n).slice(-4);
    if (/^(\d)\1{3}$/.test(pin) || '0123456789012'.indexOf(pin) > -1 || '9876543210'.indexOf(pin) > -1 || used[pin]) continue;
    used[pin] = true; return pin;
  }
  throw new Error('Could not create a PIN');
}

/** STEP 3 - run ONCE: creates the 'HSSE Officers' tab and adds the HSSE team with a random 4-digit PIN each (people already listed are left alone). */
function setupOfficers() {
  const sh = officersSheet_(), last = sh.getLastRow(), have = {}, used = {};
  if (last > 1) sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) { have[String(r[0]).trim().toLowerCase()] = true; if (String(r[1]).trim()) used[('0000' + String(r[1]).trim()).slice(-4)] = true; });
  OFFICER_SEED.forEach(function (o) {
    if (have[o[0].toLowerCase()]) { Logger.log('Already listed: ' + o[0]); return; }
    const pin = newPin_(used), r = sh.getLastRow() + 1;
    sh.getRange(r, 2).setNumberFormat('@');
    sh.getRange(r, 1, 1, 6).setValues([[o[0], pin, 'Yes', o[1], o[2], 'ALL']]);
    Logger.log('Added ' + o[0] + (o[1] ? ' (PIN will be emailed to ' + o[1] + ')' : ' - NO EMAIL: give this PIN in person: ' + pin));
  });
  Logger.log('Next: run emailPins to send each officer their PIN.');
}

function sendPin_(o) {
  MailApp.sendEmail({
    to: o.email, name: MAIL_SENDER_NAME, replyTo: MAIL_REPLY_TO,
    subject: 'Your HSSE site-familiarisation sign-off PIN',
    body: 'Hello ' + o.name + ',\n\nYour personal PIN for signing off site familiarisation in the Mentari induction app is:\n\n    ' + o.pin +
      '\n\nUse it only yourself: it records that YOU gave the site briefing and issued the sticker. Never share it or sign off for someone else.' +
      '\nIf you forget it, open any vendor certificate QR, tap the sign-off button and choose "Email me my PIN".\n\nHSSE - EcoOils Mentari Project'
  });
}

/** STEP 4 - run when ready: emails every active officer their PIN. Officers with no email are listed in the log so you can hand the PIN over in person. */
function emailPins() {
  officersSheet_();
  activeOfficers_().forEach(function (o) {
    if (!o.email) { Logger.log('NO EMAIL - give in person: ' + o.name + ' -> PIN ' + o.pin); return; }
    try { sendPin_(o); Logger.log('Emailed ' + o.name + ' at ' + o.email); } catch (e) { Logger.log('FAILED to email ' + o.name + ': ' + e); }
  });
  Logger.log('Emails left today: ' + MailApp.getRemainingDailyQuota());
}

/** 'Email me my PIN' from the sign-off screen: sends the officer's PIN to the email on file only (never shown on screen). Once per 10 minutes per officer. */
function emailPin_(d) {
  const name = String(d.officer || '').trim(), cache = CacheService.getScriptCache();
  const key = 'mail_' + name.toLowerCase().replace(/[^a-z0-9]/g, '').slice(0, 40);
  let officer = null; const list = activeOfficers_();
  for (let i = 0; i < list.length; i++) if (list[i].name === name) { officer = list[i]; break; }
  if (!officer || !officer.email) return { ok: false, error: 'NOEMAIL' };
  if (cache.get(key)) return { ok: false, error: 'WAIT' };
  if (MailApp.getRemainingDailyQuota() < 1) return { ok: false, error: 'WAIT' };
  sendPin_(officer);
  cache.put(key, '1', 600);
  const at = officer.email.indexOf('@');
  return { ok: true, sent: true, to: officer.email.slice(0, 1) + '***' + officer.email.slice(at) };
}

function pinEq_(a, b) {
  const pad = function (v) { v = String(v == null ? '' : v).trim(); return /^\d{1,8}$/.test(v) ? ('0000' + v).slice(-Math.max(4, v.length)) : ''; };   // Sheets can drop a leading 0
  const x = pad(a), y = pad(b);
  return x !== '' && x === y;
}

function famSheet_() {
  const ss = SpreadsheetApp.openById(sheetId_());
  let sh = ss.getSheetByName(FAM_SHEET);
  if (!sh) {
    sh = ss.insertSheet(FAM_SHEET);
    sh.appendRow(FAM_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, FAM_HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

function familiarise_(d) {
  const cert = String(d.certNumber || ''), code = String(d.verifyCode || '');
  if (!/^MTR-HSSE-\d{6}-\d{3}$/.test(cert) || !/^[A-HJ-NP-Z2-9]{10}$/.test(code)) return { ok: false, error: 'NOTFOUND' };
  const sh = sheet_(), last = sh.getLastRow();
  if (last < 2) return { ok: false, error: 'NOTFOUND' };
  const vals = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  let r = 0, rec = null;
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][1]) === cert && String(vals[i][COL_CODE - 1]) === code) { r = i + 2; rec = vals[i]; break; }
  }
  if (!r) return { ok: false, error: 'NOTFOUND' };
  if (String(rec[COL_FAM_BY - 1] || '')) {                     // already signed off: never overwritten
    return { ok: true, done: true, already: true, by: String(rec[COL_FAM_BY - 1]), at: String(rec[COL_FAM_AT - 1]), sticker: String(rec[COL_STICKER - 1]) };
  }
  const checks = d.checks;
  if (!Array.isArray(checks) || checks.length < 5 || !checks.every(function (x) { return x === true; })) return { ok: false, error: 'CHECKS' };
  const sticker = String(d.stickerNo || '').trim().toUpperCase();
  if (!/^[A-Z0-9][A-Z0-9\-\/]{0,15}$/.test(sticker)) return { ok: false, error: 'STICKER' };
  const cache = CacheService.getScriptCache();
  const name = String(d.officer || '').trim();
  const key = 'pin_' + name.toLowerCase().replace(/[^a-z0-9]/g, '').slice(0, 40), keyAll = 'pin_all';
  if (Number(cache.get(key) || 0) >= MAX_PIN_TRIES || Number(cache.get(keyAll) || 0) >= MAX_PIN_TRIES_ALL) return { ok: false, error: 'LOCKED' };
  let officer = null;
  const list = activeOfficers_();
  for (let i = 0; i < list.length; i++) if (list[i].name === name) { officer = list[i]; break; }
  if (!officer || !pinEq_(officer.pin, d.pin)) {
    cache.put(key, String(Number(cache.get(key) || 0) + 1), LOCK_SECS);
    cache.put(keyAll, String(Number(cache.get(keyAll) || 0) + 1), LOCK_SECS);
    return { ok: false, error: 'PIN' };
  }
  cache.remove(key);
  if (officer.scope && officer.scope !== 'ALL' && String(rec[5]).trim().toUpperCase() !== officer.scope) return { ok: false, error: 'SCOPE' };   // contractor focal points only sign off their own people
  const fs = famSheet_(), lastF = fs.getLastRow();
  if (lastF > 1) {                                             // one sticker number = one person
    const used = fs.getRange(2, 7, lastF - 1, 1).getValues();
    for (let i = 0; i < used.length; i++) if (String(used[i][0]).toUpperCase() === sticker) return { ok: false, error: 'USED' };
  }
  const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  const rowF = [now, cert, String(rec[2]), String(rec[3]), String(rec[4]), officer.name, sticker, 'Rules+emergency, site map, PPE, site rules, sticker issued: all confirmed'];
  const nr = fs.getLastRow() + 1;
  fs.getRange(nr, 1, 1, rowF.length).setNumberFormat('@').setValues([rowF]);
  sh.getRange(r, COL_FAM_BY, 1, 3).setNumberFormat('@').setValues([[officer.name, now, sticker]]);
  return { ok: true, done: true, by: officer.name, at: now, sticker: sticker };
}

function monthFolder_() {
  const root = DriveApp.getFolderById(rootId_());
  const name = Utilities.formatDate(new Date(), TZ, 'yyyy-MM');
  const it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

/* ---------- storage: the Sheet and folder live in the account that runs this script ---------- */
function sheetId_() {
  if (SHEET_ID) return SHEET_ID;
  const pr = PropertiesService.getScriptProperties();
  let id = pr.getProperty('SHEET_ID');
  if (!id) {
    const ss = SpreadsheetApp.create('Mentari Induction Records');
    id = ss.getId(); pr.setProperty('SHEET_ID', id);
    const sh = ss.getSheets()[0];
    sh.setName(SHEET_NAME); sh.appendRow(HEADERS); sh.setFrozenRows(1); sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    Logger.log('Created the Sheet: ' + ss.getUrl());
  }
  return id;
}

function rootId_() {
  if (ROOT_FOLDER_ID) return ROOT_FOLDER_ID;
  const pr = PropertiesService.getScriptProperties();
  let id = pr.getProperty('ROOT_FOLDER_ID');
  if (!id) {
    const f = DriveApp.createFolder('Mentari Induction Certificates');
    id = f.getId(); pr.setProperty('ROOT_FOLDER_ID', id);
    Logger.log('Created the Drive folder: ' + f.getUrl());
  }
  return id;
}

function last4_(s) { return String(s == null ? '' : s).replace(/[^A-Za-z0-9]/g, '').toUpperCase().slice(-4); }

function findRow_(sh, cert, idnum) {               // returns the sheet row number, or 0
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, 2, last - 1, 4).getValues();     // cols B..E: cert, name, company, ID
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === cert && last4_(vals[i][3]) === last4_(idnum)) return i + 2;
  }
  return 0;
}

function logRecord_(d) {
  const sh = sheet_();
  const found = findRow_(sh, d.certNumber, d.idnum);
  if (found) {                                     // already logged (a retry): just make sure the verify code is there
    if (d.verifyCode && !sh.getRange(found, COL_CODE).getValue()) sh.getRange(found, COL_CODE).setValue(String(d.verifyCode));
    return found;
  }
  const row = [
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
    d.certNumber, d.name, d.company, last4_(d.idnum), d.host, d.hostName, d.hostPhone, d.purpose,
    d.certify ? 'Yes' : 'No',
    d.healthCert === 'yes' ? 'Yes' : (d.healthCert === 'no' ? 'No' : ''),
    d.score, d.totalQuestions, (d.percentage != null ? d.percentage + '%' : ''),
    d.passed ? 'Pass' : 'Fail', d.attemptsTaken, d.startTime, d.completionTime,
    d.durationLabel, d.lang, '', String(d.verifyCode || ''), '', '', '', ''
  ];
  const r = sh.getLastRow() + 1;
  const range = sh.getRange(r, 1, 1, row.length);
  range.setNumberFormat('@');                      // plain text: keeps ID/phone digits, blocks formula injection
  sh.getRange(r, 12, 1, 2).setNumberFormat('0');   // score + total stay numeric
  range.setValues([row]);
  return r;
}

function setLink_(d, link) {
  const sh = sheet_();
  const r = findRow_(sh, d.certNumber, d.idnum);
  if (r && link && !sh.getRange(r, COL_LINK).getValue()) sh.getRange(r, COL_LINK).setValue(link);
}

function sheet_() {
  const ss = SpreadsheetApp.openById(sheetId_());
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  } else {
    if (sh.getRange(1, COL_CODE).getValue() !== 'Verify Code') {   // older sheet: add the new column header once
      sh.getRange(1, COL_CODE).setValue('Verify Code').setFontWeight('bold');
    }
    for (let c = COL_MCU; c <= HEADERS.length; c++) {              // v6/v7: add any missing column headers once
      if (sh.getRange(1, c).getValue() !== HEADERS[c - 1]) sh.getRange(1, c).setValue(HEADERS[c - 1]).setFontWeight('bold');
    }
    if (sh.getRange(1, 5).getValue() === 'ID/Passport') sh.getRange(1, 5).setValue('ID/Passport (last 4)');   // v5: the column now holds only the last 4
  }
  return sh;
}

/* ---------- verification (opened by scanning the QR on the certificate) ---------- */
function verifyData_(cert, code) {
  let rec = null;
  try {
    if (/^MTR-HSSE-\d{6}-\d{3}$/.test(cert) && /^[A-HJ-NP-Z2-9]{10}$/.test(code)) {
      const sh = sheet_(), last = sh.getLastRow();
      if (last > 1) {
        const vals = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
        for (let i = 0; i < vals.length; i++) {
          if (String(vals[i][1]) === cert && String(vals[i][COL_CODE - 1]) === code) { rec = vals[i]; break; }
        }
      }
    }
  } catch (err) { return { ok: false, error: 'storage: ' + String(err) }; }
  if (!rec) return { ok: true, verified: false };
  const idStr = String(rec[4] || '');
  let image = '';
  try {
    const m = /\/d\/([^\/]+)/.exec(String(rec[COL_LINK - 1] || ''));
    if (m) {
      const blob = DriveApp.getFileById(m[1]).getBlob();
      image = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
    }
  } catch (err) {}
  return { ok: true, verified: true, image: image, record: {
    cert: String(rec[1]), name: String(rec[2]), company: String(rec[3]), idLast4: idStr.slice(-4),
    host: String(rec[5]), hostName: String(rec[6]), purpose: String(rec[8]), result: String(rec[14]),
    score: String(rec[11]), total: String(rec[12]), attempts: String(rec[15]), issued: String(rec[0]),
    fam: String(rec[COL_FAM_BY - 1] || '') ? { by: String(rec[COL_FAM_BY - 1]), at: String(rec[COL_FAM_AT - 1]), sticker: String(rec[COL_STICKER - 1]) } : null } };
}

function verifyPage_(cert, code) {                 // backup page: the app's own verification screen is the main route
  const v = verifyData_(cert, code);
  const esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]; }); };
  let body;
  if (!v.verified) {
    body = '<div class="box bad"><h1>&#10007; NOT VERIFIED</h1><p>No matching induction record was found for this certificate.</p>' +
      '<p><b>Do not accept this certificate.</b><br><b>Jangan terima sertifikat ini.</b></p></div>';
  } else {
    const r = v.record, masked = r.idLast4 ? '&bull;&bull;&bull;&bull;' + esc(r.idLast4) : '';
    const row = function (k, x) { return '<tr><td>' + k + '</td><td><b>' + x + '</b></td></tr>'; };
    body = '<div class="box ok"><h1>&#10003; VERIFIED</h1><p>Induction record found / Rekaman induksi ditemukan</p></div><table>' +
      row('Certificate No.', esc(r.cert)) + row('Name', esc(r.name)) + row('Company', esc(r.company)) + row('ID (last 4)', masked) +
      row('Host / contractor', esc(r.host)) + row('Host contact', esc(r.hostName)) + row('Purpose', esc(r.purpose)) +
      row('Result', esc(r.result) + ' &mdash; ' + esc(r.score) + ' / ' + esc(r.total)) + row('Attempts', esc(r.attempts)) + row('Issued', esc(r.issued)) +
      (r.fam ? row('Site familiarisation', 'DONE by ' + esc(r.fam.by) + ' &middot; sticker ' + esc(r.fam.sticker)) : row('Site familiarisation', 'NOT YET DONE')) + '</table>' +
      '<p class="note"><b>Compare</b> the face and details with the person and the certificate they show. Self-declared record &mdash; not proof of employment or authorization to work.</p>' +
      (v.image ? '<img alt="Certificate on file" src="' + v.image + '">' : '<p class="note">Certificate image not on file yet.</p>');
  }
  const html = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,Helvetica,sans-serif;margin:0;padding:14px;background:#f4f6f3;color:#20241f}' +
    '.box{border-radius:12px;padding:14px 16px;margin-bottom:12px;text-align:center}' +
    '.ok{background:#e3f4e8;border:2px solid #2f8f52;color:#1f6b3b}.bad{background:#fbe7e1;border:2px solid #c0392b;color:#c0392b}' +
    'h1{margin:0 0 6px;font-size:26px}p{margin:6px 0}table{width:100%;border-collapse:collapse;background:#fff;border-radius:10px;overflow:hidden}' +
    'td{padding:9px 10px;border-bottom:1px solid #e1e6dd;font-size:14px;vertical-align:top}td:first-child{color:#6b7280;width:38%}' +
    '.note{font-size:12.5px;color:#4b5563;margin:12px 2px}img{width:100%;border-radius:10px;margin-top:8px;border:1px solid #e1e6dd}' +
    '</style></head><body>' + body + '</body></html>';
  return HtmlService.createHtmlOutput(html).setTitle('Certificate verification')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function clean_(s) {
  return String(s || '').replace(/[\\\/:*?"<>|\r\n]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 120);
}

function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** OPTIONAL - run ONCE, only when you are ready: trims every ID already in the sheet to its last 4 characters.
 *  This cannot be undone. Make a copy of the sheet first (File > Make a copy). Certificate images already saved in Drive are not changed. */
function maskOldIds() {
  const sh = sheet_(), last = sh.getLastRow();
  if (last < 2) { Logger.log('No rows to trim.'); return; }
  const rng = sh.getRange(2, 5, last - 1, 1), vals = rng.getValues();
  let changed = 0;
  const out = vals.map(function (r) {
    const full = String(r[0] == null ? '' : r[0]), short = last4_(full);
    if (full !== short) changed++;
    return [short];
  });
  rng.setNumberFormat('@');
  rng.setValues(out);
  Logger.log('Trimmed ' + changed + ' of ' + out.length + ' rows to the last 4 characters.');
}

/** STEP 1 - run once: tells you which account runs the script and checks every connection, one line each. Send me the whole log. */
function checkSetup() {
  const step = function (label, fn) { try { Logger.log('OK    ' + label + ': ' + fn()); } catch (e) { Logger.log('FAIL  ' + label + ': ' + e); } };
  step('1 Running as', function () { return Session.getEffectiveUser().getEmail(); });
  step('2 Sheet (open or create)', function () { const ss = SpreadsheetApp.openById(sheetId_()); return ss.getName() + ' -> ' + ss.getUrl(); });
  step('3 Induction Records tab', function () { const sh = sheet_(); return 'rows=' + sh.getLastRow() + ', headers=' + sh.getRange(1, 1, 1, 3).getValues()[0].join(' | '); });
  step('4 Drive folder (open or create)', function () { const f = DriveApp.getFolderById(rootId_()); return f.getName() + ' -> ' + f.getUrl(); });
  step('5 Month folder', function () { return monthFolder_().getName(); });
  step('6 Email permission', function () { return 'emails left today: ' + MailApp.getRemainingDailyQuota(); });
  step('7 Verify lookup (as the web address does it)', function () { return doGet({ parameter: { v: 'MTR-HSSE-000000-000', c: 'ABCDEFGHJK', f: 'json' } }).getContent().slice(0, 160); });
}

/** STEP 2 - run once: grants permissions, writes one test row + test file, then checks the verification page. */
function authorizeAndTest() {
  const fake = { certNumber: 'MTR-HSSE-000000-000', verifyCode: 'ABCDEFGHJK', name: 'TEST PERSON', company: 'TEST', idnum: '0000',
    host: 'EcoOils', hostName: 'x', hostPhone: '+620000', purpose: 'test', certify: true, healthCert: 'yes',
    score: 20, totalQuestions: 20, percentage: 100, passed: true, attemptsTaken: 1, startTime: '',
    completionTime: '', durationLabel: '0m 00s', lang: 'en',
    certImage: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==' };
  const noImg = JSON.parse(JSON.stringify(fake)); delete noImg.certImage;
  Logger.log('1) row first (no image): ' + doPost({ postData: { contents: JSON.stringify(noImg) } }).getContent());
  Logger.log('2) then the image:       ' + doPost({ postData: { contents: JSON.stringify(fake) } }).getContent());
  const withMcu = JSON.parse(JSON.stringify(fake)); delete withMcu.certImage;
  withMcu.mcuImage = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='; withMcu.mcuName = 'test.png';
  Logger.log('2b) MCU file saved:       ' + doPost({ postData: { contents: JSON.stringify(withMcu) } }).getContent());
  const page = verifyPage_('MTR-HSSE-000000-000', 'ABCDEFGHJK').getContent();
  Logger.log('3) verification page says VERIFIED: ' + (page.indexOf('VERIFIED') > -1 && page.indexOf('NOT VERIFIED') === -1));
  Logger.log('4) wrong code says NOT VERIFIED:    ' + (verifyPage_('MTR-HSSE-000000-000', 'ZZZZZZZZZZ').getContent().indexOf('NOT VERIFIED') > -1));
  const j = JSON.parse(doGet({ parameter: { v: 'MTR-HSSE-000000-000', c: 'ABCDEFGHJK', f: 'json' } }).getContent());
  Logger.log('5) JSON route (used by the app) verified: ' + (j.verified === true) + ' | name: ' + (j.record && j.record.name));
  Logger.log('5b) verification never exposes the MCU: ' + (JSON.stringify(j).indexOf('MCU') === -1));
  Logger.log('6b) email permission granted, emails left today: ' + MailApp.getRemainingDailyQuota());
  const os = officersSheet_(); os.appendRow(['TEST OFFICER', '', 'Yes', '', 'test']);
  const tr = os.getLastRow(); os.getRange(tr, 2).setNumberFormat('@').setValue('0000');
  const tf = { action: 'familiarise', certNumber: 'MTR-HSSE-000000-000', verifyCode: 'ABCDEFGHJK', officer: 'TEST OFFICER', pin: '9999', stickerNo: 'TEST-001', checks: [true, true, true, true, true] };
  const go = function (o) { return JSON.parse(doPost({ postData: { contents: JSON.stringify(o) } }).getContent()); };
  Logger.log('7) wrong PIN refused:            ' + (go(tf).error === 'PIN'));
  Logger.log('8) incomplete checklist refused:  ' + (go(Object.assign({}, tf, { pin: '0000', checks: [true, true, false, true, true] })).error === 'CHECKS'));
  Logger.log('9) sign-off with the right PIN:   ' + (go(Object.assign({}, tf, { pin: '0000' })).done === true));
  const j2 = JSON.parse(doGet({ parameter: { v: 'MTR-HSSE-000000-000', c: 'ABCDEFGHJK', f: 'json' } }).getContent());
  Logger.log('10) verification shows who signed: ' + (j2.record && j2.record.fam && j2.record.fam.by) + ' / sticker ' + (j2.record && j2.record.fam && j2.record.fam.sticker));
  Logger.log('11) officer list has names only:   ' + (JSON.stringify(JSON.parse(doGet({ parameter: { f: 'officers' } }).getContent())).indexOf('0000') === -1));
  os.getRange(tr, 3).setValue('No');                                // the test officer is switched off again
  Logger.log('6) a full number is trimmed to the last 4: ' + (last4_('1472 0429 1003 0001') === '0001' && last4_('ed3269513') === '9513'));
}