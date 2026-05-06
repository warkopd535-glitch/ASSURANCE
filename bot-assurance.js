require('dotenv').config();
const TelegramBot = require('node-telegram-bot-api');
const { google } = require('googleapis');

// === ENV VARIABLES ===
const TOKEN = process.env.TELEGRAM_TOKEN;
const SHEET_ID = process.env.SHEET_ID;
const GROUP_CHAT_ID = process.env.GROUP_CHAT_ID;
let GOOGLE_SERVICE_ACCOUNT_JSON = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;

if (!TOKEN) { console.error('❌ TELEGRAM_TOKEN not set'); process.exit(1); }
if (!SHEET_ID) { console.error('❌ SHEET_ID not set'); process.exit(1); }
if (!GOOGLE_SERVICE_ACCOUNT_JSON) { console.error('❌ GOOGLE_SERVICE_ACCOUNT_JSON not set'); process.exit(1); }
if (!GROUP_CHAT_ID) { console.warn('⚠️ GROUP_CHAT_ID not set - TTR alerts disabled'); }

// === PARSE GOOGLE SERVICE ACCOUNT ===
let serviceAccount;
try {
  let keyData = GOOGLE_SERVICE_ACCOUNT_JSON.trim();
  if (!keyData.startsWith('{')) {
    try { keyData = Buffer.from(keyData, 'base64').toString('utf-8'); } catch (e) { }
  }
  serviceAccount = JSON.parse(keyData);
  console.log('✅ Google Service Account parsed');
} catch (e) {
  console.error('❌ Failed to parse JSON:', e.message);
  process.exit(1);
}

const auth = new google.auth.GoogleAuth({
  credentials: serviceAccount,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });

// === CONSTANTS ===
const ASSURANCE_SHEET = 'PROGRES ASSURANCE';
const ORDER_ASSURANCE_SHEET = 'ORDER ASSURANCE';
const MASTER_SHEET = 'MASTER';
const SQM_SHEET = 'SQM SA SIGLI';
const MANUAL_GGN_SHEET = 'MANUAL GGN';
const UNSPEC_SHEET = 'UNSPEC';
const GAUL_ASSURANCE_SHEET = 'GAUL ASSURANCE';

const TTR_TABLE = {
  'HVC_DIAMOND': 3,
  'FFG': 3,
  'BGES HSI': 4,
  'DATIN K2': 3.6,
  'DATIN K3': 7.2,
  'HVC_PLATINUM': 6,
  'HVC_GOLD': 12,
  'REGULER': 36,
};

const BULAN_ID = {
  'januari': 1, 'februari': 2, 'maret': 3, 'april': 4,
  'mei': 5, 'juni': 6, 'juli': 7, 'agustus': 8,
  'september': 9, 'oktober': 10, 'november': 11, 'desember': 12,
};

// === HELPER: Auto-detect ORDER ASSURANCE column indices from header ===
function getOrderColumns(data) {
  const defaults = { incident: 1, teknisi: 2, ttrCustomer: 3, workzone: 4, customerType: 5, status: 9, hasilUkur: -1 };
  if (!data || data.length < 1) return defaults;

  const header = data[0] || [];
  const cols = { ...defaults };

  for (let c = 0; c < header.length; c++) {
    const h = (header[c] || '').toUpperCase().trim();
    if (h === 'INCIDENT' || h === 'NO INCIDENT') cols.incident = c;
    else if (h === 'TEKNISI') cols.teknisi = c;
    else if (h.includes('TTR') && h.includes('CUSTOMER')) cols.ttrCustomer = c;
    else if (h === 'WORKZONE' && c < 10) cols.workzone = c;
    else if (h === 'CUSTOMER TYPE' || h === 'CUSTOMER_TYPE') {
      if (c < 10) cols.customerType = c;
    }
    else if (h === 'STATUS') cols.status = c;
    else if (h === 'HASIL UKUR' || h === 'HASIL_UKUR') cols.hasilUkur = c;
  }

  console.log(`📍 ORDER columns: INC=${cols.incident}, TEKNISI=${cols.teknisi}, TTR=${cols.ttrCustomer}, WZ=${cols.workzone}, CUST_TYPE=${cols.customerType}, STATUS=${cols.status}, HASIL_UKUR=${cols.hasilUkur}`);
  return cols;
}

// === CACHING ===
const cache = {
  masterData: null, masterDataTime: 0,
  assuranceData: null, assuranceDataTime: 0,
  orderAssuranceData: null, orderAssuranceDataTime: 0,
  cacheExpiry: 5 * 60 * 1000,
};

// === STATE ===
const alertState = { warned: new Set(), expired: new Set() };
const userChatIds = {};

// === HELPER: Get sheet data with caching ===
async function getSheetData(sheetName, useCache = true) {
  try {
    if (useCache) {
      if (sheetName === MASTER_SHEET && cache.masterData && Date.now() - cache.masterDataTime < cache.cacheExpiry) {
        return cache.masterData;
      }
      if (sheetName === ASSURANCE_SHEET && cache.assuranceData && Date.now() - cache.assuranceDataTime < cache.cacheExpiry) {
        return cache.assuranceData;
      }
      if (sheetName === ORDER_ASSURANCE_SHEET && cache.orderAssuranceData && Date.now() - cache.orderAssuranceDataTime < cache.cacheExpiry) {
        return cache.orderAssuranceData;
      }
    }

    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: sheetName,
    });
    const data = res.data.values || [];

    if (sheetName === MASTER_SHEET) { cache.masterData = data; cache.masterDataTime = Date.now(); }
    else if (sheetName === ASSURANCE_SHEET) { cache.assuranceData = data; cache.assuranceDataTime = Date.now(); }
    else if (sheetName === ORDER_ASSURANCE_SHEET) { cache.orderAssuranceData = data; cache.orderAssuranceDataTime = Date.now(); }

    return data;
  } catch (error) {
    console.error(`Error reading ${sheetName}:`, error.message);
    if (sheetName === MASTER_SHEET && cache.masterData) return cache.masterData;
    if (sheetName === ASSURANCE_SHEET && cache.assuranceData) return cache.assuranceData;
    if (sheetName === ORDER_ASSURANCE_SHEET && cache.orderAssuranceData) return cache.orderAssuranceData;
    throw error;
  }
}

// === HELPER: Append to sheet ===
async function appendSheetData(sheetName, values) {
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: sheetName,
    valueInputOption: 'USER_ENTERED',
    resource: { values: [values] },
  });
}

// === HELPER: Update a single cell ===
async function updateSheetCell(sheetName, cell, value) {
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `'${sheetName}'!${cell}`,
    valueInputOption: 'USER_ENTERED',
    resource: { values: [[value]] },
  });
}

// === HELPER: Batch update multiple cells (preserves formulas in skipped columns) ===
async function batchUpdateCells(sheetName, updates) {
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SHEET_ID,
    resource: {
      valueInputOption: 'USER_ENTERED',
      data: updates.map(u => ({
        range: `'${sheetName}'!${u.range}`,
        values: [[u.value]],
      })),
    },
  });
}

// === HELPER: Find next row with empty incident column ===
function findNextEmptyRow(data, incidentColIdx) {
  for (let i = 1; i < data.length; i++) {
      if (!data[i]) continue;
    if (!(data[i][incidentColIdx] || '').trim()) {
      return i + 1; // 1-based row number
    }
  }
  return data.length + 1; // Append at end if no empty row found
}

// === HELPER: Send Telegram (with chunking) ===
async function sendTelegram(chatId, text, options = {}) {
  const maxLength = 4000;
  try {
    if (text.length <= maxLength) {
      return await bot.sendMessage(chatId, text, { parse_mode: 'HTML', ...options });
    }
    const lines = text.split('\n');
    let chunk = '';
    for (const line of lines) {
      if ((chunk + line + '\n').length > maxLength) {
        await bot.sendMessage(chatId, chunk, { parse_mode: 'HTML', ...options });
        chunk = '';
      }
      chunk += line + '\n';
    }
    if (chunk.trim()) {
      await bot.sendMessage(chatId, chunk, { parse_mode: 'HTML', ...options });
    }
  } catch (error) {
    console.error('Error sending message:', error.message);
  }
}

// === HELPER: Timeout wrapper ===
function withTimeout(promise, ms = 10000) {
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(() => reject(new Error('Timeout - Google API too slow')), ms)),
  ]);
}

// === HELPER: Get user role from MASTER ===
async function getUserRole(username) {
  try {
    const data = await getSheetData(MASTER_SHEET);
    for (let i = 1; i < data.length; i++) {
      if (!data[i]) continue;
      const sheetUser = (data[i][8] || '').replace('@', '').toLowerCase().trim();
      const inputUser = (username || '').replace('@', '').toLowerCase().trim();
      const status = (data[i][10] || '').toUpperCase().trim();
      const role = (data[i][9] || '').toUpperCase().trim();
      if (sheetUser === inputUser && status === 'AKTIF') return role;
    }
    return null;
  } catch (error) {
    console.error('Error getting user role:', error.message);
    return null;
  }
}

// === HELPER: Check authorization ===
async function checkAuthorization(username, requiredRoles = []) {
  try {
    const userRole = await withTimeout(getUserRole(username), 8000);
    if (!userRole) return { authorized: false, role: null, message: '❌ Anda tidak terdaftar di sistem.' };
    if (requiredRoles.length > 0 && !requiredRoles.includes(userRole))
      return { authorized: false, role: userRole, message: `❌ Akses ditolak. Role ${userRole} tidak memiliki izin.` };
    return { authorized: true, role: userRole };
  } catch (error) {
    return { authorized: false, role: null, message: '❌ Terjadi kesalahan saat verifikasi.' };
  }
}

// === HELPER: Get active admins from MASTER ===
async function getActiveAdmins() {
  try {
    const data = await getSheetData(MASTER_SHEET);
    const admins = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i]) continue;
      const uname = (data[i][8] || '').replace('@', '').trim();
      const role = (data[i][9] || '').toUpperCase().trim();
      const status = (data[i][10] || '').toUpperCase().trim();
      if (role === 'ADMIN' && status === 'AKTIF' && uname) admins.push(uname);
    }
    return admins;
  } catch (error) {
    console.error('Error getting admins:', error.message);
    return [];
  }
}

// === HELPER: Get workzone mappings from ORDER ASSURANCE (auto-detect columns) ===
async function getWorkzoneMappings() {
  const data = await getSheetData(ORDER_ASSURANCE_SHEET, false);
  if (!data || data.length < 2) return [];

  // Auto-detect MAPPING TEAM and WORKZONE columns from header
  const header = data[0] || [];
  let teamColIdx = -1;
  let wzColIdx = -1;

  for (let c = 0; c < header.length; c++) {
    const h = (header[c] || '').toUpperCase().trim();
    if (h.includes('MAPPING') && h.includes('TEAM')) teamColIdx = c;
    // Only match WORKZONE columns after column 10 (to skip main WORKZONE at col E)
    if (c > 10 && h.includes('WORKZONE')) wzColIdx = c;
  }

  // Fallback to Q(16) and R(17) if not found
  if (teamColIdx === -1) teamColIdx = 16;
  if (wzColIdx === -1) wzColIdx = teamColIdx + 1;

  console.log(`📍 Mapping columns detected: MAPPING TEAM=col ${teamColIdx}, WORKZONE=col ${wzColIdx}`);

  const mappings = [];
  for (let i = 1; i < data.length; i++) {
      if (!data[i]) continue;
    const team = (data[i][teamColIdx] || '').trim();
    const wz = (data[i][wzColIdx] || '').trim();
    if (team && wz) mappings.push({ team, workzone: wz });
  }

  console.log(`📍 Found ${mappings.length} workzone mappings`);

  // Remove duplicates
  const seen = new Set();
  return mappings.filter(m => {
    const key = `${m.team}|${m.workzone}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

// === HELPER: Find mapping team for a workzone ===
function findMappingTeam(workzone, mappings) {
  if (!workzone) return null;
  for (const mapping of mappings) {
    if (mapping.workzone.toUpperCase() === 'ALL') continue;
    const zones = mapping.workzone.split(/[&,]/).map(z => z.trim().toUpperCase());
    if (zones.includes(workzone.toUpperCase())) return mapping.team;
  }
  const allMap = mappings.find(m => m.workzone.toUpperCase() === 'ALL');
  return allMap ? allMap.team : null;
}

// === HELPER: Parse TTR duration string to hours ===
function parseTTRHours(ttrStr) {
  if (!ttrStr || ttrStr === '-' || ttrStr === '') return null;
  const parts = ttrStr.toString().trim().split(':');
  if (parts.length >= 2) {
    const h = parseFloat(parts[0]) || 0;
    const m = parseFloat(parts[1]) || 0;
    const s = parts.length > 2 ? (parseFloat(parts[2]) || 0) : 0;
    return h + m / 60 + s / 3600;
  }
  const num = parseFloat(ttrStr);
  return isNaN(num) ? null : num;
}

// === HELPER: Format hours to HH:MM:SS ===
function formatHours(hours) {
  const h = Math.floor(hours);
  const m = Math.floor((hours - h) * 60);
  const s = Math.floor(((hours - h) * 60 - m) * 60);
  return `${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
}

// === HELPER: Parse Indonesian date ===
function parseIndonesianDate(dateStr) {
  if (!dateStr) return null;
  const cleaned = dateStr.replace(/^[^,]*,\s*/, '').trim();
  const parts = cleaned.split(/\s+/);
  if (parts.length < 3) return null;
  const day = parseInt(parts[0]);
  const month = BULAN_ID[parts[1].toLowerCase()];
  const year = parseInt(parts[2]);
  if (!day || !month || !year) return null;
  return { day, month, year };
}

// === HELPER: Get today in Jakarta timezone ===
function getTodayJakarta() {
  const now = new Date();
  return {
    day: parseInt(now.toLocaleDateString('id-ID', { day: 'numeric', timeZone: 'Asia/Jakarta' })),
    month: parseInt(now.toLocaleDateString('id-ID', { month: 'numeric', timeZone: 'Asia/Jakarta' })),
    year: parseInt(now.toLocaleDateString('id-ID', { year: 'numeric', timeZone: 'Asia/Jakarta' })),
  };
}

// === HELPER: Parse assurance input text ===
function parseAssurance(text, username) {
  let data = {
    incidentNo: '', closeDesc: '',
    dropcore: '', patchcord: '', soc: '', pslave: '',
    passive1_8: '', passive1_4: '', pigtail: '', adaptor: '',
    roset: '', rj45: '', lan: '',
    dateCreated: new Date().toLocaleDateString('id-ID', {
      weekday: 'long', day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta',
    }),
    teknisi: (username || '').replace('@', ''),
  };

  const incidentMatch = text.match(/INC[0-9]+/i);
  if (incidentMatch) data.incidentNo = incidentMatch[0].trim().toUpperCase();

  const closeMatch = text.match(/CLOSE\s*:\s*(.+?)(?=\n|MATERIAL|$)/i);
  if (closeMatch && closeMatch[1]) data.closeDesc = closeMatch[1].trim();

  const patterns = {
    dropcore: /DROPCORE\s*:\s*([0-9\.]+)/i, patchcord: /PATCHCORD\s*:\s*([0-9\.]+)/i,
    soc: /SOC\s*:\s*([0-9\.]+)/i, pslave: /PSLAVE\s*:\s*([0-9\.]+)/i,
    passive1_8: /PASSIVE\s*1\/8\s*:\s*([0-9\.]+)/i, passive1_4: /PASSIVE\s*1\/4\s*:\s*([0-9\.]+)/i,
    pigtail: /PIGTAIL\s*:\s*([0-9\.]+)/i, adaptor: /ADAPTOR\s*:\s*([0-9\.]+)/i,
    roset: /ROSET\s*:\s*([0-9\.]+)/i, rj45: /RJ\s*45\s*:\s*([0-9\.]+)/i,
    lan: /LAN\s*:\s*([0-9\.]+)/i,
  };
  for (const [key, pattern] of Object.entries(patterns)) {
    const match = text.match(pattern);
    if (match && match[1]) data[key] = match[1].trim();
  }
  return data;
}

// === HELPER: Parse tiket baru message ===
function parseTicketBaru(text) {
  if (!text.includes('Data Tiket Baru Diterima')) return null;

  const result = {};
  const patterns = {
    teknisi: /Teknisi:\s*(.+)/i,
    incident: /Incident:\s*(INC\d+)/i,
    ttrCustomer: /TTR Customer:\s*(.+)/i,
    reportedDate: /Reported Date:\s*(.+)/i,
    tiketGamas: /Tiket gamas:\s*(.*)/i,
    contactPhone: /Contact Phone:\s*(.+)/i,
    contactName: /Contact name:\s*(.+)/i,
    customerType: /Customer Type:\s*(.+)/i,
    serviceNo: /Service no:\s*(.+)/i,
    odp: /ODP:\s*(.+)/i,
    bookingDate: /Booking date:\s*(.+)/i,
    hasilUkur: /Hasil Ukur:\s*(.+)/i,
    guaranteStatus: /Guarante Status:\s*(.+)/i,
    symtom: /Symtom:\s*(.+)/i,
  };

  for (const [key, pattern] of Object.entries(patterns)) {
    const match = text.match(pattern);
    result[key] = match ? match[1].trim() : '';
  }

  return result;
}

// === HELPER: Extract workzone from ODP string ===
// ODP-SGI-FH/33 FH/D02/33.01 → SGI
function extractWorkzoneFromODP(odp) {
  if (!odp) return '';
  const match = odp.match(/ODP-([A-Z0-9]+)-/i);
  return match ? match[1].toUpperCase() : '';
}



// === CORE: Get productivity data across all sheets ===
async function getProductivityData(dateFilterFn) {
  const safeGet = async (sheet) => { try { return await getSheetData(sheet, false); } catch (e) { console.log(`⚠️ Sheet "${sheet}" not found, skipping`); return null; } };
  const [orderData, progresData, masterData] = await Promise.all([
    getSheetData(ORDER_ASSURANCE_SHEET, false),
    getSheetData(ASSURANCE_SHEET, false),
    getSheetData(MASTER_SHEET, false),
  ]);
  const [sqmData, manualData, unspecData, gaulData] = await Promise.all([
    safeGet(SQM_SHEET),
    safeGet(MANUAL_GGN_SHEET),
    safeGet(UNSPEC_SHEET),
    safeGet(GAUL_ASSURANCE_SHEET),
  ]);

  const mappings = await getWorkzoneMappings();
  const teamConfig = getTeamConfig(masterData);
  const result = {};

  function ensureTeam(name) {
    if (!result[name]) {
      result[name] = {
        reguler: { closed: [], open: [] },
        unspec: { closed: [], open: [] },
        sqm: { closed: [], open: [] },
        manual: { closed: [] },
        gaul: { closed: [] },
      };
    }
  }

  // 1. REGULER
  const closedIncidents = new Set();
  if (progresData) {
    for (let i = 1; i < progresData.length; i++) {
      if (!progresData[i]) continue;
      const tanggal = (progresData[i][0] || '').trim();
      const incident = (progresData[i][1] || '').trim().toUpperCase();
      const d = parseIndonesianDate(tanggal);
      if (d && dateFilterFn(d) && incident) closedIncidents.add(incident);
    }
  }

  const orderCols = orderData ? getOrderColumns(orderData) : null;
  if (orderData && orderCols) {
    for (let i = 1; i < orderData.length; i++) {
      if (!orderData[i]) continue;
      const incident = (orderData[i][orderCols.incident] || '').trim();
      const teknisi = (orderData[i][orderCols.teknisi] || '').trim();
      const status = (orderData[i][orderCols.status] || '').toUpperCase().trim();
      const custType = (orderData[i][orderCols.customerType] || '').trim();
      const ttrStr = (orderData[i][orderCols.ttrCustomer] || '').trim();
      if (!incident) continue;

      const teamName = resolveTeamByTeknisi(teknisi, teamConfig);
      ensureTeam(teamName);

      if (closedIncidents.has(incident.toUpperCase())) {
        result[teamName].reguler.closed.push({ incident, teknisi, custType, ttr: ttrStr });
      } else if (status === 'OPEN') {
        result[teamName].reguler.open.push({ incident, teknisi, custType, ttr: ttrStr });
      }
    }
  }

  // 2. UNSPEC
  if (unspecData) {
    for (let i = 1; i < unspecData.length; i++) {
      if (!unspecData[i]) continue;
      const tanggal = (unspecData[i][0] || '').trim();
      const teknisi = (unspecData[i][2] || '').trim();
      const serviceNo = (unspecData[i][7] || '').trim();
      const deviceName = (unspecData[i][8] || '').trim();
      const status = (unspecData[i][9] || '').toUpperCase().trim();
      const d = parseIndonesianDate(tanggal);

      const teamName = resolveTeamByTeknisi(teknisi, teamConfig);
      ensureTeam(teamName);

      if (status === 'CLOSE' && d && dateFilterFn(d)) {
        result[teamName].unspec.closed.push({ serviceNo, teknisi, deviceName });
      } else if (status === 'OPEN') {
        result[teamName].unspec.open.push({ serviceNo, teknisi, deviceName });
      }
    }
  }

  // 3. SQM
  if (sqmData) {
    for (let i = 1; i < sqmData.length; i++) {
      if (!sqmData[i]) continue;
      const tanggal = (sqmData[i][0] || '').trim();
      const incident = (sqmData[i][1] || '').trim();
      const teknisi = (sqmData[i][2] || '').trim();
      const status = (sqmData[i][9] || '').toUpperCase().trim();
      const d = parseIndonesianDate(tanggal);

      const teamName = resolveTeamByTeknisi(teknisi, teamConfig);
      ensureTeam(teamName);

      if (status === 'CLOSE' && d && dateFilterFn(d)) {
        result[teamName].sqm.closed.push({ incident, teknisi });
      } else if (status === 'OPEN') {
        result[teamName].sqm.open.push({ incident, teknisi });
      }
    }
  }

  // 4. MANUAL GGN
  if (manualData) {
    for (let i = 1; i < manualData.length; i++) {
      if (!manualData[i]) continue;
      const tanggal = (manualData[i][0] || '').trim();
      const teknisi = (manualData[i][1] || '').trim();
      const serviceNo = (manualData[i][4] || '').trim();
      const d = parseIndonesianDate(tanggal);
      if (!d || !dateFilterFn(d)) continue;

      const teamName = resolveTeamByTeknisi(teknisi, teamConfig);
      ensureTeam(teamName);
      result[teamName].manual.closed.push({ serviceNo, teknisi, tanggal });
    }
  }

  // 5. GAUL ASSURANCE
  if (gaulData) {
    for (let i = 1; i < gaulData.length; i++) {
      if (!gaulData[i]) continue;
      const tanggal = (gaulData[i][0] || '').trim();
      const incident = (gaulData[i][1] || '').trim();
      const teknisi = (gaulData[i][2] || '').trim();
      const d = parseIndonesianDate(tanggal);
      if (!d) continue;

      const teamName = resolveTeamByTeknisi(teknisi, teamConfig);
      ensureTeam(teamName);

      if (dateFilterFn(d)) {
        result[teamName].gaul.closed.push({ incident, tanggal });
      }
    }
  }

  return { teams: result, weekday: isWeekday(), piket: getTodayPiket(getPiketSchedule(masterData)), teamConfig };
}


// === HELPER: Get team config from MASTER (username → sektor) ===
function getTeamConfig(masterData) {
  const config = new Map();
  if (!masterData || masterData.length < 3) return config;
  for (let i = 2; i < masterData.length; i++) {
      if (!masterData[i]) continue;
    const uname = (masterData[i][8] || '').replace('@', '').toLowerCase().trim();
    const sektor = (masterData[i][11] || '').trim().toUpperCase();
    const status = (masterData[i][10] || '').trim().toUpperCase();
    if (uname && sektor && status === 'AKTIF') config.set(uname, sektor);
  }
  return config;
}

// === HELPER: Get piket schedule from MASTER (T=19, U=20, V=21) ===
function getPiketSchedule(masterData) {
  const schedule = [];
  if (!masterData || masterData.length < 3) return schedule;
  for (let i = 2; i < masterData.length; i++) {
      if (!masterData[i]) continue;
    const tanggal = (masterData[i][19] || '').trim();
    const sektor = (masterData[i][20] || '').trim().toUpperCase();
    const teknisi = (masterData[i][21] || '').trim();
    if (tanggal && teknisi) schedule.push({ tanggal, sektor, teknisi });
  }
  return schedule;
}

// === HELPER: Check if today is weekday ===
function isWeekday() {
  const now = new Date();
  const dayName = now.toLocaleDateString('en-US', { weekday: 'long', timeZone: 'Asia/Jakarta' });
  return !['Saturday', 'Sunday'].includes(dayName);
}

// === HELPER: Get team display name ===
function getTeamDisplayName(sektor) {
  if (!sektor) return 'UNKNOWN';
  return `TEAM ${sektor.toUpperCase()}`;
}

// === HELPER: Shorten teknisi name ===
function shortTeknisi(name) {
  if (!name) return '-';
  return name.split(/\s*&\s*/).map(n => {
    n = n.replace('@', '').trim();
    return n.replace(/_?\d{5,}.*$/, '').replace(/_MRU$/, '');
  }).join(' & ');
}

// === HELPER: Short date format ===
function shortDate(dateStr) {
  if (!dateStr) return '-';
  const d = parseIndonesianDate(dateStr);
  if (!d) return dateStr.substring(0, 10);
  return String(d.day).padStart(2,'0') + '/' + String(d.month).padStart(2,'0');
}

// === HELPER: Build progress bar ===
function buildProgressBar(completed, total, width = 10) {
  if (total === 0) return '░'.repeat(width) + ' 0%';
  const pct = Math.round((completed / total) * 100);
  const filled = Math.round((completed / total) * width);
  return '█'.repeat(filled) + '░'.repeat(width - filled) + ` ${pct}%`;
}

// === HELPER: Resolve team from workzone via MASTER ===
function resolveTeamName(workzone, mappings, teamConfig) {
  const team = findMappingTeam(workzone, mappings);
  if (team) {
    const usernames = team.split(/\s*&\s*/).map(u => u.replace('@', '').trim().toLowerCase());
    for (const u of usernames) {
      const sektor = teamConfig.get(u);
      if (sektor && sektor !== 'ALL') return getTeamDisplayName(sektor);
    }
  }
  return workzone ? `TEAM ${workzone.toUpperCase()}` : 'UNKNOWN';
}

// === HELPER: Resolve team from TEKNISI username ===
function resolveTeamByTeknisi(teknisiStr, teamConfig) {
  if (!teknisiStr) return 'UNKNOWN';
  const TEAM_MAP = {
    'fh_demna': 'SGI', 'fh_rizqi': 'SGI',
    'fh_rayyan': 'BNN', 'fh_mhdfauzan': 'BNN',
    'fh_aqil': 'MRU', 'fh_rijal': 'MRU',
    'paman_sgi': 'MTC', 'fh_imam': 'MTC',
    'fh_ltm_ilham': 'LTM', 'fh_ltm_rozasafrian': 'LTM',
    'fh_syahrur_ltm': 'LTM', 'fh_mustaqim_ltm': 'LTM',
    'aldesgi': 'SGI',
  };
  const cleaned = teknisiStr.replace(/@/g, '').toLowerCase();
  for (const [key, team] of Object.entries(TEAM_MAP)) {
    if (cleaned.includes(key)) return getTeamDisplayName(team);
  }
  const usernames = teknisiStr.split(/[&,]/).map(u => u.replace('@', '').trim().toLowerCase());
  for (const u of usernames) {
    const sektor = teamConfig.get(u);
    if (sektor && sektor !== 'ALL') return getTeamDisplayName(sektor);
  }
  for (const u of usernames) {
    if (!u) continue;
    for (const [key, sektor] of teamConfig.entries()) {
      if (sektor && sektor !== 'ALL') {
        if (u.includes(key) || key.includes(u)) return getTeamDisplayName(sektor);
      }
    }
  }
  return 'UNKNOWN';
}

// === HELPER: Parse piket date (dd-mm-yyyy or dd.mm.yyyy) ===
function parsePiketDate(dateStr) {
  if (!dateStr) return null;
  const parts = dateStr.split(/[.\-]/);
  if (parts.length < 3) return null;
  return { day: parseInt(parts[0]), month: parseInt(parts[1]), year: parseInt(parts[2]) };
}

// === HELPER: Get today piket ===
function getTodayPiket(piketSchedule) {
  const today = getTodayJakarta();
  return piketSchedule.filter(p => {
    const d = parsePiketDate(p.tanggal);
    return d && d.day === today.day && d.month === today.month && d.year === today.year;
  });
}


// === BOT SETUP ===
const PORT = process.env.PORT || 3002;
const RAILWAY_STATIC_URL = process.env.RAILWAY_STATIC_URL;
const USE_WEBHOOK = !!RAILWAY_STATIC_URL;
let bot;

if (USE_WEBHOOK) {
  const express = require('express');
  const app = express();
  app.use(express.json());
  bot = new TelegramBot(TOKEN);
  const webhookUrl = `https://${RAILWAY_STATIC_URL}/assurance${TOKEN}`;
  bot.setWebHook(webhookUrl).then(() => console.log(`✅ Webhook set: ${webhookUrl}`)).catch(err => console.error('❌ Webhook error:', err.message));
  app.post(`/assurance${TOKEN}`, (req, res) => { bot.processUpdate(req.body); res.sendStatus(200); });
  app.get('/', (req, res) => res.send('Bot Assurance is running!'));
  app.listen(PORT, () => console.log(`✅ Server running on port ${PORT}`));
} else {
  bot = new TelegramBot(TOKEN, { polling: { interval: 300, autoStart: true, params: { timeout: 10, allowed_updates: ['message'] } } });
  console.log('✅ Bot running in polling mode');
  bot.on('polling_error', (error) => console.error(error.code === 'EFATAL' ? '❌ Polling fatal:' : '⚠️ Polling error:', error.message));
}

// ============================================================
// MONITORING FUNCTIONS (TTR Alerts + Auto-Fill Teknisi)
// ============================================================

async function autoFillTeknisi() {
  try {
    const data = await getSheetData(ORDER_ASSURANCE_SHEET, false);
    if (!data || data.length < 2) return;

    const cols = getOrderColumns(data);
    const mappings = await getWorkzoneMappings();
    let filled = 0;

    for (let i = 1; i < data.length; i++) {
      if (!data[i]) continue;
      const teknisi = (data[i][cols.teknisi] || '').trim();
      const status = (data[i][cols.status] || '').toUpperCase().trim();
      const workzone = (data[i][cols.workzone] || '').trim();

      if (teknisi || !status || status === 'CLOSE') continue;

      const team = findMappingTeam(workzone, mappings);
      if (team) {
        const rowNum = i + 1;
        await updateSheetCell(ORDER_ASSURANCE_SHEET, `C${rowNum}`, team);
        filled++;
        console.log(`🔧 Auto-fill teknisi row ${rowNum}: ${team}`);
      }
    }

    if (filled > 0) {
      cache.orderAssuranceData = null; // Invalidate cache
      console.log(`✅ Auto-fill: ${filled} teknisi filled`);
    }
  } catch (error) {
    console.error('❌ Auto-fill error:', error.message);
  }
}

async function checkTTRAlerts() {
  if (!GROUP_CHAT_ID) return;

  try {
    const data = await getSheetData(ORDER_ASSURANCE_SHEET, false);
    if (!data || data.length < 2) return;

    const cols = getOrderColumns(data);
    const mappings = await getWorkzoneMappings();
    const admins = await getActiveAdmins();
    const adminTags = admins.map(a => `@${a}`).join(' ');

    let expiredList = [];
    let warningList = [];
    let hasNewExpired = false;
    let hasNewWarning = false;

    for (let i = 1; i < data.length; i++) {
      if (!data[i]) continue;
      const incident = (data[i][cols.incident] || '').trim();
      const ttrStr = (data[i][cols.ttrCustomer] || '').trim();
      const workzone = (data[i][cols.workzone] || '').trim();
      const custType = (data[i][cols.customerType] || '').trim().toUpperCase();
      const status = (data[i][cols.status] || '').toUpperCase().trim();

      if (status !== 'OPEN' || !incident || !ttrStr) continue;

      const elapsed = parseTTRHours(ttrStr);
      if (elapsed === null) continue;

      let maxTTR = null;
      for (const [type, hours] of Object.entries(TTR_TABLE)) {
        if (type.toUpperCase() === custType) { maxTTR = hours; break; }
      }
      if (maxTTR === null) continue;

      const team = findMappingTeam(workzone, mappings) || (data[i][cols.teknisi] || '-').trim();
      const cleanTeam = team.replace(/@@/g, '@');

      // === EXPIRED ===
      if (elapsed >= maxTTR) {
        const overtime = elapsed - maxTTR;
        expiredList.push({ incident, ttrStr, custType, maxTTR, overtime, team: cleanTeam });
        if (!alertState.expired.has(incident)) {
          hasNewExpired = true;
          alertState.expired.add(incident);
          alertState.warned.delete(incident);
        }
      }
      // === WARNING (1 jam sebelum expired) ===
      else if (elapsed >= maxTTR - 1 && elapsed < maxTTR) {
        const sisa = maxTTR - elapsed;
        warningList.push({ incident, ttrStr, custType, maxTTR, sisa, team: cleanTeam });
        if (!alertState.warned.has(incident)) {
          hasNewWarning = true;
          alertState.warned.add(incident);
        }
      }
    }

    // Kirim alert EXPIRED jika ada ticket baru yang expired
    if (hasNewExpired && expiredList.length > 0) {
      expiredList.sort((a, b) => b.overtime - a.overtime);
      let msg = `🔴 EXPIRED: ${expiredList.length} tickets\n`;
      expiredList.forEach(e => {
        msg += `▸ ${e.incident} | ${e.ttrStr} | ${e.custType} (${e.maxTTR} Jam) | OT: +${formatHours(e.overtime)}\n`;
        msg += `  👷 ${e.team}\n`;
      });
      if (adminTags) msg += `\ncc bg ${adminTags}`;
      await sendTelegram(GROUP_CHAT_ID, msg);
      console.log(`🔴 TTR EXPIRED alert: ${expiredList.length} tickets`);
    }

    // Kirim alert WARNING jika ada ticket baru yang mendekati expired
    if (hasNewWarning && warningList.length > 0) {
      warningList.sort((a, b) => a.sisa - b.sisa);
      let msg = `⚠️ MENDEKATI EXPIRED:\n\n`;
      warningList.forEach((e, idx) => {
        msg += `${idx + 1}. ${e.incident}\n`;
        msg += `   ${e.custType} (Max: ${e.maxTTR} Jam)\n`;
        msg += `   Elapsed: ${e.ttrStr} | Sisa: ${formatHours(e.sisa)}\n`;
        msg += `   Teknisi: ${e.team}\n\n`;
      });
      if (adminTags) msg += `cc bg ${adminTags}`;
      await sendTelegram(GROUP_CHAT_ID, msg);
      console.log(`⚠️ TTR WARNING alert: ${warningList.length} tickets`);

      // DM ke setiap teknisi yang tiketnya mendekati expired
      for (const entry of warningList) {
        // Extract usernames dari team (e.g., "@user1 & @user2")
        const usernames = entry.team.split(/\s*&\s*/).map(u => u.replace('@', '').trim().toLowerCase()).filter(u => u);
        for (const uname of usernames) {
          const dmChatId = userChatIds[uname];
          if (dmChatId) {
            let dmMsg = `⚠️ PERINGATAN TTR MENDEKATI EXPIRED!\n\n`;
            dmMsg += `Incident: ${entry.incident}\n`;
            dmMsg += `${entry.custType} (Max: ${entry.maxTTR} Jam)\n`;
            dmMsg += `Elapsed: ${entry.ttrStr} | Sisa: ${formatHours(entry.sisa)}\n`;
            dmMsg += `\nSegera selesaikan tiket ini!`;
            try {
              await sendTelegram(dmChatId, dmMsg);
              console.log(`📩 DM warning sent to @${uname}`);
            } catch (dmErr) {
              console.log(`⚠️ Could not DM @${uname}: ${dmErr.message}`);
            }
          }
        }
      }
    }

    // Cleanup: remove closed incidents from alert state
    const openIncidents = new Set();
    for (let i = 1; i < data.length; i++) {
      if (!data[i]) continue;
      const st = (data[i][cols.status] || '').toUpperCase().trim();
      if (st === 'OPEN') openIncidents.add((data[i][cols.incident] || '').trim());
    }
    for (const inc of alertState.warned) { if (!openIncidents.has(inc)) alertState.warned.delete(inc); }
    for (const inc of alertState.expired) { if (!openIncidents.has(inc)) alertState.expired.delete(inc); }

  } catch (error) {
    console.error('❌ TTR check error:', error.message);
  }
}

// === AUTO-POST: Build sisa ticket report (reusable) ===
async function buildSisaTicketReport() {
  const data = await getSheetData(ORDER_ASSURANCE_SHEET, false);
  if (!data || data.length < 2) return null;

  const cols = getOrderColumns(data);
  const mappings = await getWorkzoneMappings();

  const now = new Date();
  const dayNameID = ['MINGGU', 'SENIN', 'SELASA', 'RABU', 'KAMIS', 'JUMAT', 'SABTU'][now.getDay()];
  const dateStr = now.toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta' }).toUpperCase();
  const timeStr = now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', timeZone: 'Asia/Jakarta' });

  let ticketsByTeam = {};
  let totalOpen = 0;
  let gamasTickets = {};
  let totalGamas = 0;

  for (let i = 1; i < data.length; i++) {
      if (!data[i]) continue;
    const incident = (data[i][cols.incident] || '-').trim();
    const ttrCustomer = (data[i][cols.ttrCustomer] || '-').trim();
    const workzone = (data[i][cols.workzone] || '').trim();
    const custType = (data[i][cols.customerType] || '').trim();
    const hasilUkur = cols.hasilUkur >= 0 ? (data[i][cols.hasilUkur] || '').trim() : '';
    const status = (data[i][cols.status] || '').toUpperCase().trim();

    if (status === 'OPEN') {
      const team = findMappingTeam(workzone, mappings) || workzone || '-';
      const cleanTeam = team.replace(/@@/g, '@');
      if (!ticketsByTeam[cleanTeam]) ticketsByTeam[cleanTeam] = [];
      ticketsByTeam[cleanTeam].push({ incident, ttr: ttrCustomer, custType, hasilUkur });
      totalOpen++;
    } else if (status === 'GAMAS') {
      const team = findMappingTeam(workzone, mappings) || workzone || '-';
      const cleanTeam = team.replace(/@@/g, '@');
      if (!gamasTickets[cleanTeam]) gamasTickets[cleanTeam] = [];
      gamasTickets[cleanTeam].push({ incident, ttr: ttrCustomer, custType, hasilUkur });
      totalGamas++;
    }
  }

  const numEmoji = ['1️⃣', '2️⃣', '3️⃣', '4️⃣', '5️⃣', '6️⃣', '7️⃣', '8️⃣', '9️⃣', '🔟'];
  const sortedTeams = Object.keys(ticketsByTeam).sort();

  let response = `🔴 <b>SISA TICKET OPEN</b>\n📅 ${dayNameID}, ${dateStr} | ${timeStr}\n\n`;
  response += `🎫 <b>OPEN : ${totalOpen} tickets</b>\n\n`;

  if (sortedTeams.length === 0) {
    response += '<i>Tidak ada ticket yang masih OPEN</i>\n';
  } else {
    sortedTeams.forEach((teamName, idx) => {
      const tickets = ticketsByTeam[teamName].sort((a, b) => (parseTTRHours(a.ttr) || 0) - (parseTTRHours(b.ttr) || 0));
      const num = idx < 10 ? numEmoji[idx] : `${idx + 1}.`;
      response += `${num} <b>${teamName}</b> [${tickets.length}]\n`;
      tickets.forEach(t => {
        const hu = t.hasilUkur ? ` | ${t.hasilUkur}` : '';
        response += `   🔹 ${t.incident} | ${t.ttr} | ${t.custType}${hu}\n`;
      });
      response += '\n';
    });
  }

  if (totalGamas > 0) {
    response += `🎫 <b>GAMAS : ${totalGamas} tickets</b>\n\n`;
    const sortedGamas = Object.keys(gamasTickets).sort();
    sortedGamas.forEach((teamName, idx) => {
      const tickets = gamasTickets[teamName].sort((a, b) => (parseTTRHours(a.ttr) || 0) - (parseTTRHours(b.ttr) || 0));
      const num = idx < 10 ? numEmoji[idx] : `${idx + 1}.`;
      response += `${num} <b>${teamName}</b> [${tickets.length}]\n`;
      tickets.forEach(t => {
        const hu = t.hasilUkur ? ` | ${t.hasilUkur}` : '';
        response += `   🔹 ${t.incident} | ${t.ttr} | ${t.custType}${hu}\n`;
      });
      response += '\n';
    });
  }

  return response;
}

// === AUTO-POST: Kirim sisa ticket ke group setiap 1 jam ===
async function autoPostSisaTicket() {
  if (!GROUP_CHAT_ID) return;
  try {
    const report = await buildSisaTicketReport();
    if (report) {
      await sendTelegram(GROUP_CHAT_ID, report);
      console.log('📊 Auto-post sisa ticket ke group');
    }
  } catch (error) {
    console.error('❌ Auto-post sisa ticket error:', error.message);
  }
}

// ============================================================
// MESSAGE HANDLER
// ============================================================
bot.on('message', async (msg) => {
  const chatId = msg.chat.id;
  const msgId = msg.message_id;
  const text = (msg.text || '').trim();
  const username = msg.from.username || '';
  const groupType = msg.chat.type;

  if (!text) return;

  // Store chat ID for potential DM (from private and group)
  if (username) {
    const unameLower = username.replace('@', '').toLowerCase();
    // Simpan user's Telegram ID untuk DM (msg.from.id selalu user ID, bukan group ID)
    if (msg.from.id) userChatIds[unameLower] = msg.from.id;
  }

  console.log(`📨 [${groupType}] [@${username}] ${text.substring(0, 60)}`);

  try {
    // ============================================================
    // AUTO-INPUT TIKET BARU ke ORDER ASSURANCE
    // Deteksi pesan "Data Tiket Baru Diterima" (non-command)
    // ============================================================
    if (text.includes('Data Tiket Baru Diterima')) {
      try {
        const tiket = parseTicketBaru(text);
        if (!tiket || !tiket.incident) {
          console.log('⚠️ Tiket baru detected but could not parse incident');
          return;
        }

        // Get ORDER ASSURANCE data
        const orderData = await getSheetData(ORDER_ASSURANCE_SHEET, false);
        const orderCols = getOrderColumns(orderData);

        // Cek duplikat incident
        for (let i = 1; i < orderData.length; i++) {
      if (!orderData[i]) continue;
          const existingInc = (orderData[i][orderCols.incident] || '').trim().toUpperCase();
          if (existingInc === tiket.incident.toUpperCase()) {
            console.log(`⚠️ Duplicate incident ${tiket.incident} - skipping`);
            return sendTelegram(chatId, `⚠️ Incident <b>${tiket.incident}</b> sudah ada di ORDER ASSURANCE.`, { reply_to_message_id: msgId });
          }
        }

        // Extract workzone from ODP (e.g., ODP-SGI-FH/33 → SGI)
        const workzone = extractWorkzoneFromODP(tiket.odp);

        // Map teknisi using workzone → MAPPING TEAM
        const mappings = await getWorkzoneMappings();
        const mappedTeknisi = findMappingTeam(workzone, mappings) || tiket.teknisi;

        // Format tanggal saat ini
        const tanggal = new Date().toLocaleDateString('id-ID', {
          weekday: 'long', day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta',
        });

        // Format timestamp saat ini (hanya waktu HH:mm:ss)
        const now = new Date();
        const jam = String(now.toLocaleString('id-ID', { hour: '2-digit', timeZone: 'Asia/Jakarta', hour12: false })).padStart(2, '0');
        const menit = String(now.toLocaleString('id-ID', { minute: '2-digit', timeZone: 'Asia/Jakarta' })).padStart(2, '0');
        const detik = String(now.toLocaleString('id-ID', { second: '2-digit', timeZone: 'Asia/Jakarta' })).padStart(2, '0');
        const timestamp = `${jam}:${menit}:${detik}`;

        // Find next row with empty incident (preserves formulas in D and L)
        const nextRow = findNextEmptyRow(orderData, orderCols.incident);

        // Batch update specific cells — SKIP kolom D (rumus) dan L (rumus)
        const cellUpdates = [
          { range: `A${nextRow}`, value: tanggal },              // A - Tanggal
          { range: `B${nextRow}`, value: tiket.incident },       // B - Incident
          { range: `C${nextRow}`, value: mappedTeknisi },        // C - Teknisi (mapped)
          // D - SKIP (rumus TTR CUSTOMER)
          { range: `E${nextRow}`, value: workzone },             // E - Workzone
          { range: `F${nextRow}`, value: 'TSEL' },               // F - Customer Segment
          { range: `G${nextRow}`, value: tiket.customerType },   // G - Customer Type
          { range: `H${nextRow}`, value: tiket.serviceNo },      // H - Service No
          { range: `I${nextRow}`, value: tiket.odp },            // I - Device Name (ODP)
          { range: `J${nextRow}`, value: 'OPEN' },               // J - Status
          { range: `K${nextRow}`, value: tiket.ttrCustomer || timestamp }, // K - TTR Dashboard (durasi, fallback ke timestamp jika kosong)
          // L - SKIP (rumus NOW)
          { range: `M${nextRow}`, value: timestamp },            // M - Timestamp
        ];

        await withTimeout(batchUpdateCells(ORDER_ASSURANCE_SHEET, cellUpdates), 10000);
        cache.orderAssuranceData = null; // Invalidate cache

        console.log(`✅ Tiket baru recorded at row ${nextRow}: ${tiket.incident} | Teknisi: ${mappedTeknisi} | WZ: ${workzone}`);

        // === GAUL DETECTION: Cek gangguan berulang (Service No sama) ===
        if (tiket.serviceNo) {
          try {
            const prevIncidents = [];
            for (let i = 1; i < orderData.length; i++) {
      if (!orderData[i]) continue;
              const existingSN = (orderData[i][orderCols.incident ? 7 : 7] || '').trim(); // H = Service No
              const existingInc = (orderData[i][orderCols.incident] || '').trim();
              const existingDate = (orderData[i][0] || '').trim(); // A = Tanggal
              // Cek kolom H (index 7) untuk Service No
              const svcNo = (orderData[i][7] || '').trim();
              if (svcNo === tiket.serviceNo && existingInc && existingInc.toUpperCase() !== tiket.incident.toUpperCase()) {
                prevIncidents.push({ incident: existingInc, tanggal: existingDate });
              }
            }

            if (prevIncidents.length > 0 && GROUP_CHAT_ID) {
              const admins = await getActiveAdmins();
              const adminTags = admins.map(a => `@${a}`).join(' ');
              let gaulMsg = `🔁 GANGGUAN BERULANG TERDETEKSI!\n\n`;
              gaulMsg += `📞 Service No: ${tiket.serviceNo}\n\n`;
              gaulMsg += `▸ Tiket Sebelumnya:\n`;
              prevIncidents.forEach(p => {
                gaulMsg += `  📅 ${p.tanggal} | ${p.incident}\n`;
              });
              gaulMsg += `\n▸ Tiket Baru:\n`;
              gaulMsg += `  📅 ${tanggal} | ${tiket.incident}\n\n`;
              gaulMsg += `👷 Teknisi: ${mappedTeknisi}\n`;
              gaulMsg += `📍 Workzone: ${workzone}\n`;
              gaulMsg += `👤 Customer: ${tiket.customerType}\n\n`;
              gaulMsg += `⚠️ Total gangguan: ${prevIncidents.length + 1}x untuk Service No ini\n\n`;
              if (adminTags) gaulMsg += `cc bg ${adminTags}`;
              await sendTelegram(GROUP_CHAT_ID, gaulMsg);
              console.log(`🔁 GAUL detected: ${tiket.serviceNo} (${prevIncidents.length + 1}x)`);
            }
          } catch (gaulErr) {
            console.error('⚠️ GAUL detection error:', gaulErr.message);
          }
        }

        let confirmMsg = `✅ <b>Data Tiket Baru berhasil disimpan!</b> (Row ${nextRow})\n\n`;
        confirmMsg += `📋 <b>Incident:</b> ${tiket.incident}\n`;
        confirmMsg += `👷 <b>Teknisi:</b> ${mappedTeknisi}\n`;
        confirmMsg += `📍 <b>Workzone:</b> ${workzone}`;

        return sendTelegram(chatId, confirmMsg, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ Tiket Baru Error:', err.message);
        return sendTelegram(chatId, `❌ Error menyimpan tiket: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // Skip non-command messages after tiket baru check
    if (!text.startsWith('/')) return;
    // ============================================================
    // /INPUT - Input data assurance + auto-close di ORDER ASSURANCE
    // ============================================================
    if (/^\/INPUT\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const inputText = text.replace(/^\/INPUT\s*/i, '').trim();
        if (!inputText) return sendTelegram(chatId, '❌ Silakan kirim data assurance setelah /INPUT.', { reply_to_message_id: msgId });

        const parsed = parseAssurance(inputText, username);
        const missing = ['incidentNo', 'closeDesc'].filter(f => !parsed[f]);
        if (missing.length > 0) return sendTelegram(chatId, `❌ Field wajib: ${missing.join(', ')}`, { reply_to_message_id: msgId });

        // Simpan ke PROGRES ASSURANCE (termasuk timestamp di kolom P)
        const inputTimestamp = new Date().toLocaleString('id-ID', {
          day: '2-digit', month: '2-digit', year: 'numeric',
          hour: '2-digit', minute: '2-digit', second: '2-digit',
          timeZone: 'Asia/Jakarta', hour12: false,
        });
        const row = [
          parsed.dateCreated, parsed.incidentNo,
          parsed.dropcore, parsed.patchcord, parsed.soc, parsed.pslave,
          parsed.passive1_8, parsed.passive1_4, parsed.pigtail, parsed.adaptor,
          parsed.roset, parsed.rj45, parsed.lan, parsed.closeDesc, parsed.teknisi,
          inputTimestamp, // P - Timestamp saat /INPUT
        ];
        await withTimeout(appendSheetData(ASSURANCE_SHEET, row), 10000);
        cache.assuranceData = null; // Invalidate cache agar /rekap_hari langsung update

        // Auto-close di ORDER ASSURANCE + cek KAWAL TTR (COMPLY/NOT COMPLY)
        let orderClosed = false;
        let kawalTTR = '';
        try {
          const orderData = await getSheetData(ORDER_ASSURANCE_SHEET, false);
          const orderCols = getOrderColumns(orderData);
          const statusColLetter = String.fromCharCode(65 + orderCols.status);
          for (let i = 1; i < orderData.length; i++) {
      if (!orderData[i]) continue;
            const incInOrder = (orderData[i][orderCols.incident] || '').trim().toUpperCase();
            if (incInOrder === parsed.incidentNo) {
              // Cek TTR untuk KAWAL TTR
              const ttrStr = (orderData[i][orderCols.ttrCustomer] || '').trim();
              const custType = (orderData[i][orderCols.customerType] || '').trim().toUpperCase();
              const elapsed = parseTTRHours(ttrStr);

              // Cari max TTR dari tabel berdasarkan customer type
              let maxTTR = null;
              for (const [type, hours] of Object.entries(TTR_TABLE)) {
                if (type.toUpperCase() === custType) { maxTTR = hours; break; }
              }

              // Tentukan COMPLY atau NOT COMPLY
              if (elapsed !== null && maxTTR !== null) {
                kawalTTR = elapsed <= maxTTR ? 'COMPLY' : 'NOT COMPLY';
              } else {
                kawalTTR = 'COMPLY'; // Default jika data tidak lengkap
              }

              // Update STATUS ke CLOSE dan KAWAL TTR ke kolom R
              await batchUpdateCells(ORDER_ASSURANCE_SHEET, [
                { range: `${statusColLetter}${i + 1}`, value: 'CLOSE' },
                { range: `R${i + 1}`, value: kawalTTR },
              ]);
              cache.orderAssuranceData = null;
              alertState.warned.delete(parsed.incidentNo);
              alertState.expired.delete(parsed.incidentNo);
              orderClosed = true;
              console.log(`✅ Auto-close ORDER: ${parsed.incidentNo} row ${i + 1} | KAWAL TTR: ${kawalTTR}`);
              break;
            }
          }
        } catch (closeErr) {
          console.error('⚠️ Auto-close error:', closeErr.message);
        }

        let confirmMsg = `✅ Data Assurance berhasil disimpan!\n\n`;
        confirmMsg += `<b>Incident:</b> ${parsed.incidentNo}\n`;
        confirmMsg += `<b>Close:</b> ${parsed.closeDesc}\n`;
        if (orderClosed) confirmMsg += `<b>Status ORDER:</b> ✅ Auto-CLOSE | <b>KAWAL TTR:</b> ${kawalTTR}\n`;
        confirmMsg += `<b>Material:</b>\n`;
        confirmMsg += `  • Dropcore: ${parsed.dropcore || '-'}\n`;
        confirmMsg += `  • Patchcord: ${parsed.patchcord || '-'}\n`;
        confirmMsg += `  • SOC: ${parsed.soc || '-'}\n`;
        confirmMsg += `  • PSLAVE: ${parsed.pslave || '-'}\n`;
        confirmMsg += `  • PASSIVE 1/8: ${parsed.passive1_8 || '-'}\n`;
        confirmMsg += `  • PASSIVE 1/4: ${parsed.passive1_4 || '-'}\n`;
        confirmMsg += `  • Pigtail: ${parsed.pigtail || '-'}\n`;
        confirmMsg += `  • Adaptor: ${parsed.adaptor || '-'}\n`;
        confirmMsg += `  • Roset: ${parsed.roset || '-'}\n`;
        confirmMsg += `  • RJ 45: ${parsed.rj45 || '-'}\n`;
        confirmMsg += `  • LAN: ${parsed.lan || '-'}`;

        return sendTelegram(chatId, confirmMsg, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /INPUT Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /sisa_ticket - Ticket OPEN grouped by mapping team
    // ============================================================
    else if (/^\/sisa_ticket\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const response = await withTimeout(buildSisaTicketReport(), 10000);
        if (!response) return sendTelegram(chatId, '<i>Tidak ada data</i>', { reply_to_message_id: msgId });

        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /sisa_ticket Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /cek_ttr - Cek TTR warning & expired secara manual
    // ============================================================
    else if (/^\/cek_ttr\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const data = await withTimeout(getSheetData(ORDER_ASSURANCE_SHEET, false), 10000);
        if (!data || data.length < 2) return sendTelegram(chatId, '<i>Tidak ada data</i>', { reply_to_message_id: msgId });

        const cols = getOrderColumns(data);
        const mappings = await getWorkzoneMappings();
        let expiredList = [];
        let warningList = [];
        let safeList = [];

        for (let i = 1; i < data.length; i++) {
      if (!data[i]) continue;
          const incident = (data[i][cols.incident] || '').trim();
          const ttrStr = (data[i][cols.ttrCustomer] || '').trim();
          const workzone = (data[i][cols.workzone] || '').trim();
          const custType = (data[i][cols.customerType] || '').trim().toUpperCase();
          const status = (data[i][cols.status] || '').toUpperCase().trim();

          if (status !== 'OPEN' || !incident || !ttrStr) continue;

          const elapsed = parseTTRHours(ttrStr);
          if (elapsed === null) continue;

          let maxTTR = null;
          for (const [type, hours] of Object.entries(TTR_TABLE)) {
            if (type.toUpperCase() === custType) { maxTTR = hours; break; }
          }
          if (maxTTR === null) continue;

          const team = findMappingTeam(workzone, mappings) || (data[i][cols.teknisi] || '-').trim();
          const cleanTeam = team.replace(/@@/g, '@');

          const entry = { incident, ttrStr, custType, maxTTR, elapsed, team: cleanTeam };

          if (elapsed >= maxTTR) {
            entry.overtime = elapsed - maxTTR;
            expiredList.push(entry);
          } else if (elapsed >= maxTTR - 1) {
            entry.sisa = maxTTR - elapsed;
            warningList.push(entry);
          } else {
            safeList.push(entry);
          }
        }

        // Sort by overtime/sisa descending
        expiredList.sort((a, b) => b.overtime - a.overtime);
        warningList.sort((a, b) => a.sisa - b.sisa);

        let response = `⏱ <b>STATUS TTR - SEMUA TICKET OPEN</b>\n\n`;

        // EXPIRED
        response += `🔴 <b>EXPIRED: ${expiredList.length} tickets</b>\n`;
        if (expiredList.length === 0) {
          response += '<i>Tidak ada ticket expired</i>\n\n';
        } else {
          expiredList.forEach(e => {
            response += `▸ ${e.incident} | ${e.ttrStr} | ${e.custType} (${e.maxTTR} Jam) | OT: +${formatHours(e.overtime)}\n`;
            response += `  👷 ${e.team}\n`;
          });
          response += '\n';
        }

        // WARNING
        response += `⚠️ <b>MENDEKATI EXPIRED: ${warningList.length} tickets</b>\n`;
        if (warningList.length === 0) {
          response += '<i>Tidak ada ticket mendekati expired</i>\n\n';
        } else {
          warningList.forEach(e => {
            response += `▸ ${e.incident} | ${e.ttrStr} | ${e.custType} (${e.maxTTR} Jam) | Sisa: ${formatHours(e.sisa)}\n`;
            response += `  👷 ${e.team}\n`;
          });
          response += '\n';
        }

        // SAFE
        response += `✅ <b>AMAN: ${safeList.length} tickets</b>`;

        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /cek_ttr Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /material_used - Total material keseluruhan
    // ============================================================
    else if (/^\/material_used\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const data = await withTimeout(getSheetData(ASSURANCE_SHEET), 10000);
        let materialMap = {
          'DROPCORE': 0, 'PATCHCORD': 0, 'SOC': 0, 'PSLAVE': 0,
          'PASSIVE 1/8': 0, 'PASSIVE 1/4': 0, 'PIGTAIL': 0, 'ADAPTOR': 0,
          'ROSET': 0, 'RJ 45': 0, 'LAN': 0,
        };
        const materialColumns = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
        const materialNames = Object.keys(materialMap);

        for (let i = 1; i < data.length; i++) {
      if (!data[i]) continue;
          materialColumns.forEach((colIdx, idx) => {
            materialMap[materialNames[idx]] += parseInt((data[i][colIdx] || '0').trim()) || 0;
          });
        }

        const entries = Object.entries(materialMap).filter(([_, c]) => c > 0).sort((a, b) => b[1] - a[1]);
        let response = `━━━━━━━━━━━━━━━━━━━━━━\n📦 <b>PENGGUNAAN MATERIAL</b>\n━━━━━━━━━━━━━━━━━━━━━━\n\n`;
        if (entries.length === 0) {
          response += '<i>Belum ada material yang dipakai</i>';
        } else {
          entries.forEach(([mat, count]) => { response += `📊 <b>${mat}</b> : ${count} unit\n`; });
        }

        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /material_used Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /rekap_hari - Rekap close HARI INI
    // ============================================================
    else if (/^\/rekap_hari\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });
        const isAdmin = authResult.role === 'ADMIN';
        const today = getTodayJakarta();
        const dateFilter = d => d.day === today.day && d.month === today.month && d.year === today.year;
        const prodData = await withTimeout(getProductivityData(dateFilter), 15000);
        const { teams } = prodData;
        const todayStr = new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta' });
        if (isAdmin) {
          let grandClosed = 0, grandOpen = 0;
          const teamStats = [];
          for (const [name, t] of Object.entries(teams)) {
            if (name === 'UNKNOWN') continue;
            const reg = t.reguler.closed.length, uns = t.unspec.closed.length, sqm = t.sqm.closed.length, man = t.manual.closed.length;
            const closed = reg + uns + sqm + man;
            const open = t.reguler.open.length + t.unspec.open.length + t.sqm.open.length;
            grandClosed += closed; grandOpen += open;
            const tekCount = {};
            [...t.reguler.closed, ...t.unspec.closed, ...t.sqm.closed, ...t.manual.closed].forEach(tk => {
              const tek = shortTeknisi(tk.teknisi || '-');
              tekCount[tek] = (tekCount[tek] || 0) + 1;
            });
            teamStats.push({ name, closed, open, reg, uns, sqm, man, tekCount });
          }
          teamStats.sort((a, b) => b.closed - a.closed);
          const medal = ['🥇', '🥈', '🥉'];
          let response = `╔══════════════════════════════════╗\n`;
          response += `║  📊 <b>REKAP CLOSE - HARI INI</b>\n`;
          response += `║  📅 ${todayStr}\n`;
          response += `╠══════════════════════════════════╣\n\n`;
          teamStats.forEach((ts, i) => {
            const icon = i < 3 ? medal[i] : '🔸';
            response += `${icon} <b>${ts.name}</b>: ${ts.closed} closed\n`;
            response += `    REG: ${ts.reg} | UNS: ${ts.uns} | SQM: ${ts.sqm} | MAN: ${ts.man}\n`;
            const tekEntries = Object.entries(ts.tekCount).sort((a,b) => b[1]-a[1]);
            if (tekEntries.length > 0) response += `    👷 ${tekEntries.map(([t,c]) => `@${t} (${c})`).join(' ')}\n`;
            response += '\n';
          });
          response += `📋 <b>Total: ${grandClosed} closed | ${grandOpen} open</b>\n╚══════════════════════════════════╝`;
          return sendTelegram(chatId, response, { reply_to_message_id: msgId });
        } else {
          let userReg = 0, userUns = 0, userSqm = 0, userMan = 0;
          for (const [, t] of Object.entries(teams)) {
            t.reguler.closed.forEach(tk => { if ((tk.teknisi||'').toLowerCase().includes(username.toLowerCase())) userReg++; });
            t.unspec.closed.forEach(tk => { if ((tk.teknisi||'').toLowerCase().includes(username.toLowerCase())) userUns++; });
            t.sqm.closed.forEach(tk => { if ((tk.teknisi||'').toLowerCase().includes(username.toLowerCase())) userSqm++; });
            t.manual.closed.forEach(tk => { if ((tk.teknisi||'').toLowerCase().includes(username.toLowerCase())) userMan++; });
          }
          const total = userReg + userUns + userSqm + userMan;
          let response = `╔══════════════════════════════════╗\n║  📊 <b>REKAP CLOSE - HARI INI</b>\n║  📅 ${todayStr} | @${username}\n╠══════════════════════════════════╣\n\n`;
          response += `📋 REGULER : ${userReg}\n📋 UNSPEC  : ${userUns}\n📋 SQM     : ${userSqm}\n📋 MANUAL  : ${userMan}\n\n📋 <b>Total: ${total} closed</b>\n╚══════════════════════════════════╝`;
          return sendTelegram(chatId, response, { reply_to_message_id: msgId });
        }
      } catch (err) {
        console.error('❌ /rekap_hari Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // /rekap_bulan
    else if (/^\/rekap_bulan\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });
        const isAdmin = authResult.role === 'ADMIN';
        const today = getTodayJakarta();
        const dateFilter = d => d.month === today.month && d.year === today.year;
        const prodData = await withTimeout(getProductivityData(dateFilter), 15000);
        const { teams } = prodData;
        const bulanStr = new Date().toLocaleDateString('id-ID', { month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta' });
        if (isAdmin) {
          let grandClosed = 0;
          const teamStats = [];
          for (const [name, t] of Object.entries(teams)) {
            if (name === 'UNKNOWN') continue;
            const reg = t.reguler.closed.length, uns = t.unspec.closed.length, sqm = t.sqm.closed.length, man = t.manual.closed.length;
            const closed = reg + uns + sqm + man;
            grandClosed += closed;
            const tekCount = {};
            [...t.reguler.closed, ...t.unspec.closed, ...t.sqm.closed, ...t.manual.closed].forEach(tk => {
              const s = shortTeknisi(tk.teknisi || '-');
              tekCount[s] = (tekCount[s] || 0) + 1;
            });
            teamStats.push({ name, closed, reg, uns, sqm, man, tekCount });
          }
          teamStats.sort((a, b) => b.closed - a.closed);
          const medal = ['🥇', '🥈', '🥉'];
          let response = `╔══════════════════════════════════╗\n║  📊 <b>REKAP CLOSE - BULAN INI</b>\n║  📅 ${bulanStr}\n╠══════════════════════════════════╣\n\n`;
          teamStats.forEach((ts, i) => {
            const icon = i < 3 ? medal[i] : '🔸';
            response += `${icon} <b>${ts.name}</b> : ${ts.closed} closed\n    REG: ${ts.reg} | UNS: ${ts.uns} | SQM: ${ts.sqm} | MAN: ${ts.man}\n`;
            const tekEntries = Object.entries(ts.tekCount).sort((a,b) => b[1]-a[1]);
            if (tekEntries.length > 0) response += `    👷 ${tekEntries.map(([t,c]) => `@${t} (${c})`).join(' ')}\n`;
            response += '\n';
          });
          response += `📋 <b>Total: ${grandClosed} closed</b>\n╚══════════════════════════════════╝`;
          return sendTelegram(chatId, response, { reply_to_message_id: msgId });
        } else {
          let userReg = 0, userUns = 0, userSqm = 0, userMan = 0;
          for (const [, t] of Object.entries(teams)) {
            t.reguler.closed.forEach(tk => { if ((tk.teknisi||'').toLowerCase().includes(username.toLowerCase())) userReg++; });
            t.unspec.closed.forEach(tk => { if ((tk.teknisi||'').toLowerCase().includes(username.toLowerCase())) userUns++; });
            t.sqm.closed.forEach(tk => { if ((tk.teknisi||'').toLowerCase().includes(username.toLowerCase())) userSqm++; });
            t.manual.closed.forEach(tk => { if ((tk.teknisi||'').toLowerCase().includes(username.toLowerCase())) userMan++; });
          }
          const total = userReg + userUns + userSqm + userMan;
          let response = `╔══════════════════════════════════╗\n║  📊 <b>REKAP CLOSE - BULAN INI</b>\n║  📅 ${bulanStr} | @${username}\n╠══════════════════════════════════╣\n\n`;
          response += `📋 REGULER : ${userReg}\n📋 UNSPEC  : ${userUns}\n📋 SQM     : ${userSqm}\n📋 MANUAL  : ${userMan}\n\n📋 <b>Total: ${total} closed</b>\n╚══════════════════════════════════╝`;
          return sendTelegram(chatId, response, { reply_to_message_id: msgId });
        }
      } catch (err) {
        console.error('❌ /rekap_bulan Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // /rekap_tahun
    else if (/^\/rekap_tahun\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });
        const isAdmin = authResult.role === 'ADMIN';
        const today = getTodayJakarta();
        const bulanNames = ['','Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
        const monthTeamData = {};
        let grandTotal = 0;
        const userMonthData = {};
        for (let m = 1; m <= today.month; m++) {
          const dateFilter = d => d.month === m && d.year === today.year;
          try {
            const prodData = await withTimeout(getProductivityData(dateFilter), 10000);
            const { teams } = prodData;
            let monthTotal = 0;
            const teamList = [];
            let uTotal = 0, uR = 0, uU = 0, uS = 0, uM = 0;
            for (const [name, t] of Object.entries(teams)) {
              if (name === 'UNKNOWN') continue;
              const reg = t.reguler.closed.length, uns = t.unspec.closed.length, sqm = t.sqm.closed.length, man = t.manual.closed.length;
              const closed = reg + uns + sqm + man;
              if (closed > 0) { monthTotal += closed; teamList.push({ name, closed, reg, uns, sqm, man }); }
              if (!isAdmin) {
                t.reguler.closed.forEach(tk => { if ((tk.teknisi||'').toLowerCase().includes(username.toLowerCase())) { uR++; uTotal++; } });
                t.unspec.closed.forEach(tk => { if ((tk.teknisi||'').toLowerCase().includes(username.toLowerCase())) { uU++; uTotal++; } });
                t.sqm.closed.forEach(tk => { if ((tk.teknisi||'').toLowerCase().includes(username.toLowerCase())) { uS++; uTotal++; } });
                t.manual.closed.forEach(tk => { if ((tk.teknisi||'').toLowerCase().includes(username.toLowerCase())) { uM++; uTotal++; } });
              }
            }
            if (monthTotal > 0) { teamList.sort((a, b) => b.closed - a.closed); monthTeamData[m] = { total: monthTotal, teams: teamList }; grandTotal += monthTotal; }
            if (uTotal > 0) userMonthData[m] = { total: uTotal, reg: uR, uns: uU, sqm: uS, man: uM };
          } catch (e) {}
        }
        if (isAdmin) {
          let response = `╔══════════════════════════════════╗\n║  📊 <b>REKAP CLOSE - TAHUN ${today.year}</b>\n╠══════════════════════════════════╣\n\n`;
          const sortedMonths = Object.keys(monthTeamData).map(Number).sort((a,b)=>a-b);
          if (sortedMonths.length === 0) response += '<i>Belum ada data</i>\n';
          else {
            sortedMonths.forEach(m => {
              const md = monthTeamData[m];
              response += `📅 <b>${bulanNames[m].toUpperCase()}</b> [${md.total} tiket]\n`;
              md.teams.forEach(ts => {
                response += `├─ ${ts.name} : ${ts.closed}\n│  REG: ${ts.reg} | UNS: ${ts.uns} | SQM: ${ts.sqm} | MAN: ${ts.man}\n`;
              });
              response += '\n';
            });
            response += `📋 <b>Grand Total: ${grandTotal} tiket</b>\n`;
          }
          response += `╚══════════════════════════════════╝`;
          return sendTelegram(chatId, response, { reply_to_message_id: msgId });
        } else {
          let userGrand = 0;
          let response = `╔══════════════════════════════════╗\n║  📊 <b>REKAP CLOSE - TAHUN ${today.year}</b>\n║  👤 @${username}\n╠══════════════════════════════════╣\n\n`;
          const sortedMonths = Object.keys(userMonthData).map(Number).sort((a,b)=>a-b);
          if (sortedMonths.length === 0) response += '<i>Belum ada data</i>\n';
          else {
            sortedMonths.forEach(m => {
              const md = userMonthData[m];
              userGrand += md.total;
              response += `📅 ${bulanNames[m].toUpperCase()} : ${md.total} closed\n    REG: ${md.reg} | UNS: ${md.uns} | SQM: ${md.sqm} | MAN: ${md.man}\n\n`;
            });
            response += `📋 <b>Grand Total: ${userGrand} tiket</b>\n`;
          }
          response += `╚══════════════════════════════════╝`;
          return sendTelegram(chatId, response, { reply_to_message_id: msgId });
        }
      } catch (err) {
        console.error('❌ /rekap_tahun Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /REKAP_JANUARI s/d /REKAP_DESEMBER - Rekap bulan spesifik
    // ============================================================
    else if (/^\/REKAP_(JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const bulanMatch = text.match(/^\/REKAP_(\w+)/i);
        const bulanName = bulanMatch[1].toLowerCase();
        const targetMonth = BULAN_ID[bulanName];
        if (!targetMonth) return sendTelegram(chatId, '❌ Bulan tidak valid.', { reply_to_message_id: msgId });

        const data = await withTimeout(getSheetData(ASSURANCE_SHEET), 10000);
        const orderData = await withTimeout(getSheetData(ORDER_ASSURANCE_SHEET), 10000);
        const today = getTodayJakarta();
        let map = {};
        let incidents = [];

        for (let i = 1; i < data.length; i++) {
      if (!data[i]) continue;
          const tanggal = (data[i][0] || '').trim();
          const teknisi = (data[i][14] || '-').trim();
          const incident = (data[i][1] || '').trim();
          const d = parseIndonesianDate(tanggal);
          if (!d) continue;
          if (d.month === targetMonth && d.year === today.year) {
            map[teknisi] = (map[teknisi] || 0) + 1;
            if (incident) incidents.push(incident.toUpperCase());
          }
        }

        let comply = 0, notComply = 0;
        for (let i = 1; i < orderData.length; i++) {
      if (!orderData[i]) continue;
          const inc = (orderData[i][1] || '').trim().toUpperCase();
          const kawal = (orderData[i][17] || '').trim().toUpperCase();
          if (incidents.includes(inc)) {
            if (kawal === 'COMPLY') comply++;
            else if (kawal === 'NOT COMPLY') notComply++;
          }
        }

        const bulanNames = ['', 'JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI', 'JULI', 'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER'];
        const entries = Object.entries(map).sort((a, b) => b[1] - a[1]);
        const total = entries.reduce((sum, [_, c]) => sum + c, 0);

        const medal = ['🥇', '🥈', '🥉'];
        let response = `━━━━━━━━━━━━━━━━━━━━━━\n📊 <b>REKAP CLOSE - ${bulanNames[targetMonth]} ${today.year}</b>\n━━━━━━━━━━━━━━━━━━━━━━\n\n`;
        if (entries.length === 0) {
          response += `<i>Belum ada data ${bulanNames[targetMonth].toLowerCase()}</i>`;
        } else {
          entries.forEach(([tek, c], i) => {
            const icon = i < 3 ? medal[i] : '🔸';
            response += `${icon} <b>${tek}</b> : ${c} tickets\n`;
          });
          response += `\n📋 <b>Total: ${total} tickets</b>\n`;
          response += `✅ COMPLY: ${comply} | ❌ NOT COMPLY: ${notComply}`;
        }

        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /REKAP_BULAN Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /unspec - Close tiket UNSPEC
    // ============================================================
    else if (/^\/unspec\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const inputText = text.replace(/^\/unspec\s*/i, '').trim();
        if (!inputText) {
          return sendTelegram(chatId, `❌ Format:\n/unspec\nCLOSE: deskripsi\nSERVICE NO: 111149103305\nWORKZONE: SLG`, { reply_to_message_id: msgId });
        }

        const closeMatch = inputText.match(/CLOSE\s*:\s*(.+)/i);
        const svcMatch = inputText.match(/SERVICE\s*NO\s*:\s*(.+)/i);
        const wzMatch = inputText.match(/WORKZONE\s*:\s*(.+)/i);

        if (!closeMatch || !svcMatch) {
          return sendTelegram(chatId, `❌ Format:\n/unspec\nCLOSE: deskripsi\nSERVICE NO: 111149103305\nWORKZONE: SLG`, { reply_to_message_id: msgId });
        }

        const closeDesc = closeMatch[1].trim();
        const serviceNo = svcMatch[1].trim();
        const wz = wzMatch ? wzMatch[1].trim().toUpperCase() : '';

        const tanggal = new Date().toLocaleDateString('id-ID', {
          weekday: 'long', day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta',
        });

        const mappings = await getWorkzoneMappings();
        const teknisi = wz ? (findMappingTeam(wz, mappings) || `@${username}`) : `@${username}`;

        // Simpan ke UNSPEC: A=tanggal, B=closeDesc, C=teknisi, D-G=(skip), H=serviceNo, I=deviceName, J=status
        const row = [tanggal, '', teknisi, '', '', '', '', serviceNo, closeDesc, 'CLOSE'];
        await withTimeout(appendSheetData(UNSPEC_SHEET, row), 10000);

        let confirmMsg = `✅ Tiket UNSPEC berhasil di-close!\n\n`;
        confirmMsg += `📋 Detail:\n`;
        confirmMsg += `📅 Tanggal: ${tanggal}\n`;
        confirmMsg += `👷 Teknisi: ${teknisi}\n`;
        confirmMsg += `📞 Service No: ${serviceNo}\n`;
        confirmMsg += `📝 Close: ${closeDesc}\n`;
        confirmMsg += `📊 Status: CLOSE`;

        return sendTelegram(chatId, confirmMsg, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /unspec Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

        // ============================================================
    // /MANUAL - Input gangguan manual ke MANUAL GGN
    // ============================================================
    else if (/^\/MANUAL\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const inputText = text.replace(/^\/MANUAL\s*/i, '').trim();
        if (!inputText) {
          return sendTelegram(chatId, `❌ Format tidak sesuai. Gunakan format:\n\n/MANUAL\nCLOSE: deskripsi perbaikan\nSERVICE NO: 111149103305\nWORKZONE: SLG`, { reply_to_message_id: msgId });
        }

        const closeMatch = inputText.match(/CLOSE\s*:\s*(.+)/i);
        const svcMatch = inputText.match(/SERVICE\s*NO\s*:\s*(.+)/i);
        const wzMatch = inputText.match(/WORKZONE\s*:\s*(.+)/i);

        if (!closeMatch || !svcMatch || !wzMatch) {
          return sendTelegram(chatId, `❌ Format tidak sesuai. Gunakan format:\n\n/MANUAL\nCLOSE: deskripsi perbaikan\nSERVICE NO: 111149103305\nWORKZONE: SLG`, { reply_to_message_id: msgId });
        }

        const closeDesc = closeMatch[1].trim();
        const serviceNo = svcMatch[1].trim();
        const wz = wzMatch[1].trim().toUpperCase();

        // Get teknisi dari mapping
        const orderData = await getSheetData(ORDER_ASSURANCE_SHEET, false);
        const mappings = await getWorkzoneMappings();
        const teknisi = findMappingTeam(wz, mappings) || `@${username}`;

        const tanggal = new Date().toLocaleDateString('id-ID', {
          weekday: 'long', day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta',
        });

        // Simpan ke MANUAL GGN: A=tanggal, B=teknisi, C=(skip), D=workzone, E=serviceNo, F=status
        const row = [tanggal, teknisi, closeDesc, wz, serviceNo, 'CLOSE'];
        await withTimeout(appendSheetData(MANUAL_GGN_SHEET, row), 10000);

        let confirmMsg = `✅ Data Gangguan Manual berhasil disimpan!\n\n`;
        confirmMsg += `📋 Detail:\n`;
        confirmMsg += `📅 Tanggal: ${tanggal}\n`;
        confirmMsg += `👷 Teknisi: ${teknisi}\n`;
        confirmMsg += `📍 Workzone: ${wz}\n`;
        confirmMsg += `📞 Service No: ${serviceNo}\n`;
        confirmMsg += `📝 Close: ${closeDesc}\n`;
        confirmMsg += `📊 Status: CLOSE`;

        return sendTelegram(chatId, confirmMsg, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /MANUAL Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /REKAP_MANUAL - Rekap gangguan manual per bulan
    // ============================================================
    else if (/^\/REKAP_MANUAL\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const isAdmin = authResult.role === 'ADMIN';
        const data = await withTimeout(getSheetData(MANUAL_GGN_SHEET), 10000);
        const today = getTodayJakarta();
        const bulanNames = ['', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

        let monthData = {};
        let grandTotal = 0;

        for (let i = 1; i < data.length; i++) {
      if (!data[i]) continue;
          const tanggal = (data[i][0] || '').trim();
          const teknisi = (data[i][1] || '').trim();
          const serviceNo = (data[i][4] || '').trim();
          const d = parseIndonesianDate(tanggal);
          if (!d || d.year !== today.year) continue;

          // Filter: user hanya lihat milik sendiri
          if (!isAdmin && !teknisi.toLowerCase().includes(username.toLowerCase())) continue;

          if (!monthData[d.month]) monthData[d.month] = [];
          monthData[d.month].push({ tanggal, serviceNo, teknisi });
          grandTotal++;
        }

        const sortedMonths = Object.keys(monthData).map(Number).sort((a, b) => a - b);
        let response = `📋 <b>DETAIL LAPORAN GANGGUAN MANUAL</b>\n📅 Tahun ${today.year}\n\n`;

        if (sortedMonths.length === 0) {
          response += '<i>Belum ada data</i>';
        } else {
          sortedMonths.forEach(month => {
            const items = monthData[month];
            response += `📅 <b>${bulanNames[month].toUpperCase()}</b> [${items.length} TIKET]\n`;
            items.forEach(item => {
              if (isAdmin) {
                response += `  ${item.tanggal} | ${item.serviceNo} | ${item.teknisi}\n`;
              } else {
                response += `  ${item.tanggal} | ${item.serviceNo}\n`;
              }
            });
            response += '\n';
          });
          response += `📋 <b>Grand Total: ${grandTotal} tiket</b>`;
        }

        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /REKAP_MANUAL Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /REKAP_UNSPEC - Rekap UNSPEC
    // ============================================================
    else if (/^\/REKAP_UNSPEC\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });
        const isAdmin = authResult.role === 'ADMIN';

        const data = await withTimeout(getSheetData(UNSPEC_SHEET), 10000);
        if (!data || data.length < 2) return sendTelegram(chatId, '<i>Tidak ada data UNSPEC</i>', { reply_to_message_id: msgId });

        const today = getTodayJakarta();
        const bulanNames = ['','Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
        let monthData = {};
        let grandTotal = 0;

        for (let i = 1; i < data.length; i++) {
          if (!data[i]) continue;
          const tanggal = (data[i][0] || '').trim();
          const teknisi = (data[i][2] || '').trim();
          const serviceNo = (data[i][7] || '').trim();
          const status = (data[i][9] || '').toUpperCase().trim();
          const d = parseIndonesianDate(tanggal);
          if (!d || d.year !== today.year || status !== 'CLOSE') continue;
          if (!isAdmin && !teknisi.toLowerCase().includes(username.toLowerCase())) continue;

          if (!monthData[d.month]) monthData[d.month] = [];
          monthData[d.month].push({ tanggal, serviceNo, teknisi });
          grandTotal++;
        }

        const sortedMonths = Object.keys(monthData).map(Number).sort((a,b) => a-b);
        let response = `╔══════════════════════════════════╗\n║  📊 <b>REKAP UNSPEC - TAHUN ${today.year}</b>\n╠══════════════════════════════════╣\n\n`;
        if (sortedMonths.length === 0) {
          response += '<i>Belum ada data</i>\n';
        } else {
          sortedMonths.forEach(month => {
            const items = monthData[month];
            response += `📅 <b>${bulanNames[month].toUpperCase()}</b> [${items.length} TIKET]\n`;
            items.forEach(item => {
              if (isAdmin) {
                response += `  ${item.tanggal} | ${item.serviceNo} | ${item.teknisi}\n`;
              } else {
                response += `  ${item.tanggal} | ${item.serviceNo}\n`;
              }
            });
            response += '\n';
          });
          response += `📋 <b>Grand Total: ${grandTotal} tiket</b>\n`;
        }
        response += `╚══════════════════════════════════╝`;
        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /REKAP_UNSPEC Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

        // ============================================================
    // /TICKET_SQM - Lihat & Pick Up tiket SQM
    // ============================================================
    else if (/^\/TICKET_SQM\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const data = await withTimeout(getSheetData(SQM_SHEET, false), 10000);
        if (!data || data.length < 2) return sendTelegram(chatId, '<i>Tidak ada data SQM</i>', { reply_to_message_id: msgId });

        // Cari tiket OPEN milik user berdasarkan username di kolom C
        const userTickets = [];
        for (let i = 1; i < data.length; i++) {
      if (!data[i]) continue;
          const teknisi = (data[i][2] || '').trim(); // C = Teknisi
          const status = (data[i][9] || '').trim().toUpperCase(); // J = Status
          const incident = (data[i][1] || '').trim(); // B = Incident
          const serviceNo = (data[i][7] || '').trim(); // H = Service No
          const deviceName = (data[i][8] || '').trim(); // I = Device Name

          if (status === 'OPEN' && teknisi.toLowerCase().includes(username.toLowerCase()) && incident) {
            userTickets.push({ incident, serviceNo, deviceName, row: i + 1 });
          }
        }

        if (userTickets.length === 0) {
          return sendTelegram(chatId, `📋 Ticket SQM Anda (@${username}):\n\n✅ Tidak ada tiket OPEN saat ini.`, { reply_to_message_id: msgId });
        }

        const numEmoji = ['1️⃣', '2️⃣', '3️⃣', '4️⃣', '5️⃣', '6️⃣', '7️⃣', '8️⃣', '9️⃣', '🔟'];
        let response = `📋 Ticket SQM Anda (@${username}):\n\n`;
        userTickets.forEach((t, idx) => {
          const num = idx < 10 ? numEmoji[idx] : `${idx + 1}.`;
          response += `${num} ${t.incident} | ${t.serviceNo} | ${t.deviceName}\n`;
        });
        response += `\n📌 Total: ${userTickets.length} tiket OPEN\n\nKirim INCIDENT yang ingin di Pick Up:\nContoh: INC46230392`;

        // Simpan state untuk pick up
        if (!global.pendingPickup) global.pendingPickup = {};
        global.pendingPickup[chatId + '_' + username] = {
          tickets: userTickets,
          timestamp: Date.now(),
        };

        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /TICKET_SQM Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /SQM - Input close SQM (mirip /INPUT tapi untuk SQM SA SIGLI)
    // ============================================================
    else if (/^\/SQM\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const inputText = text.replace(/^\/SQM\s*/i, '').trim();
        if (!inputText) return sendTelegram(chatId, '❌ Silakan kirim data setelah /SQM.', { reply_to_message_id: msgId });

        const parsed = parseAssurance(inputText, username);
        const missing = ['incidentNo', 'closeDesc'].filter(f => !parsed[f]);
        if (missing.length > 0) return sendTelegram(chatId, `❌ Field wajib: ${missing.join(', ')}`, { reply_to_message_id: msgId });

        // Simpan ke PROGRES ASSURANCE
        const inputTimestamp = new Date().toLocaleString('id-ID', {
          day: '2-digit', month: '2-digit', year: 'numeric',
          hour: '2-digit', minute: '2-digit', second: '2-digit',
          timeZone: 'Asia/Jakarta', hour12: false,
        });
        const row = [
          parsed.dateCreated, parsed.incidentNo,
          parsed.dropcore, parsed.patchcord, parsed.soc, parsed.pslave,
          parsed.passive1_8, parsed.passive1_4, parsed.pigtail, parsed.adaptor,
          parsed.roset, parsed.rj45, parsed.lan, parsed.closeDesc, parsed.teknisi,
          inputTimestamp,
        ];
        await withTimeout(appendSheetData(SQM_SHEET, row), 10000);
        cache.assuranceData = null;

        // Auto-close di SQM SA SIGLI
        let sqmClosed = false;
        try {
          const sqmData = await getSheetData(SQM_SHEET, false);
          for (let i = 1; i < sqmData.length; i++) {
      if (!sqmData[i]) continue;
            const incInSqm = (sqmData[i][1] || '').trim().toUpperCase(); // B = Incident
            if (incInSqm === parsed.incidentNo) {
              await updateSheetCell(SQM_SHEET, `J${i + 1}`, 'CLOSE'); // J = Status
              sqmClosed = true;
              console.log(`✅ Auto-close SQM: ${parsed.incidentNo} row ${i + 1}`);
              break;
            }
          }
        } catch (closeErr) {
          console.error('⚠️ SQM close error:', closeErr.message);
        }

        let confirmMsg = `✅ Data SQM berhasil disimpan!\n\nClose: ${parsed.closeDesc}`;

        return sendTelegram(chatId, confirmMsg, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /SQM Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /REKAP_SQM - Rekap SQM per bulan (INCIDENT)
    // ============================================================
    else if (/^\/REKAP_SQM\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const isAdmin = authResult.role === 'ADMIN';
        const data = await withTimeout(getSheetData(SQM_SHEET, false), 10000);
        const today = getTodayJakarta();
        const bulanNames = ['', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

        let monthData = {};
        let grandTotal = 0;

        for (let i = 1; i < data.length; i++) {
      if (!data[i]) continue;
          const tanggal = (data[i][0] || '').trim(); // A = Tanggal
          const incident = (data[i][1] || '').trim(); // B = Incident
          const teknisi = (data[i][2] || '').trim(); // C = Teknisi
          const status = (data[i][9] || '').trim().toUpperCase(); // J = Status

          if (status !== 'CLOSE') continue;

          const d = parseIndonesianDate(tanggal);
          if (!d || d.year !== today.year) continue;

          // Filter: user hanya lihat milik sendiri
          if (!isAdmin && !teknisi.toLowerCase().includes(username.toLowerCase())) continue;

          if (!monthData[d.month]) monthData[d.month] = [];
          monthData[d.month].push({ tanggal, incident, teknisi });
          grandTotal++;
        }

        const sortedMonths = Object.keys(monthData).map(Number).sort((a, b) => a - b);
        let response = `📋 <b>DETAIL LAPORAN SQM SA SIGLI</b>\n📅 Tahun ${today.year}\n\n`;

        if (sortedMonths.length === 0) {
          response += '<i>Belum ada data</i>';
        } else {
          sortedMonths.forEach(month => {
            const items = monthData[month];
            response += `📅 <b>${bulanNames[month].toUpperCase()}</b> [${items.length} TIKET]\n`;
            items.forEach(item => {
              if (isAdmin) {
                response += `  ${item.tanggal} | ${item.incident} | ${item.teknisi}\n`;
              } else {
                response += `  ${item.tanggal} | ${item.incident}\n`;
              }
            });
            response += '\n';
          });
          response += `📋 <b>Grand Total: ${grandTotal} tiket</b>`;
        }

        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /REKAP_SQM Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // Handle Pick Up response untuk /TICKET_SQM
    // ============================================================
    else if (/^INC\d+$/i.test(text) && global.pendingPickup && global.pendingPickup[chatId + '_' + username]) {
      try {
        const pending = global.pendingPickup[chatId + '_' + username];
        // Expired after 5 minutes
        if (Date.now() - pending.timestamp > 5 * 60 * 1000) {
          delete global.pendingPickup[chatId + '_' + username];
          return sendTelegram(chatId, '❌ Sesi Pick Up sudah expired. Kirim /TICKET_SQM lagi.', { reply_to_message_id: msgId });
        }

        const incidentNo = text.trim().toUpperCase();
        const ticket = pending.tickets.find(t => t.incident.toUpperCase() === incidentNo);
        if (!ticket) {
          return sendTelegram(chatId, `❌ Incident ${incidentNo} tidak ditemukan di list tiket Anda.`, { reply_to_message_id: msgId });
        }

        // Update kolom Q (index 16) = PICK UP di SQM SA SIGLI
        await updateSheetCell(SQM_SHEET, `Q${ticket.row}`, 'PICK UP');
        delete global.pendingPickup[chatId + '_' + username];

        let confirmMsg = `✅ Ticket berhasil di Pick Up!\n\n`;
        confirmMsg += `📋 Detail:\n`;
        confirmMsg += `🔹 Incident: ${ticket.incident}\n`;
        confirmMsg += `📞 Service No: ${ticket.serviceNo}\n`;
        confirmMsg += `📍 Device: ${ticket.deviceName}\n`;
        confirmMsg += `📊 Progres: PICK UP`;

        return sendTelegram(chatId, confirmMsg, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ Pick Up Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /rekap_piket - Rekap performa piket
    // ============================================================
    else if (/^\/rekap_piket\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });
        const isAdmin = authResult.role === 'ADMIN';
        const today = getTodayJakarta();
        const bulanNames = ['','Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
        const masterData = await withTimeout(getSheetData(MASTER_SHEET), 10000);
        const piketSchedule = getPiketSchedule(masterData);
        const dateFilter = d => d.month === today.month && d.year === today.year;
        const prodData = await withTimeout(getProductivityData(dateFilter), 15000);
        const { teams } = prodData;
        const allClosed = [];
        for (const [, t] of Object.entries(teams)) {
          [...t.reguler.closed, ...t.unspec.closed, ...t.sqm.closed, ...t.manual.closed].forEach(tk => {
            allClosed.push({ teknisi: (tk.teknisi || '').toLowerCase(), date: tk.tanggal || '' });
          });
        }
        const currentPiket = piketSchedule.filter(p => {
          const parts = p.tanggal.split(/[.\-]/);
          if (parts.length === 3) return parseInt(parts[1]) === today.month && parseInt(parts[2]) === today.year;
          return false;
        });
        const tekPiket = {};
        currentPiket.forEach(p => {
          const tek = shortTeknisi(p.teknisi);
          if (!tekPiket[tek]) tekPiket[tek] = { days: 0, details: [], closedOnPiket: 0 };
          tekPiket[tek].days++;
          const parts = p.tanggal.split(/[.\-]/);
          const piketDay = parseInt(parts[0]);
          let closedCount = 0;
          const tekLower = tek.toLowerCase().replace('@','');
          const origLower = p.teknisi.replace('@','').toLowerCase();
          allClosed.forEach(cl => {
            if (cl.teknisi.includes(tekLower) || cl.teknisi.includes(origLower)) {
              const d = parseIndonesianDate(cl.date);
              if (d && d.day === piketDay && d.month === today.month) closedCount++;
            }
          });
          tekPiket[tek].closedOnPiket += closedCount;
          tekPiket[tek].details.push({ date: parts[0]+'/'+parts[1], sektor: p.sektor, closed: closedCount, isPast: piketDay <= today.day });
        });
        const bulanStr = bulanNames[today.month] + ' ' + today.year;
        if (isAdmin) {
          let response = `╔══════════════════════════════════╗\n║  📅 <b>REKAP PIKET - ${bulanStr.toUpperCase()}</b>\n╠══════════════════════════════════╣\n\n`;
          const entries = Object.entries(tekPiket).sort((a,b) => b[1].closedOnPiket - a[1].closedOnPiket);
          if (entries.length === 0) response += '<i>Belum ada jadwal piket</i>\n';
          else {
            let totalDays = 0, totalClosed = 0;
            entries.forEach(([tek, data]) => {
              totalDays += data.days; totalClosed += data.closedOnPiket;
              const rate = data.days > 0 ? (data.closedOnPiket / data.days).toFixed(1) : '0';
              response += `👷 @${tek}\n├─ 🗓 Piket: ${data.days} hari\n├─ 📋 Closed: ${data.closedOnPiket}\n├─ ⚡ Rate: ${rate} /hari\n\n`;
            });
            response += `📋 <b>Total: ${totalDays} hari piket | ${totalClosed} tiket</b>\n`;
          }
          response += `╚══════════════════════════════════╝`;
          return sendTelegram(chatId, response, { reply_to_message_id: msgId });
        } else {
          let response = `╔══════════════════════════════════╗\n║  📅 <b>REKAP PIKET - ${bulanStr.toUpperCase()}</b>\n║  👤 @${username}\n╠══════════════════════════════════╣\n\n`;
          let found = false;
          for (const [tek, data] of Object.entries(tekPiket)) {
            if (tek.toLowerCase().includes(username.toLowerCase()) || username.toLowerCase().includes(tek.toLowerCase())) {
              found = true;
              response += `🗓 Piket: ${data.days} hari\n`;
              data.details.forEach(d => {
                if (d.isPast) response += `├─ ${d.date} ${d.sektor} | ${d.closed} closed ${d.closed > 0 ? '✅' : '⚠️'}\n`;
                else response += `├─ ${d.date} ${d.sektor} | (belum)\n`;
              });
              const pastDays = data.details.filter(d => d.isPast).length;
              const rate = pastDays > 0 ? (data.closedOnPiket / pastDays).toFixed(1) : '0';
              response += `\n📋 <b>Total: ${data.closedOnPiket} closed | ⚡ Rate: ${rate} /hari</b>\n`;
              break;
            }
          }
          if (!found) response += '<i>Tidak ada jadwal piket untuk Anda bulan ini</i>\n';
          response += `╚══════════════════════════════════╝`;
          return sendTelegram(chatId, response, { reply_to_message_id: msgId });
        }
      } catch (err) {
        console.error('❌ /rekap_piket Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

        // ============================================================
    // /chatid - Get chat ID (untuk setup GROUP_CHAT_ID)
    // ============================================================
    else if (/^\/chatid\b/i.test(text)) {
      return sendTelegram(chatId, `📍 <b>Chat ID:</b> <code>${chatId}</code>\n<b>Type:</b> ${msg.chat.type}`, { reply_to_message_id: msgId });
    }

    // ============================================================
    // /help or /start
    // ============================================================
    else if (/^\/(help|start)\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const helpMsg = `🤖 <b>Bot Assurance</b>

<b>📝 INPUT COMMAND:</b>
/INPUT - Input assurance (auto-close ORDER)
/SQM - Input SQM (auto-close SQM SA SIGLI)
/MANUAL - Input gangguan manual
/unspec - Close tiket UNSPEC

<b>📋 SQM:</b>
/TICKET_SQM - Lihat & Pick Up tiket SQM

<b>📊 PRODUKTIVITAS:</b>
/rekap_piket - Performa piket

<b>📊 MONITORING (ADMIN):</b>
/sisa_ticket - Ticket OPEN
/cek_ttr - Cek TTR warning & expired
/material_used - Total material

<b>📈 REKAP:</b>
/rekap_hari | /rekap_bulan | /rekap_tahun
/REKAP_JANUARI s/d /REKAP_DESEMBER
/REKAP_MANUAL | /REKAP_SQM | /REKAP_UNSPEC

<b>📋 FORMAT /MANUAL & /unspec:</b>
CLOSE: deskripsi perbaikan
SERVICE NO: 111149103305
WORKZONE: SLG

<b>⚙️ FITUR OTOMATIS:</b>
• Auto-fill teknisi & TTR monitoring
• Auto-close & deteksi GAUL
• Weekday=TEAM | Weekend=PIKET`;
        return sendTelegram(chatId, helpMsg, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /help Error:', err.message);
        return sendTelegram(chatId, '❌ Terjadi kesalahan.', { reply_to_message_id: msgId });
      }
    }

  } catch (err) {
    console.error('Error:', err.message);
    sendTelegram(chatId, '❌ Terjadi kesalahan sistem.', { reply_to_message_id: msgId });
  }
});

// === ERROR HANDLER ===
process.on('unhandledRejection', (reason) => console.error('Error:', reason));

// === STARTUP ===
console.log('\n🚀 Bot Assurance started!');
console.log(`Mode: ${USE_WEBHOOK ? 'Webhook' : 'Polling'}`);
console.log('═'.repeat(50));
console.log('✅ Auto-Cache Enabled (5 min expiry)');
console.log('✅ TTR Monitoring Enabled (5 min interval)');
console.log('✅ Auto-Post Sisa Ticket Enabled (1 jam interval)');
console.log('✅ Auto-Fill Teknisi Enabled');
console.log('✅ Timeout Protection Enabled');
console.log('═'.repeat(50));

// Start monitoring after 10 seconds delay
setTimeout(() => {
  console.log('🔄 Starting initial monitoring check...');
  autoFillTeknisi();
  checkTTRAlerts();

  // TTR check + Auto-fill setiap 5 menit
  setInterval(() => {
    autoFillTeknisi();
    checkTTRAlerts();
  }, 5 * 60 * 1000);

  // Auto-post sisa ticket ke group setiap 1 jam
  setInterval(() => {
    autoPostSisaTicket();
  }, 60 * 60 * 1000);
}, 10000);
