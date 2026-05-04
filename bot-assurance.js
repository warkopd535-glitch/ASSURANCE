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

// === HELPER: Get workzone mappings from MASTER sheet (columns X=23, Y=24) ===
async function getWorkzoneMappings() {
  const data = await getSheetData(MASTER_SHEET);
  if (!data || data.length < 3) return []; // header row 2, data from row 3

  const MAPPING_TEAM_COL = 23; // X
  const WORKZONE_COL = 24; // Y

  const mappings = [];
  for (let i = 2; i < data.length; i++) {
    const team = (data[i][MAPPING_TEAM_COL] || '').trim();
    const wz = (data[i][WORKZONE_COL] || '').trim();
    if (team && wz) mappings.push({ team, workzone: wz });
  }

  console.log(`📍 Found ${mappings.length} workzone mappings from MASTER`);

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
  // Fallback: check ALL mapping
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
  // Handle dd/mm/yyyy format (e.g. "03/04/2026")
  const slashMatch = dateStr.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (slashMatch) {
    return { day: parseInt(slashMatch[1]), month: parseInt(slashMatch[2]), year: parseInt(slashMatch[3]) };
  }
  // Handle Indonesian text format (e.g. "Senin, 1 Mei 2026")
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

// === HELPER: Get team config from MASTER (username → sektor) ===
function getTeamConfig(masterData) {
  const config = new Map();
  if (!masterData || masterData.length < 3) return config;
  for (let i = 2; i < masterData.length; i++) {
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
    const tanggal = (masterData[i][19] || '').trim();
    const sektor = (masterData[i][20] || '').trim().toUpperCase();
    const teknisi = (masterData[i][21] || '').trim();
    if (tanggal && teknisi) schedule.push({ tanggal, sektor, teknisi });
  }
  return schedule;
}

// === HELPER: Check if today is weekday (Mon-Fri) in Jakarta ===
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

// === HELPER: Build progress bar ===
function buildProgressBar(completed, total, width = 10) {
  if (total === 0) return '░'.repeat(width) + ' 0%';
  const pct = Math.round((completed / total) * 100);
  const filled = Math.round((completed / total) * width);
  return '█'.repeat(filled) + '░'.repeat(width - filled) + ` ${pct}%`;
}

// === HELPER: Resolve team name from workzone using MASTER mappings ===
function resolveTeamName(workzone, mappings, teamConfig) {
  const team = findMappingTeam(workzone, mappings);
  if (team) {
    const usernames = team.split(/\s*&\s*/).map(u => u.replace('@', '').trim().toLowerCase());
    for (const u of usernames) {
      const sektor = teamConfig.get(u);
      if (sektor && sektor !== 'ALL' && sektor !== 'LTM') return getTeamDisplayName(sektor);
    }
  }
  return workzone ? `TEAM ${workzone.toUpperCase()}` : 'UNKNOWN';
}

// === HELPER: Resolve team name from TEKNISI username directly ===
function resolveTeamByTeknisi(teknisiStr, teamConfig) {
  if (!teknisiStr) return 'UNKNOWN';
  
  // Hard-coded team mapping for reliable resolution
  const TEAM_MAP = {
    'fh_demna': 'SGI', 'fh_rizqi': 'SGI',
    'fh_rayyan': 'BNN', 'fh_mhdfauzan': 'BNN',
    'fh_aqil': 'MRU', 'fh_rijal': 'MRU',
    'paman_sgi': 'MTC', 'fh_imam': 'MTC',
  };
  
  const cleaned = teknisiStr.replace(/@/g, '').toLowerCase();
  
  // Check hard-coded map first (most reliable)
  for (const [key, team] of Object.entries(TEAM_MAP)) {
    if (cleaned.includes(key)) return getTeamDisplayName(team);
  }
  
  // Fallback to teamConfig from MASTER sheet
  const usernames = teknisiStr.split(/[&,]/).map(u => u.replace('@', '').trim().toLowerCase());
  for (const u of usernames) {
    const sektor = teamConfig.get(u);
    if (sektor && sektor !== 'ALL' && sektor !== 'LTM') return getTeamDisplayName(sektor);
  }
  for (const u of usernames) {
    if (!u) continue;
    for (const [key, sektor] of teamConfig.entries()) {
      if (sektor && sektor !== 'ALL' && sektor !== 'LTM') {
        if (u.includes(key) || key.includes(u)) return getTeamDisplayName(sektor);
      }
    }
  }
  return 'UNKNOWN';
}

// === HELPER: Parse piket date (dd.mm.yyyy) ===
function parsePiketDate(dateStr) {
  if (!dateStr) return null;
  const parts = dateStr.split('.');
  if (parts.length < 3) return null;
  return { day: parseInt(parts[0]), month: parseInt(parts[1]), year: parseInt(parts[2]) };
}

// === HELPER: Get today's piket teknisi ===// === HELPER: Get today's piket teknisi ===
function getTodayPiket(piketSchedule) {
  const today = getTodayJakarta();
  return piketSchedule.filter(p => {
    const d = parsePiketDate(p.tanggal);
    return d && d.day === today.day && d.month === today.month && d.year === today.year;
  });
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

  // 1. REGULER — closed from PROGRES ASSURANCE, open from ORDER ASSURANCE
  const closedIncidents = new Set();
  if (progresData) {
    for (let i = 1; i < progresData.length; i++) {
      const tanggal = (progresData[i][0] || '').trim();
      const incident = (progresData[i][1] || '').trim().toUpperCase();
      const d = parseIndonesianDate(tanggal);
      if (d && dateFilterFn(d) && incident) {
        closedIncidents.add(incident);
      }
    }
  }

  const orderCols = orderData ? getOrderColumns(orderData) : null;
  if (orderData && orderCols) {
    for (let i = 1; i < orderData.length; i++) {
      const incident = (orderData[i][orderCols.incident] || '').trim();
      const teknisi = (orderData[i][orderCols.teknisi] || '').trim();
      const status = (orderData[i][orderCols.status] || '').toUpperCase().trim();
      const custType = (orderData[i][orderCols.customerType] || '').trim();
      const ttrStr = (orderData[i][orderCols.ttrCustomer] || '').trim();
      if (!incident) continue;

      const teamName = resolveTeamByTeknisi(teknisi, teamConfig);
      ensureTeam(teamName);

      if (closedIncidents.has(incident.toUpperCase())) {
        result[teamName].reguler.closed.push({ incident, custType, ttr: ttrStr });
      } else if (status === 'OPEN') {
        result[teamName].reguler.open.push({ incident, custType, ttr: ttrStr });
      }
    }
  }

  // 2. UNSPEC — H=7:SERVICE NO, I=8:DEVICE NAME, J=9:STATUS
  if (unspecData) {
    for (let i = 1; i < unspecData.length; i++) {
      const tanggal = (unspecData[i][0] || '').trim();
      const teknisi = (unspecData[i][2] || '').trim();
      const serviceNo = (unspecData[i][7] || '').trim();
      const deviceName = (unspecData[i][8] || '').trim();
      const status = (unspecData[i][9] || '').toUpperCase().trim();
      const d = parseIndonesianDate(tanggal);

      const teamName = resolveTeamByTeknisi(teknisi, teamConfig);
      ensureTeam(teamName);

      if (status === 'CLOSE' && d && dateFilterFn(d)) {
        result[teamName].unspec.closed.push({ serviceNo, deviceName });
      } else if (status === 'OPEN') {
        result[teamName].unspec.open.push({ serviceNo, deviceName });
      }
    }
  }

  // 3. SQM — B=1:INCIDENT, E=4:WORKZONE, J=9:STATUS
  if (sqmData) {
    for (let i = 1; i < sqmData.length; i++) {
      const tanggal = (sqmData[i][0] || '').trim();
      const incident = (sqmData[i][1] || '').trim();
      const teknisi = (sqmData[i][2] || '').trim();
      const status = (sqmData[i][9] || '').toUpperCase().trim();
      const d = parseIndonesianDate(tanggal);

      const teamName = resolveTeamByTeknisi(teknisi, teamConfig);
      ensureTeam(teamName);

      if (status === 'CLOSE' && d && dateFilterFn(d)) {
        result[teamName].sqm.closed.push({ incident });
      } else if (status === 'OPEN') {
        result[teamName].sqm.open.push({ incident });
      }
    }
  }

  // 4. MANUAL GGN — D=3:WORKZONE, always CLOSE
  if (manualData) {
    for (let i = 1; i < manualData.length; i++) {
      const tanggal = (manualData[i][0] || '').trim();
      const teknisi = (manualData[i][1] || '').trim();
      const serviceNo = (manualData[i][4] || '').trim();
      const d = parseIndonesianDate(tanggal);
      if (!d || !dateFilterFn(d)) continue;

      const teamName = resolveTeamByTeknisi(teknisi, teamConfig);
      ensureTeam(teamName);
      result[teamName].manual.closed.push({ serviceNo, tanggal });
    }
  }

  // 5. GAUL ASSURANCE — A=0:TANGGAL, B=1:INCIDENT, D=3:WORKZONE
  if (gaulData) {
    for (let i = 1; i < gaulData.length; i++) {
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

  for (let i = 1; i < data.length; i++) {
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
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // Skip non-command messages after tiket baru check
    // Exception: INC pickup response (dari /TICKET_SQM)
    const hasPendingPickup = global.pendingPickup && global.pendingPickup[chatId + '_' + username] && /INC\d+/i.test(text);
    if (!text.startsWith('/') && !hasPendingPickup) return;
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
        const authResult = await checkAuthorization(username, ['ADMIN']);
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
        const authResult = await checkAuthorization(username, ['ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const data = await withTimeout(getSheetData(ORDER_ASSURANCE_SHEET, false), 10000);
        if (!data || data.length < 2) return sendTelegram(chatId, '<i>Tidak ada data</i>', { reply_to_message_id: msgId });

        const cols = getOrderColumns(data);
        const mappings = await getWorkzoneMappings();
        let expiredList = [];
        let warningList = [];
        let safeList = [];

        for (let i = 1; i < data.length; i++) {
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
        const authResult = await checkAuthorization(username, ['ADMIN']);
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
    // /rekap_hari - Rekap close per teknisi HARI INI
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
        const now = new Date();
        const todayStr = now.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta' });
        let grandClosed = 0, grandOpen = 0, grandGaul = 0;
        const teamStats = [];
        for (const [name, t] of Object.entries(teams)) {
          if (name === 'UNKNOWN') continue;
          const closed = t.reguler.closed.length + t.unspec.closed.length + t.sqm.closed.length + t.manual.closed.length;
          const open = t.reguler.open.length + t.unspec.open.length + t.sqm.open.length;
          grandClosed += closed; grandOpen += open; grandGaul += t.gaul.closed.length;
          teamStats.push({ name, closed, open, t });
        }
        teamStats.sort((a, b) => b.closed - a.closed);
        const medal = ['🥇', '🥈', '🥉'];
        let response = `╔══════════════════════════════════╗\n`;
        response += `║  📊 <b>REKAP CLOSE - HARI INI</b>\n`;
        response += `║  📅 ${todayStr}\n`;
        response += `╠══════════════════════════════════╣\n\n`;
        if (teamStats.length === 0) {
          response += '<i>Belum ada data hari ini</i>\n';
        } else {
          teamStats.forEach((ts, i) => {
            const icon = i < 3 ? medal[i] : '🔸';
            response += `${icon} <b>${ts.name}</b>: ${ts.closed} closed\n`;
            response += `   📋 REG:${ts.t.reguler.closed.length} ❓ UNS:${ts.t.unspec.closed.length} 📍 SQM:${ts.t.sqm.closed.length} 🛠 MAN:${ts.t.manual.closed.length}\n`;
            if (ts.t.gaul.closed.length > 0) response += `   🔄 GAUL: ${ts.t.gaul.closed.length} detected\n`;
            response += '\n';
          });
          response += `📋 <b>Total: ${grandClosed} closed | ${grandOpen} open</b>\n`;
          if (grandGaul > 0) response += `🔄 GAUL: ${grandGaul} detected\n`;
        }
        response += `╚══════════════════════════════════╝`;
        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /rekap_hari Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /rekap_bulan - Rekap close per teknisi BULAN INI
    // ============================================================
    else if (/^\/rekap_bulan\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const data = await withTimeout(getSheetData(ASSURANCE_SHEET), 10000);
        const orderData = await withTimeout(getSheetData(ORDER_ASSURANCE_SHEET), 10000);
        const today = getTodayJakarta();
        let map = {};
        let incidents = [];

        for (let i = 1; i < data.length; i++) {
          const tanggal = (data[i][0] || '').trim();
          const teknisi = (data[i][14] || '-').trim();
          const incident = (data[i][1] || '').trim();
          const d = parseIndonesianDate(tanggal);
          if (!d) continue;
          if (d.month === today.month && d.year === today.year) {
            map[teknisi] = (map[teknisi] || 0) + 1;
            if (incident) incidents.push(incident.toUpperCase());
          }
        }

        // Hitung COMPLY/NOT COMPLY
        let comply = 0, notComply = 0;
        for (let i = 1; i < orderData.length; i++) {
          const inc = (orderData[i][1] || '').trim().toUpperCase();
          const kawal = (orderData[i][17] || '').trim().toUpperCase();
          if (incidents.includes(inc)) {
            if (kawal === 'COMPLY') comply++;
            else if (kawal === 'NOT COMPLY') notComply++;
          }
        }

        const now = new Date();
        const bulanStr = now.toLocaleDateString('id-ID', { month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta' });
        const entries = Object.entries(map).sort((a, b) => b[1] - a[1]);
        const total = entries.reduce((sum, [_, c]) => sum + c, 0);

        const medal = ['🥇', '🥈', '🥉'];
        let response = `━━━━━━━━━━━━━━━━━━━━━━\n📊 <b>REKAP CLOSE - BULAN INI</b>\n📅 ${bulanStr}\n━━━━━━━━━━━━━━━━━━━━━━\n\n`;
        if (entries.length === 0) {
          response += '<i>Belum ada data bulan ini</i>';
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
        console.error('❌ /rekap_bulan Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /rekap_tahun - Rekap close per teknisi TAHUN INI
    // ============================================================
    else if (/^\/rekap_tahun\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const data = await withTimeout(getSheetData(ASSURANCE_SHEET), 10000);
        const orderData = await withTimeout(getSheetData(ORDER_ASSURANCE_SHEET), 10000);
        const today = getTodayJakarta();

        // Group by month, then by teknisi
        const bulanNames = ['', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
        let monthData = {}; // { month: { teknisi: count } }
        let allIncidents = [];
        let grandTotal = 0;

        for (let i = 1; i < data.length; i++) {
          const tanggal = (data[i][0] || '').trim();
          const teknisi = (data[i][14] || '-').trim();
          const incident = (data[i][1] || '').trim();
          const d = parseIndonesianDate(tanggal);
          if (!d) continue;
          if (d.year === today.year) {
            if (!monthData[d.month]) monthData[d.month] = {};
            monthData[d.month][teknisi] = (monthData[d.month][teknisi] || 0) + 1;
            grandTotal++;
            if (incident) allIncidents.push(incident.toUpperCase());
          }
        }

        // Hitung COMPLY/NOT COMPLY total
        let comply = 0, notComply = 0;
        for (let i = 1; i < orderData.length; i++) {
          const inc = (orderData[i][1] || '').trim().toUpperCase();
          const kawal = (orderData[i][17] || '').trim().toUpperCase();
          if (allIncidents.includes(inc)) {
            if (kawal === 'COMPLY') comply++;
            else if (kawal === 'NOT COMPLY') notComply++;
          }
        }

        const medal = ['🥇', '🥈', '🥉'];
        let response = `━━━━━━━━━━━━━━━━━━━━━━\n📊 <b>REKAP CLOSE - TAHUN ${today.year}</b>\n━━━━━━━━━━━━━━━━━━━━━━\n\n`;

        const sortedMonths = Object.keys(monthData).map(Number).sort((a, b) => a - b);
        if (sortedMonths.length === 0) {
          response += '<i>Belum ada data tahun ini</i>';
        } else {
          sortedMonths.forEach(month => {
            const entries = Object.entries(monthData[month]).sort((a, b) => b[1] - a[1]);
            const monthTotal = entries.reduce((sum, [_, c]) => sum + c, 0);
            response += `📅 <b>${bulanNames[month]} ${today.year}</b> [${monthTotal} tickets]\n`;
            entries.forEach(([tek, c], i) => {
              const icon = i < 3 ? medal[i] : '🔸';
              response += `  ${icon} ${tek} : ${c}\n`;
            });
            response += '\n';
          });
          response += `📋 <b>Grand Total: ${grandTotal} tickets</b>\n`;
          response += `✅ COMPLY: ${comply} | ❌ NOT COMPLY: ${notComply}`;
        }

        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
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
        const authResult = await checkAuthorization(username, ['ADMIN']);
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
        let header = isAdmin ? 'DETAIL LAPORAN GANGGUAN MANUAL' : `LAPORAN GANGGUAN MANUAL - @${username}`;
        let response = `📋 <b>${header}</b>\n📅 Tahun ${today.year}\n\n`;

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
    // /TICKET_SQM - Lihat & Pick Up tiket SQM
    // ============================================================
    else if (/^\/TICKET_SQM\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const isAdmin = authResult.role === 'ADMIN';
        const data = await withTimeout(getSheetData(SQM_SHEET, false), 10000);
        if (!data || data.length < 2) return sendTelegram(chatId, '<i>Tidak ada data SQM</i>', { reply_to_message_id: msgId });

        // === ADMIN: Tampilkan total per workzone ===
        if (isAdmin) {
          let wzMap = {};
          let totalOpen = 0;
          for (let i = 1; i < data.length; i++) {
            const status = (data[i][9] || '').trim().toUpperCase(); // J = Status
            const wz = (data[i][4] || '').trim().toUpperCase(); // E = Workzone
            if (status === 'OPEN' && wz) {
              wzMap[wz] = (wzMap[wz] || 0) + 1;
              totalOpen++;
            }
          }

          let response = `📋 <b>TICKET SQM - OPEN PER WORKZONE</b>\n\n`;
          if (totalOpen === 0) {
            response += '<i>Tidak ada tiket OPEN</i>';
          } else {
            const entries = Object.entries(wzMap).sort((a, b) => b[1] - a[1]);
            entries.forEach(([wz, count]) => {
              response += `📍 <b>${wz}</b> : ${count} tiket\n`;
            });
            response += `\n📋 <b>Total: ${totalOpen} tiket OPEN</b>`;
          }

          return sendTelegram(chatId, response, { reply_to_message_id: msgId });
        }

        // === USER: Tampilkan tiket OPEN milik user ===
        const userTickets = [];
        for (let i = 1; i < data.length; i++) {
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
        if (!inputText) return sendTelegram(chatId, `❌ Format tidak sesuai. Gunakan format:\n\n/SQM INC46230392\nCLOSE: deskripsi perbaikan\nDROPCORE: 0\nPATCHCORD: 0\nSOC: 0\nPSLAVE: 0\nPASSIVE 1/8: 0\nPASSIVE 1/4: 0\nPIGTAIL: 0\nADAPTOR: 0\nROSET: 0\nRJ 45: 0\nLAN: 0`, { reply_to_message_id: msgId });

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
        await withTimeout(appendSheetData(ASSURANCE_SHEET, row), 10000);
        cache.assuranceData = null;

        // Auto-close di SQM SA SIGLI (atau buat baris baru jika belum ada)
        let sqmClosed = false;
        try {
          const sqmData = await getSheetData(SQM_SHEET, false);
          let foundInSqm = false;
          for (let i = 1; i < sqmData.length; i++) {
            const incInSqm = (sqmData[i][1] || '').trim().toUpperCase(); // B = Incident
            if (incInSqm === parsed.incidentNo) {
              await updateSheetCell(SQM_SHEET, `J${i + 1}`, 'CLOSE'); // J = Status
              sqmClosed = true;
              foundInSqm = true;
              console.log(`✅ Auto-close SQM: ${parsed.incidentNo} row ${i + 1}`);
              break;
            }
          }

          // Jika INCIDENT belum ada di SQM → buat baris baru
          if (!foundInSqm) {
            const tanggal = new Date().toLocaleDateString('id-ID', {
              weekday: 'long', day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta',
            });
            const teknisi = parsed.teknisi || `@${username}`;
            // A=Tanggal, B=Incident, C=Teknisi, D-I=(skip), J=Status=CLOSE
            const newRow = [tanggal, parsed.incidentNo, teknisi, '', '', '', '', '', '', 'CLOSE'];
            await withTimeout(appendSheetData(SQM_SHEET, newRow), 10000);
            sqmClosed = true;
            console.log(`✅ SQM new row created + CLOSE: ${parsed.incidentNo}`);
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
    // Handle Pick Up response untuk /TICKET_SQM (support multi-incident)
    // ============================================================
    else if (/INC\d+/i.test(text) && global.pendingPickup && global.pendingPickup[chatId + '_' + username]) {
      try {
        const pending = global.pendingPickup[chatId + '_' + username];
        // Expired after 5 minutes
        if (Date.now() - pending.timestamp > 5 * 60 * 1000) {
          delete global.pendingPickup[chatId + '_' + username];
          return sendTelegram(chatId, '❌ Sesi Pick Up sudah expired. Kirim /TICKET_SQM lagi.', { reply_to_message_id: msgId });
        }

        // Extract semua INC dari text (support: "INC123 INC456" atau per baris)
        const incMatches = text.toUpperCase().match(/INC\d+/g) || [];
        if (incMatches.length === 0) {
          return sendTelegram(chatId, '❌ Format INC tidak valid.', { reply_to_message_id: msgId });
        }

        const pickedUp = [];
        const notFound = [];
        const alreadyPicked = [];

        // Cek kolom Q (PROGRES) langsung dari sheet untuk status PICK UP
        const sqmData = await withTimeout(getSheetData(SQM_SHEET, false), 10000);

        for (const incNo of incMatches) {
          const ticket = pending.tickets.find(t => t.incident.toUpperCase() === incNo);
          if (ticket) {
            // Cek apakah sudah di-pickup orang lain
            const sqmRow = sqmData[ticket.row - 1]; // row is 1-based, array is 0-based
            const progres = (sqmRow && sqmRow[16] || '').trim().toUpperCase(); // Q = index 16
            if (progres === 'PICK UP') {
              const pickedByTeknisi = (sqmRow[2] || '').trim(); // C = Teknisi
              alreadyPicked.push({ incident: incNo, pickedBy: pickedByTeknisi });
            } else {
              await updateSheetCell(SQM_SHEET, `Q${ticket.row}`, 'PICK UP');
              pickedUp.push(ticket);
            }
          } else {
            notFound.push(incNo);
          }
        }

        delete global.pendingPickup[chatId + '_' + username];

        let confirmMsg = '';
        if (pickedUp.length > 0) {
          confirmMsg += `✅ ${pickedUp.length} Ticket berhasil di Pick Up!\n\n`;
          pickedUp.forEach(t => {
            confirmMsg += `🔹 ${t.incident} | ${t.serviceNo} | ${t.deviceName}\n`;
          });
        }
        if (alreadyPicked.length > 0) {
          confirmMsg += `\n⚠️ Sudah di Pick Up oleh user lain:\n`;
          alreadyPicked.forEach(t => {
            confirmMsg += `🔸 ${t.incident} → ${t.pickedBy}\n`;
          });
        }
        if (notFound.length > 0) {
          confirmMsg += `\n❌ Tidak ditemukan: ${notFound.join(', ')}`;
        }

        return sendTelegram(chatId, confirmMsg, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ Pick Up Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }
    // ============================================================
    // /unspec - Close UNSPEC ticket (mirip /MANUAL)
    // ============================================================
    else if (/^\/unspec\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const inputText = text.replace(/^\/unspec\s*/i, '').trim();
        if (!inputText) {
          return sendTelegram(chatId, `❌ Format tidak sesuai. Gunakan format:\n\n/unspec\nCLOSE: deskripsi perbaikan\nSERVICE NO: 111106102101\nWORKZONE: BNN`, { reply_to_message_id: msgId });
        }

        const closeMatch = inputText.match(/CLOSE\s*:\s*(.+)/i);
        const svcMatch = inputText.match(/SERVICE\s*NO\s*:\s*(.+)/i);
        const wzMatch = inputText.match(/WORKZONE\s*:\s*(.+)/i);

        if (!closeMatch || !svcMatch || !wzMatch) {
          return sendTelegram(chatId, `❌ Format tidak sesuai. Gunakan format:\n\n/unspec\nCLOSE: deskripsi perbaikan\nSERVICE NO: 111106102101\nWORKZONE: BNN`, { reply_to_message_id: msgId });
        }

        const closeDesc = closeMatch[1].trim();
        const serviceNo = svcMatch[1].trim();
        const wz = wzMatch[1].trim().toUpperCase();

        // Get teknisi dari mapping
        const mappings = await getWorkzoneMappings();
        const teknisi = findMappingTeam(wz, mappings) || `@${username}`;

        const tanggal = new Date().toLocaleDateString('id-ID', {
          weekday: 'long', day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta',
        });

        // Auto-close di UNSPEC sheet (find by SERVICE NO)
        let unspecClosed = false;
        try {
          const unspecData = await getSheetData(UNSPEC_SHEET, false);
          for (let i = 1; i < unspecData.length; i++) {
            const existingSN = (unspecData[i][7] || '').trim(); // H = Service No
            const existingStatus = (unspecData[i][9] || '').toUpperCase().trim(); // J = Status
            if (existingSN === serviceNo && existingStatus === 'OPEN') {
              await updateSheetCell(UNSPEC_SHEET, `J${i + 1}`, 'CLOSE');
              unspecClosed = true;
              console.log(`✅ Auto-close UNSPEC: ${serviceNo} row ${i + 1}`);
              break;
            }
          }
        } catch (closeErr) {
          console.error('⚠️ UNSPEC close error:', closeErr.message);
        }

        // Simpan ke PROGRES ASSURANCE
        const inputTimestamp = new Date().toLocaleString('id-ID', {
          day: '2-digit', month: '2-digit', year: 'numeric',
          hour: '2-digit', minute: '2-digit', second: '2-digit',
          timeZone: 'Asia/Jakarta', hour12: false,
        });
        const row = [
          tanggal, `UNSPEC-${serviceNo}`,
          '', '', '', '', '', '', '', '', '', '', '',
          closeDesc, username.replace('@', ''), inputTimestamp,
        ];
        await withTimeout(appendSheetData(ASSURANCE_SHEET, row), 10000);
        cache.assuranceData = null;

        let confirmMsg = `✅ Data UNSPEC berhasil disimpan!\n\n`;
        confirmMsg += `📋 Detail:\n`;
        confirmMsg += `📅 Tanggal: ${tanggal}\n`;
        confirmMsg += `👷 Teknisi: ${teknisi}\n`;
        confirmMsg += `📍 Workzone: ${wz}\n`;
        confirmMsg += `📞 Service No: ${serviceNo}\n`;
        confirmMsg += `📝 Close: ${closeDesc}\n`;
        if (unspecClosed) confirmMsg += `📊 Status UNSPEC: ✅ Auto-CLOSE`;
        else confirmMsg += `⚠️ Service No tidak ditemukan di sheet UNSPEC`;

        return sendTelegram(chatId, confirmMsg, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /unspec Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /REKAP_UNSPEC - Rekap UNSPEC per bulan
    // ============================================================
    else if (/^\/REKAP_UNSPEC\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const isAdmin = authResult.role === 'ADMIN';
        const data = await withTimeout(getSheetData(UNSPEC_SHEET, false), 10000);
        const today = getTodayJakarta();
        const bulanNames = ['', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

        let monthData = {};
        let grandTotal = 0;

        for (let i = 1; i < data.length; i++) {
          const tanggal = (data[i][0] || '').trim();
          const teknisi = (data[i][2] || '').trim();
          const serviceNo = (data[i][7] || '').trim();
          const status = (data[i][9] || '').trim().toUpperCase();
          if (status !== 'CLOSE') continue;

          const d = parseIndonesianDate(tanggal);
          if (!d || d.year !== today.year) continue;
          if (!isAdmin && !teknisi.toLowerCase().includes(username.toLowerCase())) continue;

          if (!monthData[d.month]) monthData[d.month] = [];
          monthData[d.month].push({ tanggal, serviceNo, teknisi });
          grandTotal++;
        }

        const sortedMonths = Object.keys(monthData).map(Number).sort((a, b) => a - b);
        let header = isAdmin ? 'DETAIL LAPORAN UNSPEC' : `LAPORAN UNSPEC - @${username}`;
        let response = `📋 <b>${header}</b>\n📅 Tahun ${today.year}\n\n`;

        if (sortedMonths.length === 0) {
          response += '<i>Belum ada data</i>';
        } else {
          sortedMonths.forEach(month => {
            const items = monthData[month];
            response += `📅 <b>${bulanNames[month].toUpperCase()}</b> [${items.length} TIKET]\n`;
            items.forEach(item => {
              if (isAdmin) response += `  ${item.tanggal} | ${item.serviceNo} | ${item.teknisi}\n`;
              else response += `  ${item.tanggal} | ${item.serviceNo}\n`;
            });
            response += '\n';
          });
          response += `📋 <b>Grand Total: ${grandTotal} tiket</b>`;
        }

        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /REKAP_UNSPEC Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /produktivitas_hari - Team Daily Productivity (visual)
    // ============================================================
    else if (/^\/produktivitas_hari\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const today = getTodayJakarta();
        const dateFilter = d => d.day === today.day && d.month === today.month && d.year === today.year;
        const prodData = await withTimeout(getProductivityData(dateFilter), 15000);

        const now = new Date();
        const dayNameID = ['MINGGU', 'SENIN', 'SELASA', 'RABU', 'KAMIS', 'JUMAT', 'SABTU'];

        const dateStr = now.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta' });
        const timeStr = now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', timeZone: 'Asia/Jakarta' });

        const numEmoji = ['1️⃣', '2️⃣', '3️⃣', '4️⃣', '5️⃣', '6️⃣', '7️⃣', '8️⃣', '9️⃣', '🔟'];

        let response = `╔══════════════════════════════════╗\n`;
        response += `║  📊 <b>PRODUKTIVITAS TEAM - HARI INI</b>\n`;
        response += `║  📅 ${dateStr} | 🕐 ${timeStr}\n`;

        if (!prodData.weekday) {
          const piketList = prodData.piket;
          if (piketList.length > 0) {
            response += `║  📋 Mode: <b>PIKET</b> (Weekend)\n`;
          }
        }
        response += `╠══════════════════════════════════╣\n\n`;

        const teams = prodData.teams;
        const sortedTeams = Object.keys(teams).sort((a, b) => {
          const totalA = teams[a].reguler.closed.length + teams[a].unspec.closed.length + teams[a].sqm.closed.length + teams[a].manual.closed.length + (teams[a].gaul ? teams[a].gaul.closed.length : 0);
          const totalB = teams[b].reguler.closed.length + teams[b].unspec.closed.length + teams[b].sqm.closed.length + teams[b].manual.closed.length + (teams[b].gaul ? teams[b].gaul.closed.length : 0);
          return totalB - totalA;
        });

        if (sortedTeams.length === 0) {
          response += '<i>Belum ada data hari ini</i>\n';
        } else {
          sortedTeams.forEach((teamName, idx) => {
            const t = teams[teamName];
            const closed = t.reguler.closed.length + t.unspec.closed.length + t.sqm.closed.length + t.manual.closed.length;
            const open = t.reguler.open.length + t.unspec.open.length + t.sqm.open.length;
            const total = closed + open;

            const num = idx < 10 ? numEmoji[idx] : `${idx + 1}.`;
            response += `${num} <b>${teamName}</b>\n`;
            response += `├─ 📈 Total: ${total} | Selesai: ${closed} | Open: ${open}\n`;
            response += `├─ 📋 REGULER: ${t.reguler.closed.length} closed, ${t.reguler.open.length} open\n`;
            response += `├─ ❓ UNSPEC: ${t.unspec.closed.length} closed, ${t.unspec.open.length} open\n`;
            response += `├─ 📍 SQM: ${t.sqm.closed.length} closed, ${t.sqm.open.length} open\n`;
            response += `├─ 🛠 MANUAL GGN: ${t.manual.closed.length} closed\n`;
            response += `└─ 🔄 GAUL: ${t.gaul.closed.length} detected\n\n`;
          });
        }

        response += `╚══════════════════════════════════╝`;
        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /produktivitas_hari Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /detail_team [name] - Detailed Team Report
    // ============================================================
    else if (/^\/detail_team\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const teamArg = text.replace(/^\/detail_team\s*/i, '').trim().toUpperCase();
        if (!teamArg) {
          return sendTelegram(chatId, `❌ Gunakan format: /detail_team SGI\n\nTeam tersedia: SGI, BNN, MRU, MTC`, { reply_to_message_id: msgId });
        }

        const targetTeam = `TEAM ${teamArg}`;
        const today = getTodayJakarta();
        const dateFilter = d => d.day === today.day && d.month === today.month && d.year === today.year;
        const prodData = await withTimeout(getProductivityData(dateFilter), 15000);

        const t = prodData.teams[targetTeam];
        if (!t) {
          const available = Object.keys(prodData.teams).join(', ') || 'Belum ada data';
          return sendTelegram(chatId, `❌ Team "${targetTeam}" tidak ditemukan.\n\nTeam tersedia: ${available}`, { reply_to_message_id: msgId });
        }

        const closed = t.reguler.closed.length + t.unspec.closed.length + t.sqm.closed.length + t.manual.closed.length;
        const open = t.reguler.open.length + t.unspec.open.length + t.sqm.open.length;
        const total = closed + open;

        const dateStr = new Date().toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta' });

        let response = `╔══════════════════════════════════════╗\n`;
        response += `║  👷 <b>DETAIL ${targetTeam} - HARI INI</b>\n`;
        response += `║  📅 ${dateStr}\n`;
        response += `╠══════════════════════════════════════╣\n\n`;

        response += `📊 <b>SUMMARY</b>\n`;
        response += `├─ Total: ${total} | Completed: ${closed} (${total > 0 ? Math.round(closed / total * 100) : 0}%) | Open: ${open}\n`;
        response += `└─ ${buildProgressBar(closed, total, 14)}\n\n`;

        // REGULER
        response += `📋 <b>REGULER</b> (${t.reguler.closed.length} selesai, ${t.reguler.open.length} open)\n`;
        if (t.reguler.closed.length > 0) {
          response += `├─ CLOSED:\n`;
          t.reguler.closed.forEach(tk => { response += `│  • ${tk.incident} ✅\n`; });
        }
        if (t.reguler.open.length > 0) {
          response += `└─ OPEN:\n`;
          t.reguler.open.forEach(tk => { response += `   • ${tk.incident} | ${tk.ttr || '-'} | ${tk.custType || '-'}\n`; });
        }
        if (t.reguler.closed.length === 0 && t.reguler.open.length === 0) response += `└─ <i>Tidak ada data</i>\n`;
        response += '\n';

        // UNSPEC
        response += `❓ <b>UNSPEC</b> (${t.unspec.closed.length} selesai, ${t.unspec.open.length} open)\n`;
        if (t.unspec.closed.length > 0) {
          response += `├─ CLOSED:\n`;
          t.unspec.closed.forEach(tk => { response += `│  • ${tk.serviceNo || tk.deviceName} ✅\n`; });
        }
        if (t.unspec.open.length > 0) {
          response += `└─ OPEN:\n`;
          t.unspec.open.forEach(tk => { response += `   • ${tk.deviceName || tk.serviceNo} | OPEN\n`; });
        }
        if (t.unspec.closed.length === 0 && t.unspec.open.length === 0) response += `└─ <i>Tidak ada data</i>\n`;
        response += '\n';

        // SQM
        response += `📍 <b>SQM SA SIGLI</b> (${t.sqm.closed.length} selesai, ${t.sqm.open.length} open)\n`;
        if (t.sqm.closed.length > 0) {
          response += `├─ CLOSED:\n`;
          t.sqm.closed.forEach(tk => { response += `│  • ${tk.incident} ✅\n`; });
        }
        if (t.sqm.open.length > 0) {
          response += `└─ OPEN:\n`;
          t.sqm.open.forEach(tk => { response += `   • ${tk.incident} | OPEN\n`; });
        }
        if (t.sqm.closed.length === 0 && t.sqm.open.length === 0) response += `└─ <i>Tidak ada data</i>\n`;
        response += '\n';

        // MANUAL GGN
        response += `🛠 <b>MANUAL GGN</b> (${t.manual.closed.length} selesai)\n`;
        if (t.manual.closed.length > 0) {
          t.manual.closed.forEach(tk => { response += `  • ${tk.serviceNo} ✅\n`; });
        } else {
          response += `└─ <i>Tidak ada data</i>\n`;
        }

        response += '\n';

        // GAUL
        response += `🔄 <b>GAUL</b> (${t.gaul.closed.length} detected)\n`;
        if (t.gaul.closed.length > 0) {
          t.gaul.closed.forEach(tk => { response += `├─ ⚠️ ${tk.incident}\n`; });
        } else {
          response += `└─ <i>Tidak ada data</i>\n`;
        }

        response += `\n╚══════════════════════════════════════╝`;
        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /detail_team Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /ringkasan_produk - Weekly Summary (7 hari terakhir)
    // ============================================================
    else if (/^\/ringkasan_produk\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const today = getTodayJakarta();
        const todayDate = new Date(today.year, today.month - 1, today.day);
        const weekAgo = new Date(todayDate);
        weekAgo.setDate(weekAgo.getDate() - 6);
        const wDay = weekAgo.getDate(), wMonth = weekAgo.getMonth() + 1, wYear = weekAgo.getFullYear();

        const dateFilter = d => {
          const dDate = new Date(d.year, d.month - 1, d.day);
          return dDate >= weekAgo && dDate <= todayDate;
        };

        const prodData = await withTimeout(getProductivityData(dateFilter), 15000);
        const teams = prodData.teams;

        const dateRange = `${wDay}/${wMonth} - ${today.day}/${today.month}/${today.year}`;

        // Calculate totals per type
        let totalReguler = { closed: 0, total: 0 };
        let totalUnspec = { closed: 0, total: 0 };
        let totalSqm = { closed: 0, total: 0 };
        let totalManual = { closed: 0 };
        let grandClosed = 0, grandTotal = 0;

        const teamStats = [];
        for (const [name, t] of Object.entries(teams)) {
          const rc = t.reguler.closed.length, ro = t.reguler.open.length;
          const uc = t.unspec.closed.length, uo = t.unspec.open.length;
          const sc = t.sqm.closed.length, so = t.sqm.open.length;
          const mc = t.manual.closed.length;
          const closed = rc + uc + sc + mc;
          const total = closed + ro + uo + so;
          teamStats.push({ name, closed, total, rate: total > 0 ? (closed / total * 100) : 0, rc, ro, uc, uo, sc, so, mc });
          totalReguler.closed += rc; totalReguler.total += rc + ro;
          totalUnspec.closed += uc; totalUnspec.total += uc + uo;
          totalSqm.closed += sc; totalSqm.total += sc + so;
          totalManual.closed += mc;
          grandClosed += closed; grandTotal += total;
        }

        teamStats.sort((a, b) => b.rate - a.rate);
        const medal = ['🥇', '🥈', '🥉', '🏅'];

        let response = `╔══════════════════════════════════════════╗\n`;
        response += `║  📊 <b>RINGKASAN PRODUKTIVITAS - 7 HARI</b>\n`;
        response += `║  📅 ${dateRange}\n`;
        response += `╠══════════════════════════════════════════╣\n\n`;

        response += `🏆 <b>TEAM RANKING (BY COMPLETION RATE)</b>\n`;
        teamStats.forEach((ts, idx) => {
          const m = idx < 4 ? medal[idx] : '▸';
          response += `├─ ${m} ${ts.name}: ${Math.round(ts.rate)}% (${ts.closed}/${ts.total})\n`;
        });
        response += '\n';

        response += `📋 <b>REGULER:</b> ${totalReguler.total} total | ${totalReguler.closed} closed (${totalReguler.total > 0 ? Math.round(totalReguler.closed / totalReguler.total * 100) : 0}%)\n`;
        response += `   ${buildProgressBar(totalReguler.closed, totalReguler.total, 12)}\n\n`;

        response += `❓ <b>UNSPEC:</b> ${totalUnspec.total} total | ${totalUnspec.closed} closed (${totalUnspec.total > 0 ? Math.round(totalUnspec.closed / totalUnspec.total * 100) : 0}%)\n`;
        response += `   ${buildProgressBar(totalUnspec.closed, totalUnspec.total, 12)}\n\n`;

        response += `📍 <b>SQM:</b> ${totalSqm.total} total | ${totalSqm.closed} closed (${totalSqm.total > 0 ? Math.round(totalSqm.closed / totalSqm.total * 100) : 0}%)\n`;
        response += `   ${buildProgressBar(totalSqm.closed, totalSqm.total, 12)}\n\n`;

        response += `🛠 <b>MANUAL GGN:</b> ${totalManual.closed} closed\n`;
        let totalGaul = 0;
        for (const [_, t] of Object.entries(teams)) { totalGaul += t.gaul.closed.length; }
        response += `🔄 <b>GAUL:</b> ${totalGaul} detected\n\n`;

        response += `📈 <b>TOTAL: ${grandClosed} closed / ${grandTotal} total (${grandTotal > 0 ? Math.round(grandClosed / grandTotal * 100) : 0}%)</b>\n`;
        response += `   ${buildProgressBar(grandClosed, grandTotal, 12)}\n`;
        response += `\n╚══════════════════════════════════════════╝`;

        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /ringkasan_produk Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /rank_team - Team Ranking
    // ============================================================
    else if (/^\/rank_team\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        // Parse period: /rank_team bulan OR /rank_team (default: hari ini)
        const periodArg = text.replace(/^\/rank_team\s*/i, '').trim().toLowerCase();
        const today = getTodayJakarta();
        let dateFilter, periodLabel;

        if (periodArg === 'bulan') {
          dateFilter = d => d.month === today.month && d.year === today.year;
          const bulanStr = new Date().toLocaleDateString('id-ID', { month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta' });
          periodLabel = bulanStr;
        } else if (periodArg === 'tahun') {
          dateFilter = d => d.year === today.year;
          periodLabel = `Tahun ${today.year}`;
        } else {
          dateFilter = d => d.day === today.day && d.month === today.month && d.year === today.year;
          periodLabel = new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta' });
        }

        const prodData = await withTimeout(getProductivityData(dateFilter), 15000);
        const teams = prodData.teams;

        const teamStats = [];
        for (const [name, t] of Object.entries(teams)) {
          const closed = t.reguler.closed.length + t.unspec.closed.length + t.sqm.closed.length + t.manual.closed.length;
          const open = t.reguler.open.length + t.unspec.open.length + t.sqm.open.length;
          const total = closed + open;
          const gaul = t.gaul ? t.gaul.closed.length : 0;
          teamStats.push({ name, closed, open, total, gaul, rate: total > 0 ? (closed / total * 100) : 0 });
        }
        teamStats.sort((a, b) => b.closed - a.closed);

        const medal = ['🥇', '🥈', '🥉'];

        let response = `╔══════════════════════════════════╗\n`;
        response += `║  🏆 <b>RANKING TEAM</b>\n`;
        response += `║  📅 ${periodLabel}\n`;
        response += `╠══════════════════════════════════╣\n\n`;

        if (teamStats.length === 0) {
          response += '<i>Belum ada data</i>\n';
        } else {
          teamStats.forEach((ts, idx) => {
            const m = idx < 3 ? medal[idx] : `${idx + 1}.`;
            response += `${m} <b>${ts.name}</b>\n`;
            response += `   ${buildProgressBar(ts.closed, ts.total, 12)}\n`;
            response += `   Closed: ${ts.closed} | Open: ${ts.open} | Total: ${ts.total}\n`;
            if (ts.gaul > 0) response += `   🔄 GAUL: ${ts.gaul} detected\n`;
            response += '\n';
          });

          const totalClosed = teamStats.reduce((s, t) => s + t.closed, 0);
          const totalAll = teamStats.reduce((s, t) => s + t.total, 0);
          const totalGaul = teamStats.reduce((s, t) => s + t.gaul, 0);
          response += `📊 <b>Grand Total: ${totalClosed} closed / ${totalAll} total</b>\n`;
          if (totalGaul > 0) response += `🔄 Total GAUL: ${totalGaul} detected`;
        }

        response += `\n╚══════════════════════════════════╝`;
        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /rank_team Error:', err.message);
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

<b>📊 PRODUKTIVITAS (ADMIN):</b>
/produktivitas_hari - Produktivitas team
/detail_team SGI - Detail per team
/ringkasan_produk - Ringkasan 7 hari
/rank_team - Ranking team
/rank_team bulan - Ranking bulan ini

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
