import { google } from 'googleapis';
import { getAuth } from './googleAuth';

const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const SHEET_NAME = process.env.GOOGLE_SHEET_NAME || 'input';

// Urutan kolom di sheet HARUS sama seperti urutan ini (kolom A - M)
const HEADERS = ['id', 'tglMasuk', 'pilar', 'platform', 'brief', 'status', 'tglSelesai', 'hasilAkhir', 'tglPosting', 'deskripsiBrief', 'isReference', 'referensi', 'tglSetor'];
const LAST_COL = 'M';

function getSheetsClient() {
  if (!SHEET_ID) {
    throw new Error('GOOGLE_SHEET_ID belum diisi di .env.local');
  }
  return google.sheets({ version: 'v4', auth: getAuth() });
}

function canonicalHeader(rawHeader) {
  if (!rawHeader) return '';
  const clean = String(rawHeader).trim();
  const lower = clean.toLowerCase().replace(/[\s_-]+/g, '');
  const match = HEADERS.find((h) => h.toLowerCase().replace(/[\s_-]+/g, '') === lower);
  return match || clean;
}

async function getActualHeaders(sheets) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!1:1`,
  });
  const row = (res.data.values && res.data.values[0]) || [];
  if (row.length === 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!A1:${LAST_COL}1`,
      valueInputOption: 'RAW',
      requestBody: { values: [HEADERS] },
    });
    return [...HEADERS];
  }

  const canonicalRow = row.map((h) => canonicalHeader(h));
  const existingLower = canonicalRow.map((h) => h.toLowerCase());
  const missing = HEADERS.filter((h) => !existingLower.includes(h.toLowerCase()));
  if (missing.length > 0) {
    const updated = [...canonicalRow, ...missing];
    const lastColLetter = String.fromCharCode(65 + updated.length - 1);
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!A1:${lastColLetter}1`,
      valueInputOption: 'RAW',
      requestBody: { values: [updated] },
    });
    return updated;
  }
  return canonicalRow;
}

function rowToObj(row, actualHeaders) {
  const obj = {};
  actualHeaders.forEach((h, i) => {
    if (h) {
      obj[h] = row[i] !== undefined ? row[i] : '';
    }
  });
  return obj;
}

function objToRow(obj, actualHeaders) {
  return actualHeaders.map((h) => {
    if (!h) return '';
    if (obj[h] !== undefined && obj[h] !== null) return obj[h];
    // Fallback: case-insensitive key lookup across obj
    const lowerH = h.toLowerCase().replace(/[\s_-]+/g, '');
    const foundKey = Object.keys(obj).find((k) => k.toLowerCase().replace(/[\s_-]+/g, '') === lowerH);
    return foundKey && obj[foundKey] !== undefined && obj[foundKey] !== null ? obj[foundKey] : '';
  });
}

// Ambil semua brief dari sheet
export async function getAllBriefs() {
  const sheets = getSheetsClient();
  const actualHeaders = await getActualHeaders(sheets);
  const lastColLetter = String.fromCharCode(65 + Math.max(actualHeaders.length, 13) - 1);
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!A2:${lastColLetter}`,
  });
  const rows = res.data.values || [];
  return rows
    .map((r) => rowToObj(r, actualHeaders))
    .filter((b) => b.id)
    .map((b) => ({
      ...b,
      tglSetor: b.tglSetor || b.tglSelesai || '',
    }));
}

// Tambah brief baru (append row di bawah)
export async function addBrief(brief) {
  const sheets = getSheetsClient();
  const actualHeaders = await getActualHeaders(sheets);
  const id = 'b' + Date.now();
  const newBrief = { ...brief, id };
  const lastColLetter = String.fromCharCode(65 + Math.max(actualHeaders.length, 13) - 1);
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!A:${lastColLetter}`,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [objToRow(newBrief, actualHeaders)] },
  });
  return newBrief;
}

async function findRowNumberById(sheets, id) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!A2:A`,
  });
  const rows = res.data.values || [];
  const idx = rows.findIndex((r) => r[0] === id);
  if (idx === -1) return null;
  return idx + 2; // +2: baris 1 = header, index array mulai dari 0
}

// Update satu brief (butuh objek lengkap, bukan partial)
export async function updateBrief(id, updates) {
  const sheets = getSheetsClient();
  const actualHeaders = await getActualHeaders(sheets);
  const rowNumber = await findRowNumberById(sheets, id);
  if (rowNumber == null) throw new Error('Brief dengan id "' + id + '" tidak ditemukan di sheet');
  const merged = { ...updates, id };
  const lastColLetter = String.fromCharCode(65 + Math.max(actualHeaders.length, 13) - 1);
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!A${rowNumber}:${lastColLetter}${rowNumber}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [objToRow(merged, actualHeaders)] },
  });
  return merged;
}

async function getSheetNumericId(sheets) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const sheet = meta.data.sheets.find((s) => s.properties.title === SHEET_NAME);
  if (!sheet) throw new Error(`Tab sheet "${SHEET_NAME}" tidak ditemukan`);
  return sheet.properties.sheetId;
}

// Hapus brief (hapus baris di sheet)
export async function deleteBrief(id) {
  const sheets = getSheetsClient();
  const rowNumber = await findRowNumberById(sheets, id);
  if (rowNumber == null) return false;
  const sheetId = await getSheetNumericId(sheets);
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      requests: [
        {
          deleteDimension: {
            range: {
              sheetId,
              dimension: 'ROWS',
              startIndex: rowNumber - 1,
              endIndex: rowNumber,
            },
          },
        },
      ],
    },
  });
  return true;
}
