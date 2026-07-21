import { google } from 'googleapis';
import { getAuth } from './googleAuth';

const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const SHEET_NAME = process.env.GOOGLE_SHEET_NAME || 'input';

// Urutan kolom di sheet HARUS sama seperti urutan ini (kolom A - J)
const HEADERS = ['id', 'tglMasuk', 'pilar', 'platform', 'brief', 'status', 'tglSelesai', 'hasilAkhir', 'tglPosting', 'deskripsiBrief'];
const LAST_COL = 'J';

function getSheetsClient() {
  if (!SHEET_ID) {
    throw new Error('GOOGLE_SHEET_ID belum diisi di .env.local');
  }
  return google.sheets({ version: 'v4', auth: getAuth() });
}

async function ensureHeaders(sheets) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!A1:${LAST_COL}1`,
  });
  const row = res.data.values && res.data.values[0];
  if (!row || row.length === 0) {
    // Sheet kosong: tulis semua header dari nol.
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!A1:${LAST_COL}1`,
      valueInputOption: 'RAW',
      requestBody: { values: [HEADERS] },
    });
  } else if (row.length < HEADERS.length) {
    // Sheet lama (migrasi): kolom baru (mis. tglPosting) belum ada di header,
    // tambahkan tanpa mengubah/menghapus kolom & data yang sudah ada.
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!A1:${LAST_COL}1`,
      valueInputOption: 'RAW',
      requestBody: { values: [HEADERS] },
    });
  }
}

function rowToObj(row) {
  const obj = {};
  HEADERS.forEach((h, i) => {
    obj[h] = row[i] !== undefined ? row[i] : '';
  });
  return obj;
}

function objToRow(obj) {
  return HEADERS.map((h) => (obj[h] !== undefined && obj[h] !== null ? obj[h] : ''));
}

// Ambil semua brief dari sheet
export async function getAllBriefs() {
  const sheets = getSheetsClient();
  await ensureHeaders(sheets);
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!A2:${LAST_COL}`,
  });
  const rows = res.data.values || [];
  return rows.map(rowToObj).filter((b) => b.id);
}

// Tambah brief baru (append row di bawah)
export async function addBrief(brief) {
  const sheets = getSheetsClient();
  await ensureHeaders(sheets);
  const id = 'b' + Date.now();
  const newBrief = { ...brief, id };
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!A:${LAST_COL}`,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [objToRow(newBrief)] },
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
  const rowNumber = await findRowNumberById(sheets, id);
  if (rowNumber == null) throw new Error('Brief dengan id "' + id + '" tidak ditemukan di sheet');
  const merged = { ...updates, id };
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!A${rowNumber}:${LAST_COL}${rowNumber}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [objToRow(merged)] },
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
