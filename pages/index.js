import { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';

const PILARS = ['Ads', 'Feed', 'Story', 'Carousel', 'Video', 'Lainnya'];
const PLATFORMS = ['Instagram', 'Tiktok', 'Non sosmed'];
const STATUS_OPTIONS = ['On Going', 'Waiting Approval', 'Selesai Terupload'];

const EMPTY_FORM = {
  tglMasuk: '',
  pilar: 'Ads',
  platform: [],
  brief: '',
  deskripsiBrief: '',
  status: '',
  tglSelesai: '',
  hasilAkhir: '',
  referensi: '',
  tglPostingByPlatform: {},
  _prefillDate: '',
  isReference: false,
};

// Google Sheets (valueInputOption: USER_ENTERED) otomatis mengubah string "true"/"false"
// jadi boolean, dan saat dibaca lagi lewat API balik sebagai string "TRUE"/"FALSE" (kapital),
// bukan "true" huruf kecil. Makanya perbandingan strict === 'true' bisa gagal dan brief
// referensi ikut kehitung di Total Brief. Helper ini menangani semua variannya.
function isReferenceBrief(b) {
  if (!b) return false;
  const v = b.isReference;
  if (v === true) return true;
  return String(v || '').trim().toLowerCase() === 'true';
}

const PLATFORM_COLORS = {
  Instagram: '#af52de',
  Tiktok: '#0071e3',
  'Non sosmed': '#8e8e93',
};

// Singkatan platform khusus tampilan Kalender Konten, biar ringkas.
const PLATFORM_ABBR = { Instagram: 'IG', Tiktok: 'Tiktok', 'Non sosmed': 'Non-Sosmed' };
function platformAbbr(p) {
  return PLATFORM_ABBR[p] || p;
}

// Urutan tampil platform di Kalender Konten: Instagram selalu di atas, baru Tiktok,
// platform lain menyusul di bawahnya.
const PLATFORM_ORDER = { Instagram: 0, Tiktok: 1 };
function PLATFORM_ORDER_INDEX(platform) {
  return platform in PLATFORM_ORDER ? PLATFORM_ORDER[platform] : 99;
}

// Untuk item Instagram, urutkan lagi berdasarkan format/pilar: Story, Carousel, Feed, Reels.
// (nilai pilar yang tersimpan tetap "Video" karena dipakai bareng sama Tiktok)
const IG_PILAR_ORDER = { Story: 0, Carousel: 1, Feed: 2, Video: 3 };
function IG_PILAR_ORDER_INDEX(pilar) {
  return pilar && pilar in IG_PILAR_ORDER ? IG_PILAR_ORDER[pilar] : 99;
}

// Label pilar khusus tampilan: pilar "Video" ditampilkan sebagai "Reels" kalau platform-nya
// Instagram, tapi tetap "Video" kalau Tiktok (satu nilai data, beda istilah per platform).
function pilarDisplayLabel(platform, pilar) {
  if (platform === 'Instagram' && pilar === 'Video') return 'Reels';
  return pilar;
}

// Warna badge platform/pilar di Kalender Konten — pilar Story berwarna biru (#0071e3).
function formatColor(platform, pilar) {
  if (pilar === 'Story') return '#0071e3';
  if (platform === 'Instagram') return 'linear-gradient(135deg, #833ab4, #e1306c)';
  if (platform === 'Tiktok') return '#000000';
  return '#8e8e93';
}

// platform disimpan di Sheets sebagai string dipisah koma ("Instagram,Tiktok")
// supaya satu brief bisa menyasar lebih dari satu platform sekaligus.
function platformsOf(b) {
  if (!b) return [];
  return (b.platform || '')
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);
}
function platformLabel(b) {
  if (!b) return '-';
  const list = platformsOf(b);
  return list.length ? list.join(' + ') : '-';
}

// Pasangkan tiap platform dengan tanggal postingnya masing-masing.
// tglPosting disimpan sejajar urutan dengan platform (dipisah koma), mis.
// platform="Instagram,Tiktok" & tglPosting="2026-08-01,2026-08-03".
// Data lama yang cuma punya 1 tanggal untuk banyak platform tetap dianggap
// tanggal yang sama untuk semua platform (backward-compatible).
function platformDatePairs(b) {
  if (!b) return [];
  const plats = platformsOf(b);
  const dates = (b.tglPosting || '').split(',').map((s) => s.trim());
  if (plats.length === 0) {
    return [{ platform: '', date: dates[0] || '' }];
  }
  if (dates.length <= 1 && plats.length > 1) {
    return plats.map((p) => ({ platform: p, date: dates[0] || '' }));
  }
  return plats.map((p, i) => ({ platform: p, date: dates[i] || '' }));
}

// Warna kartu di Kalender Konten mengikuti status brief.
const CALENDAR_STATUS_STYLE = {
  'Belum Dikerjakan': { bg: '#ffffff', border: 'var(--hair)' },
  'On Going': { bg: '#fff9db', border: '#ffcc00' },
  'Waiting Approval': { bg: 'rgba(255,59,48,0.10)', border: 'var(--red)' },
  'Selesai Terupload': { bg: 'rgba(52,199,89,0.14)', border: 'var(--green)' },
};

function fmtDate(d) {
  if (!d) return '-';
  const dt = new Date(d + 'T00:00:00');
  if (isNaN(dt)) return '-';
  return dt.toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' });
}
function fmtDateShort(d) {
  if (!d) return '-';
  const dt = new Date(d + 'T00:00:00');
  if (isNaN(dt)) return '-';
  return dt.toLocaleDateString('id-ID', { weekday: 'short', day: '2-digit', month: 'short' });
}
function maxDays(pilar) {
  return pilar === 'Ads' || pilar === 'Carousel' ? 2 : 3;
}
function kpiFor(b) {
  if (!b || !b.tglSelesai || !b.tglMasuk) return null;
  const start = new Date(b.tglMasuk + 'T00:00:00');
  const end = new Date(b.tglSelesai + 'T00:00:00');
  const diffDays = Math.round((end - start) / 86400000);
  return diffDays <= maxDays(b.pilar) ? 'On Time' : 'Late';
}
function statusOf(b) {
  if (!b) return 'Belum Dikerjakan';
  return b.status ? b.status : 'Belum Dikerjakan';
}
function pillClass(status) {
  const map = {
    'Selesai Terupload': 'pill-green',
    'On Going': 'pill-orange',
    'Waiting Approval': 'pill-red',
    'Belum Dikerjakan': 'pill-grey',
  };
  return map[status] || 'pill-grey';
}

// --- Helper minggu custom ---
// Minggu 1 (index 0) = minggu pertama kerja, range-nya tidak genap 7 hari: 29 Jun - 9 Jul.
// Minggu 2 dst (index 1, 2, ...) = blok 7 hari berturut-turut mulai 10 Jul, mengikuti
// jadwal weekly report tiap hari Kamis (Jumat - Kamis).
const WEEK1_START = '2026-07-01';
const WEEK1_END = '2026-07-09';

function addDays(date, n) {
  const d = new Date(date);
  d.setDate(d.getDate() + n);
  return d;
}
function toISO(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}
function parseISODate(s) {
  return new Date(s + 'T00:00:00');
}
// Mengembalikan { start, end } (Date, inclusive) untuk index minggu ke-n (0-based).
function getWeekRange(weekIndex) {
  const week1Start = parseISODate(WEEK1_START);
  const week1End = parseISODate(WEEK1_END);
  if (weekIndex <= 0) {
    return { start: week1Start, end: week1End };
  }
  const week2Start = addDays(week1End, 1); // 10 Jul
  const start = addDays(week2Start, (weekIndex - 1) * 7);
  const end = addDays(start, 6);
  return { start, end };
}
function formatWeekLabel(weekIndex) {
  const { start, end } = getWeekRange(weekIndex);
  const sameMonth = start.getMonth() === end.getMonth();
  const optsShort = { day: '2-digit' };
  const optsFull = { day: '2-digit', month: 'short', year: 'numeric' };
  const startLabel = start.toLocaleDateString('id-ID', sameMonth ? optsShort : optsFull);
  const endLabel = end.toLocaleDateString('id-ID', optsFull);
  return `Minggu ${weekIndex + 1} · ${startLabel} - ${endLabel}`;
}
function isInWeek(isoDateStr, weekIndex) {
  if (!isoDateStr) return false;
  const d = parseISODate(isoDateStr);
  const { start, end } = getWeekRange(weekIndex);
  return d >= start && d <= end;
}
// Menentukan index minggu yang memuat tanggal hari ini.
function currentWeekIndex() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const week1End = parseISODate(WEEK1_END);
  if (today <= week1End) return 0;
  const week2Start = addDays(week1End, 1);
  const diffDays = Math.round((today - week2Start) / 86400000);
  return 1 + Math.floor(diffDays / 7);
}

// --- Helper kalender bulanan (Kalender Konten, berbasis tglPosting) ---
const DAY_LABELS = ['Min', 'Sen', 'Sel', 'Rab', 'Kam', 'Jum', 'Sab'];
function monthLabel(d) {
  return d.toLocaleDateString('id-ID', { month: 'long', year: 'numeric' });
}
// Bangun grid 6x7 (42 sel) untuk bulan dari `monthDate`, dimulai hari Minggu.
function buildMonthGrid(monthDate) {
  const year = monthDate.getFullYear();
  const month = monthDate.getMonth();
  const firstOfMonth = new Date(year, month, 1);
  const startOffset = firstOfMonth.getDay(); // 0 = Minggu
  const gridStart = addDays(firstOfMonth, -startOffset);
  const cells = [];
  for (let i = 0; i < 42; i++) {
    const date = addDays(gridStart, i);
    cells.push({ date, iso: toISO(date), inMonth: date.getMonth() === month });
  }
  return cells;
}
function isSameISODate(a, b) {
  return a === b;
}
function todayISO() {
  return toISO(new Date());
}

// --- Helper Import XLSX/CSV (transisi dari sheet kalender manual lama) ---
const CSV_DAY_HEADERS = ['senin', 'selasa', 'rabu', 'kamis', 'jumat', 'sabtu', 'minggu'];

function isCellRed(cell) {
  if (!cell || !cell.s) return false;
  const s = cell.s;
  const colorsToCheck = [];
  if (s.fgColor) colorsToCheck.push(s.fgColor);
  if (s.bgColor) colorsToCheck.push(s.bgColor);
  if (s.fill) {
    if (s.fill.fgColor) colorsToCheck.push(s.fill.fgColor);
    if (s.fill.bgColor) colorsToCheck.push(s.fill.bgColor);
  }
  for (const c of colorsToCheck) {
    if (!c) continue;
    let rgb = String(c.rgb || c.theme || '').toUpperCase();
    if (rgb.length > 6 && rgb.startsWith('FF')) {
      rgb = rgb.slice(2);
    }
    if (
      [
        'FF0000',
        'E60000',
        'CC0000',
        'FF4D4D',
        'FF3B30',
        'EA4335',
        'D93025',
        'F44336',
        'E53935',
        'D32F2F',
        'C62828',
        'B71C1C',
        'RED',
      ].includes(rgb)
    ) {
      return true;
    }
    if (/^[0-9A-F]{6}$/.test(rgb)) {
      const r = parseInt(rgb.substring(0, 2), 16);
      const g = parseInt(rgb.substring(2, 4), 16);
      const b = parseInt(rgb.substring(4, 6), 16);
      if (r > 180 && g < 100 && b < 100) {
        return true;
      }
    }
  }
  return false;
}

const MONTH_MAP = {
  jan: '01', januari: '01', january: '01',
  feb: '02', februari: '02', february: '02',
  mar: '03', maret: '03', march: '03',
  apr: '04', april: '04',
  may: '05', mei: '05',
  jun: '06', juni: '06', june: '06',
  jul: '07', juli: '07', july: '07',
  aug: '08', agu: '08', agustus: '08', august: '08',
  sep: '09', september: '09', sept: '09',
  oct: '10', okt: '10', oktober: '10', october: '10',
  nov: '11', november: '11',
  dec: '12', des: '12', desember: '12', december: '12',
};

function detectMonthYearFromText(...texts) {
  const combined = texts.filter(Boolean).join(' ').toLowerCase();
  if (!combined) return '';

  let year = new Date().getFullYear();
  const y4Match = combined.match(/\b(202\d)\b/);
  if (y4Match) {
    year = parseInt(y4Match[1], 10);
  } else {
    const y2Match = combined.match(/\b(2[4-9]|3[0-0])\b/);
    if (y2Match) {
      year = 2000 + parseInt(y2Match[1], 10);
    }
  }

  for (const [key, num] of Object.entries(MONTH_MAP)) {
    const regex = new RegExp(`\\b${key}\\b`, 'i');
    if (regex.test(combined)) {
      return `${year}-${num}`;
    }
  }
  return '';
}

function parseSheetToGrid(sheet) {
  if (!sheet || !sheet['!ref']) return [];
  const range = XLSX.utils.decode_range(sheet['!ref']);
  const grid = [];
  for (let R = range.s.r; R <= range.e.r; ++R) {
    const row = [];
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = sheet[cellAddress];
      const val = cell ? String(cell.v !== undefined ? cell.v : cell.w !== undefined ? cell.w : '').trim() : '';
      const red = cell ? isCellRed(cell) : false;
      row.push({ val, isRed: red });
    }
    grid.push(row);
  }
  return grid;
}

function splitCellTitles(raw) {
  const lines = String(raw || '')
    .split(/\r\n|\r|\n/)
    .map((s) => s.trim())
    .filter(Boolean);
  const merged = [];
  lines.forEach((line) => {
    if (line.startsWith('(') && merged.length > 0) {
      merged[merged.length - 1] += ' ' + line;
    } else {
      merged.push(line);
    }
  });
  return merged;
}

function parseWorkbookToAllBlocks(wb) {
  const allBlocks = [];
  (wb.SheetNames || []).forEach((sheetName) => {
    const sheet = wb.Sheets[sheetName];
    const grid = parseSheetToGrid(sheet);
    const blocks = detectContentPlanBlocks(grid, sheetName);
    blocks.forEach((b) => {
      allBlocks.push({ ...b, sheetName });
    });
  });
  return allBlocks;
}

function parseHeaderName(str) {
  const s = String(str || '').trim();
  if (!s) return { platform: '', label: '' };
  const lower = s.toLowerCase();

  if (/^(ig|instagram)\b/i.test(s) || lower.startsWith('ig ') || lower === 'ig' || lower.includes('instagram')) {
    return { platform: 'Instagram', label: s };
  }
  if (/^(tiktok|tok)\b/i.test(s) || lower.startsWith('tok ') || lower === 'tok' || lower.includes('tiktok')) {
    return { platform: 'Tiktok', label: s };
  }
  return { platform: '', label: s };
}

function detectContentPlanBlocks(grid, sheetName = '') {
  const blocks = [];
  let i = 0;
  while (i < grid.length) {
    const row = (grid[i] || []).slice(0, 7).map((c) => String(c.val !== undefined ? c.val : c || '').trim().toLowerCase());
    const isHeader = CSV_DAY_HEADERS.every((d, idx) => row[idx] === d);
    if (isHeader) {
      let platformName = '';
      let headerLabel = '';
      let detectedMonth = detectMonthYearFromText(sheetName);

      for (let r = 0; r < i; r++) {
        const rowCells = grid[r] || [];
        for (let c = 0; c < rowCells.length; c++) {
          const val = String(rowCells[c] && rowCells[c].val !== undefined ? rowCells[c].val : rowCells[c] || '').trim();
          if (!val) continue;
          const parsed = parseHeaderName(val);
          if (parsed.platform && !platformName) {
            platformName = parsed.platform;
            headerLabel = parsed.label;
          }
          if (!detectedMonth) {
            const m = detectMonthYearFromText(val);
            if (m) detectedMonth = m;
          }
        }
      }

      const weeks = [];
      let cursor = i + 1;
      while (cursor + 1 < grid.length) {
        const dateRow = grid[cursor] || [];
        const hasDateLike = dateRow
          .slice(0, 7)
          .some((c) => {
            const txt = String(c.val !== undefined ? c.val : c || '').trim();
            return /\d{1,2}\s*$/.test(txt) && txt !== '';
          });
        if (!hasDateLike) break;
        const titleRow = grid[cursor + 1] || [];
        const days = [];
        for (let col = 0; col < 7; col++) {
          const dateCell = dateRow[col];
          const titleCell = titleRow[col];
          const rawDateCell = String(dateCell && dateCell.val !== undefined ? dateCell.val : dateCell || '').trim();
          const dateLines = splitCellTitles(rawDateCell);
          const lastLine = dateLines[dateLines.length - 1] || '';
          const m = lastLine.match(/(\d{1,2})\s*$/);
          if (!m) {
            days.push(null);
            continue;
          }
          const dayNum = parseInt(m[1], 10);
          const leftoverOnDateLine = lastLine.replace(/(\d{1,2})\s*$/, '').trim();
          const titlesFromDateRow = [...dateLines.slice(0, -1), leftoverOnDateLine]
            .filter(Boolean)
            .map((text) => ({ text, source: 'date' }));
          
          const rawTitleCell = String(titleCell && titleCell.val !== undefined ? titleCell.val : titleCell || '').trim();
          const titlesFromTitleRow = splitCellTitles(rawTitleCell).map((text) => ({ text, source: 'title' }));
          
          const titles = [...titlesFromDateRow, ...titlesFromTitleRow];
          const isRed = (titleCell && titleCell.isRed) || (dateCell && dateCell.isRed);

          days.push({ dayNum, titles, isRed });
        }
        weeks.push(days);
        cursor += 2;
      }
      blocks.push({ headerRow: i, platformName, headerLabel, detectedMonth, weeks });
      i = cursor > i ? cursor : i + 1;
    } else {
      i++;
    }
  }
  return blocks;
}

function blocksToImportRows(blocks, blockMonths, existingBriefs = [], blockPlatforms = []) {
  const existingSet = new Set(existingBriefs.map((b) => (b.brief || '').trim().toLowerCase()));
  const rowMap = new Map();
  let rid = 0;

  blocks.forEach((block, bIdx) => {
    const my = blockMonths[bIdx] || '';
    const platformChoice = blockPlatforms[bIdx] || block.platformName || 'Instagram';
    const defaultPlatform = [platformChoice];
    const isIG = platformChoice === 'Instagram';
    const isTikTok = platformChoice === 'Tiktok';

    block.weeks.forEach((week) => {
      week.forEach((cell) => {
        if (!cell || !cell.titles || cell.titles.length === 0) return;
        let tglPosting = '';
        if (my) {
          const [y, m] = my.split('-');
          tglPosting = `${y}-${m}-${String(cell.dayNum).padStart(2, '0')}`;
        }

        cell.titles.forEach(({ text: title, source }) => {
          if (!title) return;
          const normTitle = title.trim().toLowerCase();
          const key = `${normTitle}_${tglPosting}`;

          let defaultPilar = 'Video';
          if (isIG) {
            defaultPilar = cell.isRed || source === 'date' ? 'Story' : 'Video';
          } else if (isTikTok) {
            defaultPilar = 'Video';
          } else {
            defaultPilar = cell.isRed || source === 'date' ? 'Story' : 'Video';
          }

          if (rowMap.has(key)) {
            const existing = rowMap.get(key);
            defaultPlatform.forEach((p) => {
              if (!existing.platform.includes(p)) {
                existing.platform.push(p);
              }
            });
          } else {
            const alreadyExists = existingSet.has(normTitle);
            rowMap.set(key, {
              id: `imp-${rid++}`,
              blockIdx: bIdx,
              dayNum: cell.dayNum,
              title,
              tglPosting,
              include: !alreadyExists,
              alreadyExists,
              pilar: defaultPilar,
              platform: [...defaultPlatform],
            });
          }
        });
      });
    });
  });

  return Array.from(rowMap.values());
}

function getBlockPreviewTitles(block) {
  const titles = [];
  (block.weeks || []).forEach((week) => {
    (week || []).forEach((day) => {
      if (day && day.titles) {
        day.titles.forEach((t) => {
          if (t.text && titles.length < 2) {
            titles.push(t.text);
          }
        });
      }
    });
  });
  return titles.join(', ');
}

export default function Home() {
  const [briefs, setBriefs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState('');
  const [saving, setSaving] = useState(false);

  const [form, setForm] = useState(EMPTY_FORM);
  const [editingId, setEditingId] = useState(null);
  const [formMsg, setFormMsg] = useState('');
  const [formOpen, setFormOpen] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [uploadedFileName, setUploadedFileName] = useState('');

  // null = tampilkan semua data. Selain itu: Date (Senin) minggu yang dipilih.
  const [weekFilter, setWeekFilter] = useState(null);

  // filter khusus tabel "Daftar Brief"
  const [filterPilar, setFilterPilar] = useState('');
  const [filterPlatform, setFilterPlatform] = useState('');
  const [filterStatus, setFilterStatus] = useState('');

  // tanggal yang sedang di-expand di panel "Brief Selesai per Hari"
  const [expandedDate, setExpandedDate] = useState(null);

  // 'dashboard' = tampilan lama (produksi), 'kalender' = Kalender Konten (tglPosting)
  const [view, setView] = useState('dashboard');
  const [calendarMonth, setCalendarMonth] = useState(() => {
    const d = new Date();
    return new Date(d.getFullYear(), d.getMonth(), 1);
  });

  async function loadBriefs() {
    setLoading(true);
    setLoadError('');
    try {
      const res = await fetch('/api/briefs');
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.error || 'Gagal memuat data dari Google Sheets');
      }
      const data = await res.json();
      setBriefs(data);
    } catch (err) {
      setLoadError(err.message);
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    loadBriefs();
  }, []);

  const filteredBriefs = useMemo(() => {
    const list = briefs.filter((b) => !isReferenceBrief(b));
    if (weekFilter === null) return list;
    return list.filter((b) => isInWeek(b.tglMasuk, weekFilter));
  }, [briefs, weekFilter]);

  // brief selesai, dikelompokkan per tanggal selesai (turun dari yang paling baru)
  const dailyCompleted = useMemo(() => {
    const map = {};
    filteredBriefs.forEach((b) => {
      if (statusOf(b) === 'Selesai Terupload' && b.tglSelesai) {
        if (!map[b.tglSelesai]) map[b.tglSelesai] = [];
        map[b.tglSelesai].push(b);
      }
    });
    return Object.entries(map)
      .sort((a, b) => new Date(b[0]) - new Date(a[0]))
      .map(([date, items]) => ({ date, items }));
  }, [filteredBriefs]);
  const maxDaily = Math.max(1, ...dailyCompleted.map((d) => d.items.length));

  // daftar brief di tabel: filteredBriefs + filter tambahan (pilar/platform/status)
  const tableBriefs = useMemo(() => {
    return filteredBriefs.filter((b) => {
      if (filterPilar && b.pilar !== filterPilar) return false;
      if (filterPlatform && !platformsOf(b).includes(filterPlatform)) return false;
      if (filterStatus) {
        if (filterStatus === 'Belum Dikerjakan') {
          if (statusOf(b) !== 'Belum Dikerjakan') return false;
        } else if (b.status !== filterStatus) {
          return false;
        }
      }
      return true;
    });
  }, [filteredBriefs, filterPilar, filterPlatform, filterStatus]);
  const hasTableFilter = filterPilar || filterPlatform || filterStatus;

  function openAddForm(prefill) {
    setEditingId(null);
    setForm(prefill ? { ...EMPTY_FORM, ...prefill } : EMPTY_FORM);
    setFormMsg('');
    setUploadedFileName('');
    setFormOpen(true);
  }
  function enterEditMode(id) {
    const b = briefs.find((x) => x.id === id);
    if (!b) return;
    setEditingId(id);
    const datesMap = {};
    platformDatePairs(b).forEach((pr) => {
      if (pr.platform) datesMap[pr.platform] = pr.date || '';
    });
    setForm({
      tglMasuk: b.tglMasuk,
      pilar: b.pilar,
      platform: platformsOf(b),
      brief: b.brief,
      deskripsiBrief: b.deskripsiBrief || '',
      status: b.status,
      tglSelesai: b.tglSelesai || '',
      hasilAkhir: b.hasilAkhir || '',
      referensi: b.referensi || '',
      tglPostingByPlatform: datesMap,
      _prefillDate: '',
      isReference: isReferenceBrief(b),
    });
    setFormMsg('');
    setUploadedFileName('');
    setFormOpen(true);
  }
  function closeForm() {
    setFormOpen(false);
    setEditingId(null);
    setForm(EMPTY_FORM);
    setFormMsg('');
    setUploadedFileName('');
  }

  async function handleFileUpload(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    if (file.size > 10 * 1024 * 1024) {
      setFormMsg('Ukuran file maksimal 10MB.');
      e.target.value = '';
      return;
    }
    setUploading(true);
    setFormMsg('');
    try {
      const base64Data = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result.split(',')[1]);
        reader.onerror = () => reject(new Error('Gagal membaca file'));
        reader.readAsDataURL(file);
      });
      const res = await fetch('/api/upload', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ filename: file.name, mimeType: file.type, base64Data }),
      });
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.error || 'Gagal mengunggah file');
      }
      const data = await res.json();
      setForm((f) => ({ ...f, hasilAkhir: data.url }));
      setUploadedFileName(file.name);
    } catch (err) {
      setFormMsg(err.message);
    } finally {
      setUploading(false);
      e.target.value = '';
    }
  }

  function clearHasilAkhir() {
    setForm((f) => ({ ...f, hasilAkhir: '' }));
    setUploadedFileName('');
  }

  async function handleSubmit(e) {
    e.preventDefault();
    if (!form.tglMasuk || !form.brief.trim()) {
      setFormMsg('Tanggal masuk dan judul brief wajib diisi.');
      return;
    }
    if (!form.platform || form.platform.length === 0) {
      setFormMsg('Pilih minimal satu platform.');
      return;
    }
    if (form.tglSelesai && form.tglSelesai < form.tglMasuk) {
      setFormMsg('Tanggal selesai tidak boleh sebelum tanggal masuk.');
      return;
    }
    setFormMsg('');
    setSaving(true);
    const platformDates = form.platform.map((p) => (form.tglPostingByPlatform[p] || '').trim());
    const payload = {
      tglMasuk: form.tglMasuk,
      pilar: form.pilar,
      platform: form.platform.join(','),
      brief: form.brief,
      deskripsiBrief: form.deskripsiBrief,
      status: form.status,
      tglSelesai: form.tglSelesai,
      hasilAkhir: form.hasilAkhir,
      referensi: form.referensi,
      tglPosting: platformDates.join(','),
      isReference: form.isReference ? 'true' : 'false',
    };
    try {
      if (editingId) {
        const res = await fetch(`/api/briefs/${editingId}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
        });
        if (!res.ok) {
          const data = await res.json().catch(() => ({}));
          throw new Error(data.error || 'Gagal menyimpan perubahan');
        }
      } else {
        const res = await fetch('/api/briefs', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
        });
        if (!res.ok) {
          const data = await res.json().catch(() => ({}));
          throw new Error(data.error || 'Gagal menambah brief');
        }
      }
      if (form.status === 'Selesai Terupload') {
        triggerToast(`Brief "${form.brief}" telah Selesai Terupload!`);
      }
      closeForm();
      await loadBriefs();
    } catch (err) {
      setFormMsg(err.message);
    } finally {
      setSaving(false);
    }
  }

  async function handleDelete(id) {
    if (!confirm('Apakah Anda yakin ingin menghapus brief ini?')) return;
    try {
      const res = await fetch(`/api/briefs/${id}`, { method: 'DELETE' });
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.error || 'Gagal menghapus brief');
      }
      if (editingId === id) closeForm();
      if (selectedCalendarBrief && selectedCalendarBrief.brief && selectedCalendarBrief.brief.id === id) {
        setSelectedCalendarBrief(null);
      }
      await loadBriefs();
    } catch (err) {
      alert(err.message);
    }
  }

  function exportToExcel() {
    const rows = tableBriefs
      .slice()
      .sort((a, b) => new Date(a.tglMasuk) - new Date(b.tglMasuk))
      .map((b) => ({
        'Tanggal Masuk': b.tglMasuk,
        Pilar: b.pilar,
        Platform: platformLabel(b),
        'Judul Brief': b.brief,
        'Deskripsi Brief': b.deskripsiBrief || '',
        Status: b.status,
        'Tanggal Selesai': b.tglSelesai,
        'Tanggal Posting': platformDatePairs(b)
          .filter((pr) => pr.date)
          .map((pr) => `${pr.platform || 'Posting'}: ${pr.date}`)
          .join(' | '),
        KPI: kpiFor(b) || '',
        'Sumber/Referensi': b.hasilAkhir || '',
      }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = [
      { wch: 14 },
      { wch: 12 },
      { wch: 16 },
      { wch: 30 },
      { wch: 40 },
      { wch: 18 },
      { wch: 14 },
      { wch: 14 },
      { wch: 10 },
      { wch: 30 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'input');
    XLSX.writeFile(wb, 'content-briefs.xlsx');
  }

  // ------ Turunan untuk KPI / chart (berdasarkan data yang sudah difilter minggu) ------
  const [statusModal, setStatusModal] = useState(null);

  const briefsForModal = useMemo(() => {
    if (!statusModal) return [];
    if (statusModal === 'Total Brief') return filteredBriefs;
    return filteredBriefs.filter((b) => statusOf(b) === statusModal);
  }, [statusModal, filteredBriefs]);

  function handleStatusCardClick(statusName) {
    setStatusModal(statusName);
  }

  const total = filteredBriefs.length;
  const countStatus = (s) => filteredBriefs.filter((b) => statusOf(b) === s).length;
  const statuses = [
    { name: 'Selesai Terupload', color: 'var(--green)' },
    { name: 'On Going', color: 'var(--orange)' },
    { name: 'Waiting Approval', color: 'var(--red)' },
    { name: 'Belum Dikerjakan', color: 'var(--grey)' },
  ];
  const kpiCards = [
    { label: 'Total Brief', value: total, color: 'var(--ink)' },
    ...statuses.map((s) => ({ label: s.name, value: countStatus(s.name), color: s.color })),
  ];
  const maxStatus = Math.max(1, ...statuses.map((s) => countStatus(s.name)));

  const pilarColors = {
    Ads: '#ff9500',
    Feed: '#af52de',
    Story: '#0071e3',
    Carousel: '#64d2ff',
    Reels: '#ff9500',
    Lainnya: '#8e8e93',
  };
  const pilarCounts = PILARS.map((name) => ({
    name,
    color: pilarColors[name],
    n: filteredBriefs.filter((b) => b.pilar === name && statusOf(b) === 'Selesai Terupload').length,
  }));
  const totalPilar = Math.max(1, pilarCounts.reduce((a, c) => a + c.n, 0));
  let acc = 0;
  const stops = pilarCounts
    .map((p) => {
      const start = (acc / totalPilar) * 360;
      acc += p.n;
      const end = (acc / totalPilar) * 360;
      return `${p.color} ${start}deg ${end}deg`;
    })
    .join(', ');
  const donutBg =
    totalPilar > 0
      ? `radial-gradient(circle at center, #fff 42%, transparent 43%), conic-gradient(${stops})`
      : 'var(--hair)';

  const kpiVals = {
    'On Time': filteredBriefs.filter((b) => kpiFor(b) === 'On Time').length,
    Late: filteredBriefs.filter((b) => kpiFor(b) === 'Late').length,
    'Belum Selesai': filteredBriefs.filter((b) => kpiFor(b) === null).length,
  };
  const kpiColors = { 'On Time': 'var(--green)', Late: 'var(--red)', 'Belum Selesai': 'var(--grey)' };
  const maxKpi = Math.max(1, ...Object.values(kpiVals));

  const sortedBriefs = tableBriefs.slice().sort((a, b) => new Date(b.tglMasuk) - new Date(a.tglMasuk));

  // ------ Kalender Konten: grouping brief berdasarkan tglPosting per platform ------
  const postingByDate = useMemo(() => {
    const map = {};
    briefs.forEach((b) => {
      platformDatePairs(b).forEach((pr) => {
        if (!pr.date) return;
        if (!map[pr.date]) map[pr.date] = [];
        map[pr.date].push({ brief: b, platform: pr.platform });
      });
    });
    // Urutkan item tiap tanggal: Instagram dulu (Story > Carousel > Feed > Reels),
    // baru Tiktok, sisanya menyusul.
    Object.keys(map).forEach((date) => {
      map[date].sort((a, b) => {
        if (!a || !b) return 0;
        const platformDiff = PLATFORM_ORDER_INDEX(a.platform) - PLATFORM_ORDER_INDEX(b.platform);
        if (platformDiff !== 0) return platformDiff;
        if (a.platform === 'Instagram') {
          const aPilar = a.brief ? a.brief.pilar : '';
          const bPilar = b.brief ? b.brief.pilar : '';
          return IG_PILAR_ORDER_INDEX(aPilar) - IG_PILAR_ORDER_INDEX(bPilar);
        }
        return 0;
      });
    });
    return map;
  }, [briefs]);
  const unscheduledBriefs = useMemo(
    () =>
      briefs
        .filter((b) => platformDatePairs(b).every((pr) => !pr.date))
        .sort((a, b) => new Date(b.tglMasuk) - new Date(a.tglMasuk)),
    [briefs]
  );
  const monthGrid = useMemo(() => buildMonthGrid(calendarMonth), [calendarMonth]);
  const [expandedCell, setExpandedCell] = useState(null);
  const [selectedCalendarBrief, setSelectedCalendarBrief] = useState(null);
  const [dragOverDate, setDragOverDate] = useState(null);
  const [toastMsg, setToastMsg] = useState(null);
  const today = todayISO();

  function triggerToast(msg) {
    setToastMsg(msg);
    setTimeout(() => setToastMsg(null), 5000);
  }

  function getWANotifURL(brief) {
    if (!brief) return 'https://wa.me/6285603524508';
    const msg = `🔥 PENGINGAT POSTING!\n\n📌 Judul Brief: ${brief.brief || '-'}\n📱 Platform: ${platformLabel(brief)}\n🏷️ Pilar: ${brief.pilar || '-'}\n🗓️ Tanggal Posting: ${brief.tglPosting || '-'}\n\n🔗 Link Media: ${brief.hasilAkhir || '-'}`;
    return `https://wa.me/6285603524508?text=${encodeURIComponent(msg)}`;
  }

  const [dismissedReminder, setDismissedReminder] = useState(false);

  const todayDueBriefs = useMemo(
    () =>
      briefs.filter(
        (b) =>
          b &&
          String(b.tglPosting || '').includes(today) &&
          statusOf(b) !== 'Selesai Terupload'
      ),
    [briefs, today]
  );

  async function forceDownloadFile(url, filename) {
    if (!url) return;
    try {
      const res = await fetch(url);
      const blob = await res.blob();
      const blobUrl = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = blobUrl;
      a.download = filename || 'media_asset';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(blobUrl), 10000);
    } catch (err) {
      console.error('Blob download failed, opening URL:', err);
      window.open(url, '_blank');
    }
  }

  function triggerAutoDownloadAndPlatformRedirect(brief, targetPlatform) {
    if (!brief) return;
    const mediaUrl = brief.hasilAkhir;
    if (mediaUrl) {
      let downloadUrl = mediaUrl;
      if (mediaUrl.includes('drive.google.com')) {
        const match = mediaUrl.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) || mediaUrl.match(/id=([a-zA-Z0-9_-]+)/);
        if (match && match[1]) {
          downloadUrl = `https://drive.google.com/uc?export=download&id=${match[1]}`;
        }
      }

      const isFolder = mediaUrl.includes('/folders/');
      if (isFolder) {
        window.open(mediaUrl, '_blank');
      } else {
        const cleanName = brief.brief ? brief.brief.replace(/[^a-z0-9]/gi, '_') : 'media_asset';
        const proxyUrl = `/api/download?url=${encodeURIComponent(downloadUrl)}&filename=${encodeURIComponent(cleanName)}`;
        
        // Force direct browser download using hidden iframe trick
        const iframe = document.createElement('iframe');
        iframe.style.display = 'none';
        iframe.src = proxyUrl;
        document.body.appendChild(iframe);
        setTimeout(() => {
          if (document.body.contains(iframe)) document.body.removeChild(iframe);
        }, 15000);
      }
    }

    // Open target platform in new tab after 600ms so download initiates first
    setTimeout(() => {
      const plat = (targetPlatform || platformsOf(brief)[0] || '').toLowerCase();
      const dest = plat.includes('tiktok') || plat.includes('tok') ? 'https://www.tiktok.com' : 'https://www.instagram.com';
      window.open(dest, '_blank');
    }, 600);
  }

  async function updateBriefPostingDate(brief, targetPlatform, targetDate) {
    const plats = platformsOf(brief);
    const pairs = platformDatePairs(brief);

    let newDates = [];
    if (plats.length <= 1) {
      newDates = [targetDate];
    } else {
      newDates = plats.map((p) => {
        if (p === targetPlatform) return targetDate;
        const found = pairs.find((pair) => pair.platform === p);
        return found ? found.date : '';
      });
    }

    const updatedTglPosting = newDates.join(',');

    // Optimistic UI update
    setBriefs((prev) =>
      prev.map((item) => (item.id === brief.id ? { ...item, tglPosting: updatedTglPosting } : item))
    );

    try {
      const payload = {
        ...brief,
        tglPosting: updatedTglPosting,
      };
      const res = await fetch(`/api/briefs/${brief.id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || 'Gagal menyimpan perubahan tanggal posting');
      }
    } catch (err) {
      console.error('Failed to save drop date update:', err);
      loadBriefs();
    }
  }

  function handleDropOnDate(e, targetDate) {
    try {
      const rawData = e.dataTransfer.getData('application/json');
      if (!rawData) return;
      const { briefId, platform } = JSON.parse(rawData);
      const brief = briefs.find((b) => String(b.id) === String(briefId));
      if (!brief) return;

      updateBriefPostingDate(brief, platform, targetDate);
    } catch (err) {
      console.error('Error handling drag drop:', err);
    }
  }

  // --- Import CSV (transisi dari sheet kalender manual lama) ---
  const importFileRef = useRef(null);
  const [importOpen, setImportOpen] = useState(false);
  const [importBlocks, setImportBlocks] = useState([]);
  const [importBlockMonths, setImportBlockMonths] = useState([]);
  const [importRows, setImportRows] = useState([]);
  const [importMsg, setImportMsg] = useState('');
  const [importing, setImporting] = useState(false);
  const [importProgress, setImportProgress] = useState('');
  const [importAsReference, setImportAsReference] = useState(true);
  const [showDateSummary, setShowDateSummary] = useState(false);
  const [importBlockPlatforms, setImportBlockPlatforms] = useState([]);

  function applyImportRows(blocks, months, blockPlatforms) {
    setImportRows(blocksToImportRows(blocks, months, briefs, blockPlatforms));
  }

  function updateBlockPlatform(bIdx, value) {
    const platforms = importBlockPlatforms.slice();
    platforms[bIdx] = value;
    setImportBlockPlatforms(platforms);
    applyImportRows(importBlocks, importBlockMonths, platforms);
  }

  function applyPlatformToAllBlocks(value) {
    const platforms = importBlocks.map(() => value);
    setImportBlockPlatforms(platforms);
    applyImportRows(importBlocks, importBlockMonths, platforms);
  }

  function applyStartMonthSequential(startYYYYMM) {
    if (!startYYYYMM) return;
    let [y, m] = startYYYYMM.split('-').map((n) => parseInt(n, 10));
    const newMonths = [];
    let currentY = y;
    let currentM = m;

    importBlocks.forEach((block, idx) => {
      if (idx > 0 && (block.platformName === 'Instagram' || block.sheetName !== importBlocks[idx - 1].sheetName)) {
        currentM++;
        if (currentM > 12) {
          currentM = 1;
          currentY++;
        }
      }
      const mm = String(currentM).padStart(2, '0');
      newMonths.push(`${currentY}-${mm}`);
    });

    setImportBlockMonths(newMonths);
    applyImportRows(importBlocks, newMonths, importBlockPlatforms);
  }

  function handleImportFileChange(e) {
    const file = e.target.files && e.target.files[0];
    e.target.value = '';
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: 'array', cellStyles: true });
        const blocks = parseWorkbookToAllBlocks(wb);
        if (blocks.length === 0) {
          setImportMsg('Tidak menemukan pola grid kalender (header Senin–Minggu) di file ini.');
          return;
        }
        const months = blocks.map((b) => b.detectedMonth || '');
        const platforms = blocks.map((b, idx) => {
          if (b.platformName) return b.platformName;
          const sheetBlocks = blocks.filter((x) => x.sheetName === b.sheetName);
          const blockInSheetIdx = sheetBlocks.indexOf(b);
          return blockInSheetIdx % 2 === 0 ? 'Instagram' : 'Tiktok';
        });
        setImportBlocks(blocks);
        setImportBlockMonths(months);
        setImportBlockPlatforms(platforms);
        setImportRows(blocksToImportRows(blocks, months, briefs, platforms));
        setImportMsg('');
        setImportAsReference(true);
        setImportOpen(true);
      } catch (err) {
        setImportMsg('Gagal membaca file Excel/CSV: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function updateBlockMonth(bIdx, value) {
    const months = importBlockMonths.slice();
    months[bIdx] = value;
    setImportBlockMonths(months);
    applyImportRows(importBlocks, months, importBlockPlatforms);
  }

  function updateImportRow(id, patch) {
    setImportRows((rows) => rows.map((r) => (r.id === id ? { ...r, ...patch } : r)));
  }

  function closeImportModal() {
    setImportOpen(false);
    setImportBlocks([]);
    setImportBlockMonths([]);
    setImportRows([]);
    setImportMsg('');
  }

  async function submitImportRows() {
    const existingTitles = new Set(briefs.map((b) => (b.brief || '').trim().toLowerCase()));
    const toImport = importRows.filter((r) => r.include && !existingTitles.has((r.title || '').trim().toLowerCase()));
    if (toImport.length === 0) {
      setImportMsg('Semua brief yang dipilih sudah ada di dashboard atau tidak ada brief baru untuk diimport.');
      return;
    }
    if (toImport.some((r) => !r.tglPosting)) {
      setImportMsg('Ada brief yang belum punya tanggal — pilih bulan & tahun untuk semua blok dulu.');
      return;
    }
    setImporting(true);
    setImportMsg('');
    let done = 0;
    for (const row of toImport) {
      done += 1;
      setImportProgress(`Mengimpor ${done}/${toImport.length}…`);
      try {
        await fetch('/api/briefs', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            tglMasuk: row.tglPosting || todayISO(),
            pilar: row.pilar,
            platform: row.platform.join(','),
            brief: row.title,
            deskripsiBrief: '',
            status: '',
            tglSelesai: '',
            hasilAkhir: '',
            tglPosting: row.tglPosting,
            isReference: importAsReference ? 'true' : 'false',
          }),
        });
      } catch (err) {
        // lanjut ke brief berikutnya walau satu gagal, biar tidak stuck
      }
    }
    setImporting(false);
    setImportProgress('');
    await loadBriefs();
    closeImportModal();
  }

  return (
    <div className="wrap">
      <header>
        <div>
          <h1>Content Production</h1>
          <p>Target: Ads &amp; Carousel maks. 2 hari · Reels &amp; Lainnya maks. 3 hari</p>
        </div>
        <div className="header-actions">
          <div className="view-switch">
            <button
              className={`view-switch-btn${view === 'dashboard' ? ' is-active' : ''}`}
              onClick={() => setView('dashboard')}
            >
              Dashboard
            </button>
            <button
              className={`view-switch-btn${view === 'kalender' ? ' is-active' : ''}`}
              onClick={() => setView('kalender')}
            >
              Kalender Konten
            </button>
          </div>
          <button className="btn btn-outline" onClick={loadBriefs}>Refresh</button>
          {view === 'dashboard' && (
            <button className="btn btn-primary" onClick={exportToExcel}>Export ke Excel</button>
          )}
        </div>
      </header>
      <p className="sync-note">
        Backend: <b>Google Sheets</b> (live). Setiap tambah / edit / hapus langsung tersimpan ke spreadsheet lewat
        service account, jadi data tidak hilang saat refresh.
      </p>
      {loadError && (
        <p className="sync-note error">
          Gagal memuat data: {loadError}. Cek konfigurasi <b>.env.local</b> dan pastikan sheet sudah di-share ke
          email service account.
        </p>
      )}

      {loading ? (
        <div className="loading">Memuat data dari Google Sheets…</div>
      ) : view === 'kalender' ? (
        <div className="calendar-view">
          <div className="calendar-toolbar">
            <div className="calendar-nav">
              <button
                onClick={() => setCalendarMonth((m) => new Date(m.getFullYear(), m.getMonth() - 1, 1))}
                title="Bulan sebelumnya"
              >
                ‹
              </button>
              <span className="calendar-month-label">{monthLabel(calendarMonth)}</span>
              <button
                onClick={() => setCalendarMonth((m) => new Date(m.getFullYear(), m.getMonth() + 1, 1))}
                title="Bulan berikutnya"
              >
                ›
              </button>
            </div>
            <button
              className="btn btn-ghost btn-sm"
              onClick={() => {
                const d = new Date();
                setCalendarMonth(new Date(d.getFullYear(), d.getMonth(), 1));
              }}
            >
              Bulan Ini
            </button>
            <button className="btn btn-outline btn-sm" onClick={() => importFileRef.current?.click()}>
              Import XLSX / CSV
            </button>
            <input
              type="file"
              accept=".xlsx, .xls, .csv"
              ref={importFileRef}
              style={{ display: 'none' }}
              onChange={handleImportFileChange}
            />
          </div>
          {importMsg && !importOpen && <p className="msg">{importMsg}</p>}

          <div className="calendar-grid">
            {DAY_LABELS.map((d) => (
              <div className="calendar-dow" key={d}>{d}</div>
            ))}
            {monthGrid.map((cell) => {
              const items = postingByDate[cell.iso] || [];
              const isOpen = expandedCell === cell.iso;
              const isOver = dragOverDate === cell.iso;
              return (
                <div
                  key={cell.iso}
                  className={`calendar-cell${cell.inMonth ? '' : ' is-outside'}${cell.iso === today ? ' is-today' : ''}${isOver ? ' is-drag-over' : ''}`}
                  onDoubleClick={() => openAddForm({ _prefillDate: cell.iso })}
                  onDragOver={(e) => {
                    e.preventDefault();
                    e.dataTransfer.dropEffect = 'move';
                    if (dragOverDate !== cell.iso) setDragOverDate(cell.iso);
                  }}
                  onDragLeave={(e) => {
                    if (dragOverDate === cell.iso) setDragOverDate(null);
                  }}
                  onDrop={(e) => {
                    e.preventDefault();
                    setDragOverDate(null);
                    handleDropOnDate(e, cell.iso);
                  }}
                  title="Klik dua kali untuk tambah brief di tanggal ini, atau geser (drag & drop) brief ke sini"
                >
                  <div className="calendar-cell-date">{cell.date.getDate()}</div>
                  <div className="calendar-cell-items">
                    {(isOpen ? items : items.slice(0, 3)).map((it, idx) => {
                      const b = it.brief;
                      const style = CALENDAR_STATUS_STYLE[statusOf(b)] || CALENDAR_STATUS_STYLE['Belum Dikerjakan'];
                      const label = it.platform ? platformAbbr(it.platform) : platformLabel(b);
                      const badgeColor = formatColor(it.platform, b && b.pilar);
                      const isRef = isReferenceBrief(b);
                      return (
                        <div
                          key={`${b.id}-${it.platform}-${idx}`}
                          className={`calendar-item${isRef ? ' is-reference' : ''}`}
                          style={{ background: style.bg, borderLeftColor: style.border }}
                          draggable={true}
                          onDragStart={(e) => {
                            e.stopPropagation();
                            e.dataTransfer.setData('application/json', JSON.stringify({ briefId: b.id, platform: it.platform }));
                            e.dataTransfer.effectAllowed = 'move';
                          }}
                          onClick={(e) => {
                            e.stopPropagation();
                            setSelectedCalendarBrief({ brief: b, platform: it.platform });
                          }}
                          title={`${isRef ? '[Transisi/Referensi] ' : ''}${it.platform || platformLabel(b)} · ${pilarDisplayLabel(it.platform, b.pilar)} · ${b.brief} — Klik untuk detail`}
                        >
                          <span className="calendar-item-badge" style={{ background: badgeColor }}>
                            {isRef ? '📌 ' : ''}{label} · {pilarDisplayLabel(it.platform, b.pilar)}
                          </span>
                          {cell.iso === today && (
                            <span className="due-today-badge">🔥 Hari Ini</span>
                          )}
                          <div className="calendar-item-title">{b.brief}</div>
                        </div>
                      );
                    })}
                    {!isOpen && items.length > 3 && (
                      <button className="calendar-more" onClick={() => setExpandedCell(cell.iso)}>
                        +{items.length - 3} lagi
                      </button>
                    )}
                    {isOpen && items.length > 3 && (
                      <button className="calendar-more" onClick={() => setExpandedCell(null)}>
                        Ciutkan
                      </button>
                    )}
                  </div>
                </div>
              );
            })}
          </div>

          <div className="calendar-legend">
            <span className="calendar-legend-title">Jenis &amp; Status:</span>
            <div>
              <span className="dot" style={{ background: '#0071e3' }} />
              Brief KPI Utama
            </div>
            <div>
              <span className="dot" style={{ background: '#8e8e93' }} />
              📌 Data Transisi/Referensi (Abaikan dari KPI Dashboard)
            </div>
          </div>
          <div className="calendar-legend">
            <span className="calendar-legend-title">Warna sel = status:</span>
            {Object.entries(CALENDAR_STATUS_STYLE).map(([label, style]) => (
              <div key={label}>
                <span className="dot" style={{ background: style.border }} />
                {label}
              </div>
            ))}
          </div>
          <div className="calendar-legend">
            <span className="calendar-legend-title">Warna badge = platform:</span>
            {PLATFORMS.map((p) => (
              <div key={p}>
                <span className="dot" style={{ background: formatColor(p) }} />
                {platformAbbr(p)}
              </div>
            ))}
          </div>

          <div className="list-panel" style={{ marginTop: 24 }}>
            <div className="list-head">
              <h3>Belum Dijadwalkan</h3>
              <span style={{ fontSize: 12.5, color: 'var(--sub)' }}>{unscheduledBriefs.length} brief</span>
            </div>
            {unscheduledBriefs.length === 0 ? (
              <div className="empty">Semua brief sudah punya tanggal posting.</div>
            ) : (
              <div className="unscheduled-list">
                {unscheduledBriefs.map((b) => (
                  <div
                    className="unscheduled-item"
                    key={b.id}
                    draggable={true}
                    onDragStart={(e) => {
                      e.stopPropagation();
                      const firstPlat = platformsOf(b)[0] || 'Instagram';
                      e.dataTransfer.setData('application/json', JSON.stringify({ briefId: b.id, platform: firstPlat }));
                      e.dataTransfer.effectAllowed = 'move';
                    }}
                    onClick={() => enterEditMode(b.id)}
                  >
                    {platformsOf(b).map((p) => (
                      <span key={p} className="dot" style={{ background: PLATFORM_COLORS[p] || 'var(--grey)' }} />
                    ))}
                    <span className="unscheduled-brief">{b.brief}</span>
                    <span className="tag">{b.pilar}</span>
                    <span className="tag">{platformLabel(b)}</span>
                    <span className="unscheduled-hint">Geser (drag & drop) ke tanggal kalender atau klik untuk atur</span>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      ) : (
        <>
          <div className="week-filter-panel">
            <div className="week-filter-panel-label">Filter Minggu Produksi</div>
            <div className="week-filter">
              <button
                className={`week-toggle${weekFilter ? '' : ' is-off'}`}
                onClick={() => setWeekFilter(weekFilter === null ? currentWeekIndex() : null)}
              >
                {weekFilter ? 'Filter: Mingguan' : 'Semua Minggu'}
              </button>
              {weekFilter !== null && (
                <div className="week-nav">
                  <button onClick={() => setWeekFilter(Math.max(0, weekFilter - 1))} title="Minggu sebelumnya">‹</button>
                  <span className="week-label">{formatWeekLabel(weekFilter)}</span>
                  <button onClick={() => setWeekFilter(weekFilter + 1)} title="Minggu berikutnya">›</button>
                </div>
              )}
              {weekFilter !== null && weekFilter !== currentWeekIndex() && (
                <button className="btn btn-ghost btn-sm" onClick={() => setWeekFilter(currentWeekIndex())}>
                  Minggu Ini
                </button>
              )}
            </div>
          </div>

          <div className="kpi-grid">
            {kpiCards.map((c) => {
              return (
                <div
                  className="kpi clickable"
                  key={c.label}
                  onClick={() => handleStatusCardClick(c.label)}
                  onDoubleClick={() => handleStatusCardClick(c.label)}
                  title={`Klik / Double klik untuk membuka pop-up brief ${c.label}`}
                >
                  <div className="label">{c.label}</div>
                  <div className="value" style={{ color: c.color }}>{c.value}</div>
                </div>
              );
            })}
          </div>

          <div className="panels">
            <div className="panel">
              <h3>Status Brief</h3>
              <div>
                {statuses.map((s) => {
                  const n = countStatus(s.name);
                  const pct = ((n / maxStatus) * 100).toFixed(0);
                  const isSelected = filterStatus === s.name;
                  return (
                    <div
                      className="bar-row clickable"
                      key={s.name}
                      onClick={() => handleStatusCardClick(s.name)}
                      onDoubleClick={() => handleStatusCardClick(s.name)}
                      title={`Klik / Double klik untuk membuka pop-up status ${s.name}`}
                    >
                      <div className="name" style={{ fontWeight: isSelected ? 700 : 500 }}>{s.name}</div>
                      <div className="bar-track">
                        <div className="bar-fill" style={{ width: pct + '%', background: s.color }} />
                      </div>
                      <div className="num">{n}</div>
                    </div>
                  );
                })}
              </div>
            </div>
            <div className="panel">
              <h3>Distribusi Pilar (Selesai)</h3>
              <div className="donut-wrap">
                <div className="donut" style={{ background: donutBg }} />
                <div className="legend">
                  {pilarCounts.map((p) => (
                    <div key={p.name}>
                      <span className="dot" style={{ background: p.color }} />
                      {p.name} — {p.n}
                    </div>
                  ))}
                </div>
              </div>
            </div>
            <div className="panel">
              <h3>Ketepatan Waktu</h3>
              <div>
                {Object.entries(kpiVals).map(([k, v]) => {
                  const pct = ((v / maxKpi) * 100).toFixed(0);
                  return (
                    <div className="bar-row" key={k}>
                      <div className="name">{k}</div>
                      <div className="bar-track">
                        <div className="bar-fill" style={{ width: pct + '%', background: kpiColors[k] }} />
                      </div>
                      <div className="num">{v}</div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>

          <div className="list-panel daily-panel" id="daily-panel">
            <div className="list-head">
              <h3>Brief Selesai per Hari</h3>
              <span style={{ fontSize: 12.5, color: 'var(--sub)' }}>
                {dailyCompleted.reduce((a, d) => a + d.items.length, 0)} brief selesai
                {weekFilter ? ' pada minggu ini' : ''}
              </span>
            </div>
            {dailyCompleted.length === 0 ? (
              <div className="empty">Belum ada brief yang selesai.</div>
            ) : (
              dailyCompleted.map((d) => {
                const isOpen = expandedDate === d.date;
                const pct = ((d.items.length / maxDaily) * 100).toFixed(0);
                return (
                  <div className="day-row" key={d.date}>
                    <div
                      className={`day-row-head${isOpen ? ' open' : ''}`}
                      onClick={() => setExpandedDate(isOpen ? null : d.date)}
                    >
                      <span className="chevron">▸</span>
                      <div className="day-row-date">{fmtDateShort(d.date)}</div>
                      <div className="day-row-track">
                        <div className="day-row-fill" style={{ width: pct + '%' }} />
                      </div>
                      <div className="day-row-count">{d.items.length}</div>
                    </div>
                    {isOpen && (
                      <div className="day-row-list">
                        {d.items.map((b) => (
                          <div className="day-row-item" key={b.id}>
                            <b>{b.brief}</b>
                            <span className="tag">{b.pilar}</span>
                            <span className="tag">{platformLabel(b)}</span>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                );
              })
            )}
          </div>

          <div className="list-panel" id="table-panel">
            <div className="list-head" style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
              <h3>Daftar Brief</h3>
              {filterStatus && (
                <span style={{ fontSize: 12.5, background: 'rgba(0,113,227,0.1)', color: 'var(--blue)', padding: '4px 12px', borderRadius: 20, display: 'inline-flex', alignItems: 'center', gap: 6, fontWeight: 600 }}>
                  Menampilkan: <b>{filterStatus}</b> ({sortedBriefs.length} brief)
                  <button
                    type="button"
                    onClick={() => setFilterStatus('')}
                    style={{ border: 'none', background: 'none', color: 'var(--blue)', cursor: 'pointer', fontWeight: 700, marginLeft: 4 }}
                    title="Tampilkan Semua Brief"
                  >
                    ✕ Clear
                  </button>
                </span>
              )}
            </div>

            <div className="table-filters">
              <select value={filterPilar} onChange={(e) => setFilterPilar(e.target.value)}>
                <option value="">Semua Pilar</option>
                {PILARS.map((p) => (
                  <option key={p} value={p}>{p}</option>
                ))}
              </select>
              <select value={filterPlatform} onChange={(e) => setFilterPlatform(e.target.value)}>
                <option value="">Semua Platform</option>
                {PLATFORMS.map((p) => (
                  <option key={p} value={p}>{p}</option>
                ))}
              </select>
              <select value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)}>
                <option value="">Semua Status</option>
                <option value="Belum Dikerjakan">Belum Dikerjakan</option>
                {STATUS_OPTIONS.map((s) => (
                  <option key={s} value={s}>{s}</option>
                ))}
              </select>
              {hasTableFilter && (
                <button
                  className="btn btn-ghost btn-sm"
                  onClick={() => {
                    setFilterPilar('');
                    setFilterPlatform('');
                    setFilterStatus('');
                  }}
                >
                  Reset Filter
                </button>
              )}
            </div>

            <div style={{ overflowX: 'auto' }}>
              <table>
                <thead>
                  <tr>
                    <th>Tanggal Masuk</th>
                    <th>Pilar</th>
                    <th>Platform</th>
                    <th>Judul Brief</th>
                    <th>Status</th>
                    <th>Tanggal Selesai</th>
                    <th>KPI</th>
                    <th>Hasil Final</th>
                    <th>Referensi</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {sortedBriefs.map((b) => {
                    const k = kpiFor(b);
                    return (
                      <tr key={b.id} className={b.id === editingId ? 'is-editing' : ''}>
                        <td>{fmtDate(b.tglMasuk)}</td>
                        <td>{b.pilar}</td>
                        <td>{platformLabel(b)}</td>
                        <td>{b.brief}</td>
                        <td><span className={`pill ${pillClass(statusOf(b))}`}>{statusOf(b)}</span></td>
                        <td>{fmtDate(b.tglSelesai)}</td>
                        <td>
                          {k === 'On Time' && <span className="kpi-ok">✓ On Time</span>}
                          {k === 'Late' && <span className="kpi-late">✕ Late</span>}
                          {k === null && <span className="kpi-none">Belum</span>}
                        </td>
                        <td>
                          {b.hasilAkhir ? (
                            <a className="result-link" href={b.hasilAkhir} target="_blank" rel="noopener noreferrer">
                              Hasil ↗
                            </a>
                          ) : (
                            <span className="kpi-none">-</span>
                          )}
                        </td>
                        <td>
                          {b.referensi ? (
                            <a className="result-link" href={b.referensi} target="_blank" rel="noopener noreferrer">
                              Ref ↗
                            </a>
                          ) : (
                            <span className="kpi-none">-</span>
                          )}
                        </td>
                        <td>
                          <div className="row-actions">
                            <button className="icon-btn edit" title="Edit" onClick={() => enterEditMode(b.id)}>✎</button>
                            <button className="icon-btn del" title="Hapus" onClick={() => handleDelete(b.id)}>✕</button>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            {sortedBriefs.length === 0 && (
              <div className="empty">
                {hasTableFilter
                  ? 'Tidak ada brief yang cocok dengan filter ini.'
                  : weekFilter
                  ? 'Tidak ada brief pada minggu ini.'
                  : 'Belum ada brief. Klik tombol + di kanan bawah untuk menambahkan.'}
              </div>
            )}
          </div>
        </>
      )}

      <button className="fab" onClick={openAddForm} title="Tambah Brief">+</button>

      {formOpen && (
        <div className="modal-overlay" onClick={(e) => e.target === e.currentTarget && closeForm()}>
          <div className="modal-card">
            <div className="form-head">
              <h3>{editingId ? 'Edit Brief' : 'Tambah Brief'}</h3>
              <button className="modal-close" onClick={closeForm} title="Tutup">✕</button>
            </div>
            <form onSubmit={handleSubmit}>
              <div className="form-grid">
                <div className="field">
                  <label htmlFor="tglMasuk">Tanggal Masuk</label>
                  <input
                    type="date"
                    id="tglMasuk"
                    required
                    value={form.tglMasuk}
                    onChange={(e) => setForm({ ...form, tglMasuk: e.target.value })}
                  />
                </div>
                <div className="field">
                  <label htmlFor="pilar">Pilar</label>
                  <select id="pilar" value={form.pilar} onChange={(e) => setForm({ ...form, pilar: e.target.value })}>
                    {PILARS.map((p) => (
                      <option key={p} value={p}>{p}</option>
                    ))}
                  </select>
                </div>
                <div className="field span2">
                  <label>Platform</label>
                  <div className="checkbox-row">
                    {PLATFORMS.map((p) => (
                      <label key={p} className="checkbox-chip">
                        <input
                          type="checkbox"
                          checked={form.platform.includes(p)}
                          onChange={(e) => {
                            const checked = e.target.checked;
                            const next = checked
                              ? [...form.platform, p]
                              : form.platform.filter((x) => x !== p);
                            const datesMap = { ...form.tglPostingByPlatform };
                            if (checked && !datesMap[p] && form._prefillDate) {
                              datesMap[p] = form._prefillDate;
                            }
                            setForm({ ...form, platform: next, tglPostingByPlatform: datesMap });
                          }}
                        />
                        {p}
                      </label>
                    ))}
                  </div>
                </div>
                <div className="field span3">
                  <label htmlFor="briefText">Judul Brief</label>
                  <input
                    type="text"
                    id="briefText"
                    placeholder="Judul singkat brief"
                    required
                    value={form.brief}
                    onChange={(e) => setForm({ ...form, brief: e.target.value })}
                  />
                </div>
                <div className="field span3">
                  <label htmlFor="deskripsiBrief">Deskripsi Brief</label>
                  <textarea
                    id="deskripsiBrief"
                    placeholder="Detail brief: konsep, hook, copy, catatan produksi, dll."
                    rows={4}
                    value={form.deskripsiBrief}
                    onChange={(e) => setForm({ ...form, deskripsiBrief: e.target.value })}
                  />
                </div>
                <div className="field">
                  <label htmlFor="status">Status</label>
                  <select id="status" value={form.status} onChange={(e) => setForm({ ...form, status: e.target.value })}>
                    <option value="">Belum Dikerjakan</option>
                    {STATUS_OPTIONS.map((s) => (
                      <option key={s} value={s}>{s}</option>
                    ))}
                  </select>
                </div>
                <div className="field">
                  <label htmlFor="tglSelesai">Tanggal Selesai</label>
                  <input
                    type="date"
                    id="tglSelesai"
                    value={form.tglSelesai}
                    onChange={(e) => setForm({ ...form, tglSelesai: e.target.value })}
                  />
                </div>
                <div className="field span3">
                  <label>Tanggal Posting per Platform</label>
                  {form.platform.length === 0 ? (
                    <p className="posting-date-empty">Pilih platform dulu untuk atur tanggal postingnya.</p>
                  ) : (
                    <div className="posting-date-row">
                      {form.platform.map((p) => (
                        <div className="posting-date-item" key={p}>
                          <span>{p}</span>
                          <input
                            type="date"
                            value={form.tglPostingByPlatform[p] || ''}
                            onChange={(e) =>
                              setForm({
                                ...form,
                                tglPostingByPlatform: { ...form.tglPostingByPlatform, [p]: e.target.value },
                              })
                            }
                          />
                        </div>
                      ))}
                    </div>
                  )}
                </div>
                <div className="field span3">
                  <label htmlFor="hasilAkhir">Hasil Final (link media atau file)</label>
                  <div className="hasil-akhir-row">
                    <input
                      type="text"
                      id="hasilAkhir"
                      placeholder="Tempel link Hasil Final (Drive, Canva, mp4, dll)…"
                      value={form.hasilAkhir}
                      onChange={(e) => {
                        setForm({ ...form, hasilAkhir: e.target.value });
                        setUploadedFileName('');
                      }}
                    />
                    <label className="btn btn-outline btn-sm upload-btn">
                      {uploading ? 'Mengunggah…' : 'Upload File'}
                      <input type="file" hidden onChange={handleFileUpload} disabled={uploading} />
                    </label>
                  </div>
                  {form.hasilAkhir && (
                    <div className="hasil-akhir-preview">
                      {uploadedFileName ? `File: ${uploadedFileName} — ` : ''}
                      <a href={form.hasilAkhir} target="_blank" rel="noopener noreferrer">{form.hasilAkhir}</a>
                      <button type="button" className="hasil-akhir-clear" onClick={clearHasilAkhir} title="Hapus">✕</button>
                    </div>
                  )}
                </div>
                <div className="field span3">
                  <label htmlFor="referensi">Referensi (link rujukan / benchmark)</label>
                  <input
                    type="text"
                    id="referensi"
                    placeholder="Tempel link Referensi (Instagram, TikTok, Web, dll)…"
                    value={form.referensi}
                    onChange={(e) => setForm({ ...form, referensi: e.target.value })}
                  />
                </div>
                <div className="field span3">
                  <label className="checkbox-chip" style={{ background: 'var(--bg)', border: '1px solid var(--hair)', padding: '8px 12px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 8 }}>
                    <input
                      type="checkbox"
                      checked={form.isReference}
                      onChange={(e) => setForm({ ...form, isReference: e.target.checked })}
                    />
                    <span><b>Data Transisi / Referensi Kalender</b> (Abaikan dari Total Brief &amp; KPI Dashboard)</span>
                  </label>
                </div>
              </div>
              <div className="msg">{formMsg}</div>
              <div className="form-actions">
                <button type="button" className="btn btn-ghost" onClick={closeForm}>Batal</button>
                <button type="submit" className="btn btn-primary" disabled={saving}>
                  {saving ? 'Menyimpan…' : editingId ? 'Simpan Perubahan' : 'Tambah Brief'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {importOpen && (
        <div className="modal-overlay" onClick={(e) => e.target === e.currentTarget && closeImportModal()}>
          <div className="modal-card import-modal-card">
            <div className="form-head">
              <h3>Import CSV — Content Plan Lama</h3>
              <button className="modal-close" onClick={closeImportModal} title="Tutup">✕</button>
            </div>

            <p className="import-hint">
              Terdeteksi {importBlocks.length} blok kalender di file ini. Pilih bulan &amp; tahun tiap blok
              supaya tanggal postingnya benar, lalu cek daftar di bawah sebelum diimport.
            </p>

            <div style={{ background: 'rgba(0,113,227,0.06)', border: '1px solid rgba(0,113,227,0.18)', borderRadius: 10, padding: '10px 14px', marginBottom: 16 }}>
              <label style={{ display: 'flex', alignItems: 'center', gap: 10, fontSize: 13, cursor: 'pointer', fontWeight: 500, color: 'var(--ink)' }}>
                <input
                  type="checkbox"
                  checked={importAsReference}
                  onChange={(e) => setImportAsReference(e.target.checked)}
                />
                <span>Impor sebagai <b>Data Transisi / Referensi Kalender</b> (Hanya tampil di Kalender Konten, tidak dihitung pada KPI Dashboard)</span>
              </label>
            </div>

            <div className="import-quickfill" style={{ background: 'var(--bg)', padding: '12px 14px', borderRadius: 10, border: '1px solid var(--hair)', marginBottom: 16 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', marginBottom: 8 }}>
                <span style={{ fontWeight: 500, fontSize: 13 }}>🗓️ Isi Bulan Awal (Urut Otomatis Semua Blok):</span>
                <input
                  type="month"
                  style={{ border: '1px solid var(--hair)', borderRadius: 6, padding: '4px 8px', fontSize: 13, fontFamily: 'inherit' }}
                  onChange={(e) => applyStartMonthSequential(e.target.value)}
                />
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                <span>📱 Terapkan Platform ke semua blok:</span>
                {PLATFORMS.map((p) => (
                  <button key={p} type="button" className="btn btn-outline btn-sm" onClick={() => applyPlatformToAllBlocks(p)}>
                    {p}
                  </button>
                ))}
              </div>
            </div>

            <div className="import-blocks">
              {importBlocks.map((block, bIdx) => {
                const sample = getBlockPreviewTitles(block);
                return (
                  <div className="import-block import-block-col" key={bIdx}>
                    <div className="import-block-row">
                      <span>
                        <b style={{ color: 'var(--ink)' }}>Blok {bIdx + 1}</b>
                        {block.headerLabel ? <span style={{ marginLeft: 6, fontSize: 12, fontWeight: 600, color: 'var(--ink)' }}>— Header: "{block.headerLabel}"</span> : null}
                        {block.sheetName ? <span style={{ marginLeft: 6, fontSize: 12, opacity: 0.75 }}>(Sheet "{block.sheetName}")</span> : null}
                        {block.platformName ? <span style={{ marginLeft: 6, fontSize: 12, color: '#0071e3', fontWeight: 600 }}> → {block.platformName}</span> : null}
                      </span>
                      <input
                        type="month"
                        value={importBlockMonths[bIdx] || ''}
                        onChange={(e) => updateBlockMonth(bIdx, e.target.value)}
                      />
                    </div>

                    {sample && (
                      <div style={{ fontSize: 11.5, color: 'var(--sub)', background: 'rgba(0,0,0,0.03)', padding: '4px 8px', borderRadius: 6, width: '100%' }}>
                        📝 <i>{sample}...</i>
                      </div>
                    )}

                    <div className="checkbox-row" style={{ marginTop: 2 }}>
                      <span style={{ fontSize: 12, opacity: 0.8 }}>Platform:</span>
                      {PLATFORMS.map((p) => (
                        <label key={p} className="checkbox-chip checkbox-chip-sm">
                          <input
                            type="radio"
                            name={`blockPlatform-${bIdx}`}
                            checked={importBlockPlatforms[bIdx] === p}
                            onChange={() => updateBlockPlatform(bIdx, p)}
                          />
                          {p}
                        </label>
                      ))}
                    </div>
                  </div>
                );
              })}
            </div>

            {importRows.length > 0 && (() => {
              const missingRows = importRows.filter((r) => !r.alreadyExists);
              return missingRows.length === 0 ? (
                <div className="import-verdict import-verdict-ok">
                  ✅ Semua {importRows.length} brief di CSV ini sudah cocok dengan Kalender Konten (Google Sheets) kamu.
                </div>
              ) : (
                <div className="import-verdict import-verdict-warn">
                  <p>
                    ⚠️ Ada {missingRows.length} dari {importRows.length} brief di CSV yang belum ketemu di Kalender
                    Konten (Google Sheets) kamu:
                  </p>
                  <ul>
                    {missingRows.map((r) => (
                      <li key={r.id}>
                        <b>{r.tglPosting || '(tanggal belum dipilih)'}</b> — {r.title}
                      </li>
                    ))}
                  </ul>
                </div>
              );
            })()}

            {importRows.length > 0 && (() => {
              const counts = {};
              importRows.forEach((r) => {
                const d = r.tglPosting || '(tanggal belum dipilih)';
                counts[d] = (counts[d] || 0) + 1;
              });
              const dateSummary = Object.entries(counts).sort(([a], [b]) => a.localeCompare(b));
              return (
                <div className="import-date-summary">
                  <button type="button" className="btn btn-ghost btn-sm" onClick={() => setShowDateSummary((v) => !v)}>
                    {showDateSummary ? 'Sembunyikan' : 'Lihat'} Ringkasan Jumlah Brief per Tanggal (semua platform)
                  </button>
                  {showDateSummary && (
                    <ul className="import-date-summary-list">
                      {dateSummary.map(([date, count]) => (
                        <li key={date}>
                          <span>{date}</span>
                          <b>{count} brief</b>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
              );
            })()}

            <div className="import-rows">
              {importRows.length === 0 ? (
                <div className="empty">Tidak ada judul konten yang terbaca dari file ini.</div>
              ) : (
                importRows.map((row) => (
                  <div className={`import-row${row.include ? '' : ' is-excluded'}`} key={row.id}>
                    <input
                      type="checkbox"
                      checked={row.include}
                      onChange={(e) => updateImportRow(row.id, { include: e.target.checked })}
                    />
                    <input
                      type="date"
                      value={row.tglPosting}
                      onChange={(e) => updateImportRow(row.id, { tglPosting: e.target.value })}
                    />
                    <div style={{ flex: 1, display: 'flex', alignItems: 'center', gap: 6, minWidth: 160 }}>
                      <input
                        type="text"
                        className="import-row-title"
                        style={{ flex: 1 }}
                        value={row.title}
                        onChange={(e) => updateImportRow(row.id, { title: e.target.value })}
                      />
                      {row.alreadyExists && (
                        <span style={{ fontSize: 11, background: 'rgba(255,149,0,0.12)', color: '#d97706', padding: '3px 7px', borderRadius: 6, fontWeight: 500, whiteSpace: 'nowrap' }} title="Brief ini sudah ada di dashboard sehingga otomatis tidak tercentang/dilewati">
                          ⚠️ Sudah Ada
                        </span>
                      )}
                    </div>
                    <select
                      value={row.pilar}
                      onChange={(e) => updateImportRow(row.id, { pilar: e.target.value })}
                    >
                      {PILARS.map((p) => (
                        <option key={p} value={p}>{p}</option>
                      ))}
                    </select>
                    <div className="checkbox-row">
                      {PLATFORMS.map((p) => (
                        <label key={p} className="checkbox-chip checkbox-chip-sm">
                          <input
                            type="checkbox"
                            checked={row.platform.includes(p)}
                            onChange={(e) => {
                              const next = e.target.checked
                                ? [...row.platform, p]
                                : row.platform.filter((x) => x !== p);
                              updateImportRow(row.id, { platform: next });
                            }}
                          />
                          {p}
                        </label>
                      ))}
                    </div>
                  </div>
                ))
              )}
            </div>

            {importMsg && <div className="msg">{importMsg}</div>}
            <div className="form-actions">
              <button type="button" className="btn btn-ghost" onClick={closeImportModal}>Batal</button>
              <button type="button" className="btn btn-primary" onClick={submitImportRows} disabled={importing}>
                {importing ? importProgress || 'Mengimpor…' : `Import ${importRows.filter((r) => r.include).length} Brief`}
              </button>
            </div>
          </div>
        </div>
      )}

      {selectedCalendarBrief && (
        <div className="modal-overlay" onClick={() => setSelectedCalendarBrief(null)}>
          <div className="modal-card" style={{ maxWidth: 440, padding: '24px 28px' }} onClick={(e) => e.stopPropagation()}>
            <div className="form-head" style={{ marginBottom: 16 }}>
              <h3 style={{ fontSize: 17, fontWeight: 700, margin: 0 }}>Detail Brief</h3>
              <button className="modal-close" onClick={() => setSelectedCalendarBrief(null)} title="Tutup">✕</button>
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
              <div>
                <div style={{ fontSize: 12, fontWeight: 600, color: 'var(--sub)', textTransform: 'uppercase', letterSpacing: '.03em', marginBottom: 4 }}>
                  Judul Brief
                </div>
                <div style={{ fontSize: 15, fontWeight: 600, color: 'var(--ink)', lineHeight: 1.45 }}>
                  {selectedCalendarBrief.brief.brief}
                </div>
              </div>

              <div style={{ display: 'flex', gap: 24 }}>
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 12, fontWeight: 600, color: 'var(--sub)', textTransform: 'uppercase', letterSpacing: '.03em', marginBottom: 4 }}>
                    Tanggal Brief Masuk
                  </div>
                  <div style={{ fontSize: 14, fontWeight: 500, color: 'var(--ink)' }}>
                    {selectedCalendarBrief.brief.tglMasuk || '-'}
                  </div>
                </div>

                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 12, fontWeight: 600, color: 'var(--sub)', textTransform: 'uppercase', letterSpacing: '.03em', marginBottom: 4 }}>
                    Tanggal Posting
                  </div>
                  <div style={{ fontSize: 14, fontWeight: 500, color: 'var(--ink)', display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: 4 }}>
                    <span>{selectedCalendarBrief.brief.tglPosting || '-'}</span>
                    {selectedCalendarBrief.brief && String(selectedCalendarBrief.brief.tglPosting || '').includes(today) && (
                      <span className="due-today-badge">🔥 Hari Ini</span>
                    )}
                  </div>
                </div>
              </div>

              <div>
                <div style={{ fontSize: 12, fontWeight: 600, color: 'var(--sub)', textTransform: 'uppercase', letterSpacing: '.03em', marginBottom: 6 }}>
                  Status
                </div>
                <div>
                  <span className={`pill ${pillClass(statusOf(selectedCalendarBrief.brief))}`} style={{ fontSize: 12.5, fontWeight: 600, padding: '4px 10px' }}>
                    {statusOf(selectedCalendarBrief.brief)}
                  </span>
                </div>
              </div>

              <div>
                <div style={{ fontSize: 12, fontWeight: 600, color: 'var(--sub)', textTransform: 'uppercase', letterSpacing: '.03em', marginBottom: 4 }}>
                  Hasil Final
                </div>
                <div style={{ fontSize: 13.5, color: 'var(--ink)', wordBreak: 'break-all' }}>
                  {selectedCalendarBrief.brief.hasilAkhir ? (
                    <a
                      href={selectedCalendarBrief.brief.hasilAkhir}
                      target="_blank"
                      rel="noopener noreferrer"
                      style={{ color: '#0071e3', textDecoration: 'none', fontWeight: 500, display: 'inline-flex', alignItems: 'center', gap: 4 }}
                    >
                      🎬 {selectedCalendarBrief.brief.hasilAkhir}
                    </a>
                  ) : (
                    <span style={{ color: 'var(--sub)' }}>-</span>
                  )}
                </div>
              </div>

              <div>
                <div style={{ fontSize: 12, fontWeight: 600, color: 'var(--sub)', textTransform: 'uppercase', letterSpacing: '.03em', marginBottom: 4 }}>
                  Referensi
                </div>
                <div style={{ fontSize: 13.5, color: 'var(--ink)', wordBreak: 'break-all' }}>
                  {selectedCalendarBrief.brief.referensi ? (
                    <a
                      href={selectedCalendarBrief.brief.referensi}
                      target="_blank"
                      rel="noopener noreferrer"
                      style={{ color: '#0071e3', textDecoration: 'none', fontWeight: 500, display: 'inline-flex', alignItems: 'center', gap: 4 }}
                    >
                      🔗 {selectedCalendarBrief.brief.referensi}
                    </a>
                  ) : (
                    <span style={{ color: 'var(--sub)' }}>-</span>
                  )}
                </div>
              </div>
            </div>

            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 24, paddingTop: 16, borderTop: '1px solid var(--hair)', flexWrap: 'wrap', gap: 10 }}>
              <button
                type="button"
                className="btn btn-outline"
                style={{ display: 'inline-flex', alignItems: 'center', gap: 6, padding: '8px 14px', fontSize: 13.5, fontWeight: 600, color: 'var(--red)', borderColor: 'rgba(255,59,48,0.35)' }}
                onClick={() => {
                  if (selectedCalendarBrief && selectedCalendarBrief.brief) {
                    handleDelete(selectedCalendarBrief.brief.id);
                  }
                }}
                title="Hapus Brief"
              >
                🗑️ Hapus Brief
              </button>

              <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                {selectedCalendarBrief.brief && (
                  <button
                    type="button"
                    className="btn btn-primary"
                    style={{ display: 'inline-flex', alignItems: 'center', gap: 6, padding: '8px 14px', fontSize: 13.5, fontWeight: 600 }}
                    onClick={() => {
                      triggerAutoDownloadAndPlatformRedirect(
                        selectedCalendarBrief.brief,
                        selectedCalendarBrief.platform
                      );
                    }}
                    title="Akses media & buka aplikasi platform"
                  >
                    {selectedCalendarBrief.brief.hasilAkhir && selectedCalendarBrief.brief.hasilAkhir.includes('/folders/')
                      ? '📂 Buka Folder Drive & App'
                      : '📥 Download File & App'}
                  </button>
                )}

                <button
                  type="button"
                  className="btn btn-outline"
                  style={{ display: 'inline-flex', alignItems: 'center', gap: 6, padding: '8px 14px', fontSize: 13.5, fontWeight: 600 }}
                  onClick={() => {
                    if (selectedCalendarBrief && selectedCalendarBrief.brief) {
                      const bId = selectedCalendarBrief.brief.id;
                      setSelectedCalendarBrief(null);
                      enterEditMode(bId);
                    }
                  }}
                  title="Edit Brief"
                >
                  ✏️ Edit Brief
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {statusModal && (
        <div className="modal-backdrop" onClick={() => setStatusModal(null)}>
          <div
            className="modal-box"
            style={{ maxWidth: 850, width: '92%', maxHeight: '85vh', display: 'flex', flexDirection: 'column' }}
            onClick={(e) => e.stopPropagation()}
          >
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', paddingBottom: 16, borderBottom: '1px solid var(--hair)' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                <h3 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: 'var(--ink)' }}>
                  Daftar Brief Status: {statusModal}
                </h3>
                <span className={`pill ${pillClass(statusModal)}`} style={{ fontSize: 12 }}>
                  {briefsForModal.length} Brief
                </span>
              </div>
              <button
                type="button"
                className="btn btn-ghost btn-sm"
                onClick={() => setStatusModal(null)}
                style={{ fontSize: 18, fontWeight: 700 }}
              >
                ✕
              </button>
            </div>

            <div style={{ overflowY: 'auto', flex: 1, paddingTop: 16, paddingBottom: 16 }}>
              {briefsForModal.length === 0 ? (
                <div className="empty">Tidak ada brief dengan status "{statusModal}".</div>
              ) : (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                  {briefsForModal.map((b) => (
                    <div
                      key={b.id}
                      style={{
                        background: 'var(--bg)',
                        border: '1px solid var(--hair)',
                        borderRadius: 12,
                        padding: '14px 18px',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'space-between',
                        gap: 16,
                        flexWrap: 'wrap',
                      }}
                    >
                      <div style={{ flex: '1 1 280px' }}>
                        <div style={{ fontWeight: 700, fontSize: 14.5, color: 'var(--ink)', marginBottom: 4 }}>
                          {b.brief}
                        </div>
                        <div style={{ display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
                          <span className="tag">{b.pilar}</span>
                          <span className="tag">{platformLabel(b)}</span>
                          <span style={{ fontSize: 12, color: 'var(--sub)', marginLeft: 4 }}>
                            🗓️ Posting: {b.tglPosting || 'Belum dijadwalkan'}
                          </span>
                        </div>
                      </div>

                      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        {b.hasilAkhir && (
                          <button
                            type="button"
                            className="btn btn-outline btn-sm"
                            onClick={() => triggerAutoDownloadAndPlatformRedirect(b, platformsOf(b)[0])}
                          >
                            {b.hasilAkhir.includes('/folders/') ? '📂 Drive' : '📥 Media'}
                          </button>
                        )}
                        <button
                          type="button"
                          className="btn btn-outline btn-sm"
                          onClick={() => {
                            setStatusModal(null);
                            enterEditMode(b.id);
                          }}
                        >
                          ✏️ Edit
                        </button>
                        <button
                          type="button"
                          className="btn btn-outline btn-sm"
                          style={{ color: 'var(--red)', borderColor: 'rgba(255,59,48,0.3)' }}
                          onClick={() => {
                            handleDelete(b.id);
                          }}
                        >
                          🗑️
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {toastMsg && (
        <div className="toast-notification">
          <span style={{ fontSize: 20 }}>🎉</span>
          <div>
            <div style={{ fontSize: 13.5, fontWeight: 700, color: 'var(--ink)' }}>Notifikasi Status</div>
            <div style={{ fontSize: 12.5, color: 'var(--sub)', marginTop: 2 }}>{toastMsg}</div>
          </div>
        </div>
      )}
    </div>
  );
}
