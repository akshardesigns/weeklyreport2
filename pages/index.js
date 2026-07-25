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
  tglPostingByPlatform: {},
  _prefillDate: '',
  isReference: false,
};

// Google Sheets (valueInputOption: USER_ENTERED) otomatis mengubah string "true"/"false"
// jadi boolean, dan saat dibaca lagi lewat API balik sebagai string "TRUE"/"FALSE" (kapital),
// bukan "true" huruf kecil. Makanya perbandingan strict === 'true' bisa gagal dan brief
// referensi ikut kehitung di Total Brief. Helper ini menangani semua variannya.
function isReferenceBrief(b) {
  const v = b && b.isReference;
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
  return pilar in IG_PILAR_ORDER ? IG_PILAR_ORDER[pilar] : 99;
}

// Label pilar khusus tampilan: pilar "Video" ditampilkan sebagai "Reels" kalau platform-nya
// Instagram, tapi tetap "Video" kalau Tiktok (satu nilai data, beda istilah per platform).
function pilarDisplayLabel(platform, pilar) {
  if (platform === 'Instagram' && pilar === 'Video') return 'Reels';
  return pilar;
}

// Warna badge platform di Kalender Konten — pakai warna khas brand masing-masing,
// nggak berubah walau format kontennya (Feed/Story/Carousel/Video) beda.
function formatColor(platform) {
  if (platform === 'Instagram') return 'linear-gradient(135deg, #833ab4, #e1306c)';
  if (platform === 'Tiktok') return '#000000';
  return '#8e8e93';
}

// platform disimpan di Sheets sebagai string dipisah koma ("Instagram,Tiktok")
// supaya satu brief bisa menyasar lebih dari satu platform sekaligus.
function platformsOf(b) {
  return (b.platform || '')
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);
}
function platformLabel(b) {
  const list = platformsOf(b);
  return list.length ? list.join(' + ') : '-';
}

// Pasangkan tiap platform dengan tanggal postingnya masing-masing.
// tglPosting disimpan sejajar urutan dengan platform (dipisah koma), mis.
// platform="Instagram,Tiktok" & tglPosting="2026-08-01,2026-08-03".
// Data lama yang cuma punya 1 tanggal untuk banyak platform tetap dianggap
// tanggal yang sama untuk semua platform (backward-compatible).
function platformDatePairs(b) {
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
  if (!b.tglSelesai) return null;
  const start = new Date(b.tglMasuk + 'T00:00:00');
  const end = new Date(b.tglSelesai + 'T00:00:00');
  const diffDays = Math.round((end - start) / 86400000);
  return diffDays <= maxDays(b.pilar) ? 'On Time' : 'Late';
}
function statusOf(b) {
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

// --- Helper Import CSV (transisi dari sheet kalender manual lama) ---
// Format sumbernya: grid mingguan (header Senin..Minggu), lalu baris angka tanggal,
// lalu baris judul konten, berulang per minggu. Ini best-effort karena formatnya
// ditulis manual (kadang ada anotasi nama nempel di angka tanggal, mis. "thomas3").
const CSV_DAY_HEADERS = ['senin', 'selasa', 'rabu', 'kamis', 'jumat', 'sabtu', 'minggu'];

function parseCsvToGrid(text) {
  const wb = XLSX.read(text, { type: 'string' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
}

// Pecah isi 1 sel jadi beberapa judul brief terpisah. Satu sel bisa berisi lebih
// dari 1 brief kalau user menekan Enter/Alt+Enter di dalam sel yang sama di Sheets
// (jadi ada line break literal di dalam cell saat diexport ke CSV).
// Perhatian: baris lanjutan yang diawali tanda kurung, mis. "(AUDIO DI POV CUSTOMER)",
// dianggap keterangan/anotasi dari judul di atasnya (bukan brief baru) dan digabung balik.
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

// Tebak platform dari label blok (baris judul tab di atas header Senin..Minggu,
// mis. "IG JUL" atau "TIKTOK JUL"). Kosong kalau tidak match apa-apa (biar user isi manual).
function guessPlatformFromLabel(label) {
  const l = String(label || '').toLowerCase();
  if (/\btiktok\b|\btt\b/.test(l)) return 'Tiktok';
  if (/\big\b|instagram/.test(l)) return 'Instagram';
  return '';
}

function detectContentPlanBlocks(grid) {
  const blocks = [];
  let i = 0;
  while (i < grid.length) {
    const row = (grid[i] || []).slice(0, 7).map((c) => String(c || '').trim().toLowerCase());
    const isHeader = CSV_DAY_HEADERS.every((d, idx) => row[idx] === d);
    if (isHeader) {
      // Cari label tab di baris-baris terdekat di atas header "Senin..Minggu"
      // (mis. "IG JUL" / "TIKTOK JUL"), dipakai untuk menebak platform otomatis.
      // Nggak cuma cek 1 baris tepat di atasnya — kadang ada baris kosong
      // pemisah, jadi kita cari sampai 4 baris ke atas.
      let label = '';
      for (let offset = 1; offset <= 4; offset++) {
        const candidateRow = grid[i - offset] || [];
        const text = candidateRow.map((c) => String(c || '').trim()).filter(Boolean).join(' ');
        if (text) {
          label = text;
          break;
        }
      }
      const weeks = [];
      let cursor = i + 1;
      while (cursor + 1 < grid.length) {
        const dateRow = grid[cursor] || [];
        const hasDateLike = dateRow
          .slice(0, 7)
          .some((c) => /\d{1,2}\s*$/.test(String(c || '').trim()) && String(c || '').trim() !== '');
        if (!hasDateLike) break;
        const titleRow = grid[cursor + 1] || [];
        const days = [];
        for (let col = 0; col < 7; col++) {
          const rawDateCell = String(dateRow[col] || '');
          // Ambil baris terakhir dari sel tanggal untuk cari angka tanggalnya
          // (biar tetap kebaca walau sel tanggal juga sempat ditumpuk multi-baris).
          const dateLines = splitCellTitles(rawDateCell);
          const lastLine = dateLines[dateLines.length - 1] || '';
          const m = lastLine.match(/(\d{1,2})\s*$/);
          if (!m) {
            days.push(null);
            continue;
          }
          const dayNum = parseInt(m[1], 10);
          // Sisa teks di baris terakhir (sebelum angka tanggal) — kalau ada,
          // dihitung sebagai 1 brief tambahan dari data lama yang belum dirapikan.
          // source: 'date' = brief singkat yang nempel di baris tanggal -> pilar Story.
          const leftoverOnDateLine = lastLine.replace(/(\d{1,2})\s*$/, '').trim();
          const titlesFromDateRow = [...dateLines.slice(0, -1), leftoverOnDateLine]
            .filter(Boolean)
            .map((text) => ({ text, source: 'date' }));
          // source: 'title' = brief utama di baris judul -> pilar Video.
          const titlesFromTitleRow = splitCellTitles(titleRow[col]).map((text) => ({ text, source: 'title' }));
          const titles = [...titlesFromDateRow, ...titlesFromTitleRow];
          days.push({ dayNum, titles });
        }
        weeks.push(days);
        cursor += 2;
      }
      blocks.push({ headerRow: i, label, platform: guessPlatformFromLabel(label), weeks });
      i = cursor > i ? cursor : i + 1;
    } else {
      i++;
    }
  }
  return blocks;
}

function blocksToImportRows(blocks, blockMonths, existingBriefs = []) {
  const existingSet = new Set(existingBriefs.map((b) => (b.brief || '').trim().toLowerCase()));
  const rows = [];
  let rid = 0;
  blocks.forEach((block, bIdx) => {
    const my = blockMonths[bIdx] || '';
    block.weeks.forEach((week) => {
      week.forEach((cell) => {
        if (!cell || !cell.titles || cell.titles.length === 0) return;
        let tglPosting = '';
        if (my) {
          const [y, m] = my.split('-');
          tglPosting = `${y}-${m}-${String(cell.dayNum).padStart(2, '0')}`;
        }
        cell.titles.forEach(({ text: title, source }) => {
          const alreadyExists = existingSet.has(title.trim().toLowerCase());
          rows.push({
            id: `imp-${rid++}`,
            blockIdx: bIdx,
            dayNum: cell.dayNum,
            title,
            tglPosting,
            include: !alreadyExists,
            alreadyExists,
            // Brief yang nempel di baris tanggal (mis. "UKV Periksa Visa Dulu (o) 8") -> Story.
            // Brief di baris judul utama -> Video. Bisa diubah manual di form sebelum import.
            pilar: source === 'date' ? 'Story' : 'Video',
            platform: block.platform ? [block.platform] : [],
          });
        });
      });
    });
  });
  return rows;
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
      closeForm();
      await loadBriefs();
    } catch (err) {
      setFormMsg(err.message);
    } finally {
      setSaving(false);
    }
  }

  async function handleDelete(id) {
    if (!confirm('Hapus brief ini?')) return;
    try {
      const res = await fetch(`/api/briefs/${id}`, { method: 'DELETE' });
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.error || 'Gagal menghapus brief');
      }
      if (editingId === id) closeForm();
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
    Ads: '#0071e3',
    Feed: '#af52de',
    Story: '#ff375f',
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
        const platformDiff = PLATFORM_ORDER_INDEX(a.platform) - PLATFORM_ORDER_INDEX(b.platform);
        if (platformDiff !== 0) return platformDiff;
        if (a.platform === 'Instagram') {
          return IG_PILAR_ORDER_INDEX(a.brief.pilar) - IG_PILAR_ORDER_INDEX(b.brief.pilar);
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
  const today = todayISO();

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

  function handleImportFileChange(e) {
    const file = e.target.files && e.target.files[0];
    e.target.value = '';
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const grid = parseCsvToGrid(String(reader.result));
        const blocks = detectContentPlanBlocks(grid);
        if (blocks.length === 0) {
          setImportMsg('Tidak menemukan pola grid kalender (header Senin–Minggu) di file ini.');
          return;
        }
        const months = blocks.map(() => '');
        setImportBlocks(blocks);
        setImportBlockMonths(months);
        setImportRows(blocksToImportRows(blocks, months, briefs));
        setImportMsg('');
        setImportAsReference(true);
        setImportOpen(true);
      } catch (err) {
        setImportMsg('Gagal membaca file CSV: ' + err.message);
      }
    };
    reader.readAsText(file);
  }

  function updateBlockMonth(bIdx, value) {
    const months = importBlockMonths.slice();
    months[bIdx] = value;
    setImportBlockMonths(months);
    setImportRows(blocksToImportRows(importBlocks, months, briefs));
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
              Import CSV
            </button>
            <input
              type="file"
              accept=".csv"
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
              return (
                <div
                  key={cell.iso}
                  className={`calendar-cell${cell.inMonth ? '' : ' is-outside'}${cell.iso === today ? ' is-today' : ''}`}
                  onDoubleClick={() => openAddForm({ _prefillDate: cell.iso })}
                  title="Klik dua kali untuk tambah brief di tanggal ini"
                >
                  <div className="calendar-cell-date">{cell.date.getDate()}</div>
                  <div className="calendar-cell-items">
                    {(isOpen ? items : items.slice(0, 3)).map((it, idx) => {
                      const b = it.brief;
                      const style = CALENDAR_STATUS_STYLE[statusOf(b)] || CALENDAR_STATUS_STYLE['Belum Dikerjakan'];
                      const label = it.platform ? platformAbbr(it.platform) : platformLabel(b);
                      const badgeColor = formatColor(it.platform);
                      const isRef = isReferenceBrief(b);
                      return (
                        <div
                          key={`${b.id}-${it.platform}-${idx}`}
                          className={`calendar-item${isRef ? ' is-reference' : ''}`}
                          style={{ background: style.bg, borderLeftColor: style.border }}
                          onClick={(e) => {
                            e.stopPropagation();
                            enterEditMode(b.id);
                          }}
                          title={`${isRef ? '[Transisi/Referensi] ' : ''}${it.platform || platformLabel(b)} · ${pilarDisplayLabel(it.platform, b.pilar)} · ${b.brief} — ${statusOf(b)}`}
                        >
                          <span className="calendar-item-badge" style={{ background: badgeColor }}>
                            {isRef ? '📌 ' : ''}{label} · {pilarDisplayLabel(it.platform, b.pilar)}
                          </span>
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
                  <div className="unscheduled-item" key={b.id} onClick={() => enterEditMode(b.id)}>
                    {platformsOf(b).map((p) => (
                      <span key={p} className="dot" style={{ background: PLATFORM_COLORS[p] || 'var(--grey)' }} />
                    ))}
                    <span className="unscheduled-brief">{b.brief}</span>
                    <span className="tag">{b.pilar}</span>
                    <span className="tag">{platformLabel(b)}</span>
                    <span className="unscheduled-hint">Klik untuk atur tanggal posting</span>
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
              const clickable = c.label === 'Selesai Terupload';
              return (
                <div
                  className={`kpi${clickable ? ' clickable' : ''}`}
                  key={c.label}
                  onClick={
                    clickable
                      ? () => document.getElementById('daily-panel')?.scrollIntoView({ behavior: 'smooth', block: 'start' })
                      : undefined
                  }
                  title={clickable ? 'Lihat rincian per hari' : undefined}
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
                  return (
                    <div className="bar-row" key={s.name}>
                      <div className="name">{s.name}</div>
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

          <div className="list-panel">
            <div className="list-head">
              <h3>Daftar Brief</h3>
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
                    <th>Sumber/Referensi</th>
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
                              Buka ↗
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
                  <label htmlFor="hasilAkhir">Sumber/Referensi (link atau file)</label>
                  <div className="hasil-akhir-row">
                    <input
                      type="text"
                      id="hasilAkhir"
                      placeholder="Tempel link (Drive, Canva, dll)…"
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

            <div className="import-blocks">
              {importBlocks.map((block, bIdx) => (
                <div className="import-block" key={bIdx}>
                  <span>
                    Blok {bIdx + 1} ({block.weeks.length} minggu){block.label ? ` — "${block.label}"` : ''}
                    {block.platform ? ` → ${block.platform}` : ' → platform tidak terdeteksi, isi manual per baris'}
                  </span>
                  <input
                    type="month"
                    value={importBlockMonths[bIdx] || ''}
                    onChange={(e) => updateBlockMonth(bIdx, e.target.value)}
                  />
                </div>
              ))}
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
    </div>
  );
}
