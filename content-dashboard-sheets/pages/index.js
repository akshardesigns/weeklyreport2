import { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';

const PILARS = ['Ads', 'Feed', 'Story', 'Carousel', 'Video', 'Lainnya'];
const PLATFORMS = ['Instagram', 'Tiktok', 'Non sosmed'];
const STATUS_OPTIONS = ['On Going', 'Waiting Approval', 'Selesai Terupload'];

const EMPTY_FORM = {
  tglMasuk: '',
  pilar: 'Ads',
  platform: 'Instagram',
  brief: '',
  status: '',
  tglSelesai: '',
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

// --- Helper minggu (Senin - Minggu) ---
function startOfWeek(date) {
  const d = new Date(date);
  const day = d.getDay(); // 0=Minggu
  const diff = day === 0 ? -6 : 1 - day; // geser ke Senin
  d.setDate(d.getDate() + diff);
  d.setHours(0, 0, 0, 0);
  return d;
}
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
function formatWeekLabel(weekStart) {
  const weekEnd = addDays(weekStart, 6);
  const sameMonth = weekStart.getMonth() === weekEnd.getMonth();
  const optsShort = { day: '2-digit' };
  const optsFull = { day: '2-digit', month: 'short', year: 'numeric' };
  const startLabel = weekStart.toLocaleDateString('id-ID', sameMonth ? optsShort : optsFull);
  const endLabel = weekEnd.toLocaleDateString('id-ID', optsFull);
  return `${startLabel} - ${endLabel}`;
}
function isInWeek(isoDateStr, weekStart) {
  if (!isoDateStr) return false;
  const d = new Date(isoDateStr + 'T00:00:00');
  const weekEnd = addDays(weekStart, 7);
  return d >= weekStart && d < weekEnd;
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

  // null = tampilkan semua data. Selain itu: Date (Senin) minggu yang dipilih.
  const [weekFilter, setWeekFilter] = useState(null);

  // filter khusus tabel "Daftar Brief"
  const [filterPilar, setFilterPilar] = useState('');
  const [filterPlatform, setFilterPlatform] = useState('');
  const [filterStatus, setFilterStatus] = useState('');

  // tanggal yang sedang di-expand di panel "Brief Selesai per Hari"
  const [expandedDate, setExpandedDate] = useState(null);

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
    if (!weekFilter) return briefs;
    return briefs.filter((b) => isInWeek(b.tglMasuk, weekFilter));
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
      if (filterPlatform && b.platform !== filterPlatform) return false;
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

  function openAddForm() {
    setEditingId(null);
    setForm(EMPTY_FORM);
    setFormMsg('');
    setFormOpen(true);
  }
  function enterEditMode(id) {
    const b = briefs.find((x) => x.id === id);
    if (!b) return;
    setEditingId(id);
    setForm({
      tglMasuk: b.tglMasuk,
      pilar: b.pilar,
      platform: b.platform,
      brief: b.brief,
      status: b.status,
      tglSelesai: b.tglSelesai || '',
    });
    setFormMsg('');
    setFormOpen(true);
  }
  function closeForm() {
    setFormOpen(false);
    setEditingId(null);
    setForm(EMPTY_FORM);
    setFormMsg('');
  }

  async function handleSubmit(e) {
    e.preventDefault();
    if (!form.tglMasuk || !form.brief.trim()) {
      setFormMsg('Tanggal masuk dan nama brief wajib diisi.');
      return;
    }
    if (form.tglSelesai && form.tglSelesai < form.tglMasuk) {
      setFormMsg('Tanggal selesai tidak boleh sebelum tanggal masuk.');
      return;
    }
    setFormMsg('');
    setSaving(true);
    try {
      if (editingId) {
        const res = await fetch(`/api/briefs/${editingId}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(form),
        });
        if (!res.ok) {
          const data = await res.json().catch(() => ({}));
          throw new Error(data.error || 'Gagal menyimpan perubahan');
        }
      } else {
        const res = await fetch('/api/briefs', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(form),
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
        Platform: b.platform,
        Brief: b.brief,
        Status: b.status,
        'Tanggal Selesai': b.tglSelesai,
        KPI: kpiFor(b) || '',
      }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = [{ wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 40 }, { wch: 18 }, { wch: 14 }, { wch: 10 }];
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
    Video: '#ff9500',
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

  return (
    <div className="wrap">
      <header>
        <div>
          <h1>Content Production</h1>
          <p>Target: Ads &amp; Carousel maks. 2 hari · Video &amp; Lainnya maks. 3 hari</p>
        </div>
        <div className="header-actions">
          <button className="btn btn-outline" onClick={loadBriefs}>Refresh</button>
          <button className="btn btn-primary" onClick={exportToExcel}>Export ke Excel</button>
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
      ) : (
        <>
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
                            <span className="tag">{b.platform}</span>
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
              <div className="week-filter">
                <button
                  className={`week-toggle${weekFilter ? '' : ' is-off'}`}
                  onClick={() => setWeekFilter(weekFilter ? null : startOfWeek(new Date()))}
                >
                  {weekFilter ? 'Filter: Mingguan' : 'Semua Minggu'}
                </button>
                {weekFilter && (
                  <div className="week-nav">
                    <button onClick={() => setWeekFilter(addDays(weekFilter, -7))} title="Minggu sebelumnya">‹</button>
                    <span className="week-label">{formatWeekLabel(weekFilter)}</span>
                    <button onClick={() => setWeekFilter(addDays(weekFilter, 7))} title="Minggu berikutnya">›</button>
                  </div>
                )}
                {weekFilter && toISO(weekFilter) !== toISO(startOfWeek(new Date())) && (
                  <button className="btn btn-ghost btn-sm" onClick={() => setWeekFilter(startOfWeek(new Date()))}>
                    Minggu Ini
                  </button>
                )}
              </div>
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
                    <th>Brief</th>
                    <th>Status</th>
                    <th>Tanggal Selesai</th>
                    <th>KPI</th>
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
                        <td>{b.platform}</td>
                        <td>{b.brief}</td>
                        <td><span className={`pill ${pillClass(statusOf(b))}`}>{statusOf(b)}</span></td>
                        <td>{fmtDate(b.tglSelesai)}</td>
                        <td>
                          {k === 'On Time' && <span className="kpi-ok">✓ On Time</span>}
                          {k === 'Late' && <span className="kpi-late">✕ Late</span>}
                          {k === null && <span className="kpi-none">Belum</span>}
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
                <div className="field">
                  <label htmlFor="platform">Platform</label>
                  <select id="platform" value={form.platform} onChange={(e) => setForm({ ...form, platform: e.target.value })}>
                    {PLATFORMS.map((p) => (
                      <option key={p} value={p}>{p}</option>
                    ))}
                  </select>
                </div>
                <div className="field span3">
                  <label htmlFor="briefText">Brief</label>
                  <input
                    type="text"
                    id="briefText"
                    placeholder="Nama / deskripsi brief"
                    required
                    value={form.brief}
                    onChange={(e) => setForm({ ...form, brief: e.target.value })}
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
    </div>
  );
}
