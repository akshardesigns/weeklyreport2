import { updateBrief, deleteBrief } from '../../../lib/googleSheets';

export default async function handler(req, res) {
  const { id } = req.query;

  try {
    if (req.method === 'PUT') {
      const { tglMasuk, pilar, platform, brief, status, tglSelesai, hasilAkhir, tglPosting, deskripsiBrief } = req.body || {};
      if (!tglMasuk || !brief) {
        return res.status(400).json({ error: 'tglMasuk dan brief wajib diisi' });
      }
      const updated = await updateBrief(id, { tglMasuk, pilar, platform, brief, status, tglSelesai, hasilAkhir, tglPosting, deskripsiBrief });
      return res.status(200).json(updated);
    }

    if (req.method === 'DELETE') {
      const ok = await deleteBrief(id);
      return res.status(ok ? 200 : 404).json({ deleted: ok });
    }

    res.setHeader('Allow', ['PUT', 'DELETE']);
    return res.status(405).json({ error: `Method ${req.method} tidak diizinkan` });
  } catch (err) {
    console.error('API /briefs/[id] error:', err);
    return res.status(500).json({ error: err.message || 'Terjadi kesalahan pada server' });
  }
}
