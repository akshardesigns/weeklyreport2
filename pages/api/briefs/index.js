import { getAllBriefs, addBrief } from '../../../lib/googleSheets';

export default async function handler(req, res) {
  try {
    if (req.method === 'GET') {
      const briefs = await getAllBriefs();
      return res.status(200).json(briefs);
    }

    if (req.method === 'POST') {
      const { tglMasuk, pilar, platform, brief, status, tglSelesai, hasilAkhir } = req.body || {};
      if (!tglMasuk || !brief) {
        return res.status(400).json({ error: 'tglMasuk dan brief wajib diisi' });
      }
      const created = await addBrief({ tglMasuk, pilar, platform, brief, status, tglSelesai, hasilAkhir });
      return res.status(201).json(created);
    }

    res.setHeader('Allow', ['GET', 'POST']);
    return res.status(405).json({ error: `Method ${req.method} tidak diizinkan` });
  } catch (err) {
    console.error('API /briefs error:', err);
    return res.status(500).json({ error: err.message || 'Terjadi kesalahan pada server' });
  }
}
