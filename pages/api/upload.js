import { uploadFileToDrive } from '../../lib/googleDrive';

export const config = {
  api: {
    bodyParser: {
      sizeLimit: '15mb',
    },
  },
};

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.setHeader('Allow', ['POST']);
    return res.status(405).json({ error: `Method ${req.method} tidak diizinkan` });
  }

  try {
    const { filename, mimeType, base64Data } = req.body || {};
    if (!filename || !base64Data) {
      return res.status(400).json({ error: 'File tidak valid' });
    }
    const uploaded = await uploadFileToDrive({ filename, mimeType, base64Data });
    return res.status(201).json(uploaded);
  } catch (err) {
    console.error('API /upload error:', err);
    const msg = /invalid_grant|Drive API/i.test(err.message)
      ? err.message
      : err.message || 'Gagal mengunggah file';
    return res.status(500).json({ error: msg });
  }
}
