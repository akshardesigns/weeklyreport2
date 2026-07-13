import { google } from 'googleapis';
import { Readable } from 'stream';
import { getAuth } from './googleAuth';

const FOLDER_ID = process.env.GOOGLE_DRIVE_FOLDER_ID || undefined;

function getDriveClient() {
  return google.drive({ version: 'v3', auth: getAuth() });
}

// Upload file (base64) ke Google Drive milik service account, lalu buka aksesnya
// jadi "siapa saja yang punya link bisa lihat" supaya link bisa dibuka dari dashboard.
export async function uploadFileToDrive({ filename, mimeType, base64Data }) {
  if (!filename || !base64Data) {
    throw new Error('File dan nama file wajib diisi');
  }
  const drive = getDriveClient();
  const buffer = Buffer.from(base64Data, 'base64');

  const file = await drive.files.create({
    requestBody: {
      name: filename,
      parents: FOLDER_ID ? [FOLDER_ID] : undefined,
    },
    media: {
      mimeType: mimeType || 'application/octet-stream',
      body: Readable.from(buffer),
    },
    fields: 'id, webViewLink, webContentLink',
  });

  await drive.permissions.create({
    fileId: file.data.id,
    requestBody: { role: 'reader', type: 'anyone' },
  });

  const meta = await drive.files.get({
    fileId: file.data.id,
    fields: 'webViewLink',
  });

  return { id: file.data.id, url: meta.data.webViewLink };
}
