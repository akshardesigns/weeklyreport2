import { google } from 'googleapis';

// Auth bersama untuk Google Sheets & Google Drive (dipakai fitur upload file "Hasil Akhir")
export function getAuth() {
  const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  const rawKey = process.env.GOOGLE_PRIVATE_KEY;

  if (!email || !rawKey) {
    throw new Error(
      'Konfigurasi belum lengkap. Pastikan GOOGLE_SERVICE_ACCOUNT_EMAIL dan GOOGLE_PRIVATE_KEY sudah diisi di .env.local'
    );
  }

  // private key dari .env biasanya "\n" literal, perlu diubah jadi newline asli
  const privateKey = rawKey.replace(/\\n/g, '\n');

  return new google.auth.JWT(email, undefined, privateKey, [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file',
  ]);
}
