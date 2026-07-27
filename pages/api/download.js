export default async function handler(req, res) {
  let { url, filename } = req.query;

  if (!url) {
    return res.status(400).json({ error: 'Missing url parameter' });
  }

  // Convert Google Drive view link to direct download link
  if (url.includes('drive.google.com')) {
    const match = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) || url.match(/id=([a-zA-Z0-9_-]+)/);
    if (match && match[1]) {
      url = `https://drive.google.com/uc?export=download&id=${match[1]}`;
    }
  }

  try {
    const response = await fetch(url, { redirect: 'follow' });
    if (!response.ok) {
      return res.redirect(302, url);
    }

    const contentType = response.headers.get('content-type') || 'application/octet-stream';

    // If response is HTML (e.g. Google Drive preview or login page), fallback to direct URL
    if (contentType.includes('text/html')) {
      return res.redirect(302, url);
    }

    let ext = '';
    if (contentType.includes('video/mp4')) ext = '.mp4';
    else if (contentType.includes('video/quicktime')) ext = '.mov';
    else if (contentType.includes('image/png')) ext = '.png';
    else if (contentType.includes('image/jpeg')) ext = '.jpg';

    const baseName = (filename || 'media_asset').replace(/[^a-z0-9_-]/gi, '_');
    const safeName = ext && !baseName.endsWith(ext) ? `${baseName}${ext}` : baseName;

    res.setHeader('Content-Type', contentType);
    res.setHeader('Content-Disposition', `attachment; filename="${safeName}"`);

    const arrayBuffer = await response.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    return res.send(buffer);
  } catch (err) {
    console.error('Download proxy error:', err);
    return res.redirect(302, url);
  }
}
