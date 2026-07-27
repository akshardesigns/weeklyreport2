export default async function handler(req, res) {
  const { url, filename } = req.query;

  if (!url) {
    return res.status(400).json({ error: 'Missing url parameter' });
  }

  try {
    const response = await fetch(url);
    if (!response.ok) {
      return res.redirect(302, url);
    }

    const contentType = response.headers.get('content-type') || 'application/octet-stream';
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
