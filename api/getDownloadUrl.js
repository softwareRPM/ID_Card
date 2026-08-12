export default async function handler(req, res) {
  // Add CORS headers for every response
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    const url = (req.query && (req.query.url || req.query.u)) || (req.url && new URL(req.url, `http://${req.headers.host}`).searchParams.get('url'));
    if (!url) return res.status(400).send('Missing url query parameter (url=shareLink)');

    const shareLink = decodeURIComponent(url);

    // Build shareId required by Microsoft Graph for shared links
    const rawB64 = Buffer.from(shareLink, 'utf8').toString('base64');
    const safeB64 = rawB64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
    const shareId = 'u!' + safeB64;

    const tenant = process.env.TENANT_ID;
    const clientId = process.env.CLIENT_ID;
    const clientSecret = process.env.CLIENT_SECRET;

    if (!tenant || !clientId || !clientSecret) {
      return res.status(500).send('Server not configured. Set TENANT_ID, CLIENT_ID, and CLIENT_SECRET environment variables.');
    }

    // Acquire token (client credentials)
    const params = new URLSearchParams();
    params.append('client_id', clientId);
    params.append('client_secret', clientSecret);
    params.append('scope', 'https://graph.microsoft.com/.default');
    params.append('grant_type', 'client_credentials');

    const tokenRes = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params.toString(),
    });

    const tokenJson = await tokenRes.json();
    if (!tokenJson.access_token) {
      return res.status(500).json({ error: 'Failed to obtain access token', details: tokenJson });
    }
    const accessToken = tokenJson.access_token;

    // Call Graph to get the driveItem from the share link
    const driveRes = await fetch(`https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem`, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });

    if (!driveRes.ok) {
      const txt = await driveRes.text();
      return res.status(driveRes.status).send(txt);
    }

    const driveJson = await driveRes.json();
    const downloadUrl = driveJson['@microsoft.graph.downloadUrl'];
    if (!downloadUrl) return res.status(500).json({ error: 'No downloadUrl returned', details: driveJson });

    // Fetch the PDF and stream it back with CORS headers
    const pdfRes = await fetch(downloadUrl);
    if (!pdfRes.ok) return res.status(pdfRes.status).send('Failed to fetch PDF');

    const arrayBuf = await pdfRes.arrayBuffer();
    const buffer = Buffer.from(arrayBuf);

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');
    res.setHeader('Access-Control-Allow-Origin', '*');

    return res.status(200).send(buffer);
  } catch (err) {
    console.error(err);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.status(500).json({ error: 'Server error', message: err.message });
  }
}
