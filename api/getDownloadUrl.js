export default async function handler(req, res) {
  // Always allow CORS for this API (trusted proxy)
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
      return res.status(500).json({ error: 'Server not configured. Set TENANT_ID, CLIENT_ID, and CLIENT_SECRET environment variables.' });
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

    let tokenJson;
    try {
      tokenJson = await tokenRes.json();
    } catch (e) {
      return res.status(500).json({ error: 'Failed to parse token response', status: tokenRes.status, bodyText: await tokenRes.text() });
    }

    if (!tokenJson || !tokenJson.access_token) {
      return res.status(500).json({ error: 'Failed to obtain access token', status: tokenRes.status, details: tokenJson });
    }

    const accessToken = tokenJson.access_token;

    // Call Graph to get the driveItem from the share link
    const driveRes = await fetch(`https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem`, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });

    let driveJson;
    try {
      // Graph normally returns JSON here
      driveJson = await driveRes.json();
    } catch (e) {
      return res.status(500).json({ error: 'Failed to parse driveItem response', status: driveRes.status, bodyText: await driveRes.text() });
    }

    if (!driveRes.ok) {
      return res.status(driveRes.status).json({ error: 'Graph /shares request failed', status: driveRes.status, body: driveJson });
    }

    const downloadUrl = driveJson['@microsoft.graph.downloadUrl'];
    if (!downloadUrl) return res.status(500).json({ error: 'No downloadUrl returned', details: driveJson });

    // If debug is requested, return diagnostics (redact access_token)
    if (req.query.debug === '1' || req.query.debug === 'true') {
      const redactedToken = Object.assign({}, tokenJson);
      if (redactedToken.access_token) redactedToken.access_token = 'REDACTED';
      return res.status(200).json({ ok: true, shareId, token: redactedToken, driveItem: driveJson });
    }

    // Fetch the PDF from the ephemeral download URL and stream bytes back
    const pdfRes = await fetch(downloadUrl);
    if (!pdfRes.ok) {
      const body = await pdfRes.text();
      return res.status(pdfRes.status).json({ error: 'Failed to fetch PDF from downloadUrl', status: pdfRes.status, body });
    }

    const arrayBuf = await pdfRes.arrayBuffer();
    const buffer = Buffer.from(arrayBuf);

    res.setHeader('Content-Type', 'application/pdf');
    // Prevent browser caching so new uploads show immediately
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');
    res.setHeader('Access-Control-Allow-Origin', '*');

    return res.status(200).send(buffer);
  } catch (err) {
    console.error('API error', err);
    // Ensure CORS header on error responses
    res.setHeader('Access-Control-Allow-Origin', '*');
    return res.status(500).json({ error: 'Server error', message: err.message });
  }
}
