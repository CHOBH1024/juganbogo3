export default function handler(req, res) {
  const clientId = process.env.GOOGLE_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
  const appUrl = process.env.APP_URL || `https://${req.headers.host}`;
  const callbackUrl = `${appUrl}/api/google-auth/callback`;

  if (!clientId || !clientSecret) {
    return res.status(400).send(`<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">
      <title>설정 필요</title><style>body{font-family:sans-serif;max-width:600px;margin:60px auto;padding:24px}
      code{background:#f1f5f9;padding:3px 8px;border-radius:4px}</style></head><body>
      <h2>⚠️ 환경변수 미설정</h2>
      <p>Vercel → Settings → Environment Variables에 아래 두 값을 등록하고 Redeploy 하세요.</p>
      <p><code>GOOGLE_CLIENT_ID</code></p><p><code>GOOGLE_CLIENT_SECRET</code></p>
      </body></html>`);
  }

  const params = new URLSearchParams({
    client_id: clientId,
    redirect_uri: callbackUrl,
    response_type: 'code',
    scope: 'https://www.googleapis.com/auth/drive.file',
    access_type: 'offline',
    prompt: 'consent',
  });

  res.redirect(`https://accounts.google.com/o/oauth2/v2/auth?${params.toString()}`);
}
