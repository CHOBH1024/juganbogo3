export default async function handler(req, res) {
  const { code, error } = req.query;
  const clientId = process.env.GOOGLE_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
  const appUrl = process.env.APP_URL || `https://${req.headers.host}`;
  const callbackUrl = `${appUrl}/api/google-auth/callback`;

  if (error || !code) {
    return res.status(400).send(`<h2>인증 실패: ${error || '코드 없음'}</h2>`);
  }

  try {
    const tokenRes = await fetch('https://oauth2.googleapis.com/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        code,
        client_id: clientId,
        client_secret: clientSecret,
        redirect_uri: callbackUrl,
        grant_type: 'authorization_code',
      }).toString(),
    });

    const tokens = await tokenRes.json();
    const refreshToken = tokens.refresh_token;

    res.send(`<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><title>연동 완료</title>
      <style>body{font-family:system-ui,sans-serif;max-width:680px;margin:60px auto;padding:24px;color:#1e293b;line-height:1.6}
      h2{color:#16a34a}.card{border-radius:12px;padding:20px;margin:16px 0}
      .green{background:#f0fdf4;border:1px solid #bbf7d0}.blue{background:#eff6ff;border:1px solid #bfdbfe}
      .step{display:flex;gap:12px;margin:10px 0}
      .num{width:24px;height:24px;border-radius:50%;background:#6366f1;color:#fff;font-size:11px;font-weight:900;display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:2px}
      code{background:#f1f5f9;padding:4px 10px;border-radius:6px;font-family:monospace;font-size:13px;word-break:break-all;display:block;margin-top:6px}
      button{background:#6366f1;color:#fff;border:none;padding:10px 20px;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer;margin-top:10px}
      p{margin:6px 0;font-size:14px}</style></head><body>
      <h2>✅ 구글 계정 인증 완료!</h2>
      <p style="color:#64748b">아래 Refresh Token을 Vercel 환경변수에 등록하세요.</p>
      <div class="card green"><p><strong>📋 Refresh Token</strong></p>
        ${refreshToken
          ? `<code id="tok">${refreshToken}</code>
             <button onclick="navigator.clipboard.writeText(document.getElementById('tok').textContent).then(()=>this.textContent='✅ 복사됨!')">클립보드에 복사</button>`
          : `<p style="color:#92400e">⚠️ 토큰 미발급 — <a href="/api/google-auth">재인증</a>하세요.</p>
             <p style="font-size:12px;color:#64748b">응답: ${JSON.stringify(tokens)}</p>`}
      </div>
      <div class="card blue"><p><strong>📌 등록 방법</strong></p>
        <div class="step"><div class="num">1</div><p>Vercel → juganbogo3 → <b>Settings → Environment Variables</b></p></div>
        <div class="step"><div class="num">2</div><p>Key: <b>GOOGLE_REFRESH_TOKEN</b> / Value: 위 토큰 → Add</p></div>
        <div class="step"><div class="num">3</div><p><b>Deployments</b> → 최신 배포 ··· → <b>Redeploy</b></p></div>
      </div></body></html>`);
  } catch (err) {
    res.status(500).send(`<h2>오류 발생</h2><pre>${err.message}</pre>`);
  }
}
