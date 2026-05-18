import { google } from 'googleapis';

export default async function handler(req, res) {
  const { code, error } = req.query;
  const clientId = process.env.GOOGLE_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
  const appUrl = process.env.APP_URL || `https://${req.headers.host}`;
  const callbackUrl = `${appUrl}/api/google-auth/callback`;

  if (error || !code) {
    return res.status(400).send(`<h2>인증 실패: ${error || '코드 없음'}</h2>`);
  }

  const oauth2Client = new google.auth.OAuth2(clientId, clientSecret, callbackUrl);

  try {
    const { tokens } = await oauth2Client.getToken(code);
    const refreshToken = tokens.refresh_token;

    res.send(`<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">
      <title>Google Drive 연동 완료</title>
      <style>
        *{box-sizing:border-box}
        body{font-family:system-ui,sans-serif;max-width:680px;margin:60px auto;padding:24px;color:#1e293b;line-height:1.6}
        h2{color:#16a34a}
        .card{border-radius:12px;padding:20px;margin:16px 0}
        .green{background:#f0fdf4;border:1px solid #bbf7d0}
        .blue{background:#eff6ff;border:1px solid #bfdbfe}
        .step{display:flex;gap:12px;margin:10px 0;align-items:flex-start}
        .num{width:24px;height:24px;border-radius:50%;background:#6366f1;color:#fff;font-size:11px;font-weight:900;display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:2px}
        code{background:#f1f5f9;padding:4px 10px;border-radius:6px;font-family:monospace;font-size:13px;word-break:break-all;display:block;margin-top:6px}
        button{background:#6366f1;color:#fff;border:none;padding:10px 20px;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer;margin-top:10px}
        button:hover{background:#4f46e5}
        p{margin:6px 0;font-size:14px}
        .warn{background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:12px;color:#92400e;font-size:13px}
      </style></head><body>
      <h2>✅ 구글 계정 인증 완료!</h2>
      <p style="color:#64748b">아래 Refresh Token을 Vercel 환경변수에 등록해야 자동 업로드가 작동합니다.</p>

      <div class="card green">
        <p><strong>📋 복사할 Refresh Token</strong></p>
        ${refreshToken
          ? `<code id="tok">${refreshToken}</code>
             <button onclick="navigator.clipboard.writeText(document.getElementById('tok').textContent).then(()=>this.textContent='✅ 복사됨!')">클립보드에 복사</button>`
          : `<p class="warn">⚠️ Refresh Token이 발급되지 않았습니다.<br>
             이미 한 번 인증한 계정이면 Google 계정 설정에서 앱 권한을 해제 후 다시 시도하세요.<br>
             또는 <a href="/api/google-auth">여기서 재인증</a>하세요.</p>`}
      </div>

      <div class="card blue">
        <p><strong>📌 Vercel에 등록하는 방법</strong></p>
        <div class="step"><div class="num">1</div><p><a href="https://vercel.com/dashboard" target="_blank">vercel.com</a> → juganbogo3 프로젝트 클릭</p></div>
        <div class="step"><div class="num">2</div><p><b>Settings → Environment Variables</b></p></div>
        <div class="step"><div class="num">3</div><p>Key: <b>GOOGLE_REFRESH_TOKEN</b><br>Value: 위 토큰 붙여넣기 → Add</p></div>
        <div class="step"><div class="num">4</div><p><b>Deployments</b> 탭 → 최신 배포 <code style="display:inline;padding:2px 6px">···</code> → <b>Redeploy</b></p></div>
        <div class="step"><div class="num">5</div><p>앱 접속 → 관리자 콘솔 → 배너가 🟢 <b>연결됨</b> 확인</p></div>
      </div>

      <p style="font-size:13px;color:#94a3b8">이 창은 닫아도 됩니다.</p>
      </body></html>`);
  } catch (err) {
    console.error('[GoogleDrive] Token exchange failed:', err);
    res.status(500).send(`<h2>토큰 교환 실패</h2><pre>${err.message}</pre>`);
  }
}
