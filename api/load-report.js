import { google } from 'googleapis';

function getDriveClient() {
  const { GOOGLE_CLIENT_ID: clientId, GOOGLE_CLIENT_SECRET: clientSecret, GOOGLE_REFRESH_TOKEN: refreshToken } = process.env;
  if (!clientId || !clientSecret || !refreshToken) return null;
  const auth = new google.auth.OAuth2(clientId, clientSecret);
  auth.setCredentials({ refresh_token: refreshToken });
  return google.drive({ version: 'v3', auth });
}

function isHyeohoe(parish) {
  return parish === '협회' || parish === '협회본부';
}
function getDisplayParish(parish) {
  if (isHyeohoe(parish)) return '협회본부';
  return parish.endsWith('교구') ? parish : `${parish}교구`;
}
function toKey(church) {
  return church.replace(/[ ()]/g, '_');
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'GET') return res.status(405).json({ error: 'Method not allowed' });

  const drive = getDriveClient();
  if (!drive) return res.status(200).json({ found: false, reason: 'Drive not configured' });

  const { parish, church } = req.query;
  if (!parish || !church) return res.status(400).json({ error: 'parish, church 파라미터 필요' });

  try {
    // 루트 폴더 탐색
    const rootId = process.env.GOOGLE_DRIVE_FOLDER_ID || await (async () => {
      const r = await drive.files.list({
        q: `name='주간보고_제출현황' and mimeType='application/vnd.google-apps.folder' and trashed=false`,
        fields: 'files(id)', spaces: 'drive',
      });
      return r.data.files[0]?.id || null;
    })();
    if (!rootId) return res.status(200).json({ found: false });

    // 교구 폴더 탐색
    let parishFolderName;
    if (isHyeohoe(parish)) {
      parishFolderName = '협회본부';
    } else {
      parishFolderName = getDisplayParish(parish);
    }

    const parishRes = await drive.files.list({
      q: `'${rootId}' in parents and name='${parishFolderName.replace(/'/g,"\\'")}' and mimeType='application/vnd.google-apps.folder' and trashed=false`,
      fields: 'files(id)', spaces: 'drive',
    });
    if (!parishRes.data.files.length) return res.status(200).json({ found: false });
    let folderId = parishRes.data.files[0].id;

    // 협회인 경우 국 폴더 한 단계 더
    if (isHyeohoe(parish)) {
      const gukRes = await drive.files.list({
        q: `'${folderId}' in parents and name='${church.replace(/'/g,"\\'")}' and mimeType='application/vnd.google-apps.folder' and trashed=false`,
        fields: 'files(id)', spaces: 'drive',
      });
      if (!gukRes.data.files.length) return res.status(200).json({ found: false });
      folderId = gukRes.data.files[0].id;
    }

    // JSON 파일 탐색
    const cleanChurch = toKey(church);
    const fileRes = await drive.files.list({
      q: `'${folderId}' in parents and name='${cleanChurch}_data.json' and trashed=false`,
      fields: 'files(id,name,modifiedTime)', spaces: 'drive',
    });
    if (!fileRes.data.files.length) return res.status(200).json({ found: false });

    const fileId = fileRes.data.files[0].id;
    const modifiedTime = fileRes.data.files[0].modifiedTime;

    const download = await drive.files.get({ fileId, alt: 'media' }, { responseType: 'arraybuffer' });
    const payload = JSON.parse(Buffer.from(download.data).toString('utf-8'));

    res.status(200).json({ found: true, payload, modifiedTime });
  } catch (err) {
    console.error('[/api/load-report]', err.message);
    res.status(500).json({ error: err.message });
  }
}
