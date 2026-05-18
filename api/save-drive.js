import { google } from 'googleapis';
import { Readable } from 'stream';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ImageRun } from 'docx';

function toRoman(num) {
  const lookup = { M:1000,CM:900,D:500,CD:400,C:100,XC:90,L:50,XL:40,X:10,IX:9,V:5,IV:4,I:1 };
  let roman = '';
  for (let i in lookup) { while (num >= lookup[i]) { roman += i; num -= lookup[i]; } }
  return roman;
}
function toCircled(num) {
  const c = ['①','②','③','④','⑤','⑥','⑦','⑧','⑨','⑩','⑪','⑫','⑬','⑭','⑮'];
  return c[num - 1] || `(${num})`;
}

function isHyeohoe(parish) {
  return parish === '협회' || parish === '협회본부';
}

function getDisplayParish(parish) {
  if (isHyeohoe(parish)) return '협회본부';
  return parish.endsWith('교구') ? parish : `${parish}교구`;
}

function getDisplayChurch(church) {
  if (church === '교구본부' || church.endsWith('국')) return church;
  if (church.endsWith('교회') || church.endsWith('학사') || church.endsWith('센터') ||
      church.endsWith('대학') || church.endsWith('전도소') || church.endsWith('글로벌')) return church;
  return `${church}교회`;
}

function getDriveClient() {
  const { GOOGLE_CLIENT_ID: clientId, GOOGLE_CLIENT_SECRET: clientSecret, GOOGLE_REFRESH_TOKEN: refreshToken } = process.env;
  if (!clientId || !clientSecret || !refreshToken) return null;
  const auth = new google.auth.OAuth2(clientId, clientSecret);
  auth.setCredentials({ refresh_token: refreshToken });
  return google.drive({ version: 'v3', auth });
}

async function getOrCreateFolder(drive, name, parentId) {
  const safe = name.replace(/'/g, "\\'");
  const q = parentId
    ? `name='${safe}' and mimeType='application/vnd.google-apps.folder' and '${parentId}' in parents and trashed=false`
    : `name='${safe}' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
  const res = await drive.files.list({ q, fields: 'files(id)', spaces: 'drive' });
  if (res.data.files.length > 0) return res.data.files[0].id;
  const created = await drive.files.create({
    requestBody: {
      name,
      mimeType: 'application/vnd.google-apps.folder',
      parents: parentId ? [parentId] : (process.env.GOOGLE_DRIVE_FOLDER_ID ? [process.env.GOOGLE_DRIVE_FOLDER_ID] : []),
    },
    fields: 'id',
  });
  return created.data.id;
}

async function uploadOrUpdate(drive, filename, buffer, mimeType, folderId) {
  const q = `name='${filename}' and '${folderId}' in parents and trashed=false`;
  const existing = await drive.files.list({ q, fields: 'files(id)', spaces: 'drive' });
  const stream = Readable.from(buffer);
  if (existing.data.files.length > 0) {
    await drive.files.update({ fileId: existing.data.files[0].id, media: { mimeType, body: stream } });
  } else {
    await drive.files.create({ requestBody: { name: filename, parents: [folderId] }, media: { mimeType, body: stream }, fields: 'id' });
  }
}

async function fetchImageBuffer(imageStr) {
  try {
    if (imageStr.startsWith('http')) {
      const r = await fetch(imageStr);
      return Buffer.from(await r.arrayBuffer());
    } else if (imageStr.startsWith('data:')) {
      return Buffer.from(imageStr.split(',')[1], 'base64');
    }
  } catch (e) {
    console.error('[Drive] Image fetch failed:', e.message);
  }
  return null;
}

// 드라이브 폴더 구조:
// 교구:  주간보고_제출현황/{교구명교구}/{교회명}_주간보고_YYYY-MM-DD.docx
// 협회:  주간보고_제출현황/협회본부/{국명}/{국명}_주간보고_YYYY-MM-DD.docx
async function resolveDocxFolder(drive, parish, church, rootId) {
  if (isHyeohoe(parish)) {
    const hyeohoeId = await getOrCreateFolder(drive, '협회본부', rootId);
    return await getOrCreateFolder(drive, church, hyeohoeId);
  }
  const parishDisplay = getDisplayParish(parish);
  return await getOrCreateFolder(drive, parishDisplay, rootId);
}

async function buildWordBuffer(parish, church, payload) {
  const children = [];
  const title = `${getDisplayParish(parish)} - ${getDisplayChurch(church)} 주간보고서`;
  children.push(new Paragraph({
    children: [new TextRun({ text: title, bold: true, size: 32, color: '1D4ED8', font: '맑은 고딕' })],
    spacing: { after: 300 },
  }));
  children.push(new Paragraph({
    children: [new TextRun({ text: `제출 시각: ${new Date().toLocaleString('ko-KR')}`, size: 18, color: '64748B', font: '맑은 고딕' })],
    spacing: { after: 500 },
  }));

  const counters = [0,0,0,0];
  for (const item of (payload.data || [])) {
    let prefix = '';
    if (item.level===0){counters[0]++;counters[1]=0;counters[2]=0;counters[3]=0;prefix=toRoman(counters[0])+'. ';}
    else if(item.level===1){counters[1]++;counters[2]=0;counters[3]=0;prefix=counters[1]+'. ';}
    else if(item.level===2){counters[2]++;counters[3]=0;prefix=counters[2]+') ';}
    else if(item.level===3){counters[3]++;prefix=toCircled(counters[3])+' ';}

    const color = item.level <= 1 ? '1D4ED8' : '000000';
    const indent = { left: item.level === 0 ? 0 : (item.level - 1) * 360 };
    const lines = `${prefix}${item.text || ''}`.split('\n');
    children.push(new Paragraph({
      children: lines.map((line,idx) => new TextRun({ text: line, break: idx>0?1:0, bold: item.level<=1, color, font:'맑은 고딕', size:22 })),
      indent,
      spacing: { before: item.level===0 ? 360 : 120 },
    }));

    // 사진 삽입
    if (item.image && item.imageWidth && item.imageHeight) {
      const imgBuf = await fetchImageBuffer(item.image);
      if (imgBuf) {
        const maxPx = 620; // A4 본문 너비(px 기준)
        const ratio = Math.min(1, maxPx / item.imageWidth);
        const w = Math.round(item.imageWidth * ratio);
        const h = Math.round(item.imageHeight * ratio);
        const type = item.image.startsWith('data:image/png') ? 'png' : 'jpg';
        try {
          children.push(new Paragraph({
            children: [new ImageRun({ data: imgBuf, transformation: { width: w, height: h }, type })],
            indent,
            spacing: { before: 120, after: 120 },
          }));
        } catch (e) {
          console.error('[Drive] ImageRun failed:', e.message);
        }
      }
    }

    // 테이블 삽입
    if (item.tableData?.length > 0) {
      const skipped = new Set();
      children.push(new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: item.tableData.map((row, rIdx) => {
          const cells = [];
          row.forEach((cellText, cIdx) => {
            if (skipped.has(`${rIdx},${cIdx}`)) return;
            const spanDef = item.tableSpans?.[rIdx]?.[cIdx];
            const colSpan = (typeof spanDef === 'number' ? spanDef : spanDef?.colspan) || 1;
            const rowSpan = (typeof spanDef === 'number' ? 1 : spanDef?.rowspan) || 1;
            for (let r=0;r<rowSpan;r++) for(let c=0;c<colSpan;c++){if(r===0&&c===0)continue;skipped.add(`${rIdx+r},${cIdx+c}`);}
            const isH = item.tableHighlights?.[rIdx]?.[cIdx];
            const align = item.tableAlignments?.[rIdx]?.[cIdx] || 'left';
            cells.push(new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: cellText||'', bold: isH, color: isH?'1D4ED8':'000000', font:'맑은 고딕', size:18 })], alignment: align })],
              columnSpan: colSpan, rowSpan,
              shading: isH ? { fill: 'EFF6FF' } : undefined,
            }));
          });
          return new TableRow({ children: cells });
        }),
      }));
    }
  }

  const doc = new Document({ sections: [{ properties: {}, children }] });
  return Packer.toBuffer(doc);
}

function buildTextContent(parish, church, payload) {
  let txt = `=========================================\n교구/협회: ${getDisplayParish(parish)}\n교회/부서: ${getDisplayChurch(church)}\n제출상태: ${payload.status==='submitted'?'제출 완료':'임시 저장'}\n업데이트: ${new Date().toLocaleString('ko-KR')}\n=========================================\n\n`;
  const counters = [0,0,0,0];
  for (const item of (payload.data || [])) {
    let prefix = '';
    if(item.level===0){counters[0]++;counters[1]=0;counters[2]=0;counters[3]=0;prefix=toRoman(counters[0])+'. ';}
    else if(item.level===1){counters[1]++;counters[2]=0;counters[3]=0;prefix=counters[1]+'. ';}
    else if(item.level===2){counters[2]++;counters[3]=0;prefix=counters[2]+') ';}
    else if(item.level===3){counters[3]++;prefix=toCircled(counters[3])+' ';}
    txt += `${'  '.repeat(item.level)}${prefix}${item.text||''}\n`;
    if (item.image) txt += `${'  '.repeat(item.level+1)}[사진 첨부됨]\n`;
    if (item.tableData?.length > 0) {
      txt += '\n[데이터 테이블]\n';
      item.tableData.forEach((row,rIdx) => {
        txt += '| ' + row.map(c => String(c).trim().replace(/\n/g,' ')).join(' | ') + ' |\n';
        if (rIdx === 0) txt += '| ' + row.map(() => '---').join(' | ') + ' |\n';
      });
      txt += '\n';
    }
  }
  return txt;
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const drive = getDriveClient();
  if (!drive) return res.status(200).json({ skipped: true, reason: 'Drive not configured' });

  const { id, payload } = req.body;
  if (!id?.startsWith('report_')) return res.status(200).json({ skipped: true });
  const parts = id.split('_');
  if (parts.length < 3) return res.status(200).json({ skipped: true });
  const parish = parts[1];
  const church = parts[2];
  const cleanChurch = church.replace(/[/\\?%*:|"<>. ]/g, '_');
  const cleanParish = parish.replace(/[/\\?%*:|"<>. ]/g, '_');

  try {
    const rootId = process.env.GOOGLE_DRIVE_FOLDER_ID || await getOrCreateFolder(drive, '주간보고_제출현황', null);

    // 실시간 텍스트 (저장/제출 모두): 루트에 저장
    const txtBuffer = Buffer.from(buildTextContent(parish, church, payload), 'utf-8');
    await uploadOrUpdate(drive, `[${cleanParish}_${cleanChurch}]_주간보고.txt`, txtBuffer, 'text/plain', rootId);

    // Word 파일 (제출 확정 시만): 교구/국별 폴더 구조로 저장
    if (payload?.status === 'submitted') {
      const targetFolderId = await resolveDocxFolder(drive, parish, church, rootId);
      const wordBuffer = await buildWordBuffer(parish, church, payload);
      const filename = `${cleanChurch}_주간보고_${new Date().toISOString().slice(0,10)}.docx`;
      await uploadOrUpdate(drive, filename, wordBuffer, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', targetFolderId);
    }

    res.json({ success: true });
  } catch (err) {
    console.error('[Drive] Upload failed:', err.message);
    res.status(500).json({ error: err.message });
  }
}
