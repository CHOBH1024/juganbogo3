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
  const c = ["①","②","③","④","⑤","⑥","⑦","⑧","⑨","⑩","⑪","⑫","⑬","⑭","⑮"];
  return c[num - 1] || `(${num})`;
}

function getDriveClient() {
  const clientId = process.env.GOOGLE_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
  const refreshToken = process.env.GOOGLE_REFRESH_TOKEN;
  if (!clientId || !clientSecret || !refreshToken) return null;

  const auth = new google.auth.OAuth2(clientId, clientSecret);
  auth.setCredentials({ refresh_token: refreshToken });
  return google.drive({ version: 'v3', auth });
}

async function getOrCreateFolder(drive, name, parentId) {
  const q = parentId
    ? `name='${name}' and mimeType='application/vnd.google-apps.folder' and '${parentId}' in parents and trashed=false`
    : `name='${name}' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
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

async function buildWordBuffer(parish, church, payload) {
  const children = [];
  children.push(new Paragraph({ children: [new TextRun({ text: `${parish} - ${church} 주간보고서`, bold: true, size: 32, color: '1D4ED8', font: '맑은 고딕' })], spacing: { after: 300 } }));
  children.push(new Paragraph({ children: [new TextRun({ text: `제출 시각: ${new Date().toLocaleString('ko-KR')}`, size: 18, color: '64748B', font: '맑은 고딕' })], spacing: { after: 500 } }));

  const counters = [0, 0, 0, 0];
  for (const item of (payload.data || [])) {
    let prefix = '';
    if (item.level === 0) { counters[0]++; counters[1]=0; counters[2]=0; counters[3]=0; prefix = toRoman(counters[0]) + '. '; }
    else if (item.level === 1) { counters[1]++; counters[2]=0; counters[3]=0; prefix = counters[1] + '. '; }
    else if (item.level === 2) { counters[2]++; counters[3]=0; prefix = counters[2] + ') '; }
    else if (item.level === 3) { counters[3]++; prefix = toCircled(counters[3]) + ' '; }

    const color = item.level <= 1 ? '1D4ED8' : '000000';
    const lines = `${prefix}${item.text || ''}`.split('\n');
    children.push(new Paragraph({
      children: lines.map((line, idx) => new TextRun({ text: line, break: idx > 0 ? 1 : 0, bold: item.level <= 1, color, font: '맑은 고딕', size: 22 })),
      indent: { left: item.level === 0 ? 0 : (item.level - 1) * 360 },
      spacing: { before: item.level === 0 ? 360 : 120 },
    }));

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
            for (let r = 0; r < rowSpan; r++) for (let c = 0; c < colSpan; c++) { if (r===0&&c===0) continue; skipped.add(`${rIdx+r},${cIdx+c}`); }
            const isH = item.tableHighlights?.[rIdx]?.[cIdx];
            const align = item.tableAlignments?.[rIdx]?.[cIdx] || 'left';
            cells.push(new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: cellText||'', bold: isH, color: isH?'1D4ED8':'000000', font:'맑은 고딕', size:18 })], alignment: align })], columnSpan: colSpan, rowSpan, shading: isH?{fill:'EFF6FF'}:undefined }));
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
  let txt = `=========================================\n교구: ${parish}\n교회: ${church}\n제출상태: ${payload.status==='submitted'?'제출 완료':'임시 저장'}\n업데이트: ${new Date().toLocaleString('ko-KR')}\n=========================================\n\n`;
  const counters = [0,0,0,0];
  for (const item of (payload.data || [])) {
    let prefix = '';
    if (item.level===0){counters[0]++;counters[1]=0;counters[2]=0;counters[3]=0;prefix=toRoman(counters[0])+'. ';}
    else if(item.level===1){counters[1]++;counters[2]=0;counters[3]=0;prefix=counters[1]+'. ';}
    else if(item.level===2){counters[2]++;counters[3]=0;prefix=counters[2]+') ';}
    else if(item.level===3){counters[3]++;prefix=toCircled(counters[3])+' ';}
    txt += `${'  '.repeat(item.level)}${prefix}${item.text||''}\n`;
    if (item.tableData?.length>0) {
      txt += '\n[데이터 테이블]\n';
      item.tableData.forEach((row,rIdx)=>{
        txt += '| '+row.map(c=>c.trim().replace(/\n/g,' ')).join(' | ')+' |\n';
        if(rIdx===0) txt += '| '+row.map(()=>'---').join(' | ')+' |\n';
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
  const cleanChurch = church.replace(/[\/\\?%*:|"<>. ]/g, '_');

  try {
    const rootFolderId = process.env.GOOGLE_DRIVE_FOLDER_ID || await getOrCreateFolder(drive, '주간보고_제출현황', null);

    // 텍스트 파일 (실시간 — 저장/제출 모두)
    const txtContent = buildTextContent(parish, church, payload);
    const txtBuffer = Buffer.from(txtContent, 'utf-8');
    const cleanParish = parish.replace(/[\/\\?%*:|"<>. ]/g, '_');
    await uploadOrUpdate(drive, `[${cleanParish}_${cleanChurch}]_주간보고.txt`, txtBuffer, 'text/plain', rootFolderId);

    // Word 파일 (제출 확정 시만)
    if (payload?.status === 'submitted') {
      const parishFolderId = await getOrCreateFolder(drive, parish, rootFolderId);
      const wordBuffer = await buildWordBuffer(parish, church, payload);
      const filename = `${cleanChurch}_주간보고_${new Date().toISOString().slice(0,10)}.docx`;
      await uploadOrUpdate(drive, filename, wordBuffer, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', parishFolderId);
    }

    res.json({ success: true });
  } catch (err) {
    console.error('[Drive] Upload failed:', err.message);
    res.status(500).json({ error: err.message });
  }
}
