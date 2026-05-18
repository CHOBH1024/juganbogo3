import express from 'express';
import { google } from 'googleapis';
import { GoogleGenAI } from '@google/genai';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ImageRun } from 'docx';
import { Readable } from 'stream';

const app = express();
const PORT = process.env.PORT || 8080;

// 시작 시 필수 환경변수 확인
const REQUIRED_VARS = ['GOOGLE_CLIENT_ID', 'GOOGLE_CLIENT_SECRET', 'GOOGLE_REFRESH_TOKEN'];
const MISSING = REQUIRED_VARS.filter(v => !process.env[v]);
if (MISSING.length) {
  console.warn(`⚠️  환경변수 미설정: ${MISSING.join(', ')}`);
  console.warn('   Drive 연동 기능이 동작하지 않습니다.');
}
if (!process.env.API_KEY && !process.env.GEMINI_API_KEY) {
  console.warn('⚠️  API_KEY (Gemini) 미설정 — AI 검토 기능이 동작하지 않습니다.');
}

app.use(express.json({ limit: '50mb' }));
app.use(express.static('public'));

// ── Google Drive 클라이언트 ──────────────────────────────────────────────────
function getDriveClient() {
  const auth = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET
  );
  auth.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
  return google.drive({ version: 'v3', auth });
}

// Drive 폴더/파일 탐색 헬퍼
async function listFiles(drive, folderId, nameFilter) {
  const q = nameFilter
    ? `'${folderId}' in parents and name contains '${nameFilter}' and trashed=false`
    : `'${folderId}' in parents and trashed=false`;
  const res = await drive.files.list({ q, fields: 'files(id,name,mimeType)', spaces: 'drive' });
  return res.data.files;
}

async function downloadJson(drive, fileId) {
  const res = await drive.files.get({ fileId, alt: 'media' }, { responseType: 'arraybuffer' });
  return JSON.parse(Buffer.from(res.data).toString('utf-8'));
}

// ── API: Drive에서 보고서 목록 불러오기 ────────────────────────────────────────
app.get('/api/reports', async (req, res) => {
  try {
    const drive = getDriveClient();
    let folderId = process.env.GOOGLE_DRIVE_FOLDER_ID;

    if (!folderId) {
      const found = await drive.files.list({
        q: `name='주간보고_제출현황' and mimeType='application/vnd.google-apps.folder' and trashed=false`,
        fields: 'files(id)',
        spaces: 'drive',
      });
      if (!found.data.files.length) return res.json({ parishes: [] });
      folderId = found.data.files[0].id;
    }

    // 교구 폴더 목록
    const parishFolders = await listFiles(drive, folderId, null);
    const parishes = [];

    for (const folder of parishFolders) {
      if (folder.mimeType !== 'application/vnd.google-apps.folder') continue;

      // 협회본부: 국별 하위 폴더
      if (folder.name === '협회본부') {
        const gukFolders = await listFiles(drive, folder.id, null);
        for (const gukFolder of gukFolders) {
          if (gukFolder.mimeType !== 'application/vnd.google-apps.folder') continue;
          const dataFiles = await listFiles(drive, gukFolder.id, '_data.json');
          const churches = [];
          for (const f of dataFiles) {
            try {
              const payload = await downloadJson(drive, f.id);
              const church = f.name.replace('_data.json', '').replace(/_/g, ' ');
              churches.push({ church, status: payload.status || 'draft', payload });
            } catch (e) { /* skip */ }
          }
          if (churches.length) parishes.push({ parish: '협회본부', folder: gukFolder.name, churches });
        }
        continue;
      }

      // 일반 교구
      const dataFiles = await listFiles(drive, folder.id, '_data.json');
      const churches = [];
      for (const f of dataFiles) {
        try {
          const payload = await downloadJson(drive, f.id);
          const church = f.name.replace('_data.json', '').replace(/_/g, ' ');
          churches.push({ church, status: payload.status || 'draft', payload });
        } catch (e) { /* skip */ }
      }
      if (churches.length) parishes.push({ parish: folder.name, folder: folder.name, churches });
    }

    res.json({ parishes });
  } catch (err) {
    console.error('[/api/reports]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── API: Gemini AI 검토 ────────────────────────────────────────────────────────
app.post('/api/ai-review', async (req, res) => {
  try {
    const { parishes } = req.body;
    const ai = new GoogleGenAI({ apiKey: process.env.API_KEY || process.env.GEMINI_API_KEY });

    // 검토용 텍스트 구성
    let fullText = '';
    for (const p of parishes) {
      fullText += `\n\n===== ${p.parish} ${p.folder !== p.parish ? '/ ' + p.folder : ''} =====\n`;
      for (const c of p.churches) {
        if (c.status !== 'submitted') continue;
        fullText += `\n[${c.church}]\n`;
        for (const item of (c.payload?.data || [])) {
          const indent = '  '.repeat(item.level);
          fullText += `${indent}${item.text || ''}\n`;
          if (item.tableData?.length) {
            item.tableData.forEach(row => { fullText += `${indent}| ${row.join(' | ')} |\n`; });
          }
        }
      }
    }

    const prompt = `당신은 전국 교구 주간업무보고를 총괄 검토하는 전문 편집자입니다.
아래 보고서들의 맞춤법과 문장을 교정하고, 교정 결과를 아래 JSON 형식으로 반환하세요.
각 항목은 원문(original)과 교정문(corrected), 교정 이유(reason)를 포함해야 합니다.
교정이 불필요한 항목은 original과 corrected를 동일하게 하세요.

응답 형식:
[
  { "parish": "교구명", "church": "교회명", "original": "원문", "corrected": "교정문", "reason": "이유" },
  ...
]

검토할 보고서:
${fullText}`;

    const response = await ai.models.generateContent({
      model: 'gemini-2.0-flash',
      contents: prompt,
    });

    let text = response.text();
    // JSON 추출
    const match = text.match(/\[[\s\S]*\]/);
    const corrections = match ? JSON.parse(match[0]) : [];

    res.json({ corrections, rawText: text });
  } catch (err) {
    console.error('[/api/ai-review]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── 헬퍼: Word 생성 ───────────────────────────────────────────────────────────
function toRoman(num) {
  const lookup = { M:1000,CM:900,D:500,CD:400,C:100,XC:90,L:50,XL:40,X:10,IX:9,V:5,IV:4,I:1 };
  let r = '';
  for (let i in lookup) { while (num >= lookup[i]) { r += i; num -= lookup[i]; } }
  return r;
}
function toCircled(n) {
  return ['①','②','③','④','⑤','⑥','⑦','⑧','⑨','⑩'][n-1] || `(${n})`;
}

async function fetchImageBuffer(url) {
  try {
    if (url.startsWith('http')) return Buffer.from(await (await fetch(url)).arrayBuffer());
    if (url.startsWith('data:')) return Buffer.from(url.split(',')[1], 'base64');
  } catch { return null; }
  return null;
}

async function buildParishWord(parishData, corrections) {
  const corrMap = {};
  for (const c of (corrections || [])) {
    corrMap[`${c.parish}__${c.church}__${c.original}`] = c.corrected;
  }

  const children = [];
  children.push(new Paragraph({
    children: [new TextRun({ text: `${parishData.parish} 주간업무보고서`, bold: true, size: 40, color: '1E3A8A', font: '맑은 고딕' })],
    spacing: { after: 200 },
  }));
  children.push(new Paragraph({
    children: [new TextRun({ text: `취합일: ${new Date().toLocaleDateString('ko-KR')}`, size: 20, color: '64748B', font: '맑은 고딕' })],
    spacing: { after: 600 },
  }));

  let secNum = 0;
  for (const ch of parishData.churches) {
    if (ch.status !== 'submitted') continue;
    secNum++;
    children.push(new Paragraph({
      children: [new TextRun({ text: `${toRoman(secNum)}. ${ch.church}`, bold: true, size: 28, color: '1E3A8A', font: '맑은 고딕' })],
      spacing: { before: 600, after: 200 },
      border: { bottom: { style: 'single', size: 6, color: 'BFDBFE' } },
    }));

    const KOR_CONS = ['가','나','다','라','마','바','사','아','자','차','카','타','파','하'];
    const counters = [0,0,0,0,0,0];
    for (const item of (ch.payload?.data || [])) {
      let prefix = '';
      if (item.level===0){counters[0]++;counters[1]=0;counters[2]=0;counters[3]=0;counters[4]=0;counters[5]=0;prefix=toRoman(counters[0])+'. ';}
      else if(item.level===1){counters[1]++;counters[2]=0;counters[3]=0;counters[4]=0;counters[5]=0;prefix=counters[1]+'. ';}
      else if(item.level===2){counters[2]++;counters[3]=0;counters[4]=0;counters[5]=0;prefix=counters[2]+') ';}
      else if(item.level===3){counters[3]++;counters[4]=0;counters[5]=0;prefix=toCircled(counters[3])+' ';}
      else if(item.level===4){counters[4]++;counters[5]=0;prefix=(KOR_CONS[counters[4]-1]||`(${counters[4]})`)+'. ';}
      else if(item.level===5){counters[5]++;prefix=String.fromCharCode(96+counters[5])+'. ';}

      const raw = item.text || '';
      const key = `${parishData.parish}__${ch.church}__${raw}`;
      const text = corrMap[key] || raw;
      const color = item.level === 0 ? '1D4ED8' : '000000';
      const indent = { left: item.level === 0 ? 0 : item.level * 360 };

      children.push(new Paragraph({
        children: (prefix + text).split('\n').map((line, idx) =>
          new TextRun({ text: line, break: idx>0?1:0, bold: item.level===0, color, font:'맑은 고딕', size:20 })
        ),
        indent,
        spacing: { before: item.level===0 ? 240 : 80 },
      }));

      if (item.image && item.imageWidth && item.imageHeight) {
        const buf = await fetchImageBuffer(item.image);
        if (buf) {
          const ratio = Math.min(1, 580 / item.imageWidth);
          try {
            children.push(new Paragraph({
              children: [new ImageRun({
                data: buf,
                transformation: { width: Math.round(item.imageWidth * ratio), height: Math.round(item.imageHeight * ratio) },
                type: item.image.startsWith('data:image/png') ? 'png' : 'jpg',
              })],
              indent,
              spacing: { before: 100, after: 100 },
            }));
          } catch { /* skip */ }
        }
      }

      if (item.tableData?.length > 0) {
        const skipped = new Set();
        children.push(new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: item.tableData.map((row, rIdx) => {
            const cells = [];
            row.forEach((cellText, cIdx) => {
              if (skipped.has(`${rIdx},${cIdx}`)) return;
              const spanDef = item.tableSpans?.[rIdx]?.[cIdx];
              const colSpan = (typeof spanDef==='number'?spanDef:spanDef?.colspan)||1;
              const rowSpan = (typeof spanDef==='number'?1:spanDef?.rowspan)||1;
              for(let r=0;r<rowSpan;r++)for(let c=0;c<colSpan;c++){if(r===0&&c===0)continue;skipped.add(`${rIdx+r},${cIdx+c}`);}
              const isH = item.tableHighlights?.[rIdx]?.[cIdx];
              cells.push(new TableCell({
                children: [new Paragraph({ children: [new TextRun({ text: cellText||'', bold: isH, color: isH?'1D4ED8':'000000', font:'맑은 고딕', size:18 })] })],
                columnSpan: colSpan, rowSpan,
                shading: isH ? { fill: 'EFF6FF' } : undefined,
              }));
            });
            return new TableRow({ children: cells });
          }),
        }));
      }
    }
  }

  return Packer.toBuffer(new Document({ sections: [{ properties: {}, children }] }));
}

// ── API: Word 다운로드 ────────────────────────────────────────────────────────
app.post('/api/download-word', async (req, res) => {
  try {
    const { parishData, corrections } = req.body;
    const buf = await buildParishWord(parishData, corrections);
    const filename = encodeURIComponent(`${parishData.parish}_주간보고_${new Date().toISOString().slice(0,10)}.docx`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
    res.send(buf);
  } catch (err) {
    console.error('[/api/download-word]', err.message);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => console.log(`AI Review App running on port ${PORT}`));
