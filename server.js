import express from 'express';
import fs from 'fs';
import path from 'path';
import os from 'os';
import { fileURLToPath } from 'url';
import { Readable } from 'stream';
import { spawn, exec } from 'child_process';
import { createClient } from '@supabase/supabase-js';

const SUPABASE_URL = 'https://xootqaeuixpsszcejhev.supabase.co';
const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inhvb3RxYWV1aXhwc3N6Y2VqaGV2Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg3ODk3MDgsImV4cCI6MjA5NDM2NTcwOH0.W2h7M1zUZFNG6KjtQm92CfG3ixcllhhW2_Az6loxYJI';
const sb = createClient(SUPABASE_URL, SUPABASE_KEY);
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ImageRun } from 'docx';
import { google } from 'googleapis';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = 5000;

// ─── Google Drive API ───────────────────────────────────────────────────────

const GOOGLE_CLIENT_ID = process.env.GOOGLE_CLIENT_ID || '';
const GOOGLE_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET || '';
const GOOGLE_REFRESH_TOKEN = process.env.GOOGLE_REFRESH_TOKEN || '';
// 업로드 대상 Drive 폴더 ID (비워두면 내 드라이브 루트에 생성)
const GOOGLE_DRIVE_FOLDER_ID = process.env.GOOGLE_DRIVE_FOLDER_ID || '';
const APP_URL = process.env.APP_URL || `http://localhost:${5000}`;

let oauth2Client = null;
let driveClient = null;

if (GOOGLE_CLIENT_ID && GOOGLE_CLIENT_SECRET) {
  oauth2Client = new google.auth.OAuth2(
    GOOGLE_CLIENT_ID,
    GOOGLE_CLIENT_SECRET,
    `${APP_URL}/api/google-auth/callback`
  );
  if (GOOGLE_REFRESH_TOKEN) {
    oauth2Client.setCredentials({ refresh_token: GOOGLE_REFRESH_TOKEN });
    driveClient = google.drive({ version: 'v3', auth: oauth2Client });
    console.log('[GoogleDrive] Drive API client initialized with stored refresh token.');
  }
}

// Drive에 폴더를 조회 또는 생성하여 ID를 반환
async function getOrCreateDriveFolder(name, parentId) {
  const q = parentId
    ? `name='${name}' and mimeType='application/vnd.google-apps.folder' and '${parentId}' in parents and trashed=false`
    : `name='${name}' and mimeType='application/vnd.google-apps.folder' and trashed=false`;

  const res = await driveClient.files.list({ q, fields: 'files(id)', spaces: 'drive' });
  if (res.data.files.length > 0) return res.data.files[0].id;

  const created = await driveClient.files.create({
    requestBody: {
      name,
      mimeType: 'application/vnd.google-apps.folder',
      parents: parentId ? [parentId] : (GOOGLE_DRIVE_FOLDER_ID ? [GOOGLE_DRIVE_FOLDER_ID] : [])
    },
    fields: 'id'
  });
  return created.data.id;
}

// Drive에 파일 업로드 (동명 파일이 있으면 덮어씀)
async function uploadFileToDrive(filename, buffer, mimeType, folderId) {
  if (!driveClient) return null;

  // 동명 파일 존재 확인
  const q = `name='${filename}' and '${folderId}' in parents and trashed=false`;
  const existing = await driveClient.files.list({ q, fields: 'files(id)', spaces: 'drive' });

  const stream = Readable.from(buffer);

  if (existing.data.files.length > 0) {
    // 기존 파일 업데이트 (내용 교체)
    const fileId = existing.data.files[0].id;
    await driveClient.files.update({
      fileId,
      media: { mimeType, body: stream }
    });
    console.log(`[GoogleDrive] Updated: ${filename} (id=${fileId})`);
    return fileId;
  } else {
    const res = await driveClient.files.create({
      requestBody: { name: filename, parents: [folderId] },
      media: { mimeType, body: stream },
      fields: 'id'
    });
    console.log(`[GoogleDrive] Uploaded: ${filename} (id=${res.data.id})`);
    return res.data.id;
  }
}

// ───────────────────────────────────────────────────────────────────────────

// JSON body parser with 50MB limit for large report payloads
app.use(express.json({ limit: '50mb' }));

// CORS headers to allow Vite client on port 3000 to call the server
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS, PUT, PATCH, DELETE');
  res.setHeader('Access-Control-Allow-Headers', 'X-Requested-With,Content-Type,X-Filename,ngrok-skip-browser-warning,Authorization');
  if (req.method === 'OPTIONS') {
    return res.sendStatus(200);
  }
  next();
});

// Ensure storage folders exist
const DB_DIR = path.join(__dirname, 'local_db');
const UPLOADS_DIR = path.join(__dirname, 'uploads');

if (!fs.existsSync(DB_DIR)) {
  fs.mkdirSync(DB_DIR, { recursive: true });
}
if (!fs.existsSync(UPLOADS_DIR)) {
  fs.mkdirSync(UPLOADS_DIR, { recursive: true });
}

// Serve uploaded images/PDFs statically
app.use('/uploads', express.static(UPLOADS_DIR));

// Word 문서 자동 컴파일 도구 (Google Drive 스트리밍 지원)
function toRoman(num) {
  const lookup = { M:1000,CM:900,D:500,CD:400,C:100,XC:90,L:50,XL:40,X:10,IX:9,V:5,IV:4,I:1 };
  let roman = '';
  for (let i in lookup) {
    while (num >= lookup[i]) {
      roman += i;
      num -= lookup[i];
    }
  }
  return roman;
}

function toCircled(num) {
  const circles = ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "⑪", "⑫", "⑬", "⑭", "⑮"];
  return circles[num - 1] || `(${num})`;
}

async function generateWordDocument(id, payload) {
  if (!id.startsWith('report_')) return;
  const parts = id.split('_');
  if (parts.length < 3) return;
  const parish = parts[1];
  const church = parts[2];
  
  // 1. Google Drive 스트리밍 경로 찾기
  let gDriveDir = null;
  const possiblePaths = [
    'C:\\Users\\note\\Google Drive 스트리밍\\내 드라이브',
    'G:\\내 드라이브',
    'G:\\My Drive',
    path.join(__dirname, 'google_drive_sync')
  ];

  for (const p of possiblePaths) {
    if (fs.existsSync(p)) {
      gDriveDir = path.join(p, '주간보고_제출현황');
      break;
    }
  }
  if (!gDriveDir) {
    gDriveDir = path.join(__dirname, 'google_drive_sync');
  }

  if (!fs.existsSync(gDriveDir)) {
    fs.mkdirSync(gDriveDir, { recursive: true });
  }

  // 2. Word 문서 빌드 시작
  const docChildren = [];

  // 타이틀
  docChildren.push(new Paragraph({
    children: [
      new TextRun({
        text: `${parish} - ${church} 주간보고서`,
        bold: true,
        size: 32, // 16pt
        color: "1D4ED8",
        font: "맑은 고딕"
      })
    ],
    spacing: { after: 300 }
  }));

  // 시간 기록
  const nowStr = new Date().toLocaleString('ko-KR');
  docChildren.push(new Paragraph({
    children: [
      new TextRun({
        text: `제출 시각: ${nowStr}`,
        size: 18, // 9pt
        color: "64748B",
        font: "맑은 고딕"
      })
    ],
    spacing: { after: 500 }
  }));

  const dataToUse = payload.data || [];
  const counters = [0, 0, 0, 0];

  for (const item of dataToUse) {
    let prefix = "";
    if (item.level === 0) {
      counters[0]++; counters[1] = 0; counters[2] = 0; counters[3] = 0;
      prefix = toRoman(counters[0]) + ". ";
    } else if (item.level === 1) {
      counters[1]++; counters[2] = 0; counters[3] = 0;
      prefix = counters[1] + ". ";
    } else if (item.level === 2) {
      counters[2]++; counters[3] = 0;
      prefix = counters[2] + ") ";
    } else if (item.level === 3) {
      counters[3]++;
      prefix = toCircled(counters[3]) + " ";
    }

    let color = "000000";
    if (item.level <= 1) color = "1D4ED8"; // blue-700
    
    const lines = `${prefix}${item.text || ''}`.split('\n');

    docChildren.push(new Paragraph({
      children: lines.map((line, idx) => 
        new TextRun({
          text: line,
          break: idx > 0 ? 1 : 0,
          bold: item.level <= 1,
          color: color,
          font: "맑은 고딕",
          size: 22, // 11pt
        })
      ),
      indent: {
        left: Math.max(0, (item.level === 0 ? 0 : (item.level - 1) * 360))
      },
      spacing: {
        before: item.level === 0 ? 360 : 120, 
      }
    }));

    // 표(Table) 구현
    if (item.tableData && item.tableData.length > 0) {
      const skippedCells = new Set();
      const docTable = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: item.tableData.map((row, rIdx) => {
          const cells = [];
          row.forEach((cellText, cIdx) => {
            if (skippedCells.has(`${rIdx},${cIdx}`)) return;
            
            const spanDef = item.tableSpans?.[rIdx]?.[cIdx];
            const colSpan = (typeof spanDef === 'number' ? spanDef : spanDef?.colspan) || 1;
            const rowSpan = (typeof spanDef === 'number' ? 1 : spanDef?.rowspan) || 1;
            
            for (let r = 0; r < rowSpan; r++) {
              for (let c = 0; c < colSpan; c++) {
                if (r === 0 && c === 0) continue;
                skippedCells.add(`${rIdx + r},${cIdx + c}`);
              }
            }
            
            const align = item.tableAlignments?.[rIdx]?.[cIdx] || 'left';
            const isHighlighted = item.tableHighlights?.[rIdx]?.[cIdx];
            
            cells.push(new TableCell({
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: cellText || '',
                      bold: isHighlighted,
                      color: isHighlighted ? "1D4ED8" : "000000",
                      font: "맑은 고딕",
                      size: 18
                    })
                  ],
                  alignment: align === 'center' ? 'center' : align === 'right' ? 'right' : 'left'
                })
              ],
              columnSpan: colSpan,
              rowSpan: rowSpan,
              shading: isHighlighted ? { fill: "EFF6FF" } : undefined
            }));
          });
          return new TableRow({ children: cells });
        })
      });
      docChildren.push(docTable);
    }
    
    // 사진(Image) 구현
    if (item.image && item.imageWidth && item.imageHeight) {
      try {
        let imageBuffer = null;
        if (item.image.startsWith("http")) {
          const urlParts = item.image.split('/uploads/');
          if (urlParts.length > 1) {
            const localPath = path.join(UPLOADS_DIR, urlParts[1]);
            if (fs.existsSync(localPath)) {
              imageBuffer = fs.readFileSync(localPath);
            }
          }
        } else if (item.image.startsWith("data:image")) {
          const base64Data = item.image.split(",")[1];
          imageBuffer = Buffer.from(base64Data, 'base64');
        }
        
        if (imageBuffer) {
          const targetWidth = 500;
          const ratio = Math.min(1, targetWidth / item.imageWidth);
          const finalWidth = item.imageWidth * ratio;
          const finalHeight = item.imageHeight * ratio;

          docChildren.push(new Paragraph({
            children: [
              new ImageRun({
                data: imageBuffer,
                transformation: {
                  width: finalWidth,
                  height: finalHeight
                },
                type: "png"
              })
            ],
            indent: {
              left: Math.max(0, (item.level === 0 ? 0 : (item.level - 1) * 360))
            },
            spacing: { before: 120 }
          }));
        }
      } catch (e) {
        console.error("[Word] Failed to embed image:", e);
      }
    }
  }

  // 3. Word 파일 물리 저장
  const doc = new Document({
    sections: [{
      properties: {},
      children: docChildren
    }]
  });

  try {
    const buffer = await Packer.toBuffer(doc);
    const cleanChurchName = church.replace(/[\/\\?%*:|"<>. ]/g, '_');
    const filename = `${cleanChurchName}_주간보고_${new Date().toISOString().slice(0, 10)}.docx`;

    // 로컬 경로 저장 (Google Drive 스트리밍 / 로컬 백업)
    const parishPath = path.join(gDriveDir, parish);
    if (!fs.existsSync(parishPath)) fs.mkdirSync(parishPath, { recursive: true });
    const finalFilePath = path.join(parishPath, filename);
    fs.writeFileSync(finalFilePath, buffer);
    console.log(`[GoogleDriveSync] Saved locally: ${finalFilePath}`);

    // Google Drive API 업로드
    if (driveClient) {
      try {
        const rootFolderId = GOOGLE_DRIVE_FOLDER_ID || await getOrCreateDriveFolder('주간보고_제출현황', null);
        const parishFolderId = await getOrCreateDriveFolder(parish, rootFolderId);
        await uploadFileToDrive(filename, buffer, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', parishFolderId);
        console.log(`[GoogleDrive] Word 파일 업로드 완료: ${parish}/${filename}`);
      } catch (driveErr) {
        console.error('[GoogleDrive] Word 업로드 실패:', driveErr.message);
      }
    }
  } catch (error) {
    console.error("[GoogleDriveSync] Failed to write Word file:", error);
  }
}

// 구글 드라이브 통합 폴더 저장을 위한 실시간 개별 텍스트 스트리밍 도구
async function saveTextDocument(id, payload) {
  if (!id.startsWith('report_')) return;
  const parts = id.split('_');
  if (parts.length < 3) return;
  const parish = parts[1];
  const church = parts[2];

  // 1. Google Drive 스트리밍 경로 찾기
  let gDriveDir = null;
  const possiblePaths = [
    'C:\\Users\\note\\Google Drive 스트리밍\\내 드라이브',
    'G:\\내 드라이브',
    'G:\\My Drive',
    path.join(__dirname, 'google_drive_sync')
  ];

  for (const p of possiblePaths) {
    if (fs.existsSync(p)) {
      gDriveDir = path.join(p, '주간보고_제출현황');
      break;
    }
  }
  if (!gDriveDir) {
    gDriveDir = path.join(__dirname, 'google_drive_sync');
  }

  if (!fs.existsSync(gDriveDir)) {
    fs.mkdirSync(gDriveDir, { recursive: true });
  }

  // 2. 가공이 쉬운 텍스트 파일 포맷 생성 (Gemini가 즉시 분석 가능하게 마크다운 접목)
  let textContent = `=========================================\n`;
  textContent += `교구: ${parish}\n`;
  textContent += `교회: ${church}\n`;
  textContent += `제출상태: ${payload.status === 'submitted' ? '제출 완료' : '임시 저장'}\n`;
  textContent += `업데이트시각: ${new Date().toLocaleString('ko-KR')}\n`;
  textContent += `=========================================\n\n`;

  const dataToUse = payload.data || [];
  const counters = [0, 0, 0, 0];
  let imageCounter = 0;

  for (const item of dataToUse) {
    let prefix = "";
    if (item.level === 0) {
      counters[0]++; counters[1] = 0; counters[2] = 0; counters[3] = 0;
      prefix = toRoman(counters[0]) + ". ";
    } else if (item.level === 1) {
      counters[1]++; counters[2] = 0; counters[3] = 0;
      prefix = counters[1] + ". ";
    } else if (item.level === 2) {
      counters[2]++; counters[3] = 0;
      prefix = counters[2] + ") ";
    } else if (item.level === 3) {
      counters[3]++;
      prefix = toCircled(counters[3]) + " ";
    }

    const indent = "  ".repeat(item.level);
    textContent += `${indent}${prefix}${item.text || ''}\n`;

    // 사진(Image) 추출 및 구글 드라이브 폴더 일대일 물리 저장 (Gemini가 파일 업로드 시 텍스트와 이미지 매핑 가능하도록 처리)
    if (item.image) {
      try {
        let imageBuffer = null;
        let ext = 'png';
        if (item.image.startsWith("http")) {
          const urlParts = item.image.split('/uploads/');
          if (urlParts.length > 1) {
            const localPath = path.join(UPLOADS_DIR, urlParts[1]);
            if (fs.existsSync(localPath)) {
              imageBuffer = fs.readFileSync(localPath);
              const extPart = urlParts[1].split('.');
              if (extPart.length > 1) ext = extPart[extPart.length - 1];
            }
          }
        } else if (item.image.startsWith("data:image")) {
          const mimePart = item.image.split(';')[0];
          if (mimePart.includes('jpeg') || mimePart.includes('jpg')) ext = 'jpg';
          else if (mimePart.includes('gif')) ext = 'gif';
          
          const base64Data = item.image.split(",")[1];
          imageBuffer = Buffer.from(base64Data, 'base64');
        }

        if (imageBuffer) {
          imageCounter++;
          const cleanParishName = parish.replace(/[\/\\?%*:|"<>. ]/g, '_');
          const cleanChurchName = church.replace(/[\/\\?%*:|"<>. ]/g, '_');
          const imgFilename = `[${cleanParishName}_${cleanChurchName}]_사진_${imageCounter}.${ext}`;
          const imgFilePath = path.join(gDriveDir, imgFilename);
          fs.writeFileSync(imgFilePath, imageBuffer);
          
          // 텍스트 파일 내에 사진 파일명을 마크다운 이미지 링크 포맷으로 박아두어 Gemini가 사진과 텍스트 문맥을 100% 매칭하도록 지원
          textContent += `${indent}![첨부사진](./${imgFilename})\n`;
          textContent += `${indent}[참부 사진 파일명: ${imgFilename}]\n`;
        }
      } catch (imgError) {
        console.error("[GoogleDriveSync] Failed to save inline image:", imgError);
      }
    }

    // 표 데이터 텍스트 포맷 (MarkDown 표 규격 준수 - Gemini 분석 정확도 100%)
    if (item.tableData && item.tableData.length > 0) {
      textContent += `\n${indent}[데이터 테이블]\n`;
      item.tableData.forEach((row, rIdx) => {
        let rowStr = indent + "| " + row.map(cell => cell.trim().replace(/\n/g, " ")).join(" | ") + " |";
        textContent += rowStr + "\n";
        if (rIdx === 0) {
          textContent += indent + "| " + row.map(() => "---").join(" | ") + " |\n";
        }
      });
      textContent += `\n`;
    }
  }

  try {
    const cleanParishName = parish.replace(/[\/\\?%*:|"<>. ]/g, '_');
    const cleanChurchName = church.replace(/[\/\\?%*:|"<>. ]/g, '_');
    const filename = `[${cleanParishName}_${cleanChurchName}]_주간보고.txt`;

    // 로컬 저장
    const finalFilePath = path.join(gDriveDir, filename);
    fs.writeFileSync(finalFilePath, textContent, 'utf-8');
    console.log(`[GoogleDriveSync] Saved locally: ${finalFilePath}`);

    // Google Drive API 업로드 (텍스트 파일 → AI Studio에서 바로 열기 가능)
    if (driveClient) {
      try {
        const rootFolderId = GOOGLE_DRIVE_FOLDER_ID || await getOrCreateDriveFolder('주간보고_제출현황', null);
        const textBuffer = Buffer.from(textContent, 'utf-8');
        await uploadFileToDrive(filename, textBuffer, 'text/plain', rootFolderId);
        console.log(`[GoogleDrive] 텍스트 파일 업로드 완료: ${filename}`);
      } catch (driveErr) {
        console.error('[GoogleDrive] 텍스트 업로드 실패:', driveErr.message);
      }
    }
  } catch (error) {
    console.error("[GoogleDriveSync] Failed to write text file:", error);
  }
}

// 1. Save Report Data
app.post('/api/save-data', (req, res) => {
  const { id, payload } = req.body;
  if (!id) {
    return res.status(400).json({ error: 'Missing ID' });
  }
  try {
    const filePath = path.join(DB_DIR, `${id}.json`);
    fs.writeFileSync(filePath, JSON.stringify(payload, null, 2), 'utf-8');
    console.log(`[DB] Saved data for ID: ${id}`);
    
    // 실시간 작성 또는 제출 시 비동기로 개별 텍스트 파일 저장해 구글 드라이브 동기화
    saveTextDocument(id, payload).catch(err => {
      console.error('[GoogleDriveSync] Async text generation failed:', err);
    });

    // 제출 확정 시 비동기로 Word 문서 자동 빌드하여 구글 드라이브 동기화 폴더로 전송
    if (payload && payload.status === 'submitted') {
      generateWordDocument(id, payload).catch(err => {
        console.error('[GoogleDriveSync] Async generation failed:', err);
      });
    }

    res.json({ success: true });
  } catch (error) {
    console.error('Failed to save local data:', error);
    res.status(500).json({ error: 'Failed to save data' });
  }
});

// 2. Load Report Data
app.get('/api/load-data/:id', (req, res) => {
  const { id } = req.params;
  try {
    const filePath = path.join(DB_DIR, `${id}.json`);
    if (fs.existsSync(filePath)) {
      const raw = fs.readFileSync(filePath, 'utf-8');
      res.json(JSON.parse(raw));
    } else {
      res.status(404).json({ error: 'Not found' });
    }
  } catch (error) {
    console.error('Failed to load local data:', error);
    res.status(500).json({ error: 'Failed to load data' });
  }
});

// 3. Raw Upload for Images/PDFs (Zero-dependency binary handler)
app.post(
  '/api/upload-image',
  express.raw({ type: ['image/jpeg', 'image/png', 'image/gif', 'application/pdf'], limit: '20mb' }),
  (req, res) => {
    const contentType = req.headers['content-type'];
    const clientFilename = req.headers['x-filename'];
    
    let extension = 'jpg';
    if (contentType === 'application/pdf') extension = 'pdf';
    else if (contentType === 'image/png') extension = 'png';
    else if (contentType === 'image/gif') extension = 'gif';

    const filename = clientFilename 
      ? String(clientFilename) 
      : `upload_${Date.now()}.${extension}`;

    const filePath = path.join(UPLOADS_DIR, filename);

    try {
      fs.writeFileSync(filePath, req.body);
      
      // Serve via actual server IP or hostname dynamically
      const host = req.headers.host || `localhost:${PORT}`;
      const url = `http://${host.split(':')[0]}:${PORT}/uploads/${filename}`;
      
      console.log(`[Upload] File saved to ${filePath} -> ${url}`);
      res.json({ url });
    } catch (error) {
      console.error('Upload failed:', error);
      res.status(500).json({ error: 'Failed to write upload file' });
    }
  }
);

// Ping — 로컬 서버 자동 감지용
app.get('/api/ping', (req, res) => res.json({ ok: true, mode: 'local' }));

// 4a. Claude Code CLI Proxy (Claude Pro 구독자용 — claude -p 비대화형 실행)
app.post('/api/claude-chat', async (req, res) => {
  const { prompt } = req.body;
  if (!prompt) return res.status(400).json({ error: 'Missing prompt' });

  // 임시 파일에 프롬프트 저장 후 stdin 리다이렉트 (특수문자/한글/줄바꿈 안전)
  const tmpFile = path.join(os.tmpdir(), `claude_prompt_${Date.now()}.txt`);
  try {
    fs.writeFileSync(tmpFile, prompt, 'utf8');
  } catch (e) {
    return res.status(500).json({ error: '임시 파일 생성 실패' });
  }

  // Windows: cmd /c "type file | claude --print"
  // Unix:    sh -c "claude --print < file"
  const isWin = process.platform === 'win32';
  // chcp 65001 = UTF-8 코드페이지로 전환 후 실행
  const cmd = isWin
    ? `cmd /c "chcp 65001 > nul && type "${tmpFile}" | claude --print"`
    : `claude --print < "${tmpFile}"`;

  const cwd = os.tmpdir();
  exec(cmd, { timeout: 120000, maxBuffer: 10 * 1024 * 1024, cwd }, (err, stdout, stderr) => {
    try { fs.unlinkSync(tmpFile); } catch {}
    if (err) {
      console.error('[ClaudeCode] error:', stderr || err.message);
      return res.status(500).json({ error: stderr || err.message });
    }
    console.log('[ClaudeCode] success, length:', stdout.length);
    res.json({ text: stdout.trim() });
  });
});

// 4. Dynamic Ollama AI Proxy (handles CORS & automatic model selection)
app.post('/api/ollama-chat', async (req, res) => {
  const { prompt } = req.body;
  if (!prompt) {
    return res.status(400).json({ error: 'Missing prompt' });
  }

  // A. List local Ollama models and find the best match
  let modelToUse = 'gemma2:9b'; // Default fallback
  try {
    const tagsRes = await fetch('http://localhost:11434/api/tags');
    if (tagsRes.ok) {
      const tagsData = await tagsRes.json();
      const models = tagsData.models || [];
      if (models.length > 0) {
        const modelNames = models.map(m => m.name);
        // Preference list: gemma2, qwen2.5, llama3.1, llama3, etc.
        const preferred = ['gemma2:9b', 'gemma2', 'qwen2.5:7b-instruct', 'qwen2.5', 'llama3.1', 'llama3', 'llama2', 'mistral'];
        const found = preferred.find(pref => modelNames.some(name => name.startsWith(pref)));
        if (found) {
          const exactMatch = modelNames.find(name => name.startsWith(found));
          modelToUse = exactMatch;
        } else {
          // If no preferred model, use the first model found
          modelToUse = modelNames[0];
        }
      }
    }
  } catch (err) {
    console.warn('[Ollama] Could not query models list, using default model: gemma2:9b');
  }

  console.log(`[Ollama] Selected model for request: "${modelToUse}"`);

  // B. Call local Ollama chat completions endpoint
  try {
    const ollamaResponse = await fetch('http://localhost:11434/v1/chat/completions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: modelToUse,
        messages: [{ role: 'user', content: prompt }],
        response_format: { type: 'json_object' },
        temperature: 0.1 // low temperature for strict json output consistency
      })
    });

    if (!ollamaResponse.ok) {
      const errText = await ollamaResponse.text();
      return res.status(ollamaResponse.status).json({ error: `Ollama error: ${errText}` });
    }

    const result = await ollamaResponse.json();
    const textResult = result.choices?.[0]?.message?.content || '{}';
    res.json({ text: textResult, model: modelToUse });
  } catch (error) {
    console.error('[Ollama] Proxy request failed:', error);
    res.status(500).json({ error: 'Could not connect to Ollama. Make sure Ollama is running on your PC (http://localhost:11434).' });
  }
});

// ─── Google Drive OAuth2 설정 엔드포인트 ────────────────────────────────────

// GET /api/google-auth/status  → 현재 Drive 연동 상태 + 설정에 필요한 URL 반환
app.get('/api/google-auth/status', (req, res) => {
  const callbackUrl = `${APP_URL}/api/google-auth/callback`;
  const authUrl = `${APP_URL}/api/google-auth`;
  res.json({
    configured: !!(GOOGLE_CLIENT_ID && GOOGLE_CLIENT_SECRET),
    authenticated: !!driveClient,
    hasRefreshToken: !!GOOGLE_REFRESH_TOKEN,
    folderId: GOOGLE_DRIVE_FOLDER_ID || null,
    appUrl: APP_URL,
    callbackUrl,
    authUrl,
  });
});

// GET /api/google-auth  → Google OAuth 동의 화면으로 이동 (최초 1회 인증)
app.get('/api/google-auth', (req, res) => {
  if (!oauth2Client) {
    return res.status(400).send(`<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">
      <title>설정 필요</title>
      <style>
        body{font-family:'Pretendard',sans-serif;max-width:640px;margin:60px auto;padding:24px;color:#1e293b}
        .card{background:#fef2f2;border:1px solid #fecaca;border-radius:12px;padding:24px;margin-bottom:16px}
        .step{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:16px;margin:8px 0}
        code{background:#f1f5f9;padding:3px 8px;border-radius:4px;font-family:monospace;font-size:13px;word-break:break-all}
        h2{color:#dc2626} h3{color:#475569;font-size:14px;margin:0 0 8px}
        p{font-size:14px;line-height:1.6;margin:6px 0}
      </style></head><body>
      <h2>⚠️ Google Cloud 인증 정보 미설정</h2>
      <div class="card">
        <p>Google AI Studio의 <strong>Secrets 패널</strong>에 아래 두 값을 먼저 등록하세요.</p>
      </div>
      <div class="step"><h3>등록해야 할 Secret 이름</h3>
        <p><code>GOOGLE_CLIENT_ID</code> — Google Cloud OAuth 클라이언트 ID</p>
        <p><code>GOOGLE_CLIENT_SECRET</code> — Google Cloud OAuth 클라이언트 보안 비밀번호</p>
      </div>
      <div class="step"><h3>등록 위치</h3>
        <p>Google AI Studio → 앱 편집기 우측 상단 <strong>Secrets</strong> 탭 → + 추가</p>
      </div>
      <p style="margin-top:20px;font-size:13px;color:#64748b">Secret 등록 후 앱을 재배포하고 다시 이 주소를 열어주세요.</p>
      </body></html>`);
  }
  const authUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    prompt: 'consent',
    scope: ['https://www.googleapis.com/auth/drive.file']
  });
  res.redirect(authUrl);
});

// GET /api/google-auth/callback  → 인증 후 Google이 리다이렉트하는 콜백
app.get('/api/google-auth/callback', async (req, res) => {
  const { code, error } = req.query;
  if (error || !code) {
    return res.status(400).send(`<h2>인증 실패: ${error || '코드 없음'}</h2>`);
  }
  try {
    const { tokens } = await oauth2Client.getToken(code);
    oauth2Client.setCredentials(tokens);
    driveClient = google.drive({ version: 'v3', auth: oauth2Client });

    const refreshToken = tokens.refresh_token;
    res.send(`<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">
      <title>Google Drive 연동 완료</title>
      <style>
        body{font-family:'Pretendard',sans-serif;max-width:680px;margin:60px auto;padding:24px;color:#1e293b}
        h2{color:#16a34a;margin-bottom:8px}
        .green{background:#f0fdf4;border:1px solid #bbf7d0;border-radius:12px;padding:20px;margin:16px 0}
        .blue{background:#eff6ff;border:1px solid #bfdbfe;border-radius:12px;padding:20px;margin:16px 0}
        .step{display:flex;gap:12px;align-items:flex-start;margin:10px 0}
        .num{width:24px;height:24px;border-radius:50%;background:#6366f1;color:#fff;font-size:11px;font-weight:900;display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:2px}
        code{background:#f1f5f9;padding:4px 10px;border-radius:6px;font-family:monospace;font-size:13px;word-break:break-all;display:block;margin-top:6px}
        button{background:#6366f1;color:#fff;border:none;padding:8px 16px;border-radius:8px;font-size:13px;font-weight:700;cursor:pointer;margin-top:8px}
        button:hover{background:#4f46e5}
        p{font-size:14px;line-height:1.7;margin:4px 0}
        .warn{color:#92400e;background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:12px;font-size:13px;margin-top:8px}
      </style></head><body>
      <h2>✅ Google 계정 인증 완료!</h2>
      <p style="color:#475569">이제 아래 Refresh Token을 Google AI Studio Secrets에 등록해야 자동 업로드가 작동합니다.</p>

      <div class="green">
        <p><strong>📋 복사할 Refresh Token</strong></p>
        ${refreshToken
          ? `<code id="token">${refreshToken}</code>
             <button onclick="navigator.clipboard.writeText('${refreshToken}').then(()=>this.textContent='✅ 복사됨!')">클립보드 복사</button>`
          : `<p class="warn">⚠️ Refresh Token이 발급되지 않았습니다. <br>이미 이 계정으로 인증한 적이 있으면 <a href="/api/google-auth">여기서 다시 인증</a>하세요 (consent 화면에서 계정 재선택).</p>`
        }
      </div>

      <div class="blue">
        <p><strong>📌 AI Studio Secrets에 등록하는 방법</strong></p>
        <div class="step"><div class="num">1</div><p>Google AI Studio → 현재 앱 편집기 열기</p></div>
        <div class="step"><div class="num">2</div><p>우측 상단 <strong>Secrets</strong> 탭 클릭</p></div>
        <div class="step"><div class="num">3</div><p>이름: <code>GOOGLE_REFRESH_TOKEN</code><br>값: 위 토큰 붙여넣기 → 저장</p></div>
        <div class="step"><div class="num">4</div><p>(선택) 업로드할 Drive 폴더 ID가 있으면:<br><code>GOOGLE_DRIVE_FOLDER_ID</code> 도 함께 등록</p></div>
        <div class="step"><div class="num">5</div><p>앱 <strong>재배포</strong> → 관리자 콘솔에서 연동 상태 🟢 확인</p></div>
      </div>

      <p style="font-size:13px;color:#94a3b8;margin-top:16px">이 창은 닫아도 됩니다.</p>
      </body></html>`);
  } catch (err) {
    console.error('[GoogleDrive] Token exchange failed:', err);
    res.status(500).send(`<h2>토큰 교환 실패</h2><pre>${err.message}</pre>`);
  }
});

// ─────────────────────────────────────────────────────────────────────────────

// 4.5 Open Sync Folder inside File Explorer
app.post('/api/open-folder', async (req, res) => {
  try {
    const { exec } = await import('child_process');
    
    // Find Google Drive folder
    let gDriveDir = null;
    const possiblePaths = [
      'C:\\Users\\note\\Google Drive 스트리밍\\내 드라이브',
      'G:\\내 드라이브',
      'G:\\My Drive',
      path.join(__dirname, 'google_drive_sync')
    ];

    for (const p of possiblePaths) {
      if (fs.existsSync(p)) {
        gDriveDir = path.join(p, '주간보고_제출현황');
        break;
      }
    }
    if (!gDriveDir) {
      gDriveDir = path.join(__dirname, 'google_drive_sync');
    }

    if (!fs.existsSync(gDriveDir)) {
      fs.mkdirSync(gDriveDir, { recursive: true });
    }

    // Run platform-specific shell command
    let command = '';
    if (process.platform === 'win32') {
      command = `explorer.exe "${gDriveDir}"`;
    } else if (process.platform === 'darwin') {
      command = `open "${gDriveDir}"`;
    } else {
      command = `xdg-open "${gDriveDir}"`;
    }

    exec(command, (err) => {
      if (err) {
        console.error('[OpenFolder] Failed to open folder:', err);
        return res.status(500).json({ error: 'Folder opening failed' });
      }
      res.json({ success: true, path: gDriveDir });
    });
  } catch (error) {
    console.error('[OpenFolder] Error:', error);
    res.status(500).json({ error: error.message });
  }
});

// 5. Statically serve the React build (for high-performance single-port local production)
const DIST_DIR = path.join(__dirname, 'dist');
if (fs.existsSync(DIST_DIR)) {
  app.use(express.static(DIST_DIR));
  app.get('*', (req, res, next) => {
    // API and uploads routes should bypass index.html static serving
    if (req.path.startsWith('/api/') || req.path.startsWith('/uploads/')) {
      return next();
    }
    res.sendFile(path.join(DIST_DIR, 'index.html'));
  });
  console.log(`[Static] Serving React production build from ./dist`);
}

app.listen(PORT, '0.0.0.0', async () => {
  console.log(`=================================================`);
  console.log(`🏠 Weekly Report Local Server is running on:`);
  console.log(`   - Local PC:   http://localhost:${PORT}`);
  console.log(`=================================================`);

  // ngrok URL 자동 감지 후 Supabase에 등록
  const registerNgrok = async () => {
    for (let i = 0; i < 10; i++) {
      try {
        const res = await fetch('http://localhost:4040/api/tunnels');
        if (res.ok) {
          const data = await res.json();
          const url = data?.tunnels?.[0]?.public_url;
          if (url) {
            const payload = { id: 'SYSTEM_SERVER_URL', url, updated_at: new Date().toISOString() };
            const json = JSON.stringify(payload);
            await sb.storage.from('images').upload('db_reports/SYSTEM_SERVER_URL.json', json, { upsert: true, contentType: 'application/json' });
            console.log(`✅ ngrok URL 등록됨: ${url}`);
            return;
          }
        }
      } catch {}
      await new Promise(r => setTimeout(r, 2000));
    }
    // ngrok 없으면 로컬 IP만 등록
    const payload = { id: 'SYSTEM_SERVER_URL', url: null, updated_at: new Date().toISOString() };
    await sb.storage.from('images').upload('db_reports/SYSTEM_SERVER_URL.json', JSON.stringify(payload), { upsert: true, contentType: 'application/json' }).catch(() => {});
    console.log('⚠️ ngrok 미감지 — 로컬 전용 모드');
  };

  registerNgrok();
});
