import { google } from 'googleapis';
import { Readable } from 'stream';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ImageRun } from 'docx';

// App.tsx의 PARISH_CHURCH_MAP과 동일 — 교구별 교회 순서 기준
const PARISH_CHURCH_MAP = {
  "협회": ["기획국", "미래인재국", "가정행복국", "총무국", "문화홍보국", "사회공헌국", "대외협력국"],
  "천원특별": ["교구본부", "천원궁 천원", "가평", "청평", "북면가정", "조종가정", "상면가정"],
  "서울북부": ["교구본부", "천원궁 천승", "천승(1구역)", "천승(2구역)", "천승(3구역)", "광진", "노원", "동대문", "성북", "안암학사", "신촌학사", "한양학사", "은평", "서대문", "중구", "종로", "중랑", "강북", "장안", "도봉", "청파", "광화문학사", "HJ글로벌"],
  "서울남부": ["교구본부", "강남", "영등포", "강동", "양천", "강서", "관악", "구로", "금천", "명일", "송파", "흑석동작"],
  "경기북부": ["교구본부", "구리", "남양주", "일산", "금촌", "덕양", "동두천", "양주", "연천", "의정부", "파주", "포천", "화도", "양평"],
  "인천경기서부": ["교구본부", "인천", "강화", "계양", "김포", "남부천", "부천", "부평", "서구", "주안", "안산", "광명", "군포"],
  "경기남부": ["교구본부", "수원", "수원학사", "이천", "일신", "평택", "과천", "광주(경기남부)", "기흥", "성남", "안성", "안양", "야목화성", "여주", "오산", "용인", "하남"],
  "강원": ["교구본부", "춘천", "춘천전도소", "원주", "양구", "강릉", "고성(강원)", "동해", "삼척", "속초", "양양", "영월", "인제", "정선", "철원", "태백", "평창", "홍천", "화천", "횡성"],
  "대전충남": ["교구본부", "천안", "선문대학", "선문학사", "공주", "당진", "병천", "보령", "부여", "서산", "서천", "아산", "예산", "청양", "태안", "홍성", "대전", "금산", "논산", "대덕", "대전중앙", "세종", "유성"],
  "충북": ["교구본부", "청주", "충주", "괴산", "금왕", "남일", "단양", "미원", "보은", "영동", "옥천", "음성", "제천", "증평", "진천", "청원", "광혜원전도소"],
  "전북": ["교구본부", "전주", "남원", "익산", "고창", "군산", "김제", "무주", "부안", "순창", "완주", "임실", "장수", "정읍", "진안"],
  "광주전남제주": ["교구본부", "광주", "광주청년센터", "광주학사", "나주", "목포", "화순", "강진", "곡성", "광산", "구례", "남광주", "담양", "무안", "보성", "영광", "영암", "완도", "장성", "장흥", "진도", "함평", "해남", "해양여수", "광양", "거문도", "고흥", "서순천", "순천", "제주", "제주학사", "서귀포"],
  "대구경북": ["교구본부", "대구", "경주", "동대구", "경산", "고령", "구미", "군위", "김천", "달성", "문경", "봉화", "상주", "선산", "성주", "수성", "안강", "안동", "영덕", "영양", "영주", "영천", "예천", "울릉", "울진", "의성", "청도", "청송", "칠곡", "포항"],
  "경남": ["교구본부", "창원", "동창원", "마산", "거제", "거창", "고성", "김해", "남해", "밀양", "사천", "산청", "양산", "의령", "진주", "진주학사", "진해", "창녕", "통영", "하동", "함안", "함양", "합천"],
  "부산울산": ["교구본부", "부산", "부산청년센터", "부산학사", "남부산", "동부산", "북부산", "서부산", "울산", "동울산", "울주"],
};

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
// 교회명 → 파일 키: 공백과 괄호를 _로 치환
function toKey(church) {
  return church.replace(/[ ()]/g, '_');
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
  const safe = filename.replace(/'/g, "\\'");
  const stream = Readable.from(buffer);

  if (filename.endsWith('.docx')) {
    // Word 파일은 날짜가 달라도 같은 교구/교회 파일이면 기존 것 전부 삭제 후 새로 업로드
    const prefix = filename.replace(/_\d{4}-\d{2}-\d{2}\.docx$/, '');
    const q = `name contains '${prefix.replace(/'/g, "\\'")}' and '${folderId}' in parents and trashed=false`;
    const existing = await drive.files.list({ q, fields: 'files(id)', spaces: 'drive' });
    await Promise.all(existing.data.files.map(f => drive.files.delete({ fileId: f.id }).catch(() => {})));
    await drive.files.create({ requestBody: { name: filename, parents: [folderId] }, media: { mimeType, body: stream }, fields: 'id' });
  } else {
    const q = `name='${safe}' and '${folderId}' in parents and trashed=false`;
    const existing = await drive.files.list({ q, fields: 'files(id)', spaces: 'drive' });
    if (existing.data.files.length > 0) {
      await drive.files.update({ fileId: existing.data.files[0].id, media: { mimeType, body: stream } });
    } else {
      await drive.files.create({ requestBody: { name: filename, parents: [folderId] }, media: { mimeType, body: stream }, fields: 'id' });
    }
  }
}

async function downloadJson(drive, fileId) {
  const res = await drive.files.get({ fileId, alt: 'media' }, { responseType: 'arraybuffer' });
  return JSON.parse(Buffer.from(res.data).toString('utf-8'));
}

// 교구 폴더의 모든 _data.json 파일을 읽어 { key: payload } 맵 반환
async function loadParishSubmissions(drive, parishFolderId) {
  const q = `name contains '_data.json' and '${parishFolderId}' in parents and trashed=false`;
  const res = await drive.files.list({ q, fields: 'files(id,name)', spaces: 'drive' });
  const submissions = {};
  for (const file of res.data.files) {
    const key = file.name.replace('_data.json', '');
    try {
      submissions[key] = await downloadJson(drive, file.id);
    } catch (e) {
      console.error(`[Drive] Failed to load ${file.name}:`, e.message);
    }
  }
  return submissions;
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

// 교회 1개 섹션의 단락 배열 반환
async function buildChurchSection(sectionNum, church, payload) {
  const paragraphs = [];

  // 교회 헤더
  paragraphs.push(new Paragraph({
    children: [new TextRun({
      text: `${toRoman(sectionNum)}. ${getDisplayChurch(church)}`,
      bold: true, size: 28, color: '1E3A8A', font: '맑은 고딕',
    })],
    spacing: { before: 600, after: 200 },
    border: { bottom: { style: 'single', size: 6, color: 'BFDBFE' } },
  }));

  const KOR_CONSONANTS = ['가','나','다','라','마','바','사','아','자','차','카','타','파','하'];
  const counters = [0, 0, 0, 0, 0, 0];
  for (const item of (payload.data || [])) {
    let prefix = '';
    if (item.level === 0) { counters[0]++; counters[1]=0; counters[2]=0; counters[3]=0; counters[4]=0; counters[5]=0; prefix = toRoman(counters[0]) + '. '; }
    else if (item.level === 1) { counters[1]++; counters[2]=0; counters[3]=0; counters[4]=0; counters[5]=0; prefix = counters[1] + '. '; }
    else if (item.level === 2) { counters[2]++; counters[3]=0; counters[4]=0; counters[5]=0; prefix = counters[2] + ') '; }
    else if (item.level === 3) { counters[3]++; counters[4]=0; counters[5]=0; prefix = toCircled(counters[3]) + ' '; }
    else if (item.level === 4) { counters[4]++; counters[5]=0; prefix = (KOR_CONSONANTS[counters[4]-1]||`(${counters[4]})`) + '. '; }
    else if (item.level === 5) { counters[5]++; prefix = String.fromCharCode(96+counters[5]) + '. '; }

    const color = item.level === 0 ? '1D4ED8' : '000000';
    const indent = { left: item.level === 0 ? 0 : item.level * 360 };
    const lines = `${prefix}${item.text || ''}`.split('\n');
    paragraphs.push(new Paragraph({
      children: lines.map((line, idx) => new TextRun({ text: line, break: idx>0?1:0, bold: item.level===0, color, font:'맑은 고딕', size:20 })),
      indent,
      spacing: { before: item.level===0 ? 240 : 80 },
    }));

    if (item.image && item.imageWidth && item.imageHeight) {
      const imgBuf = await fetchImageBuffer(item.image);
      if (imgBuf) {
        const maxPx = 580;
        const ratio = Math.min(1, maxPx / item.imageWidth);
        const w = Math.round(item.imageWidth * ratio);
        const h = Math.round(item.imageHeight * ratio);
        const type = item.image.startsWith('data:image/png') ? 'png' : 'jpg';
        try {
          paragraphs.push(new Paragraph({
            children: [new ImageRun({ data: imgBuf, transformation: { width: w, height: h }, type })],
            indent,
            spacing: { before: 100, after: 100 },
          }));
        } catch (e) {
          console.error('[Drive] ImageRun failed:', e.message);
        }
      }
    }

    if (item.tableData?.length > 0) {
      const skipped = new Set();
      paragraphs.push(new Table({
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
  return paragraphs;
}

// 교구 통합 Word: 제출된 교회들을 PARISH_CHURCH_MAP 순서로 번호 붙여 이어쓰기
async function buildParishWordBuffer(parish, submissions) {
  const displayParish = getDisplayParish(parish);
  const children = [];
  const today = new Date().toLocaleDateString('ko-KR');
  const churches = PARISH_CHURCH_MAP[parish] || [];
  const submittedCount = churches.filter(c => submissions[toKey(c)]).length;

  children.push(new Paragraph({
    children: [new TextRun({ text: `${displayParish} 주간업무보고서`, bold: true, size: 40, color: '1E3A8A', font: '맑은 고딕' })],
    spacing: { after: 200 },
  }));
  children.push(new Paragraph({
    children: [new TextRun({ text: `취합일: ${today}　|　제출 현황: ${submittedCount}/${churches.length}개`, size: 20, color: '64748B', font: '맑은 고딕' })],
    spacing: { after: 800 },
  }));

  let sectionNum = 0;
  for (const church of churches) {
    const payload = submissions[toKey(church)];
    if (!payload) continue;
    sectionNum++;
    const section = await buildChurchSection(sectionNum, church, payload);
    children.push(...section);
  }

  const doc = new Document({ sections: [{ properties: {}, children }] });
  return Packer.toBuffer(doc);
}

// 협회 국 단독 Word 문서
async function buildGukWordBuffer(guk, payload) {
  const children = [];
  children.push(new Paragraph({
    children: [new TextRun({ text: `협회본부 ${guk} 주간업무보고서`, bold: true, size: 36, color: '1E3A8A', font: '맑은 고딕' })],
    spacing: { after: 300 },
  }));
  children.push(new Paragraph({
    children: [new TextRun({ text: `제출 시각: ${new Date().toLocaleString('ko-KR')}`, size: 18, color: '64748B', font: '맑은 고딕' })],
    spacing: { after: 500 },
  }));
  const section = await buildChurchSection(1, guk, payload);
  children.push(...section);
  const doc = new Document({ sections: [{ properties: {}, children }] });
  return Packer.toBuffer(doc);
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
  // payload에 parish/church가 명시돼 있으면 직접 사용, 없으면 id에서 추출
  const parish = payload?.parish || id.replace(/^report_/, '').split('_')[0];
  const church = payload?.church || id.replace(/^report_/, '').split('_').slice(1).join('_');
  if (!parish || !church) return res.status(200).json({ skipped: true });
  const cleanChurch = toKey(church);

  try {
    const rootId = process.env.GOOGLE_DRIVE_FOLDER_ID || await getOrCreateFolder(drive, '주간보고_제출현황', null);

    if (isHyeohoe(parish)) {
      // 협회: 국별 독립 폴더 + 단독 Word
      const hyeohoeId = await getOrCreateFolder(drive, '협회본부', rootId);
      const gukId = await getOrCreateFolder(drive, church, hyeohoeId);

      await uploadOrUpdate(drive, `${cleanChurch}_data.json`, Buffer.from(JSON.stringify(payload), 'utf-8'), 'application/json', gukId);

      if (payload?.status === 'submitted') {
        const wordBuffer = await buildGukWordBuffer(church, payload);
        const filename = `협회본부_${cleanChurch}_주간보고_${new Date().toISOString().slice(0,10)}.docx`;
        await uploadOrUpdate(drive, filename, wordBuffer, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', gukId);
      }
    } else {
      // 교구: 교회별 JSON 저장 + 교구 단일 통합 Word 재생성
      const parishDisplay = getDisplayParish(parish);
      const parishFolderId = await getOrCreateFolder(drive, parishDisplay, rootId);

      // 교회 JSON 실시간 저장 (저장/제출 모두)
      await uploadOrUpdate(drive, `${cleanChurch}_data.json`, Buffer.from(JSON.stringify(payload), 'utf-8'), 'application/json', parishFolderId);

      // 제출 시: 교구 전체 교회 데이터로 통합 Word 재생성
      if (payload?.status === 'submitted') {
        const submissions = await loadParishSubmissions(drive, parishFolderId);
        submissions[cleanChurch] = payload; // 방금 제출분 즉시 반영

        const wordBuffer = await buildParishWordBuffer(parish, submissions);
        const filename = `${parishDisplay}_주간보고_${new Date().toISOString().slice(0,10)}.docx`;
        await uploadOrUpdate(drive, filename, wordBuffer, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', parishFolderId);
      }
    }

    res.json({ success: true });
  } catch (err) {
    console.error('[Drive] Upload failed:', err.message);
    res.status(500).json({ error: err.message });
  }
}
