import React, { useState, useEffect, useRef } from 'react';
import { Plus, X, ArrowLeft, ArrowRight, ArrowUp, ArrowDown, FileJson, Copy, Check, Save, Download, Bot, Clock, AlertCircle, RefreshCw, Image as ImageIcon, Crop as CropIcon, Table as TableIcon, BarChart2, Trash2, Highlighter, BookOpen, AlignLeft, AlignCenter, AlignRight, Settings, Key, Bell, Upload, FileText, Sparkles, Folder } from 'lucide-react';
import { GoogleGenAI, Type } from '@google/genai';
import { Document, Packer, Paragraph, TextRun, HeadingLevel, ImageRun, Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign } from "docx";
import TextareaAutosize from 'react-textarea-autosize';
import ReactCrop, { Crop, PixelCrop, centerCrop, makeAspectCrop } from 'react-image-crop';
import 'react-image-crop/dist/ReactCrop.css';
import { BarChart, Bar, LineChart, Line, PieChart, Pie, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, Cell } from 'recharts';
import { parseHtmlTable } from './lib/tableParser';
import { supabase } from './lib/supabase';

const getLocalServerUrl = () => {
  const customUrl = localStorage.getItem('LOCAL_SERVER_URL');
  if (customUrl) return customUrl;
  return 'http://localhost:5000';
};

// RLS 우회를 위한 Storage 기반 JSON DB 헬퍼 (로컬 및 클라우드 동시 지원)
const fetchDbData = async (id: string) => {
  const isLocal = localStorage.getItem('IS_LOCAL_MODE') === 'true';
  if (isLocal) {
    try {
      const serverUrl = getLocalServerUrl();
      const res = await fetch(`${serverUrl}/api/load-data/${id}`);
      if (res.ok) {
        return await res.json();
      }
    } catch (e) {
      console.warn("Local DB load failed, trying localStorage fallback:", e);
    }
    return null;
  }

  if (!supabase) return null;
  const { data, error } = await supabase.storage.from('images').download(`db_reports/${id}.json`);
  if (!error && data) {
    try {
      const text = await data.text();
      return JSON.parse(text);
    } catch(e) {}
  }
  return null;
};

const saveDbData = async (id: string, payload: any) => {
  const isLocal = localStorage.getItem('IS_LOCAL_MODE') === 'true';
  if (isLocal) {
    try {
      const serverUrl = getLocalServerUrl();
      await fetch(`${serverUrl}/api/save-data`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id, payload })
      });
    } catch (e) {
      console.error("Local DB save failed:", e);
    }
    return;
  }

  // 클라우드 모드: Supabase 저장 + Google Drive 업로드
  if (supabase) {
    try {
      const jsonString = JSON.stringify(payload);
      await supabase.storage.from('images').upload(`db_reports/${id}.json`, jsonString, { upsert: true, contentType: 'application/json' });
    } catch(e) {
      console.error("Storage DB save failed:", e);
    }
  }

  // Vercel 서버리스 함수로 Drive 업로드 — report_ 접두사 포함해서 전달
  const driveId = id.startsWith('report_') ? id : `report_${id}`;
  try {
    const driveRes = await fetch('/api/save-drive', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id: driveId, payload })
    });
    const driveJson = await driveRes.json();
    return driveJson?.success ? 'ok' : driveJson?.skipped ? 'skipped' : 'error';
  } catch {
    return 'error';
  }
};


interface ReportItem {
  id: number;
  text: string;
  level: number;
  isFixed?: boolean;
  image?: string;
  imageWidth?: number;
  imageHeight?: number;
  tableData?: string[][];
  tableHighlights?: boolean[][];
  tableSpans?: any[][];
  tableAlignments?: ('left' | 'center' | 'right' | undefined)[][];
  chartType?: 'none' | 'bar' | 'line' | 'pie';
  chartColumnSelection?: boolean[];
}

const PARISH_CHURCH_MAP: Record<string, string[]> = {
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
  "부산울산": ["교구본부", "부산", "부산청년센터", "부산학사", "남부산", "동부산", "북부산", "서부산", "울산", "동울산", "울주"]
};

// Helper functions for display
const getDisplayParish = (parishName: string) => {
  if (parishName === '협회') return '협회본부';
  return parishName.endsWith('교구') ? parishName : `${parishName}교구`;
};
const getDisplayChurch = (churchName: string) => {
  if (churchName === '교구본부' || churchName.endsWith('국')) return churchName;
  return churchName.endsWith('교회') || churchName.endsWith('학사') || churchName.endsWith('글로벌') || churchName.endsWith('센터') || churchName.endsWith('대학') || churchName.endsWith('전도소') ? churchName : `${churchName}교회`;
};

const DEFAULT_REPORT: ReportItem[] = [
  { id: 1, text: "전주 결과보고", level: 0 },
  { id: 2, text: "", level: 1 },
  { id: 3, text: "금주 계획 및 보고", level: 0 },
  { id: 4, text: "", level: 1 }
];

const toRoman = (num: number) => {
  const roman = ["", "Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ", "Ⅶ", "Ⅷ", "Ⅸ", "Ⅹ"];
  return roman[num] || num.toString();
};

const toCircled = (num: number) => {
  if (num >= 1 && num <= 15) return String.fromCharCode(9311 + num);
  return `(${num})`;
};

const KOR_SYLLABLES = ['가','나','다','라','마','바','사','아','자','차','카','타','파','하'];
const toKorSyllable = (num: number) => KOR_SYLLABLES[num - 1] || `(${num})`;
const toLatin = (num: number) => String.fromCharCode(96 + num) || `(${num})`;

function getCleanData(data: ReportItem[]) {
  const result: ReportItem[] = [];
  for (let i = 0; i < data.length; i++) {
    const item = data[i];
    if (item.level === 0) {
      let hasChildren = false;
      for (let j = i + 1; j < data.length; j++) {
        if (data[j].level === 0) break;
        if (data[j].text.trim() !== "" || data[j].image || (data[j].tableData && data[j].tableData!.length > 0)) {
          hasChildren = true;
          break;
        }
      }
      if (hasChildren) {
        result.push(item);
      }
    } else {
      if (item.text.trim() !== "" || item.image || (item.tableData && item.tableData.length > 0)) {
        result.push(item);
      }
    }
  }
  return result;
}

function buildTree(flatData: ReportItem[]) {
  const tree: any[] = [];
  const path: any[] = [];

  flatData.forEach((item) => {
    const node = { id: item.id, text: item.text, level: item.level, children: [] };
    if (item.level === 1) {
      tree.push(node);
      path[1] = node;
    } else {
      let parentLevel = item.level - 1;
      while (parentLevel > 0 && !path[parentLevel]) {
        parentLevel--;
      }
      if (parentLevel > 0) {
        path[parentLevel].children.push(node);
      } else {
        tree.push(node);
      }
      path[item.level] = node;
    }
  });
  return tree;
}

export default function App() {

  const [isLocalMode, setIsLocalMode] = useState(() => localStorage.getItem('IS_LOCAL_MODE') === 'true');
  const [parish, setParish] = useState("천원특별");
  const [church, setChurch] = useState(PARISH_CHURCH_MAP["천원특별"][0]);
  
  const [reportData, setReportData] = useState<ReportItem[]>(DEFAULT_REPORT);
  const [lastSaved, setLastSaved] = useState<string | null>(null);
  const [tableSelection, setTableSelection] = useState<{ id: number, start: { r: number, c: number }, end: { r: number, c: number }, isDragging: boolean } | null>(null);
  const [tableContextMenu, setTableContextMenu] = useState<{ x: number, y: number, id: number, r: number, c: number } | null>(null);
  
  const [nextId, setNextId] = useState(5);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [showAiModal, setShowAiModal] = useState(false);
  const [showGuideModal, setShowGuideModal] = useState(false);
  const [copied, setCopied] = useState(false);
  const [jsonFormat, setJsonFormat] = useState<'flat' | 'tree'>('flat');

  const [isSaving, setIsSaving] = useState(false);
  const [driveSaveResult, setDriveSaveResult] = useState<'ok' | 'skipped' | 'error' | 'saving' | null>(null);
  const [driveSavedAt, setDriveSavedAt] = useState<string | null>(null);
  const [isCheckingAI, setIsCheckingAI] = useState(false);
  const [aiCorrections, setAiCorrections] = useState<any[] | null>(null);

  const [cropModalOpen, setCropModalOpen] = useState(false);
  const [cropImageSrc, setCropImageSrc] = useState<string>('');
  const [cropItemId, setCropItemId] = useState<number | null>(null);
  const [crop, setCrop] = useState<Crop>();
  const [completedCrop, setCompletedCrop] = useState<PixelCrop>();
  const imgRef = useRef<HTMLImageElement>(null);
  const isLoadingDataRef = useRef(false);

  const [status, setStatus] = useState<'draft' | 'submitted'>('draft');
  const [parishStats, setParishStats] = useState<Record<string, 'empty' | 'draft' | 'submitted'>>({});

  const [activeTab, setActiveTab] = useState<'report' | 'association' | 'notice_write' | 'notice'>('report');
  const [noticeTitle, setNoticeTitle] = useState('');
  const [noticePdfUrl, setNoticePdfUrl] = useState<string | null>(null);
  const [mobileView, setMobileView] = useState<'editor' | 'preview'>('editor');

  const handleNoticeWriteTab = () => {
    if (activeTab === 'notice_write') return;
    const pwd = prompt('공지사항 작성 비밀번호를 입력하세요:');
    if (pwd === 'skmt0909!') {
      setActiveTab('notice_write');
      setReportData([]);
      setNoticeTitle('');
      setNoticePdfUrl(null);
    } else {
      if (pwd !== null) alert('비밀번호가 일치하지 않습니다.');
    }
  };

  const handleAssociationTab = () => {
    if (activeTab === 'association') return;
    const pwd = prompt('협회 보고용 비밀번호를 입력하세요:');
    if (pwd === '20252027') {
      setActiveTab('association');
      setParish('협회');
      setChurch(PARISH_CHURCH_MAP['협회'][0]);
    } else {
      if (pwd !== null) alert('비밀번호가 일치하지 않습니다.');
    }
  };
  const [notices, setNotices] = useState<any[]>([]);
  const [isAdmin, setIsAdmin] = useState(false);
  const [isDownloadUnlocked, setIsDownloadUnlocked] = useState(() => sessionStorage.getItem('download_unlocked') === '1');
  const [activeNotice, setActiveNotice] = useState<any | null>(null);
  const [isUploadingNotice, setIsUploadingNotice] = useState(false);

  // 관리자 통합 취합 콘솔용 상태
  const [isAdminCheckingAI, setIsAdminCheckingAI] = useState(false);
  const [adminAiCorrections, setAdminAiCorrections] = useState<any[] | null>(null);
  const [adminReportStatusMap, setAdminReportStatusMap] = useState<Record<string, 'empty' | 'draft' | 'submitted'>>({});
  const [adminActiveParish, setAdminActiveParish] = useState<string>('전체');
  const [adminCompilationProgress, setAdminCompilationProgress] = useState<string>('');
  const [adminSelectedCorrections, setAdminSelectedCorrections] = useState<Record<string, boolean>>({});

  // Google Drive 연동 상태
  const [driveStatus, setDriveStatus] = useState<{
    configured: boolean;
    authenticated: boolean;
    hasRefreshToken: boolean;
    folderId: string | null;
    appUrl?: string;
    callbackUrl?: string;
    authUrl?: string;
  } | null>(null);
  const [driveStatusLoading, setDriveStatusLoading] = useState(false);
  const [showAiStudioGuide, setShowAiStudioGuide] = useState(false);
  const [copiedField, setCopiedField] = useState<string | null>(null);

  const copyToClipboard = (text: string, field: string) => {
    navigator.clipboard.writeText(text).then(() => {
      setCopiedField(field);
      setTimeout(() => setCopiedField(null), 2000);
    });
  };

  const checkDriveStatus = async () => {
    setDriveStatusLoading(true);
    try {
      const isLocal = localStorage.getItem('IS_LOCAL_MODE') === 'true';
      const url = isLocal
        ? `${getLocalServerUrl()}/api/google-auth/status`
        : '/api/google-auth/status';
      const res = await fetch(url);
      if (res.ok) setDriveStatus(await res.json());
    } catch {
      setDriveStatus(null);
    } finally {
      setDriveStatusLoading(false);
    }
  };

  const loadNotices = async () => {
    try {
      const parsed = await fetchDbData('SYSTEM_NOTICES');
      if (parsed && parsed.data) {
        setNotices(parsed.data);
      }
    } catch (e) { console.error(e); }
  };

  useEffect(() => {
    if (activeTab === 'notice') loadNotices();
  }, [activeTab]);

  const handleAdminLogin = () => {
    const pwd = prompt('관리자 비밀번호를 입력하세요:');
    if (pwd === 'skmt0909!') {
      setIsAdmin(true);
      alert('관리자로 로그인되었습니다.');
    } else {
      if (pwd !== null) alert('비밀번호가 일치하지 않습니다.');
    }
  };

  const handleNoticeUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (file.type !== 'application/pdf') {
      alert('PDF 파일만 업로드 가능합니다.');
      return;
    }
    const title = prompt('공지사항 제목을 입력하세요:');
    if (!title) return;

    setIsUploadingNotice(true);
    try {
      let pdfUrl = '';
      if (isLocalMode) {
        const filename = `notice_${Date.now()}.pdf`;
        const serverUrl = getLocalServerUrl();
        const res = await fetch(`${serverUrl}/api/upload-image`, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/pdf',
            'X-Filename': filename
          },
          body: file
        });
        if (res.ok) {
          const resData = await res.json();
          pdfUrl = resData.url;
        } else {
          throw new Error('Local PDF upload failed');
        }
      } else {
        const fileExt = file.name.split('.').pop();
        const fileName = `notice_${Date.now()}.${fileExt}`;
        const { error: uploadError } = await supabase.storage.from('images').upload(fileName, file);
        if (uploadError) throw uploadError;

        const { data: publicUrlData } = supabase.storage.from('images').getPublicUrl(fileName);
        pdfUrl = publicUrlData.publicUrl;
      }

      const newNotice = { id: Date.now().toString(), title, pdfUrl, created_at: new Date().toISOString() };
      const newNotices = [newNotice, ...notices];

      await saveDbData('SYSTEM_NOTICES', { id: 'SYSTEM_NOTICES', data: newNotices, updated_at: new Date().toISOString() });
      setNotices(newNotices);
      alert('공지사항이 등록되었습니다.');
    } catch (err) {
      console.error(err);
      alert('업로드 중 오류가 발생했습니다.');
    } finally {
      setIsUploadingNotice(false);
    }
  };

  const handleNoticePdfUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (file.type !== 'application/pdf') {
      alert('PDF 파일만 업로드 가능합니다.');
      return;
    }

    setIsUploadingNotice(true);
    try {
      let pdfUrl = '';
      if (isLocalMode) {
        const filename = `notice_pdf_${Date.now()}.pdf`;
        const serverUrl = getLocalServerUrl();
        const res = await fetch(`${serverUrl}/api/upload-image`, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/pdf',
            'X-Filename': filename
          },
          body: file
        });
        if (res.ok) {
          const resData = await res.json();
          pdfUrl = resData.url;
        } else {
          throw new Error('Local PDF upload failed');
        }
      } else {
        const fileExt = file.name.split('.').pop();
        const fileName = `notice_pdf_${Date.now()}.${fileExt}`;
        const { error: uploadError } = await supabase.storage.from('images').upload(fileName, file);
        if (uploadError) throw uploadError;

        const { data: publicUrlData } = supabase.storage.from('images').getPublicUrl(fileName);
        pdfUrl = publicUrlData.publicUrl;
      }
      setNoticePdfUrl(pdfUrl);
      alert('PDF가 첨부되었습니다.');
    } catch (err) {
      console.error(err);
      alert('PDF 업로드 중 오류가 발생했습니다.');
    } finally {
      setIsUploadingNotice(false);
    }
  };

  const handlePublishNotice = async () => {
    if (!noticeTitle.trim()) { alert('공지사항 제목을 입력해주세요.'); return; }
    if (reportData.length === 0 && !noticePdfUrl) { alert('내용을 작성하거나 PDF를 첨부해주세요.'); return; }
    
    setIsUploadingNotice(true);
    try {
      const newNotice = { 
        id: Date.now().toString(), 
        title: noticeTitle, 
        pdfUrl: noticePdfUrl,
        data: reportData.length > 0 ? reportData : null, 
        created_at: new Date().toISOString() 
      };
      
      const parsed = await fetchDbData('SYSTEM_NOTICES');
      const existingNotices = (parsed && parsed.data) ? parsed.data : [];
      
      const newNotices = [newNotice, ...existingNotices];
      await saveDbData('SYSTEM_NOTICES', { id: 'SYSTEM_NOTICES', data: newNotices, updated_at: new Date().toISOString() });
      
      setNotices(newNotices);
      alert('공지사항이 성공적으로 등록되었습니다.');
      setNoticeTitle('');
      setReportData([]);
      setNoticePdfUrl(null);
      setActiveTab('notice');
    } catch (e) {
      console.error(e);
      alert('등록 중 오류가 발생했습니다.');
    } finally {
      setIsUploadingNotice(false);
    }
  };

  const deleteNotice = async (id: string) => {
    if (!window.confirm('이 공지사항을 삭제하시겠습니까?')) return;
    try {
      const newNotices = notices.filter(n => n.id !== id);
      await saveDbData('SYSTEM_NOTICES', { id: 'SYSTEM_NOTICES', data: newNotices, updated_at: new Date().toISOString() });
      setNotices(newNotices);
      if (activeNotice?.id === id) setActiveNotice(null);
    } catch (e) {
      console.error(e);
      alert('삭제 중 오류가 발생했습니다.');
    }
  };

  // --- 관리자 콘솔 전용 핵심 로직 ---
  const getReportDataFor = async (p: string, c: string) => {
    const key = `report_${p}_${c}`;

    // 1. 세션 캐시 (Drive 호출 스킵)
    const session = sessionStorage.getItem(key);
    if (session) { try { return JSON.parse(session); } catch(e){} }

    // 2. localStorage
    const local = localStorage.getItem(key);
    if (local) {
      try {
        const parsed = JSON.parse(local);
        sessionStorage.setItem(key, local);
        return parsed;
      } catch(e){}
    }
    if (isLocalMode) {
      try {
        const serverUrl = getLocalServerUrl();
        const res = await fetch(`${serverUrl}/api/load-data/${key}`);
        if (res.ok) return await res.json();
      } catch(e){}
    }
    if (supabase) {
      try {
        const { data, error } = await supabase.storage.from('images').download(`db_reports/${p}_${c}.json`);
        if (!error && data) {
          const text = await data.text();
          const parsed = JSON.parse(text);
          localStorage.setItem(key, JSON.stringify(parsed));
          sessionStorage.setItem(key, JSON.stringify(parsed));
          return parsed;
        }
      } catch(e){}
    }
    // 3. Drive 폴백
    try {
      const res = await fetch(`/api/load-report?parish=${encodeURIComponent(p)}&church=${encodeURIComponent(c)}`);
      if (res.ok) {
        const { found, payload } = await res.json();
        if (found && payload) {
          localStorage.setItem(key, JSON.stringify(payload));
          sessionStorage.setItem(key, JSON.stringify(payload));
          return payload;
        }
      }
    } catch(e){}
    return null;
  };

  const loadAllReportsStatus = async () => {
    const stats: Record<string, 'empty' | 'draft' | 'submitted'> = {};
    for (const p of Object.keys(PARISH_CHURCH_MAP)) {
      for (const c of PARISH_CHURCH_MAP[p]) {
        const key = `${p}_${c}`;
        
        // 1. 현재 로드된 교회 체크
        if (p === parish && c === church) {
          if (status === 'submitted') {
            stats[key] = 'submitted';
          } else {
            stats[key] = getCleanData(reportData).length > 0 ? 'draft' : 'empty';
          }
          continue;
        }

        // 2. localStorage 체크
        const saved = localStorage.getItem(`report_${p}_${c}`);
        if (saved) {
          try {
            const parsed = JSON.parse(saved);
            if (parsed.status === 'submitted') {
              stats[key] = 'submitted';
            } else {
              const d = parsed.data || [];
              stats[key] = getCleanData(d).length > 0 ? 'draft' : 'empty';
            }
          } catch(e) {
            stats[key] = 'empty';
          }
        } else {
          stats[key] = 'empty';
        }
      }
    }
    setAdminReportStatusMap(stats);
  };

  useEffect(() => {
    if (activeTab === 'admin_console') {
      loadAllReportsStatus();
      checkDriveStatus();
      // 관리자 콘솔 진입 시 전 교구·전 교회 데이터를 백그라운드로 미리 캐시
      Object.entries(PARISH_CHURCH_MAP).forEach(([p, churches]) => {
        (churches as string[]).forEach((c: string) => {
          const key = `report_${p}_${c}`;
          if (sessionStorage.getItem(key) || localStorage.getItem(key)) return;
          fetch(`/api/load-report?parish=${encodeURIComponent(p)}&church=${encodeURIComponent(c)}`)
            .then(r => r.ok ? r.json() : null)
            .then(json => {
              if (json?.found && json?.payload) {
                const str = JSON.stringify(json.payload);
                localStorage.setItem(key, str);
                sessionStorage.setItem(key, str);
              }
            })
            .catch(() => {});
        });
      });
    }
  }, [activeTab, parish, church, status, reportData]);

  const startAdminAiReview = async () => {
    setIsAdminCheckingAI(true);
    setAdminCompilationProgress("전국 교구 및 협회 보고서 실시간 취합 중...");
    setAdminAiCorrections(null);

    try {
      const parishes = Object.keys(PARISH_CHURCH_MAP);
      setAdminCompilationProgress("전국 교구 및 협회 보고서 Drive에서 병렬 취합 중...");

      // 교구별 병렬, 교구 내 교회도 병렬로 Drive 취합
      const parishResults = await Promise.all(
        parishes.map(async (p) => {
          const churchResults = await Promise.all(
            PARISH_CHURCH_MAP[p].map(async (c) => {
              if (p === parish && c === church) return { p, c, data: getCleanData(reportData) };
              const report = await getReportDataFor(p, c);
              return { p, c, data: report?.data ? getCleanData(report.data) : [] };
            })
          );
          return churchResults;
        })
      );

      const allAdminPayload: any[] = [];
      for (const churchResults of parishResults) {
        for (const { p, c, data } of churchResults) {
          data.filter((item: ReportItem) => item.text.trim() !== "").forEach((item: ReportItem) => {
            allAdminPayload.push({ parish: p, church: c, id: item.id, text: item.text });
          });
        }
      }

      if (allAdminPayload.length === 0) {
        alert("취합된 주간보고 내용이 없습니다. 각 교구 및 부서의 보고서 작성 현황을 확인해 주세요.");
        setIsAdminCheckingAI(false);
        return;
      }

      setAdminCompilationProgress(`총 ${allAdminPayload.length}개 업무보고 항목 취합 성공. AI 종합 편집 및 문맥 검토 분석 중...`);

      const adminAiPrompt = `당신은 전체 교구 및 협회 주간업무보고를 총괄 검토하는 전문 수석 편집자입니다.
아래 제공된 데이터의 텍스트(text)를 검토하세요. 각 항목은 교구(parish), 교회(church), 항목 ID(id)를 가지고 있습니다.
1. 오타가 있거나 2. 문맥상 어색하거나 3. 주간보고 개조식 형식(~함, ~예정 등)에 맞지 않는 항목들을 찾아 완벽하게 교정해 주세요.
반드시 아래 JSON 배열 형태로만 정확히 응답하세요. (백틱이나 markdown 코드 블록 없이 순수 JSON만 반환해야 합니다.)
[{ "parish": "교구이름", "church": "교회이름", "id": 1, "original": "원래 텍스트", "corrected": "완벽히 교정한 텍스트", "reason": "교정 이유" }]`;

      let text = "";

      if (isLocalMode) {
        const serverUrl = getLocalServerUrl();
        const res = await fetch(`${serverUrl}/api/ollama-chat`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            prompt: adminAiPrompt + `\n\n데이터:\n${JSON.stringify(allAdminPayload, null, 2)}`
          })
        });
        if (res.ok) {
          const data = await res.json();
          text = data.text;
        } else {
          const errData = await res.json().catch(() => ({}));
          throw new Error(errData.error || "Ollama API failed");
        }
      } else {
        const openRouterKey = localStorage.getItem('OPENROUTER_API_KEY');
        if (openRouterKey) {
          const res = await fetch("https://openrouter.ai/api/v1/chat/completions", {
            method: "POST",
            headers: {
              "Authorization": `Bearer ${openRouterKey}`,
              "Content-Type": "application/json"
            },
            body: JSON.stringify({
              model: "google/gemini-2.0-flash-lite-preview-02-05:free",
              response_format: { type: "json_object" },
              messages: [{ role: "user", content: adminAiPrompt + `\n\n데이터:\n${JSON.stringify(allAdminPayload, null, 2)}` }]
            })
          });
          if (res.ok) {
            const json = await res.json();
            text = json.choices[0].message.content;
          }
        }
      }

      if (!text) throw new Error("AI 검토 응답을 수신하지 못했습니다.");

      let cleanText = text.replace(/```json/gi, '').replace(/```/gi, '').trim();
      const match = cleanText.match(/\[\s*\{[\s\S]*\}\s*\]/);
      if (match) cleanText = match[0];
      
      const corrections = JSON.parse(cleanText);
      setAdminAiCorrections(corrections);
      
      // 기본적으로 모든 교정 사항을 체크(적용) 상태로 활성화
      const initialSelected: Record<string, boolean> = {};
      corrections.forEach((c: any) => {
        const key = `${c.parish}_${c.church}_${c.id}`;
        initialSelected[key] = true;
      });
      setAdminSelectedCorrections(initialSelected);

      setAdminCompilationProgress("");
    } catch (err: any) {
      console.error(err);
      alert(`AI 통합 검토 중 오류가 발생했습니다: ${err.message}`);
    } finally {
      setIsAdminCheckingAI(false);
    }
  };

  const toggleAdminCorrectionSelected = (parishName: string, churchName: string, id: number) => {
    const key = `${parishName}_${churchName}_${id}`;
    setAdminSelectedCorrections(prev => ({
      ...prev,
      [key]: !prev[key]
    }));
  };

  const applySelectedAdminCorrections = async () => {
    if (!adminAiCorrections) return;

    const selected = adminAiCorrections.filter(c => {
      const key = `${c.parish}_${c.church}_${c.id}`;
      return !!adminSelectedCorrections[key];
    });

    if (selected.length === 0) {
      alert("적용할 교정 사항이 선택되지 않았습니다.");
      return;
    }

    setAdminCompilationProgress("선택된 교정 사항을 데이터베이스에 실시간 반영 중...");

    const grouped: Record<string, Record<string, {id: number, text: string}[]>> = {};
    selected.forEach(c => {
      if (!grouped[c.parish]) grouped[c.parish] = {};
      if (!grouped[c.parish][c.church]) grouped[c.parish][c.church] = [];
      grouped[c.parish][c.church].push({ id: c.id, text: c.corrected });
    });

    for (const [pName, churchesMap] of Object.entries(grouped)) {
      for (const [cName, updates] of Object.entries(churchesMap)) {
        const updateMap = new Map(updates.map(u => [u.id, u.text]));
        
        if (pName === parish && cName === church) {
          setReportData(data => data.map(item => updateMap.has(item.id) ? { ...item, text: updateMap.get(item.id)! } : item));
        } else {
          const key = `report_${pName}_${cName}`;
          const saved = localStorage.getItem(key);
          if (saved) {
            try {
              const parsed = JSON.parse(saved);
              if (parsed.data) {
                parsed.data = parsed.data.map((item: any) => updateMap.has(item.id) ? { ...item, text: updateMap.get(item.id)! } : item);
                localStorage.setItem(key, JSON.stringify(parsed));
                await saveDbData(`${pName}_${cName}`, { id: `${pName}_${cName}`, ...parsed, updated_at: new Date().toISOString() });
              }
            } catch(e){}
          }
        }
      }
    }

    const appliedKeys = new Set(selected.map(c => `${c.parish}_${c.church}_${c.id}`));
    setAdminAiCorrections(prev => prev ? prev.filter(c => !appliedKeys.has(`${c.parish}_${c.church}_${c.id}`)) : null);
    
    setAdminCompilationProgress("");
    alert(`${selected.length}개의 AI 교정 제안이 성공적으로 반영되었습니다!`);
    loadAllReportsStatus();
  };

  const exportMasterToWord = async () => {
    setAdminCompilationProgress("전체 교구 및 협회 데이터를 최종 취합하여 워드(Word) 문서를 생성하는 중...");
    
    try {
      let allChildren: (Paragraph | Table)[] = [];
      
      allChildren.push(
        new Paragraph({
          children: [
            new TextRun({
              text: "전국 교구 및 협회 통합 주간업무보고서",
              size: 40,
              bold: true,
              color: "1D4ED8",
              font: "맑은 고딕"
            })
          ],
          alignment: "center" as const,
          spacing: { before: 400, after: 400 },
          border: {
            bottom: { color: "1D4ED8", space: 15, size: 24, style: BorderStyle.SINGLE },
          }
        })
      );

      const parishes = Object.keys(PARISH_CHURCH_MAP);

      for (const p of parishes) {
        let parishHasData = false;
        const tempChildren: (Paragraph | Table)[] = [];

        tempChildren.push(
          new Paragraph({
            children: [
              new TextRun({
                text: `■ ${getDisplayParish(p)} 업무보고`,
                size: 28,
                bold: true,
                color: "1D4ED8",
                font: "맑은 고딕"
              })
            ],
            spacing: { before: 400, after: 200 }
          })
        );

        const churches = PARISH_CHURCH_MAP[p];
        let chIndex = 1;

        for (const c of churches) {
          let dataToUse: ReportItem[] = [];
          if (p === parish && c === church) {
            dataToUse = getCleanData(reportData);
          } else {
            const report = await getReportDataFor(p, c);
            if (report && report.data) {
              dataToUse = getCleanData(report.data);
            }
          }

          if (dataToUse.length === 0) continue;
          parishHasData = true;

          tempChildren.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: `${chIndex}. ${getDisplayChurch(c)}`,
                  size: 24,
                  bold: true,
                  color: "1E3A8A",
                  font: "맑은 고딕"
                })
              ],
              spacing: { before: 240, after: 120 }
            })
          );
          chIndex++;

          const counters = [0, 0, 0, 0, 0, 0];
          for (const item of dataToUse) {
            let prefix = "";
            if (item.level === 0) {
              counters[0]++; counters[1] = 0; counters[2] = 0; counters[3] = 0; counters[4] = 0; counters[5] = 0;
              prefix = toRoman(counters[0]) + ". ";
            } else if (item.level === 1) {
              counters[1]++; counters[2] = 0; counters[3] = 0; counters[4] = 0; counters[5] = 0;
              prefix = counters[1] + ". ";
            } else if (item.level === 2) {
              counters[2]++; counters[3] = 0; counters[4] = 0; counters[5] = 0;
              prefix = counters[2] + ") ";
            } else if (item.level === 3) {
              counters[3]++; counters[4] = 0; counters[5] = 0;
              prefix = toCircled(counters[3]) + " ";
            } else if (item.level === 4) {
              counters[4]++; counters[5] = 0;
              prefix = toKorSyllable(counters[4]) + ". ";
            } else if (item.level === 5) {
              counters[5]++;
              prefix = toLatin(counters[5]) + ". ";
            }

            let color = "000000";
            if (item.level <= 1) color = "1D4ED8";
            
            const lines = `${prefix}${item.text || ''}`.split('\n');

            tempChildren.push(new Paragraph({
              children: lines.map((line, idx) => 
                new TextRun({
                  text: line,
                  break: idx > 0 ? 1 : 0,
                  bold: item.level <= 1,
                  color: color,
                  font: "맑은 고딕",
                  size: 22,
                })
              ),
              indent: {
                left: Math.max(0, (item.level === 0 ? 0 : (item.level - 1) * 360))
              },
              spacing: {
                before: item.level === 0 ? 360 : 120, 
              }
            }));

            if (item.tableData && item.tableData.length > 0) {
              const skippedCells = new Set<string>();
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
              tempChildren.push(docTable);
            }

            if (item.image && item.imageWidth && item.imageHeight) {
              try {
                let imageBuffer: ArrayBuffer;
                if (item.image.startsWith("http")) {
                  const res = await fetch(item.image);
                  imageBuffer = await res.arrayBuffer();
                } else {
                  const base64Data = item.image.split(",")[1];
                  imageBuffer = Uint8Array.from(atob(base64Data), char => char.charCodeAt(0));
                }
                
                const targetWidth = 500;
                const ratio = Math.min(1, targetWidth / item.imageWidth);
                const finalWidth = item.imageWidth * ratio;
                const finalHeight = item.imageHeight * ratio;

                tempChildren.push(new Paragraph({
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
              } catch (e) {
                console.error("[Word] Failed to embed image:", e);
              }
            }
          }
        }

        if (parishHasData) {
          allChildren = [...allChildren, ...tempChildren];
        }
      }

      if (allChildren.length <= 1) {
        alert("취합할 보고서 데이터가 존재하지 않습니다.");
        setAdminCompilationProgress("");
        return;
      }

      const doc = new Document({
        sections: [{
          properties: {
            page: {
              size: {
                width: 11906,
                height: 16838,
              }
            }
          },
          children: allChildren,
        }],
      });

      const blob = await Packer.toBlob(doc);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `전국교구_및_협회_종합_업무보고_${new Date().toISOString().split('T')[0]}.docx`;
      a.click();
      URL.revokeObjectURL(url);

      setAdminCompilationProgress("");
      alert("종합 주간업무보고 워드 문서 다운로드가 완료되었습니다!");
    } catch(err: any) {
      console.error(err);
      alert(`문서 생성 중 오류 발생: ${err.message}`);
      setAdminCompilationProgress("");
    }
  };

  const updateParishStats = () => {
    const stats: Record<string, 'empty' | 'draft' | 'submitted'> = {};
    const churches = PARISH_CHURCH_MAP[parish];
    churches.forEach(c => {
      if (c === church) {
        // Evaluate current
        if (status === 'submitted') {
          stats[c] = 'submitted';
        } else {
          stats[c] = getCleanData(reportData).length > 0 ? 'draft' : 'empty';
        }
      } else {
        const key = `report_${parish}_${c}`;
        const saved = localStorage.getItem(key);
        if (saved) {
          try {
            const parsed = JSON.parse(saved);
            if (parsed.status === 'submitted') {
              stats[c] = 'submitted';
            } else {
              const d = parsed.data || DEFAULT_REPORT;
              stats[c] = getCleanData(d).length > 0 ? 'draft' : 'empty';
            }
          } catch (e) {
            stats[c] = 'empty';
          }
        } else {
          stats[c] = 'empty';
        }
      }
    });
    setParishStats(stats);
  };

  useEffect(() => {
    updateParishStats();
  }, [parish, church, status, reportData]);

  useEffect(() => {
    const handleMouseUp = () => {
      setTableSelection(prev => prev ? { ...prev, isDragging: false } : null);
    };
    const handleClick = () => setTableContextMenu(null);
    window.addEventListener('mouseup', handleMouseUp);
    window.addEventListener('click', handleClick);
    return () => {
      window.removeEventListener('mouseup', handleMouseUp);
      window.removeEventListener('click', handleClick);
    };
  }, []);

  // Load data when parish or church changes
  useEffect(() => {
    if (activeTab === 'notice_write' || activeTab === 'notice') return;

    const loadData = async () => {
      isLoadingDataRef.current = true;
      const key = `report_${parish}_${church}`;
      const savedLocal = localStorage.getItem(key);
      let localParsed: any = null;
      if (savedLocal) {
        try { localParsed = JSON.parse(savedLocal); } catch(e){}
      }

      if (localParsed) {
         setReportData(localParsed.data && localParsed.data.length > 0 ? localParsed.data : DEFAULT_REPORT);
         setLastSaved(localParsed.lastSaved || null);
         setStatus(localParsed.status || 'draft');
         const maxId = Math.max(4, ...(localParsed.data || DEFAULT_REPORT).map((d: any) => d.id));
         setNextId(maxId + 1);
      } else {
         setReportData(DEFAULT_REPORT);
         setLastSaved(null);
         setStatus('draft');
      }

      // Fetch from Supabase (Storage DB) to get latest if exists
      if (supabase) {
         try {
            const supaData = await fetchDbData(`${parish}_${church}`);

            if (supaData) {
               setReportData(supaData.data && supaData.data.length > 0 ? supaData.data : DEFAULT_REPORT);
               setLastSaved(supaData.lastSaved || null);
               setStatus(supaData.status || 'draft');
               const maxId = Math.max(4, ...(supaData.data || DEFAULT_REPORT).map((d: any) => d.id));
               setNextId(maxId + 1);
               localStorage.setItem(key, JSON.stringify(supaData));
               setAiCorrections(null);
               return;
            }
         } catch(e) {
            console.error("Supabase load failed", e);
         }
      }

      // Drive에서 불러오기 (localStorage/Supabase에 없을 때)
      try {
        const driveRes = await fetch(`/api/load-report?parish=${encodeURIComponent(parish)}&church=${encodeURIComponent(church)}`);
        if (driveRes.ok) {
          const { found, payload } = await driveRes.json();
          if (found && payload) {
            setReportData(payload.data && payload.data.length > 0 ? payload.data : DEFAULT_REPORT);
            setLastSaved(payload.lastSaved || null);
            setStatus(payload.status || 'draft');
            const maxId = Math.max(4, ...(payload.data || DEFAULT_REPORT).map((d: any) => d.id));
            setNextId(maxId + 1);
            localStorage.setItem(key, JSON.stringify(payload));
            setDriveSaveResult('ok');
            setDriveSavedAt('Drive에서 불러옴');
          }
        }
      } catch(e) {
        console.error("Drive load failed", e);
      }

      setAiCorrections(null);

      isLoadingDataRef.current = false;
      // 백그라운드 프리로드: 같은 교구 내 다른 교회를 미리 Drive에서 캐시
      const churches = PARISH_CHURCH_MAP[parish] || [];
      churches.forEach((c: string) => {
        if (c === church) return;
        const sessionKey = `report_${parish}_${c}`;
        if (sessionStorage.getItem(sessionKey) || localStorage.getItem(sessionKey)) return;
        // 이미 캐시된 교회는 스킵, 없는 교회만 백그라운드 로드
        fetch(`/api/load-report?parish=${encodeURIComponent(parish)}&church=${encodeURIComponent(c)}`)
          .then(r => r.ok ? r.json() : null)
          .then(json => {
            if (json?.found && json?.payload) {
              const str = JSON.stringify(json.payload);
              localStorage.setItem(sessionKey, str);
              sessionStorage.setItem(sessionKey, str);
            }
          })
          .catch(() => {});
      });
    };
    loadData();
  }, [parish, church, activeTab, isLocalMode]);

  // Silent auto-save on data change
  useEffect(() => {
    if (activeTab === 'notice_write' || activeTab === 'notice') return;
    if (reportData === DEFAULT_REPORT && !lastSaved) return;
    if (isLoadingDataRef.current) return;
    const key = `report_${parish}_${church}`;
    const timestamp = lastSaved || new Date().toLocaleString('ko-KR');
    
    // 자동저장은 로컬만 (클라우드는 제출 확정 시에만)
    const saveData = { data: reportData, lastSaved: timestamp, status };
    localStorage.setItem(key, JSON.stringify(saveData));
  }, [reportData, parish, church, status, lastSaved, activeTab]);

  const handleImageUpload = (e: React.ChangeEvent<HTMLInputElement>, id: number) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      setCropImageSrc(event.target?.result as string);
      setCropItemId(id);
      setCrop(undefined);
      setCompletedCrop(undefined);
      setCropModalOpen(true);
    };
    reader.readAsDataURL(file);
    e.target.value = '';
  };

  const applyCrop = () => {
    if (!imgRef.current || !cropItemId) {
      setCropModalOpen(false);
      return;
    }

    const image = imgRef.current;
    
    let cropWidth = image.naturalWidth;
    let cropHeight = image.naturalHeight;
    let cropX = 0;
    let cropY = 0;

    if (completedCrop && completedCrop.width > 0 && completedCrop.height > 0) {
      const scaleX = image.naturalWidth / image.width;
      const scaleY = image.naturalHeight / image.height;
      cropWidth = completedCrop.width * scaleX;
      cropHeight = completedCrop.height * scaleY;
      cropX = completedCrop.x * scaleX;
      cropY = completedCrop.y * scaleY;
    }

    const tempCanvas = document.createElement('canvas');
    tempCanvas.width = cropWidth;
    tempCanvas.height = cropHeight;
    const tempCtx = tempCanvas.getContext('2d');
    
    if (tempCtx) {
      tempCtx.drawImage(
        image,
        cropX,
        cropY,
        cropWidth,
        cropHeight,
        0,
        0,
        cropWidth,
        cropHeight
      );
    }

    const MAX_WIDTH = 800;
    const MAX_HEIGHT = 800;
    let finalWidth = cropWidth;
    let finalHeight = cropHeight;

    if (finalWidth > finalHeight) {
      if (finalWidth > MAX_WIDTH) {
        finalHeight *= MAX_WIDTH / finalWidth;
        finalWidth = MAX_WIDTH;
      }
    } else {
      if (finalHeight > MAX_HEIGHT) {
        finalWidth *= MAX_HEIGHT / finalHeight;
        finalHeight = MAX_HEIGHT;
      }
    }

    const canvas = document.createElement('canvas');
    canvas.width = finalWidth;
    canvas.height = finalHeight;
    const ctx = canvas.getContext('2d');
    if (ctx) {
      ctx.fillStyle = 'white';
      ctx.fillRect(0, 0, finalWidth, finalHeight);
      ctx.drawImage(tempCanvas, 0, 0, finalWidth, finalHeight);
    }

    const dataUrl = canvas.toDataURL('image/jpeg', 0.8);
    setCropModalOpen(false);

    // Optimistic update
    setReportData(data => data.map(item => item.id === cropItemId ? { 
      ...item, 
      image: dataUrl,
      imageWidth: finalWidth,
      imageHeight: finalHeight
    } : item));

    // Upload to local storage or Supabase
    canvas.toBlob(async (blob) => {
      if (!blob) return;

      if (isLocalMode) {
        try {
          const filename = `${parish}_${church}_${Date.now()}.jpg`.replace(/\s+/g, '_');
          const serverUrl = getLocalServerUrl();
          
          const res = await fetch(`${serverUrl}/api/upload-image`, {
            method: 'POST',
            headers: {
              'Content-Type': 'image/jpeg',
              'X-Filename': filename
            },
            body: blob
          });
          
          if (res.ok) {
            const uploadData = await res.json();
            setReportData(data => data.map(item => item.id === cropItemId ? { 
              ...item, 
              image: uploadData.url
            } : item));
          } else {
            console.error("Local image upload API failed");
          }
        } catch (e) {
          console.error("Local image upload failed:", e);
        }
        return;
      }

      if (supabase) {
        const timestamp = new Date().getTime();
        const filePath = `${parish}/${church}/${timestamp}.jpg`;
        const { data, error } = await supabase.storage
          .from('images')
          .upload(filePath, blob, { contentType: 'image/jpeg' });
          
        if (!error) {
          const { data: { publicUrl } } = supabase.storage
            .from('images')
            .getPublicUrl(filePath);
            
          setReportData(data => data.map(item => item.id === cropItemId ? { 
            ...item, 
            image: publicUrl
          } : item));
        } else {
          console.error("Image upload failed", error);
        }
      }
    }, 'image/jpeg', 0.8);
  };

  const removeImage = (id: number) => {
    setReportData(data => data.map(item => {
      if (item.id === id) {
        const { image, imageWidth, imageHeight, ...rest } = item;
        return rest;
      }
      return item;
    }));
  };

  const handlePaste = (e: React.ClipboardEvent<HTMLInputElement | HTMLTextAreaElement>, id: number) => {
    // 이미지 붙여넣기 (클립보드에서 직접 Ctrl+V)
    const imageItem = Array.from(e.clipboardData.items).find(item => item.type.startsWith('image/'));
    if (imageItem) {
      e.preventDefault();
      const file = imageItem.getAsFile();
      if (file) {
        const reader = new FileReader();
        reader.onload = (ev) => {
          const dataUrl = ev.target?.result as string;
          const img = new Image();
          img.onload = () => {
            setReportData(data => data.map(item => item.id === id ? {
              ...item, image: dataUrl, imageWidth: img.naturalWidth, imageHeight: img.naturalHeight
            } : item));
          };
          img.src = dataUrl;
        };
        reader.readAsDataURL(file);
      }
      return;
    }

    const htmlInfo = e.clipboardData.getData('text/html');
    if (htmlInfo) {
      const parsed = parseHtmlTable(htmlInfo);
      if (parsed && parsed.tableData.length > 0) {
        e.preventDefault();
        setReportData(data => data.map(item => item.id === id ? { 
          ...item, 
          tableData: parsed.tableData,
          tableSpans: parsed.tableSpans,
          tableHighlights: parsed.tableData.map(r => r.map(() => false)),
          chartType: item.chartType || 'none' 
        } : item));
        return;
      }
    }

    const text = e.clipboardData.getData('text/plain');
    if (text.includes('\t')) {
      const rows = text.split(/\r?\n/).map(row => row.split('\t').map(cell => cell.trim())).filter(row => row.some(cell => cell !== ''));
      if (rows.length > 1 || (rows.length === 1 && rows[0].length > 1)) {
        e.preventDefault();
        setReportData(data => data.map(item => item.id === id ? { 
          ...item, 
          tableData: rows, 
          tableHighlights: rows.map(r => r.map(() => false)),
          chartType: item.chartType || 'none' 
        } : item));
        return;
      }
    }

    const lines = text.split(/\r?\n/).filter(line => line.trim() !== '');
    if (lines.length > 1) {
      e.preventDefault();
      setReportData(data => {
        const newData = [...data];
        const targetIndex = newData.findIndex(item => item.id === id);
        if (targetIndex !== -1) {
          const currentItem = newData[targetIndex];
          const isCurrentEmpty = !currentItem.text.trim();
          newData[targetIndex] = { ...currentItem, text: isCurrentEmpty ? lines[0] : currentItem.text + ' ' + lines[0] };
          
          let maxId = Math.max(0, ...data.map(d => d.id));
          const itemsToAdd = lines.slice(1).map((line, i) => ({
            id: maxId + 1 + i,
            text: line.trim(),
            level: currentItem.level
          }));
          newData.splice(targetIndex + 1, 0, ...itemsToAdd);
        }
        return newData;
      });
    }
  };

  const addEmptyTable = (id: number) => {
    setReportData(data => data.map(item => item.id === id ? { 
      ...item, 
      tableData: [["", ""], ["", ""]], 
      tableHighlights: [[false, false], [false, false]],
      chartType: 'none' 
    } : item));
  };

  const updateTableCell = (id: number, rowIndex: number, colIndex: number, value: string) => {
    setReportData(data => data.map(item => {
      if (item.id === id && item.tableData) {
        // Handle tabular data paste
        if (value.includes('\t') || value.includes('\n')) {
          const rows = value.split(/\r?\n/).filter(r => r.trim() !== '');
          if (rows.length > 0) {
            const newTable = [...item.tableData.map(r => [...r])];
            const newHighlights = [...(item.tableHighlights || newTable.map(r => r.map(() => false))).map(r => [...r])];
            
            // Expand rows if needed
            while (newTable.length < rowIndex + rows.length) {
              const cols = newTable[0]?.length || 1;
              newTable.push(Array(cols).fill(""));
              newHighlights.push(Array(cols).fill(false));
            }
            
            rows.forEach((r, rOff) => {
              const cells = r.split('\t');
              // Expand columns if needed
              while (newTable[rowIndex + rOff].length < colIndex + cells.length) {
                newTable.forEach((row, idx) => {
                  row.push("");
                  newHighlights[idx].push(false);
                });
              }
              cells.forEach((c, cOff) => {
                newTable[rowIndex + rOff][colIndex + cOff] = c;
              });
            });
            return { ...item, tableData: newTable, tableHighlights: newHighlights };
          }
        }

        const newTable = [...item.tableData];
        newTable[rowIndex] = [...newTable[rowIndex]];
        newTable[rowIndex][colIndex] = value;
        return { ...item, tableData: newTable };
      }
      return item;
    }));
  };

  const addTableRow = (id: number) => {
    setReportData(data => data.map(item => {
      if (item.id === id && item.tableData) {
        const cols = item.tableData[0]?.length || 1;
        return { 
          ...item, 
          tableData: [...item.tableData, Array(cols).fill("")],
          tableHighlights: [...(item.tableHighlights || item.tableData.map(r => r.map(() => false))), Array(cols).fill(false)]
        };
      }
      return item;
    }));
  };

  const addTableCol = (id: number) => {
    setReportData(data => data.map(item => {
      if (item.id === id && item.tableData) {
        return { 
          ...item, 
          tableData: item.tableData.map(row => [...row, ""]),
          tableHighlights: (item.tableHighlights || item.tableData.map(r => r.map(() => false))).map(row => [...row, false])
        };
      }
      return item;
    }));
  };

  const removeTableRow = (id: number, rowIndex: number) => {
    setReportData(data => data.map(item => {
      if (item.id === id && item.tableData && item.tableData.length > 1) {
        return { 
          ...item, 
          tableData: item.tableData.filter((_, idx) => idx !== rowIndex),
          tableHighlights: item.tableHighlights ? item.tableHighlights.filter((_, idx) => idx !== rowIndex) : undefined
        };
      }
      return item;
    }));
  };

  const removeTableCol = (id: number, colIndex: number) => {
    setReportData(data => data.map(item => {
      if (item.id === id && item.tableData && item.tableData[0].length > 1) {
        return { 
          ...item, 
          tableData: item.tableData.map(row => row.filter((_, idx) => idx !== colIndex)),
          tableHighlights: item.tableHighlights ? item.tableHighlights.map(row => row.filter((_, idx) => idx !== colIndex)) : undefined
        };
      }
      return item;
    }));
  };

  const toggleCellHighlight = (id: number, rIdx: number, cIdx: number) => {
    setReportData(data => data.map(item => {
      if (item.id === id && item.tableData) {
        const newHighlights = item.tableHighlights ? item.tableHighlights.map(row => [...row]) :
          item.tableData.map(row => Array(row.length).fill(false));
        
        while (newHighlights.length <= rIdx) {
          newHighlights.push(Array(item.tableData[0].length).fill(false));
        }
        while (newHighlights[rIdx].length <= cIdx) {
          newHighlights[rIdx].push(false);
        }
        
        newHighlights[rIdx][cIdx] = !newHighlights[rIdx][cIdx];
        return { ...item, tableHighlights: newHighlights };
      }
      return item;
    }));
  };

  const toggleRowHighlight = (id: number, rIdx: number) => {
    setReportData(data => data.map(item => {
      if (item.id === id && item.tableData) {
        const newHighlights = item.tableHighlights ? item.tableHighlights.map(row => [...row]) :
          item.tableData.map(row => Array(row.length).fill(false));
        
        while (newHighlights.length <= rIdx) {
          newHighlights.push(Array(item.tableData[0].length).fill(false));
        }
        
        const isCurrentlyHighlighted = newHighlights[rIdx].some(h => h);
        for (let c = 0; c < item.tableData[rIdx].length; c++) {
          newHighlights[rIdx][c] = !isCurrentlyHighlighted;
        }
        return { ...item, tableHighlights: newHighlights };
      }
      return item;
    }));
  };

  const setCellAlignment = (id: number, rIdx: number, cIdx: number, align: 'left' | 'center' | 'right' | undefined) => {
    setReportData(data => data.map(item => {
      if (item.id === id && item.tableData) {
        const newAlignments = item.tableAlignments ? item.tableAlignments.map(row => [...row]) :
          item.tableData.map(row => Array(row.length).fill(undefined));
        
        while (newAlignments.length <= rIdx) {
          newAlignments.push(Array(item.tableData[0].length).fill(undefined));
        }
        while (newAlignments[rIdx].length <= cIdx) {
          newAlignments[rIdx].push(undefined);
        }
        
        // If selection is active, apply to all selected cells
        if (tableSelection?.id === id && tableSelection.start && tableSelection.end) {
          const minR = Math.min(tableSelection.start.r, tableSelection.end.r);
          const maxR = Math.max(tableSelection.start.r, tableSelection.end.r);
          const minC = Math.min(tableSelection.start.c, tableSelection.end.c);
          const maxC = Math.max(tableSelection.start.c, tableSelection.end.c);
          
          for (let r = minR; r <= maxR; r++) {
            for (let c = minC; c <= maxC; c++) {
              newAlignments[r][c] = align;
            }
          }
        } else {
          newAlignments[rIdx][cIdx] = align;
        }

        return { ...item, tableAlignments: newAlignments };
      }
      return item;
    }));
  };

  const removeTable = (id: number) => {
    setReportData(data => data.map(item => {
      if (item.id === id) {
        const { tableData, chartType, ...rest } = item;
        return rest;
      }
      return item;
    }));
  };

  const setChartView = (id: number, type: 'none' | 'bar' | 'line' | 'pie') => {
    setReportData(data => data.map(item => item.id === id ? { ...item, chartType: type } : item));
  };

  const toggleChartColumnSelection = (id: number, colIndex: number) => {
    setReportData(data => data.map(item => {
      if (item.id === id) {
        if (!item.tableData || item.tableData.length === 0) return item;
        const columnCount = item.tableData[0].length;
        let newSelection = item.chartColumnSelection ? [...item.chartColumnSelection] : Array(columnCount).fill(true);
        if (newSelection.length < columnCount) {
            newSelection = [...newSelection, ...Array(columnCount - newSelection.length).fill(true)];
        }
        newSelection[colIndex] = !newSelection[colIndex];
        return { ...item, chartColumnSelection: newSelection };
      }
      return item;
    }));
  };

  const handleMergeSelection = (id: number) => {
    if (!tableSelection || tableSelection.id !== id) return;
    setReportData(data => data.map(item => {
      if (item.id === id && item.tableData) {
        let currentSpans = item.tableSpans ? item.tableSpans.map(r => [...r]) : item.tableData.map(r => r.map(() => 1));
        const minR = Math.min(tableSelection.start.r, tableSelection.end.r);
        const maxR = Math.max(tableSelection.start.r, tableSelection.end.r);
        const minC = Math.min(tableSelection.start.c, tableSelection.end.c);
        const maxC = Math.max(tableSelection.start.c, tableSelection.end.c);
        
        const rowSpan = maxR - minR + 1;
        const colSpan = maxC - minC + 1;
        
        if (rowSpan > 1 || colSpan > 1) {
            currentSpans[minR][minC] = { rowspan: rowSpan, colspan: colSpan };
            // Clear spans for the rest of the merged area to avoid overlapping merges
            for (let r = minR; r <= maxR; r++) {
               for (let c = minC; c <= maxC; c++) {
                  if (r === minR && c === minC) continue;
                  currentSpans[r][c] = 1;
               }
            }
            return { ...item, tableSpans: currentSpans };
        }
      }
      return item;
    }));
    setTableContextMenu(null);
    setTableSelection(null);
  };

  const unmergeCell = (id: number, rIdx: number, cIdx: number) => {
    setReportData(data => data.map(item => {
      if (item.id === id && item.tableSpans) {
        let currentSpans = item.tableSpans.map(r => [...r]);
        if (currentSpans[rIdx]) {
           currentSpans[rIdx][cIdx] = 1;
        }
        return { ...item, tableSpans: currentSpans };
      }
      return item;
    }));
    setTableContextMenu(null);
  };

  const handleParishChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    handleSave(false);
    const newParish = e.target.value;
    setParish(newParish);
    setChurch(PARISH_CHURCH_MAP[newParish][0]);
  };
  
  const handleChurchChange = (newChurch: string) => {
    handleSave(false);
    setChurch(newChurch);
  };

  const getCellSelectionStyle = (id: number, rIdx: number, cIdx: number) => {
    if (!tableSelection || tableSelection.id !== id || !tableSelection.start || !tableSelection.end) return '';
    const minR = Math.min(tableSelection.start.r, tableSelection.end.r);
    const maxR = Math.max(tableSelection.start.r, tableSelection.end.r);
    const minC = Math.min(tableSelection.start.c, tableSelection.end.c);
    const maxC = Math.max(tableSelection.start.c, tableSelection.end.c);
    if (rIdx >= minR && rIdx <= maxR && cIdx >= minC && cIdx <= maxC) {
      return 'ring-2 ring-inset ring-blue-500 bg-blue-50/50';
    }
    return '';
  };

  const handleReset = () => {
    if (window.confirm("현재 작성된 내용이 모두 삭제되고 초기화됩니다. 계속하시겠습니까?")) {
      setReportData(DEFAULT_REPORT);
      setLastSaved(null);
      setStatus('draft');
      const key = `report_${parish}_${church}`;
      localStorage.removeItem(key);
    }
  };

  const handleSave = async (isSubmit: boolean = false) => {
    setIsSaving(true);
    const key = `report_${parish}_${church}`;
    const timestamp = new Date().toLocaleString('ko-KR');
    const newStatus = isSubmit ? 'submitted' : 'draft';
    const saveData = { data: reportData, lastSaved: timestamp, status: newStatus };

    // 항상 로컬 저장
    localStorage.setItem(key, JSON.stringify(saveData));
    setLastSaved(timestamp);
    setStatus(newStatus);

    // 저장/제출 모두 클라우드 업로드
    if (supabase) {
      try {
        setDriveSaveResult('saving');
        const result = await saveDbData(`${parish}_${church}`, {
          id: `${parish}_${church}`,
          parish,
          church,
          ...saveData,
          updated_at: new Date().toISOString()
        });
        setDriveSaveResult(result as any || null);
        if (result === 'ok') setDriveSavedAt(new Date().toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit', second: '2-digit' }));
      } catch (e) {
        console.error("Manual save failed", e);
        setDriveSaveResult('error');
      }
    }
    setTimeout(() => setIsSaving(false), 600);
  };

  const updateText = (id: number, text: string) => {
    setReportData(data => data.map(item => item.id === id ? { ...item, text } : item));
  };

  const changeLevel = (id: number, delta: number) => {
    setReportData(data => data.map(item => {
      if (item.id === id && !item.isFixed) {
        const newLevel = Math.max(1, Math.min(5, item.level + delta));
        return { ...item, level: newLevel };
      }
      return item;
    }));
  };

  const moveL0Block = (index: number, direction: 'up' | 'down') => {
    const blocks: ReportItem[][] = [];
    let currentBlock: ReportItem[] = [];

    reportData.forEach((item) => {
      if (item.level === 0) {
        if (currentBlock.length > 0) {
          blocks.push(currentBlock);
        }
        currentBlock = [item];
      } else {
        currentBlock.push(item);
      }
    });
    if (currentBlock.length > 0) {
      blocks.push(currentBlock);
    }

    const targetItem = reportData[index];
    const blockIndex = blocks.findIndex(b => b[0]?.id === targetItem.id);

    if (blockIndex === -1) return;

    if (direction === 'up' && blockIndex > 0) {
      const temp = blocks[blockIndex];
      blocks[blockIndex] = blocks[blockIndex - 1];
      blocks[blockIndex - 1] = temp;
    } else if (direction === 'down' && blockIndex < blocks.length - 1) {
      const temp = blocks[blockIndex];
      blocks[blockIndex] = blocks[blockIndex + 1];
      blocks[blockIndex + 1] = temp;
    } else {
      return;
    }

    const newReportData = blocks.flat();
    setReportData(newReportData);
  };

  const addNewItem = (insertIndex: number = -1, defaultLevel: number = 1) => {
    const newItem = { id: nextId, text: "", level: defaultLevel };
    setNextId(prev => prev + 1);
    
    if (insertIndex === -1) {
      setReportData(data => [...data, newItem]);
    } else {
      setReportData(data => {
        const newData = [...data];
        newData.splice(insertIndex + 1, 0, newItem);
        return newData;
      });
    }
    
    setTimeout(() => {
      document.getElementById(`input-${newItem.id}`)?.focus();
    }, 10);
  };

  const removeItem = (id: number) => {
    setReportData(data => data.filter(item => item.id !== id));
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>, id: number, index: number) => {
    if (e.key === 'Tab') {
      e.preventDefault();
      changeLevel(id, e.shiftKey ? -1 : 1);
    } else if (e.key === 'Enter') {
      e.preventDefault();
      addNewItem(index, reportData[index].level);
    } else if (e.key === 'Backspace' && reportData[index].text === '') {
      e.preventDefault();
      removeItem(id);
      
      let prevIndex = index - 1;
      while (prevIndex >= 0 && reportData[prevIndex].isFixed) {
          prevIndex--;
      }
      if (prevIndex >= 0) {
        setTimeout(() => {
          document.getElementById(`input-${reportData[prevIndex].id}`)?.focus();
        }, 10);
      }
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      let prevIndex = index - 1;
      while (prevIndex >= 0 && reportData[prevIndex].isFixed) {
          prevIndex--;
      }
      if (prevIndex >= 0) {
        document.getElementById(`input-${reportData[prevIndex].id}`)?.focus();
      }
    } else if (e.key === 'ArrowDown') {
      e.preventDefault();
      let nextIndex = index + 1;
      while (nextIndex < reportData.length && reportData[nextIndex].isFixed) {
          nextIndex++;
      }
      if (nextIndex < reportData.length) {
        document.getElementById(`input-${reportData[nextIndex].id}`)?.focus();
      }
    }
  };

  const checkWithAI = async () => {
    setIsCheckingAI(true);
    setShowAiModal(true);
    setAiCorrections(null);

    try {
      const churches = PARISH_CHURCH_MAP[parish] || [church];

      // 제출 완료 교회만 + 세션 캐시 우선으로 병렬 취합
      const results = await Promise.all(
        churches.map(async (c) => {
          if (c === church) {
            if (status !== 'submitted') return { c, data: [] }; // 현재 교회도 미제출이면 스킵
            return { c, data: getCleanData(reportData) };
          }
          const report = await getReportDataFor(parish, c);
          if (!report || report.status !== 'submitted') return { c, data: [] }; // 미제출 스킵
          return { c, data: report.data ? getCleanData(report.data) : [] };
        })
      );

      const allPayload: any[] = [];
      for (const { c, data } of results) {
        data.filter((item: ReportItem) => item.text.trim() !== "").forEach((item: ReportItem) => {
          allPayload.push({ church: c, id: item.id, text: item.text });
        });
      }

      const aiPrompt = `당신은 교구 주간업무보고서를 검토하는 전문 편집자입니다.
아래 제공된 데이터의 텍스트(text)와 구조(level)를 함께 검토하세요. 대충 작성된 텍스트라도 맥락을 파악하세요.

문서 서식 규칙 (level 값):
- level 0: 대분류 제목 (예: "전주 결과보고", "금주 계획") — Ⅰ. Ⅱ. 형식
- level 1: 중분류 항목 (세부 내용의 소제목) — 1. 2. 형식
- level 2: 세부 내용 (구체적인 실행 내용) — 1) 2) 형식
- level 3: 부가 설명 — ① ② 형식
- level 4: 세부 사항 — 가. 나. 형식
- level 5: 최하위 항목 — a. b. 형식

검토 사항:
1. 오타 또는 문맥상 어색한 텍스트
2. 주간보고 양식(~함, ~예정 등 개조식)에 맞지 않는 항목
3. level이 내용에 맞지 않는 항목 (잘못된 계층 구조 → 올바른 level로 교정)

반드시 아래 JSON 배열 형태로만 응답하세요. (백틱이나 markdown 없이 순수 JSON만)
수정이 필요한 항목만 포함하세요. level과 text 중 변경 없는 필드는 원본 그대로 유지하세요.
[{ "church": "교회이름", "id": 1, "original": "원래 텍스트", "corrected": "교정된 텍스트", "level": 1, "reason": "이유" }]`;

      let text = "";

      if (isLocalMode) {
        try {
          const serverUrl = getLocalServerUrl();
          
          const res = await fetch(`${serverUrl}/api/ollama-chat`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              prompt: aiPrompt + `\n\n데이터:\n${JSON.stringify(allPayload, null, 2)}`
            })
          });
          
          if (res.ok) {
            const data = await res.json();
            text = data.text;
          } else {
            const errData = await res.json().catch(() => ({}));
            throw new Error(errData.error || "Ollama API failed");
          }
        } catch (e: any) {
          console.error("Local Ollama AI check failed:", e);
          alert(`로컬 AI 검토에 실패했습니다.\n사유: ${e.message}\nPC에서 Ollama가 켜져 있는지 확인해 주세요.\n(Ollama가 실행 중이 아니라면 우측 상단의 '클라우드 모드'로 전환하여 검토하실 수 있습니다.)`);
          setIsCheckingAI(false);
          setShowAiModal(false);
          return;
        }
      } else {
        let googleApiKey = localStorage.getItem('GEMINI_KEY') || 'AIzaSyAZBlFO30dN6Y1kOOmH1I24wCDqQi-xm-M';
        
        // 1. 구글 Gemini REST API (가장 빠르고 확실한 방법)
        try {
          const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${googleApiKey}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              contents: [{ parts: [{ text: aiPrompt + `\n\n데이터:\n${JSON.stringify(allPayload, null, 2)}` }] }],
              generationConfig: { responseMimeType: "application/json" }
            })
          });
          
          if (response.ok) {
            const data = await response.json();
            text = data.candidates?.[0]?.content?.parts?.[0]?.text || "";
          } else if (response.status === 429) {
            throw new Error("RATE_LIMIT");
          } else {
            throw new Error("API_ERROR");
          }
        } catch (e: any) {
          console.warn("Gemini REST API failed", e);
          if (e.message === "RATE_LIMIT") {
              const newKey = prompt("기본 제공된 구글 AI 키의 사용량이 초과되었습니다 (오류 429).\n본인의 구글 AI Studio API 키를 입력해주시면 계속 사용 가능합니다.\n취소하시면 우회 채널(OpenRouter)로 재시도합니다.");
              if (newKey) {
                  localStorage.setItem('GEMINI_KEY', newKey);
                  alert("키가 저장되었습니다. 다시 검토 버튼을 눌러주세요.");
                  setShowAiModal(false);
                  setIsCheckingAI(false);
                  return;
              }
          }
        }

        // 2. OpenRouter API 폴발 (우회 경로)
        if (!text) {
          let openRouterKey = localStorage.getItem('OPENROUTER_KEY');
          if (!openRouterKey) {
            const inputKey = prompt("Gemini에 연결할 수 없습니다.\n오픈라우터(OpenRouter) API 키를 입력하시면 무료 모델로 우회합니다.");
            if (inputKey) {
              localStorage.setItem('OPENROUTER_KEY', inputKey);
              openRouterKey = inputKey;
            }
          }

          if (openRouterKey) {
            try {
              const res = await fetch("https://openrouter.ai/api/v1/chat/completions", {
                method: "POST",
                headers: {
                  "Authorization": `Bearer ${openRouterKey}`,
                  "Content-Type": "application/json"
                },
                body: JSON.stringify({
                  model: "google/gemini-2.0-flash-lite-preview-02-05:free",
                  response_format: { type: "json_object" },
                  messages: [{ role: "user", content: aiPrompt + `\n\n데이터:\n${JSON.stringify(allPayload, null, 2)}` }]
                })
              });
              if (res.ok) {
                const json = await res.json();
                text = json.choices[0].message.content;
              }
            } catch(e) {
              console.error("OpenRouter fallback failed", e);
            }
          }
        }
      }
      
      if (!text) throw new Error("AI 응답을 받지 못했습니다.");
      
      let cleanText = text.replace(/```json/gi, '').replace(/```/gi, '').trim();
      const match = cleanText.match(/\[\s*\{[\s\S]*\}\s*\]/);
      if (match) cleanText = match[0];
      
      const corrections = JSON.parse(cleanText);
      setAiCorrections(corrections);
    } catch (err) {
      console.error(err);
      alert("AI 검토 중 한도 초과 또는 오류가 발생했습니다. 잠시 후 시도해주세요.");
      setShowAiModal(false);
    } finally {
      setIsCheckingAI(false);
    }
  };

  const applyCorrection = async (churchName: string, id: number, correctedText: string, correctedLevel?: number) => {
    if (churchName === church) {
      setReportData(data => data.map(item => item.id === id ? { ...item, text: correctedText, ...(correctedLevel !== undefined ? { level: correctedLevel } : {}) } : item));
    } else {
      const key = `report_${parish}_${churchName}`;
      const saved = localStorage.getItem(key);
      if (saved) {
        try {
           const parsed = JSON.parse(saved);
           if (parsed.data) {
             parsed.data = parsed.data.map((item: any) => item.id === id ? { ...item, text: correctedText, ...(correctedLevel !== undefined ? { level: correctedLevel } : {}) } : item);
             localStorage.setItem(key, JSON.stringify(parsed));
             if (supabase) {
                await saveDbData(`${parish}_${churchName}`, { id: `${parish}_${churchName}`, ...parsed, updated_at: new Date().toISOString() });
             }
           }
        } catch(e){}
      }
    }
    setAiCorrections(prev => prev ? prev.filter(c => !(c.id === id && c.church === churchName)) : null);
  };

  const applyAllCorrections = async () => {
    if (!aiCorrections) return;

    const byChurch: Record<string, {id: number, text: string, level?: number}[]> = {};
    aiCorrections.forEach(c => {
       if (!byChurch[c.church]) byChurch[c.church] = [];
       byChurch[c.church].push({ id: c.id, text: c.corrected, level: c.level });
    });

    for (const [cName, updates] of Object.entries(byChurch)) {
       const updateMap = new Map(updates.map(u => [u.id, u]));
       if (cName === church) {
          setReportData(data => data.map(item => {
            const u = updateMap.get(item.id);
            if (!u) return item;
            return { ...item, text: u.text, ...(u.level !== undefined ? { level: u.level } : {}) };
          }));
       } else {
          const key = `report_${parish}_${cName}`;
          const saved = localStorage.getItem(key);
          if (saved) {
             try {
                const parsed = JSON.parse(saved);
                if (parsed.data) {
                   parsed.data = parsed.data.map((item: any) => {
                     const u = updateMap.get(item.id);
                     if (!u) return item;
                     return { ...item, text: u.text, ...(u.level !== undefined ? { level: u.level } : {}) };
                   });
                   localStorage.setItem(key, JSON.stringify(parsed));
                   if (supabase) {
                      await saveDbData(`${parish}_${cName}`, { id: `${parish}_${cName}`, ...parsed, updated_at: new Date().toISOString() });
                   }
                }
             } catch(e){}
          }
       }
    }

    setAiCorrections([]);
  };

  const exportToWord = async () => {
    const churches = PARISH_CHURCH_MAP[parish];
    let allChildren: (Paragraph | Table)[] = [];
    
    // Add Parish Header
    allChildren.push(
      new Paragraph({
        children: [
          new TextRun({
            text: `<${getDisplayParish(parish)}> 주간업무보고`,
            size: 36, // 18pt
            bold: true,
          })
        ],
        alignment: "center" as const,
        spacing: { before: 200, after: 200 },
        border: {
          top: { color: "1E88E5", space: 10, size: 12, style: BorderStyle.SINGLE },
          bottom: { color: "1E88E5", space: 10, size: 24, style: BorderStyle.SINGLE },
        }
      })
    );

    let churchIndex = 1;

    for (const c of churches) {
      let dataToUse: ReportItem[] = [];
      if (c === church) {
        dataToUse = getCleanData(reportData);
      } else {
        const key = `report_${parish}_${c}`;
        const saved = localStorage.getItem(key);
        if (saved) {
          try {
            const parsed = JSON.parse(saved);
            dataToUse = getCleanData(parsed.data || []);
          } catch (e) {
             dataToUse = [];
          }
        }
      }

      if (dataToUse.length === 0) continue;

      let counters = { 0: 0, 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };

      // Church Header
      allChildren.push(
        new Paragraph({
          children: [
            new TextRun({
              text: `${churchIndex}. ${getDisplayChurch(c)}`,
              size: 28, // 14pt
              bold: true,
              color: "0000FF",
              font: "맑은 고딕"
            })
          ],
          spacing: { before: 400, after: 200 }
        })
      );
      churchIndex++;

      const paragraphs: (Paragraph | Table)[] = [];
      for (const item of dataToUse) {
        let prefix = "";
        if (item.level === 0) {
            counters[0]++; counters[1] = 0; counters[2] = 0; counters[3] = 0; counters[4] = 0; counters[5] = 0;
            prefix = toRoman(counters[0]) + ". ";
        } else if (item.level === 1) {
            counters[1]++; counters[2] = 0; counters[3] = 0; counters[4] = 0; counters[5] = 0;
            prefix = counters[1] + ". ";
        } else if (item.level === 2) {
            counters[2]++; counters[3] = 0; counters[4] = 0; counters[5] = 0;
            prefix = counters[2] + ") ";
        } else if (item.level === 3) {
            counters[3]++; counters[4] = 0; counters[5] = 0;
            prefix = toCircled(counters[3]) + " ";
        } else if (item.level === 4) {
            counters[4]++; counters[5] = 0;
            prefix = toKorSyllable(counters[4]) + ". ";
        } else if (item.level === 5) {
            counters[5]++;
            prefix = toLatin(counters[5]) + ". ";
        }

        let color = "000000";
        if (item.level <= 1) color = "1D4ED8"; // Tailwind blue-700
        
        const paras: (Paragraph | Table)[] = [];

        const lines = `${prefix}${item.text || '(빈 항목)'}`.split('\n');

        paras.push(new Paragraph({
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

        if (item.image && item.imageWidth && item.imageHeight) {
          try {
            let imageBuffer: ArrayBuffer;
            if (item.image.startsWith("http")) {
              const res = await fetch(item.image);
              imageBuffer = await res.arrayBuffer();
            } else {
              const base64Data = item.image.split(",")[1];
              imageBuffer = Uint8Array.from(atob(base64Data), char => char.charCodeAt(0));
            }
            
            const targetWidth = 500;
            const ratio = Math.min(1, targetWidth / item.imageWidth);
            const finalWidth = item.imageWidth * ratio;
            const finalHeight = item.imageHeight * ratio;

            paras.push(new Paragraph({
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
              spacing: {
                before: 120, 
              }
            }));
          } catch (e) {
            console.error("Failed to add image to word document", e);
          }
        }

        if (item.tableData && item.tableData.length > 0) {
          const skippedCells = new Set<string>();
          const table = new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: item.tableData.map((row, rIdx) => {
              const cells: TableCell[] = [];
              row.forEach((cell, cIdx) => {
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
                const isHighlighted = item.tableHighlights?.[rIdx]?.[cIdx];
                const align = item.tableAlignments?.[rIdx]?.[cIdx] || (colSpan > 1 ? 'center' : 'left');
                cells.push(new TableCell({
                  columnSpan: colSpan > 1 ? colSpan : undefined,
                  rowSpan: rowSpan > 1 ? rowSpan : undefined,
                  children: cell.split('\n').map(line => new Paragraph({ 
                    children: [new TextRun({
                      text: line,
                      font: "맑은 고딕",
                      size: 22, // 11pt
                      bold: !!isHighlighted,
                    })],
                    alignment: align === 'center' ? "center" : align === 'right' ? "right" : "left",
                  })),
                  verticalAlign: VerticalAlign.CENTER,
                  margins: { top: 100, bottom: 100, left: 100, right: 100 },
                  shading: isHighlighted ? { fill: "E6F2FF" } : undefined,
                  borders: {
                    top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                    bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                    left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                    right: { style: BorderStyle.SINGLE, size: 1, color: "000000" }
                  }
                }));
              });
              return new TableRow({ children: cells });
            })
          });
          
          paras.push(new Paragraph({ spacing: { before: 120 } }));
          paras.push(table);
        }

        paragraphs.push(...paras);
      }

      allChildren.push(...paragraphs);
    }

    if (allChildren.length <= 1) {
      alert("출력할 교구 데이터가 없습니다. 문서를 작성해주세요.");
      return;
    }

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            size: {
              width: 11906, // A4 width
              height: 16838, // A4 height
            }
          }
        },
        children: allChildren,
      }],
    });

    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${getDisplayParish(parish)}_주간보고_${new Date().toISOString().split('T')[0]}.docx`;
    a.click();
    URL.revokeObjectURL(url);
    
    // Save state as completed
    handleSave();
  };

const renderPreviewLines = () => {
    let counters = { 0: 0, 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
    const cleanData = getCleanData(reportData);

    if (cleanData.length === 0) {
      return (
        <div className="flex flex-col items-center justify-center p-12 text-slate-400 border-2 border-dashed border-slate-200 rounded-xl mt-8">
          <AlertCircle className="w-12 h-12 mb-3 text-slate-300" />
          <p className="font-medium text-slate-500">작성된 내용이 없습니다</p>
          <p className="text-sm mt-1 text-slate-400 text-center">좌측 에디터에서 내용을 작성하시면<br/>이곳에 문서가 미리보기로 표시됩니다.</p>
        </div>
      );
    }
    
    return cleanData.map(item => {
      let prefix = "";
      if (item.level === 0) {
          counters[0]++; counters[1] = 0; counters[2] = 0; counters[3] = 0; counters[4] = 0; counters[5] = 0;
          prefix = toRoman(counters[0]) + ". ";
      } else if (item.level === 1) {
          counters[1]++; counters[2] = 0; counters[3] = 0; counters[4] = 0; counters[5] = 0;
          prefix = counters[1] + ". ";
      } else if (item.level === 2) {
          counters[2]++; counters[3] = 0; counters[4] = 0; counters[5] = 0;
          prefix = counters[2] + ") ";
      } else if (item.level === 3) {
          counters[3]++; counters[4] = 0; counters[5] = 0;
          prefix = toCircled(counters[3]) + " ";
      } else if (item.level === 4) {
          counters[4]++; counters[5] = 0;
          prefix = toKorSyllable(counters[4]) + ". ";
      } else if (item.level === 5) {
          counters[5]++;
          prefix = toLatin(counters[5]) + ". ";
      }

      let colorClass = "";
      if (item.level === 0) colorClass = "text-blue-700 font-bold underline text-[1.05rem]";
      else if (item.level === 1) colorClass = "text-blue-700 ml-2 font-bold";
      else if (item.level === 2) colorClass = "text-slate-800 ml-6";
      else if (item.level === 3) colorClass = "text-slate-700 ml-10";
      else if (item.level === 4) colorClass = "text-slate-600 ml-14";
      else if (item.level === 5) colorClass = "text-slate-600 ml-16";

      return (
        <div key={item.id} className={`leading-snug mb-0.5 ${colorClass}`}>
          <div className="whitespace-pre-line">{prefix}{item.text || ''}</div>
          {item.image && (
            <div className={`mt-2 ${item.level === 0 ? 'ml-0' : item.level === 1 ? 'ml-2' : item.level === 2 ? 'ml-8' : 'ml-12'}`}>
              <img src={item.image} alt="첨부됨" className="max-w-full max-h-[400px] object-contain inline-block rounded border border-slate-200 shadow-sm" />
            </div>
          )}
          {item.tableData && item.tableData.length > 0 && (
            <div className={`mt-3 mb-2 overflow-x-auto ${item.level === 0 ? 'ml-0' : item.level === 1 ? 'ml-2' : item.level === 2 ? 'ml-8' : 'ml-12'}`}>
              {item.chartType !== 'none' && item.chartType ? (
                <div className="h-64 mt-2 p-4 bg-white border border-slate-200 shadow-sm rounded">
                  <ResponsiveContainer width="100%" height="100%">
                    {(() => {
                        const headers = item.tableData![0];
                        const data = item.tableData!.slice(1).map(row => {
                            const obj: any = { name: row[0] };
                            for (let i = 1; i < row.length; i++) {
                                obj[headers[i] || `col${i}`] = Number(row[i]) || 0;
                            }
                            return obj;
                        });
                        const selectedColIndices = item.chartColumnSelection 
                            ? item.chartColumnSelection.map((sel, idx) => sel && idx > 0 ? idx : -1).filter(idx => idx !== -1)
                            : headers.map((_, idx) => idx > 0 ? idx : -1).filter(idx => idx !== -1);
                        if (selectedColIndices.length === 0 && headers.length > 1) selectedColIndices.push(1);
                        const keys = selectedColIndices.map(idx => headers[idx] || `col${idx}`);
                        const colors = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6'];

                        if (item.chartType === 'bar') {
                            return (
                                <BarChart data={data}>
                                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                                    <XAxis dataKey="name" tick={{fontSize: 12}} stroke="#94a3b8" />
                                    <YAxis tick={{fontSize: 12}} stroke="#94a3b8" />
                                    <Tooltip contentStyle={{ fontSize: 12, borderRadius: 4, padding: 8 }} />
                                    <Legend wrapperStyle={{fontSize: "12px"}} />
                                    {keys.map((k, i) => <Bar key={k} dataKey={k} fill={colors[i % colors.length]} radius={[2,2,0,0]} />)}
                                </BarChart>
                            );
                        } else if (item.chartType === 'line') {
                            return (
                                <LineChart data={data}>
                                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                                    <XAxis dataKey="name" tick={{fontSize: 12}} stroke="#94a3b8" />
                                    <YAxis tick={{fontSize: 12}} stroke="#94a3b8" />
                                    <Tooltip contentStyle={{ fontSize: 12, borderRadius: 4, padding: 8 }} />
                                    <Legend wrapperStyle={{fontSize: "12px"}} />
                                    {keys.map((k, i) => <Line key={k} type="monotone" dataKey={k} stroke={colors[i % colors.length]} strokeWidth={2} />)}
                                </LineChart>
                            );
                        } else {
                            const pieData = data.map(d => ({ name: d.name, value: d[keys[0]] || 0 }));
                            return (
                                <PieChart>
                                    <Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80} label={({name, percent}) => `${name} ${(percent * 100).toFixed(0)}%`} labelLine={false} style={{fontSize: 12}}>
                                        {pieData.map((entry, index) => (
                                            <Cell key={`cell-${index}`} fill={colors[index % colors.length]} />
                                        ))}
                                    </Pie>
                                    <Tooltip contentStyle={{ fontSize: 12, borderRadius: 4, padding: 8 }} />
                                    <Legend wrapperStyle={{fontSize: "12px"}} />
                                </PieChart>
                            );
                        }
                    })()}
                  </ResponsiveContainer>
                </div>
              ) : (
                <table className="w-full border-collapse text-sm my-2">
                  <tbody>
                    {(() => {
                      const skippedCells = new Set<string>();
                      return item.tableData.map((row, rIdx) => (
                        <tr key={rIdx}>
                          {row.map((cell, cIdx) => {
                            if (skippedCells.has(`${rIdx},${cIdx}`)) return null;
                            
                            const spanDef = item.tableSpans?.[rIdx]?.[cIdx];
                            const colSpan = (typeof spanDef === 'number' ? spanDef : spanDef?.colspan) || 1;
                            const rowSpan = (typeof spanDef === 'number' ? 1 : spanDef?.rowspan) || 1;
                            
                            for (let r = 0; r < rowSpan; r++) {
                              for (let c = 0; c < colSpan; c++) {
                                if (r === 0 && c === 0) continue;
                                skippedCells.add(`${rIdx + r},${cIdx + c}`);
                              }
                            }
                            const isHighlighted = item.tableHighlights?.[rIdx]?.[cIdx];
                            const align = item.tableAlignments?.[rIdx]?.[cIdx] || (colSpan > 1 ? 'center' : 'left');
                            const alignClass = align === 'center' ? 'text-center' : align === 'right' ? 'text-right' : 'text-left';
                            return (
                              <td key={cIdx} colSpan={colSpan} rowSpan={rowSpan} className={`border border-slate-400 p-2 whitespace-pre-line break-words ${isHighlighted ? 'bg-blue-50 font-bold text-slate-900 border-blue-300' : 'bg-white text-slate-800'} ${alignClass}`}>
                                {cell}
                              </td>
                            );
                          })}
                        </tr>
                      ));
                    })()}
                  </tbody>
                </table>
              )}
            </div>
          )}
        </div>
      );
    });
  }
                      const resetActiveTabParishData = async () => {
    const password = prompt("초기화 비밀번호를 입력하세요:");
    if (password !== "skmt0909!") {
      if (password !== null) alert("비밀번호가 일치하지 않습니다.");
      return;
    }

    let targetParishes: string[] = [];
    let confirmMsg = "";
    
    if (activeTab === 'report') {
      confirmMsg = "정말로 [모든 교구]의 데이터를 초기화하시겠습니까?\n(협회 데이터는 유지되며, 이 작업은 복구할 수 없습니다!)";
      targetParishes = Object.keys(PARISH_CHURCH_MAP).filter(p => p !== '협회');
    } else if (activeTab === 'association') {
      confirmMsg = "정말로 [협회]의 전체 데이터를 초기화하시겠습니까?\n(이 작업은 복구할 수 없습니다!)";
      targetParishes = ['협회'];
    } else {
      return;
    }

    if (window.confirm(confirmMsg)) {
      const defaultData = { data: DEFAULT_REPORT, lastSaved: null, status: 'draft' };
      
      for (const p of targetParishes) {
        for (const c of PARISH_CHURCH_MAP[p]) {
          const key = `report_${p}_${c}`;
          localStorage.setItem(key, JSON.stringify(defaultData));
          
          if (supabase) {
            try {
              await supabase.from('reports').upsert({
                id: `${p}_${c}`,
                ...defaultData,
                updated_at: new Date().toISOString()
              });
            } catch (e) {
              console.error(e);
            }
          }
        }
      }
      setReportData(DEFAULT_REPORT);
      setLastSaved(null);
      setStatus('draft');
      updateParishStats();
      alert("데이터가 성공적으로 초기화되었습니다.");
    }
  };

  const getParishData = (jsonFormat: 'flat' | 'tree' = 'flat') => {
    const parishData: Record<string, any> = {};
    const churches = PARISH_CHURCH_MAP[parish] || [];
    
    churches.forEach(c => {
      let dataToUse: ReportItem[] = [];
      
      // 현재 선택된 교회는 입력 중인 최신 상태(reportData)를 사용
      if (c === church) {
        dataToUse = reportData;
      } else {
        // 다른 교회들은 LocalStorage에서 불러오기
        const key = `report_${parish}_${c}`;
        const saved = localStorage.getItem(key);
        if (saved) {
          try {
            const parsed = JSON.parse(saved);
            dataToUse = parsed.data || DEFAULT_REPORT;
          } catch (e) {
            dataToUse = DEFAULT_REPORT;
          }
        } else {
          dataToUse = DEFAULT_REPORT;
        }
      }

      if (jsonFormat === 'flat') {
        parishData[c] = getCleanData(dataToUse);
      } else {
        parishData[c] = buildTree(getCleanData(dataToUse));
      }
    });

    return { 
      parish: parish,
      exportDate: new Date().toISOString(),
      reports: parishData 
    };
  };

  const handleCopyJson = () => {
    const dataToExport = getParishData();
    navigator.clipboard.writeText(JSON.stringify(dataToExport, null, 2));
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const resetAllData = async () => {
    const password = prompt("전체 초기화 비밀번호를 입력하세요:");
    if (password !== "skmt0909!") {
      if (password !== null) alert("비밀번호가 일치하지 않습니다.");
      return;
    }

    if (window.confirm("정말로 모든 교구/교회의 데이터를 초기화하시겠습니까?\n이 작업은 복구할 수 없습니다!")) {
      const defaultData = { data: DEFAULT_REPORT, lastSaved: null, status: 'draft' };
      
      for (const p of Object.keys(PARISH_CHURCH_MAP)) {
        for (const c of PARISH_CHURCH_MAP[p]) {
          const key = `report_${p}_${c}`;
          localStorage.setItem(key, JSON.stringify(defaultData));
          
          if (supabase) {
            try {
              await supabase.from('reports').upsert({
                id: `${p}_${c}`,
                ...defaultData,
                updated_at: new Date().toISOString()
              });
            } catch (e) {
              console.error(e);
            }
          }
        }
      }
      setReportData(DEFAULT_REPORT);
      setLastSaved(null);
      setStatus('draft');
      updateParishStats();
      alert("전 교구 데이터가 성공적으로 초기화되었습니다.");
    }
  };

  return (
    <div className="min-h-screen bg-slate-100 p-2 sm:p-4 md:p-6 font-sans text-slate-800 flex flex-col">
      <div className="w-full max-w-full px-1 sm:px-4 lg:px-8 mx-auto mb-4 flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex gap-2 overflow-x-auto pb-2 md:pb-0 scrollbar-hide snap-x whitespace-nowrap">
           <button onClick={() => { setActiveTab('report'); setParish('천원특별'); setChurch(PARISH_CHURCH_MAP['천원특별'][0]); }} className={`shrink-0 snap-start px-4 sm:px-5 py-2.5 font-bold rounded-lg transition-colors flex items-center gap-2 shadow-sm text-sm sm:text-base ${activeTab === 'report' ? 'bg-blue-600 text-white' : 'bg-white text-slate-600 hover:bg-slate-50 border border-slate-200'}`}><BookOpen className="w-4 h-4"/> 교구 업무보고 작성</button>
           <button onClick={handleAssociationTab} className={`shrink-0 snap-start px-4 sm:px-5 py-2.5 font-bold rounded-lg transition-colors flex items-center gap-2 shadow-sm text-sm sm:text-base ${activeTab === 'association' ? 'bg-indigo-600 text-white' : 'bg-white text-slate-600 hover:bg-slate-50 border border-slate-200'}`}><BookOpen className="w-4 h-4"/> 협회 업무보고 작성</button>
           <button onClick={handleNoticeWriteTab} className={`shrink-0 snap-start px-4 sm:px-5 py-2.5 font-bold rounded-lg transition-colors flex items-center gap-2 shadow-sm text-sm sm:text-base ${activeTab === 'notice_write' ? 'bg-emerald-600 text-white' : 'bg-white text-slate-600 hover:bg-slate-50 border border-slate-200'}`}><FileText className="w-4 h-4"/> 공지사항 올리기</button>
           <button onClick={() => setActiveTab('notice')} className={`shrink-0 snap-start px-4 sm:px-5 py-2.5 font-bold rounded-lg transition-colors flex items-center gap-2 shadow-sm text-sm sm:text-base ${activeTab === 'notice' ? 'bg-blue-600 text-white' : 'bg-white text-slate-600 hover:bg-slate-50 border border-slate-200'}`}><Bell className="w-4 h-4"/> 공지사항 확인</button>
            <button 
              onClick={() => {
                const pwd = prompt("관리자 비밀번호를 입력하세요:");
                if (pwd === "skmt0909!") {
                  setActiveTab('admin_console');
                } else if (pwd !== null) {
                  alert("비밀번호가 일치하지 않습니다.");
                }
              }} 
              className={`shrink-0 snap-start px-4 sm:px-5 py-2.5 font-bold rounded-lg transition-colors flex items-center gap-2 shadow-sm text-sm sm:text-base ${activeTab === 'admin_console' ? 'bg-purple-600 text-white font-extrabold shadow-purple-200' : 'bg-white text-slate-600 hover:bg-slate-50 border border-slate-200'}`}
            >
              <Settings className="w-4 h-4"/> 관리자 취합 콘솔
            </button>
        </div>
        
        <div className="shrink-0 flex items-center gap-2 bg-white px-3.5 py-2 rounded-lg border border-slate-200 shadow-sm text-xs font-black self-end md:self-auto select-none">
          <span className="text-blue-600 font-extrabold">☁️ 클라우드</span>
        </div>
      </div>

      {activeTab === 'notice' && (
        <div className="w-full max-w-full px-2 sm:px-4 lg:px-8 mx-auto flex-1 flex flex-col h-[calc(100vh-8rem)]">
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 flex flex-col md:flex-row h-full overflow-hidden">
            <div className="w-full md:w-1/3 lg:w-1/4 border-b md:border-b-0 md:border-r border-slate-200 flex flex-col bg-slate-50 shrink-0">
              <div className="p-4 border-b border-slate-200 bg-white flex justify-between items-center">
                <h2 className="text-lg font-bold text-slate-800 flex items-center gap-2"><Bell className="w-5 h-5 text-blue-500"/> 공지사항</h2>
                {!isAdmin ? (
                  <button onClick={handleAdminLogin} className="text-xs text-slate-500 hover:text-blue-600 font-medium bg-slate-100 px-2 py-1 rounded">관리자 로그인</button>
                ) : (
                  <div className="flex gap-2">
                    <button 
                      onClick={async () => {
                        if(window.confirm('정말로 모든 공지사항을 삭제하시겠습니까? (복구 불가)')) {
                          await supabase.from('reports').upsert({ id: 'SYSTEM_NOTICES', data: [], updated_at: new Date().toISOString() });
                          setNotices([]);
                          setActiveNotice(null);
                          alert('모든 공지사항이 초기화되었습니다.');
                        }
                      }}
                      className="text-xs bg-red-50 hover:bg-red-100 text-red-600 border border-red-200 px-2.5 py-1.5 rounded font-bold transition-colors"
                    >
                      전체 초기화
                    </button>
                  </div>
                )}
              </div>
              <div className="flex-1 overflow-y-auto p-2 space-y-1">
                {notices.length === 0 ? (
                  <div className="text-center text-slate-400 py-10 text-sm">등록된 공지사항이 없습니다.</div>
                ) : (
                  notices.map(notice => (
                    <div key={notice.id} onClick={() => setActiveNotice(notice)} className={`p-3 rounded-lg cursor-pointer transition-colors border ${activeNotice?.id === notice.id ? 'bg-blue-50 border-blue-200 shadow-sm' : 'bg-white border-transparent hover:border-slate-200'}`}>
                      <div className="flex justify-between items-start">
                        <div className="font-medium text-slate-800 text-sm leading-snug">{notice.title}</div>
                        {isAdmin && (
                          <button onClick={(e) => { e.stopPropagation(); deleteNotice(notice.id); }} className="text-red-400 hover:text-red-600 p-1 shrink-0"><X className="w-3.5 h-3.5"/></button>
                        )}
                      </div>
                      <div className="text-xs text-slate-400 mt-2">{new Date(notice.created_at).toLocaleDateString()}</div>
                    </div>
                  ))
                )}
              </div>
            </div>
            <div className="w-full md:w-2/3 lg:w-3/4 bg-slate-200 flex-1 relative flex flex-col h-full overflow-y-auto">
              {activeNotice ? (
                <div className="flex-1 flex flex-col h-full">
                  {activeNotice.pdfUrl && (
                    <iframe src={activeNotice.pdfUrl} className="w-full border-0 flex-1 min-h-[1000px]" title="PDF Viewer" />
                  )}
                  {activeNotice.data && activeNotice.data.length > 0 && (
                    <div className="p-6 md:p-10 bg-white font-serif text-slate-900 mx-auto w-full max-w-4xl shadow-sm min-h-full">
                      <h1 className="text-2xl md:text-3xl font-black mb-8 pb-4 border-b-2 border-blue-600 text-center text-slate-800">{activeNotice.title}</h1>
                      {activeNotice.data.map((item: any) => (
                        <div key={item.id} className="mb-3" style={{ paddingLeft: `${item.level * 24}px` }}>
                          {item.level === 0 ? (
                            <div className="font-bold text-lg md:text-xl text-blue-800 border-b border-blue-100 pb-1 mt-6 mb-2">{item.text}</div>
                          ) : (
                            <div className="flex items-start gap-2 text-sm md:text-base text-slate-800">
                              <span className="text-blue-500 mt-0.5">•</span>
                              <div className="whitespace-pre-wrap leading-relaxed flex-1">{item.text}</div>
                            </div>
                          )}
                          {item.image && <img src={item.image} alt="첨부" className="mt-3 max-h-64 object-contain rounded-lg border border-slate-200 shadow-sm" />}
                          {item.tableData && (
                            <div className="mt-3 overflow-x-auto w-full">
                              <table className="w-full border-collapse border border-slate-300 text-xs md:text-sm text-center bg-white shadow-sm">
                                <tbody>
                                  {item.tableData.map((row: any, rIdx: number) => (
                                    <tr key={rIdx} className={rIdx === 0 ? "bg-slate-100 font-bold border-b-2 border-slate-300" : "hover:bg-slate-50 transition-colors"}>
                                      {row.map((cell: any, cIdx: number) => (
                                        <td key={cIdx} className={`border border-slate-300 p-2 md:p-3 ${cIdx === 0 ? 'bg-slate-50 font-bold text-slate-700' : 'text-slate-600'}`}>{cell}</td>
                                      ))}
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          )}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              ) : (
                <div className="flex-1 flex flex-col items-center justify-center text-slate-400">
                  <FileText className="w-16 h-16 mb-4 text-slate-300 opacity-50" />
                  <p>좌측에서 공지사항을 선택하시면 내용이 표시됩니다.</p>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {(activeTab === 'report' || activeTab === 'association' || activeTab === 'notice_write') && (
      <div className="w-full max-w-full px-1 sm:px-4 lg:px-8 mx-auto flex flex-col flex-1 min-h-0 overflow-hidden">
        
        {/* Mobile Tab Switcher + 제출 버튼 */}
        <div className="flex xl:hidden mb-3 gap-2">
          <div className="flex flex-1 bg-white rounded-lg shadow-sm p-1 border border-slate-200">
            <button onClick={() => setMobileView('editor')} className={`flex-1 py-2.5 text-sm font-bold rounded-md transition-colors ${mobileView === 'editor' ? 'bg-blue-600 text-white shadow-sm' : 'text-slate-500 hover:bg-slate-50'}`}>📝 작성창</button>
            <button onClick={() => setMobileView('preview')} className={`flex-1 py-2.5 text-sm font-bold rounded-md transition-colors ${mobileView === 'preview' ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-500 hover:bg-slate-50'}`}>👀 미리보기</button>
          </div>
          {activeTab !== 'notice_write' && (
            <button
              onClick={() => handleSave(status !== 'submitted')}
              disabled={isSaving}
              className={`shrink-0 px-4 py-2 rounded-lg text-sm font-bold shadow-sm transition-colors ${status === 'submitted' ? 'bg-emerald-100 text-emerald-800 border border-emerald-300' : 'bg-emerald-600 text-white hover:bg-emerald-700'}`}
            >
              {isSaving ? '저장중' : status === 'submitted' ? '✅ 제출됨' : '제출 확정'}
            </button>
          )}
        </div>

        <div className="grid grid-cols-1 xl:grid-cols-[65%_35%] gap-4 lg:gap-6 flex-1 min-h-0">
          {/* Editor Panel */}
          <div className={`bg-white p-4 sm:p-6 rounded-xl shadow-sm border border-slate-200 flex-col h-[calc(100vh-8rem)] xl:h-[calc(100vh-3rem)] ${mobileView === 'editor' ? 'flex' : 'hidden xl:flex'}`}>
          
          <div className="flex flex-col mb-4 pb-4 border-b border-slate-200 gap-3">
            <div className="flex items-center justify-between">
              <h2 className="text-xl font-bold flex items-center gap-2">
                {activeTab === 'association' ? '협회 보고서 작성' : activeTab === 'notice_write' ? '공지사항 작성 에디터' : '교구 보고서 작성'}
                {lastSaved && (
                  <span className="text-xs font-normal text-slate-500 flex items-center gap-1 bg-slate-100 px-2 py-1 rounded-full">
                    <Clock className="w-3 h-3" /> 최근 저장: {lastSaved}
                  </span>
                )}
              </h2>
              <div className="flex gap-2">
                <span className="flex items-center gap-1.5 text-sm px-3 py-1.5 rounded-md bg-emerald-50 border border-emerald-200 text-emerald-700">
                  <Check className="w-4 h-4" /> AI 준비됨
                </span>
                <button 
                  onClick={() => setShowGuideModal(true)}
                  className="flex items-center gap-1.5 text-sm bg-blue-50 border border-blue-200 hover:bg-blue-100 text-blue-700 font-bold px-3 py-1.5 rounded-md transition-colors shadow-sm"
                  title="작성 방법 안내"
                >
                  <BookOpen className="w-4 h-4" />
                  작성법 / 가이드
                </button>
                <button 
                  onClick={() => setShowJsonModal(true)}
                  className="flex items-center gap-2 text-sm bg-slate-100 border border-slate-200 hover:bg-slate-200 text-slate-700 px-3 py-1.5 rounded-md transition-colors shadow-sm"
                  title="데이터 구조 확인"
                >
                  <FileJson className="w-4 h-4" />
                  추출
                </button>
                <button 
                  onClick={resetAllData}
                  className="flex items-center gap-1.5 text-sm bg-red-50 border border-red-200 hover:bg-red-100 text-red-700 font-bold px-3 py-1.5 rounded-md transition-colors shadow-sm"
                  title="모든 교구/교회 데이터 초기화"
                >
                  <Trash2 className="w-4 h-4" />
                  전체 초기화
                </button>
              </div>
            </div>
            
            <div className="flex flex-col gap-3">
              <div className="flex gap-3">
                {activeTab === 'notice_write' ? (
                  <div className="flex-1 flex gap-2 items-end">
                    <div className="flex-1">
                      <label className="block text-xs font-bold text-slate-500 mb-1">공지사항 제목</label>
                      <input type="text" value={noticeTitle} onChange={e => setNoticeTitle(e.target.value)} className="w-full px-3 py-2 border border-slate-300 rounded-md text-sm font-medium" placeholder="제목을 입력하세요..." />
                    </div>
                    <label className="text-xs bg-slate-50 hover:bg-slate-100 text-slate-700 border border-slate-200 px-3 py-2.5 rounded cursor-pointer font-bold flex items-center gap-1 transition-colors shrink-0 h-[38px]">
                      {isUploadingNotice ? <RefreshCw className="w-3.5 h-3.5 animate-spin" /> : <Upload className="w-3.5 h-3.5" />}
                      {isUploadingNotice ? 'PDF 첨부중...' : 'PDF 첨부 (선택)'}
                      <input type="file" accept="application/pdf" className="hidden" onChange={handleNoticePdfUpload} disabled={isUploadingNotice} />
                    </label>
                    {noticePdfUrl && <span className="text-xs text-blue-600 font-bold mb-2 shrink-0 self-center">PDF 첨부됨</span>}
                  </div>
                ) : (
                  <>
                    {activeTab !== 'association' && (
                      <div className="flex-1">
                        <label htmlFor="parishSelect" className="block text-xs font-bold text-slate-500 mb-1">교구</label>
                        <select 
                          id="parishSelect" 
                          value={parish}
                          onChange={handleParishChange}
                          className="w-full px-3 py-2 border border-slate-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm font-medium"
                        >
                          {Object.keys(PARISH_CHURCH_MAP).filter(p => p !== '협회').map(p => (
                            <option key={p} value={p}>{getDisplayParish(p)}</option>
                          ))}
                        </select>
                      </div>
                    )}
                    <div className="flex-1">
                      <label htmlFor="churchSelect" className="block text-xs font-bold text-slate-500 mb-1">{activeTab === 'association' ? '협회 부서' : '교회'}</label>
                      <select 
                        id="churchSelect" 
                        value={church}
                        onChange={(e) => handleChurchChange(e.target.value)}
                        className="w-full px-3 py-2 border border-slate-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm font-medium"
                      >
                        {(PARISH_CHURCH_MAP[parish] || []).map(c => (
                          <option key={c} value={c}>{getDisplayChurch(c)}</option>
                        ))}
                      </select>
                    </div>
                  </>
                )}
              </div>
              <div className="bg-slate-50 border border-slate-200 rounded-lg p-3">
                <div className="flex justify-between items-center mb-2">
                  <span className="text-xs font-bold text-slate-500">교회별 제출 현황</span>
                  <div className="flex gap-3 text-[10px] text-slate-500">
                    <span className="flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>제출 확정</span>
                    <span className="flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full bg-blue-500"></span>작성 중</span>
                    <span className="flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full bg-slate-300"></span>미작성</span>
                  </div>
                </div>
                <div className="flex flex-wrap gap-2 text-xs">
                  {PARISH_CHURCH_MAP[parish].map(c => {
                    const stat = parishStats[c] || 'empty';
                    let textColor = '';
                    
                    if (stat === 'submitted') {
                      textColor = 'text-slate-700 font-bold';
                    } else if (stat === 'draft') {
                      textColor = 'text-slate-700 font-bold';
                    } else {
                      textColor = 'text-slate-400';
                    }
                    
                    return (
                      <div 
                        key={c}
                        className={`flex items-center gap-1.5 px-2 py-1 rounded-md border cursor-pointer hover:shadow-sm transition-all ${c === church ? 'ring-2 ring-blue-500 bg-white border-transparent' : 'bg-white border-slate-200'}`}
                        onClick={() => handleChurchChange(c)}
                      >
                        <span className={`w-2 h-2 rounded-full ${stat === 'submitted' ? 'bg-emerald-500' : stat === 'draft' ? 'bg-blue-500' : 'bg-slate-300'}`}></span>
                        <span className={textColor}>{getDisplayChurch(c)}</span>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
          </div>



          <div className="flex-1 overflow-y-auto pr-2 space-y-2 pb-4">
            {(() => {
              const editorCounters = [0, 0, 0, 0, 0, 0];
              return reportData.map((item, index) => {
                if (item.level === 0) {
                  editorCounters[0]++; editorCounters[1]=0; editorCounters[2]=0; editorCounters[3]=0; editorCounters[4]=0; editorCounters[5]=0;
                  return (
                    <div key={item.id} className="flex flex-col gap-2 py-3 mt-4 first:mt-0 group">
                      <div className="font-bold text-lg text-blue-800 border-b-2 border-blue-100 w-full pb-1 flex items-center gap-2">
                        <span className="shrink-0">{toRoman(editorCounters[0])}.</span>
                        <TextareaAutosize
                          id={`input-${item.id}`}
                          value={item.text}
                          onChange={(e) => updateText(item.id, e.target.value)}
                          onKeyDown={(e) => handleKeyDown(e, item.id, index)}
                          onPaste={(e) => handlePaste(e, item.id)}
                          className="bg-transparent border-none outline-none focus:ring-0 flex-1 font-bold text-lg text-blue-800 p-0 m-0 w-full placeholder-blue-300 resize-none"
                          placeholder="대항목 제목 입력 (붙여넣기로 표 생성 가능)"
                        />
                        <div className="flex items-center shrink-0 gap-1 opacity-70 hover:opacity-100 transition-opacity bg-white/50 px-1 py-0.5 rounded">
                          <button onClick={() => moveL0Block(index, 'up')} className="p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-600 rounded" title="위로 이동">
                            <ArrowUp className="w-4 h-4" />
                          </button>
                          <button onClick={() => moveL0Block(index, 'down')} className="p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-600 rounded" title="아래로 이동">
                            <ArrowDown className="w-4 h-4" />
                          </button>
                          <button onClick={() => addEmptyTable(item.id)} className="p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-600 rounded" title="표 삽입">
                            <TableIcon className="w-4 h-4" />
                          </button>
                          <label className="p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-600 rounded cursor-pointer" title="사진 첨부">
                            <input type="file" accept="image/*" className="hidden" onChange={(e) => handleImageUpload(e, item.id)} />
                            <ImageIcon className="w-4 h-4" />
                          </label>
                          <button onClick={() => addNewItem(index, 1)} className="p-1 text-blue-500 hover:bg-blue-50 hover:text-blue-600 font-medium text-xs rounded flex items-center gap-1" title="세부 항목 추가">
                              <Plus className="w-3.5 h-3.5" /> 세부 항목
                          </button>
                          <button onClick={() => removeItem(item.id)} className="p-1 text-red-300 hover:bg-red-50 hover:text-red-500 rounded-md shrink-0 transition-colors" title="삭제">
                            <X className="w-4 h-4" />
                          </button>
                        </div>
                      </div>
                      
                      {item.image && (
                        <div className="flex items-center gap-2 mt-1">
                          <div className="relative inline-block group/img">
                            <img src={item.image} alt="첨부" className="h-24 object-contain rounded border border-slate-200" />
                            <button 
                              onClick={() => removeImage(item.id)}
                              className="absolute top-1 right-1 bg-red-500 text-white rounded-full p-1 opacity-0 group-hover/img:opacity-100 transition-opacity shadow"
                            >
                              <X className="w-3 h-3" />
                            </button>
                          </div>
                        </div>
                      )}
                      
                      {item.tableData && (
                        <div className="mt-2 text-slate-700">
                          <div className="border border-slate-200 rounded-lg bg-white overflow-hidden shadow-sm">
                            <div className="flex justify-between items-center bg-slate-50 px-2 py-1.5 border-b border-slate-200">
                              <div className="flex items-center gap-2">
                                 <span className="text-xs font-bold text-slate-500 flex items-center gap-1"><TableIcon className="w-3.5 h-3.5"/> 데이터 테이블</span>
                                 {item.chartType !== undefined && (
                                    <select value={item.chartType} onChange={e => setChartView(item.id, e.target.value as any)} className="text-xs bg-white border border-slate-300 rounded px-1 min-w-[80px]">
                                      <option value="none">표로 보기</option>
                                      <option value="bar">막대 차트</option>
                                      <option value="line">선 차트</option>
                                      <option value="pie">원형 차트</option>
                                    </select>
                                 )}
                              </div>
                              <div className="flex gap-1">
                                  <button onClick={() => addTableRow(item.id)} className="px-2 py-1 bg-white border border-slate-200 hover:bg-slate-50 rounded text-[10px] font-medium text-slate-600">행+</button>
                                  <button onClick={() => addTableCol(item.id)} className="px-2 py-1 bg-white border border-slate-200 hover:bg-slate-50 rounded text-[10px] font-medium text-slate-600">열+</button>
                                  <button onClick={() => removeTable(item.id)} className="px-1.5 py-1 bg-red-50 text-red-500 hover:bg-red-100 rounded text-[10px]" title="표 삭제"><Trash2 className="w-3.5 h-3.5"/></button>
                              </div>
                            </div>
                            {item.chartType === 'none' || !item.chartType ? (
                              <div className="overflow-x-auto p-2">
                                <table className="w-full border-collapse">
                                    <tbody>
                                        {(() => {
                                      const skippedCells = new Set<string>();
                                      return item.tableData.map((row, rIdx) => (
                                              <tr key={rIdx}>
                                                  {row.map((cell, cIdx) => {
                                                      if (skippedCells.has(`${rIdx},${cIdx}`)) return null;
                                                      
                                                      const spanDef = item.tableSpans?.[rIdx]?.[cIdx];
                                                      const colSpan = (typeof spanDef === 'number' ? spanDef : spanDef?.colspan) || 1;
                                                      const rowSpan = (typeof spanDef === 'number' ? 1 : spanDef?.rowspan) || 1;
                                                      
                                                      for (let r = 0; r < rowSpan; r++) {
                                                        for (let c = 0; c < colSpan; c++) {
                                                          if (r === 0 && c === 0) continue;
                                                          skippedCells.add(`${rIdx + r},${cIdx + c}`);
                                                        }
                                                      }
                                                      return (
                                                        <td 
                                                          key={cIdx} 
                                                          colSpan={colSpan} 
                                                          rowSpan={rowSpan} 
                                                          className={`border border-slate-200 p-0 relative group/td min-w-[80px] ${getCellSelectionStyle(item.id, rIdx, cIdx)}`}
                                                          onMouseDown={(e) => {
                                                            if (e.button !== 0 && e.button !== 2) return;
                                                            if (e.button === 0) {
                                                              setTableSelection({ id: item.id, start: { r: rIdx, c: cIdx }, end: { r: rIdx, c: cIdx }, isDragging: true });
                                                              setTableContextMenu(null);
                                                            }
                                                          }}
                                                          onMouseEnter={() => {
                                                            if (tableSelection?.isDragging && tableSelection.id === item.id) {
                                                              setTableSelection({ ...tableSelection, end: { r: rIdx, c: cIdx } });
                                                            }
                                                          }}
                                                          onContextMenu={(e) => {
                                                            e.preventDefault();
                                                            // If we click outside the current selection, update selection
                                                            const isInsideSelection = tableSelection && tableSelection.id === item.id && 
                                                              rIdx >= Math.min(tableSelection.start.r, tableSelection.end.r) && 
                                                              rIdx <= Math.max(tableSelection.start.r, tableSelection.end.r) && 
                                                              cIdx >= Math.min(tableSelection.start.c, tableSelection.end.c) && 
                                                              cIdx <= Math.max(tableSelection.start.c, tableSelection.end.c);
                                                            
                                                            if (!isInsideSelection) {
                                                              setTableSelection({ id: item.id, start: { r: rIdx, c: cIdx }, end: { r: rIdx, c: cIdx }, isDragging: false });
                                                            }
                                                            setTableContextMenu({ x: e.clientX, y: e.clientY, id: item.id, r: rIdx, c: cIdx });
                                                          }}
                                                        >
                                                            {rIdx === 0 && cIdx > 0 && (
                                                              <div className="absolute top-0 right-0 -mt-5 flex justify-end w-full whitespace-nowrap z-10">
                                                                <label className="flex items-center gap-1 text-[10px] text-slate-500 bg-white px-1 leading-none cursor-pointer hover:bg-slate-50 border border-transparent hover:border-slate-200 rounded">
                                                                  <input type="checkbox" checked={item.chartColumnSelection ? item.chartColumnSelection[cIdx] : true} onChange={() => toggleChartColumnSelection(item.id, cIdx)} className="w-2.5 h-2.5" />
                                                                  차트에 포함
                                                                </label>
                                                              </div>
                                                            )}
                                                            <input 
                                                                type="text" 
                                                                value={cell} 
                                                                onChange={(e) => updateTableCell(item.id, rIdx, cIdx, e.target.value)}
                                                                className={`w-full text-xs outline-none px-2 py-1.5 focus:bg-blue-50 bg-transparent transition-colors ${tableSelection?.isDragging ? 'pointer-events-none' : ''} ${colSpan > 1 ? 'text-center' : ''}`}
                                                            />
                                                            {rIdx === 0 && row.length > 1 && (
                                                              <button onClick={() => removeTableCol(item.id, cIdx)} className="absolute -top-2 left-1/2 -translate-x-1/2 bg-red-500 text-white rounded-full p-0.5 opacity-0 group-hover/td:opacity-100 z-10 scale-75 hover:scale-100 transition-all"><X className="w-3 h-3"/></button>
                                                            )}
                                                            {cIdx === 0 && item.tableData!.length > 1 && (
                                                              <button onClick={() => removeTableRow(item.id, rIdx)} className="absolute top-1/2 -left-2 -translate-y-1/2 bg-red-500 text-white rounded-full p-0.5 opacity-0 group-hover/td:opacity-100 z-10 scale-75 hover:scale-100 transition-all"><X className="w-3 h-3"/></button>
                                                            )}
                                                        </td>
                                                      );
                                                  })}
                                              </tr>
                                            ));
                                        })()}
                                    </tbody>
                                </table>
                              </div>
                            ) : (
                              <div className="p-2 h-44 w-full">
                                <ResponsiveContainer width="100%" height="100%">
                                  {(() => {
                                    const headers = item.tableData[0];
                                    const data = item.tableData.slice(1).map(row => {
                                      const obj: any = { name: row[0] };
                                      for(let i=1; i<row.length; i++) {
                                        obj[headers[i] || `col${i}`] = Number(row[i]) || 0;
                                      }
                                      return obj;
                                    });
                                    const selectedColIndices = item.chartColumnSelection 
                                        ? item.chartColumnSelection.map((sel, idx) => sel && idx > 0 ? idx : -1).filter(idx => idx !== -1)
                                        : headers.map((_, idx) => idx > 0 ? idx : -1).filter(idx => idx !== -1);
                                    if (selectedColIndices.length === 0 && headers.length > 1) selectedColIndices.push(1);
                                    const keys = selectedColIndices.map(idx => headers[idx] || `col${idx}`);
                                    const colors = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6'];
                                    
                                    if (item.chartType === 'bar') {
                                      return (
                                        <BarChart data={data} margin={{ top: 5, right: 5, left: -20, bottom: 0 }}>
                                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                                          <XAxis dataKey="name" tick={{fontSize: 10}} stroke="#94a3b8" />
                                          <YAxis tick={{fontSize: 10}} stroke="#94a3b8" />
                                          <Tooltip contentStyle={{ fontSize: 10, borderRadius: 4, padding: 4 }} />
                                          <Legend wrapperStyle={{ fontSize: 10 }} />
                                          {keys.map((k, i) => <Bar key={k} dataKey={k} fill={colors[i % colors.length]} radius={[2,2,0,0]} />)}
                                        </BarChart>
                                      );
                                    } else if (item.chartType === 'line') {
                                      return (
                                        <LineChart data={data} margin={{ top: 5, right: 5, left: -20, bottom: 0 }}>
                                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                                          <XAxis dataKey="name" tick={{fontSize: 10}} stroke="#94a3b8" />
                                          <YAxis tick={{fontSize: 10}} stroke="#94a3b8" />
                                          <Tooltip contentStyle={{ fontSize: 10, borderRadius: 4, padding: 4 }} />
                                          <Legend wrapperStyle={{ fontSize: 10 }} />
                                          {keys.map((k, i) => <Line key={k} type="monotone" dataKey={k} stroke={colors[i % colors.length]} strokeWidth={2} />)}
                                        </LineChart>
                                      );
                                    } else if (item.chartType === 'pie') {
                                      const pieData = data.map(d => ({ name: d.name, value: d[keys[0]] || 0 }));
                                      return (
                                        <PieChart margin={{ top: 0, right: 0, left: 0, bottom: 0 }}>
                                          <Pie data={pieData} cx="50%" cy="50%" outerRadius={60} fill="#8884d8" dataKey="value" label={({name, percent}) => `${name} ${(percent * 100).toFixed(0)}%`} labelLine={false} style={{fontSize: 10}}>
                                            {pieData.map((entry, index) => <Cell key={`cell-${index}`} fill={colors[index % colors.length]} />)}
                                          </Pie>
                                          <Tooltip contentStyle={{ fontSize: 10, borderRadius: 4, padding: 4 }} />
                                        </PieChart>
                                      );
                                    }
                                    return null;
                                  })()}
                                </ResponsiveContainer>
                              </div>
                            )}
                          </div>
                        </div>
                      )}
                    </div>
                  );
                }

                return (() => {
                  if (item.level === 1) { editorCounters[1]++; editorCounters[2]=0; editorCounters[3]=0; editorCounters[4]=0; editorCounters[5]=0; }
                  else if (item.level === 2) { editorCounters[2]++; editorCounters[3]=0; editorCounters[4]=0; editorCounters[5]=0; }
                  else if (item.level === 3) { editorCounters[3]++; editorCounters[4]=0; editorCounters[5]=0; }
                  else if (item.level === 4) { editorCounters[4]++; editorCounters[5]=0; }
                  else if (item.level === 5) { editorCounters[5]++; }
                  const lvlPrefix =
                    item.level === 1 ? `${editorCounters[1]}.` :
                    item.level === 2 ? `${editorCounters[2]})` :
                    item.level === 3 ? toCircled(editorCounters[3]) :
                    item.level === 4 ? `${toKorSyllable(editorCounters[4])}.` :
                    item.level === 5 ? `${toLatin(editorCounters[5])}.` : '';
                  return (
                  <div key={item.id} className="flex flex-col gap-1 group">
                    <div className="flex items-center gap-1">
                    <div style={{ width: `${(item.level - 1) * 20}px` }} className="shrink-0 transition-all duration-200" />
                    <div className={`flex items-center flex-1 rounded-md border transition-colors focus-within:border-blue-400 focus-within:ring-2 focus-within:ring-blue-100 focus-within:bg-white bg-slate-50 border-slate-200`}>
                      <span className={`pl-2.5 pr-1 py-2.5 text-sm font-bold shrink-0 select-none whitespace-nowrap ${item.level === 1 ? 'text-blue-700' : item.level === 2 ? 'text-slate-700' : 'text-slate-500'}`}>
                        {lvlPrefix}
                      </span>
                      <TextareaAutosize
                        id={`input-${item.id}`}
                        value={item.text}
                        onChange={(e) => updateText(item.id, e.target.value)}
                        onKeyDown={(e) => handleKeyDown(e, item.id, index)}
                        onPaste={(e) => handlePaste(e, item.id)}
                        placeholder="항목 내용을 입력하세요"
                        minRows={1}
                        className="flex-1 pr-3 py-2.5 text-sm bg-transparent border-none outline-none focus:ring-0 resize-none leading-relaxed"
                      />
                    </div>
                  </div>

                  <div
                    className="flex items-center flex-wrap gap-2 transition-all duration-200 opacity-0 max-h-0 overflow-hidden group-focus-within:opacity-100 group-focus-within:max-h-24 group-focus-within:mt-1 group-hover:opacity-100 group-hover:max-h-24 group-hover:mt-1"
                    style={{ paddingLeft: `${(item.level - 1) * 20 + 4}px` }}
                  >
                    <button onClick={() => changeLevel(item.id, -1)} className="flex items-center gap-1 px-2 py-1 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded text-xs font-medium transition-colors" title="상위 수준 (Shift+Tab)">
                      <ArrowLeft className="w-3.5 h-3.5" /> 내어쓰기
                    </button>
                    <button onClick={() => changeLevel(item.id, 1)} className="flex items-center gap-1 px-2 py-1 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded text-xs font-medium transition-colors" title="하위 수준 (Tab)">
                      <ArrowRight className="w-3.5 h-3.5" /> 들여쓰기
                    </button>
                    <button onClick={() => addNewItem(index, item.level)} className="flex items-center gap-1 px-2 py-1 bg-blue-50 border border-blue-100 hover:bg-blue-100 text-blue-700 rounded text-xs font-medium transition-colors" title="항목 추가 (Enter)">
                      <Plus className="w-3.5 h-3.5" /> 항목 추가
                    </button>
                    <button onClick={() => addEmptyTable(item.id)} className="flex items-center gap-1 px-2 py-1 bg-indigo-50 border border-indigo-100 hover:bg-indigo-100 text-indigo-700 rounded text-xs font-medium cursor-pointer transition-colors" title="표 삽입">
                      <TableIcon className="w-3.5 h-3.5" /> 표 삽입
                    </button>
                    <label className="flex items-center gap-1 px-2 py-1 bg-emerald-50 border border-emerald-100 hover:bg-emerald-100 text-emerald-700 rounded text-xs font-medium cursor-pointer transition-colors" title="사진 첨부">
                      <input type="file" accept="image/*" className="hidden" onChange={(e) => handleImageUpload(e, item.id)} />
                      <ImageIcon className="w-3.5 h-3.5" /> 사진 첨부
                    </label>
                    <button onClick={() => removeItem(item.id)} className="flex items-center gap-1 px-2 py-1 bg-red-50 hover:bg-red-100 text-red-600 rounded text-xs font-medium transition-colors" title="삭제">
                      <X className="w-3.5 h-3.5" /> 삭제
                    </button>
                  </div>
                  {item.image && (
                    <div className="flex items-center gap-2 mt-1">
                      <div style={{ width: `${(item.level - 1) * 20 + 4}px` }} className="shrink-0" />
                      <div className="relative inline-block group/img">
                        <img src={item.image} alt="첨부" className="h-24 object-contain rounded border border-slate-200" />
                        <button 
                          onClick={() => removeImage(item.id)}
                          className="absolute top-1 right-1 bg-red-500 text-white rounded-full p-1 opacity-0 group-hover/img:opacity-100 transition-opacity shadow"
                        >
                          <X className="w-3 h-3" />
                        </button>
                      </div>
                    </div>
                  )}
                  {item.tableData && (
                    <div className="mt-2 text-slate-700" style={{ paddingLeft: `${(item.level - 1) * 20 + 4}px` }}>
                      <div className="border border-slate-200 rounded-lg bg-white overflow-hidden shadow-sm">
                        <div className="flex justify-between items-center bg-slate-50 px-2 py-1.5 border-b border-slate-200">
                          <div className="flex items-center gap-2">
                             <span className="text-xs font-bold text-slate-500 flex items-center gap-1"><TableIcon className="w-3.5 h-3.5"/> 데이터 테이블</span>
                             {item.chartType !== undefined && (
                                <select value={item.chartType} onChange={e => setChartView(item.id, e.target.value as any)} className="text-xs bg-white border border-slate-300 rounded px-1 min-w-[80px]">
                                  <option value="none">표로 보기</option>
                                  <option value="bar">막대 차트</option>
                                  <option value="line">선 차트</option>
                                  <option value="pie">원형 차트</option>
                                </select>
                             )}
                          </div>
                          <div className="flex gap-1">
                              <button onClick={() => addTableRow(item.id)} className="px-2 py-1 bg-white border border-slate-200 hover:bg-slate-50 rounded text-[10px] font-medium text-slate-600">행+</button>
                              <button onClick={() => addTableCol(item.id)} className="px-2 py-1 bg-white border border-slate-200 hover:bg-slate-50 rounded text-[10px] font-medium text-slate-600">열+</button>
                              <button onClick={() => removeTable(item.id)} className="px-1.5 py-1 bg-red-50 text-red-500 hover:bg-red-100 rounded text-[10px]" title="표 삭제"><Trash2 className="w-3.5 h-3.5"/></button>
                          </div>
                        </div>
                        {item.chartType === 'none' || !item.chartType ? (
                          <div className="overflow-x-auto p-4 pl-8">
                            <table className="w-full border-collapse">
                                <tbody>
                                    {(() => {
                                      const skippedCells = new Set<string>();
                                      return item.tableData.map((row, rIdx) => (
                                        <tr key={rIdx} className="group/tr relative">
                                            {row.map((cell, cIdx) => {
                                                if (skippedCells.has(`${rIdx},${cIdx}`)) return null;
                                                
                                                const spanDef = item.tableSpans?.[rIdx]?.[cIdx];
                                                const colSpan = (typeof spanDef === 'number' ? spanDef : spanDef?.colspan) || 1;
                                                const rowSpan = (typeof spanDef === 'number' ? 1 : spanDef?.rowspan) || 1;
                                                
                                                for (let r = 0; r < rowSpan; r++) {
                                                  for (let c = 0; c < colSpan; c++) {
                                                    if (r === 0 && c === 0) continue;
                                                    skippedCells.add(`${rIdx + r},${cIdx + c}`);
                                                  }
                                                }
                                                const isHighlighted = item.tableHighlights?.[rIdx]?.[cIdx];
                                                const align = item.tableAlignments?.[rIdx]?.[cIdx] || (colSpan > 1 ? 'center' : 'left');
                                                const alignClass = align === 'center' ? 'text-center' : align === 'right' ? 'text-right' : 'text-left';
                                                return (
                                                <td 
                                                  key={cIdx} 
                                                  colSpan={colSpan} 
                                                  rowSpan={rowSpan} 
                                                  className={`border ${isHighlighted ? 'border-blue-300 bg-blue-50/60' : 'border-slate-300 bg-white hover:bg-slate-50 focus-within:bg-blue-50/30'} p-0 relative group/td min-w-[120px] align-top ${getCellSelectionStyle(item.id, rIdx, cIdx)}`}
                                                  onMouseDown={(e) => {
                                                    if (e.button !== 0 && e.button !== 2) return;
                                                    if (e.button === 0) {
                                                      setTableSelection({ id: item.id, start: { r: rIdx, c: cIdx }, end: { r: rIdx, c: cIdx }, isDragging: true });
                                                      setTableContextMenu(null);
                                                    }
                                                  }}
                                                  onMouseEnter={() => {
                                                    if (tableSelection?.isDragging && tableSelection.id === item.id) {
                                                      setTableSelection({ ...tableSelection, end: { r: rIdx, c: cIdx } });
                                                    }
                                                  }}
                                                  onContextMenu={(e) => {
                                                    e.preventDefault();
                                                    const isInsideSelection = tableSelection && tableSelection.id === item.id && 
                                                      rIdx >= Math.min(tableSelection.start.r, tableSelection.end.r) && 
                                                      rIdx <= Math.max(tableSelection.start.r, tableSelection.end.r) && 
                                                      cIdx >= Math.min(tableSelection.start.c, tableSelection.end.c) && 
                                                      cIdx <= Math.max(tableSelection.start.c, tableSelection.end.c);
                                                    
                                                    if (!isInsideSelection) {
                                                      setTableSelection({ id: item.id, start: { r: rIdx, c: cIdx }, end: { r: rIdx, c: cIdx }, isDragging: false });
                                                    }
                                                    setTableContextMenu({ x: e.clientX, y: e.clientY, id: item.id, r: rIdx, c: cIdx });
                                                  }}
                                                >
                                                    {rIdx === 0 && cIdx > 0 && (
                                                      <div className="absolute top-0 right-0 -mt-5 flex justify-end w-full whitespace-nowrap z-10">
                                                        <label className="flex items-center gap-1 text-[10px] text-slate-500 bg-white px-1 leading-none cursor-pointer hover:bg-slate-50 border border-transparent hover:border-slate-200 rounded">
                                                          <input type="checkbox" checked={item.chartColumnSelection ? item.chartColumnSelection[cIdx] : true} onChange={() => toggleChartColumnSelection(item.id, cIdx)} className="w-2.5 h-2.5" />
                                                          차트에 포함
                                                        </label>
                                                      </div>
                                                    )}
                                                    <TextareaAutosize 
                                                        value={cell} 
                                                        onChange={(e) => updateTableCell(item.id, rIdx, cIdx, e.target.value)}
                                                        className={`w-full text-xs outline-none px-3 py-2 bg-transparent transition-colors resize-none leading-relaxed ${tableSelection?.isDragging ? 'pointer-events-none select-none' : ''} ${isHighlighted ? 'font-bold text-slate-800' : 'text-slate-700'} ${alignClass}`}
                                                    />
                                                    <div className="absolute top-1 right-1 flex gap-1 opacity-0 group-hover/td:opacity-100 transition-opacity z-10">
                                                        <button onClick={() => toggleCellHighlight(item.id, rIdx, cIdx)} className={`p-1 rounded shadow-sm border ${isHighlighted ? 'bg-blue-500 text-white border-blue-600' : 'bg-white text-slate-500 border-slate-200 hover:bg-slate-100'}`} title="셀 강조 (볼드/배경색)">
                                                            <Highlighter className="w-3 h-3"/>
                                                        </button>
                                                    </div>
                                                    {rIdx === 0 && row.length > 1 && (
                                                      <button onClick={() => removeTableCol(item.id, cIdx)} className="absolute -top-3 left-1/2 -translate-x-1/2 bg-red-500 text-white rounded p-0.5 opacity-0 group-hover/td:opacity-100 z-10 scale-75 hover:scale-100 transition-all shadow-sm" title="열 삭제"><X className="w-3 h-3"/></button>
                                                    )}
                                                    {cIdx === 0 && (
                                                      <button onClick={() => toggleRowHighlight(item.id, rIdx)} className="absolute top-1/2 -left-4 -translate-y-1/2 bg-blue-500 text-white shadow-sm border border-white rounded p-0.5 opacity-0 group-hover/tr:opacity-100 z-10 scale-75 hover:scale-100 transition-all" title="행 전체 강조"><Highlighter className="w-3 h-3"/></button>
                                                    )}
                                                    {cIdx === 0 && item.tableData!.length > 1 && (
                                                      <button onClick={() => removeTableRow(item.id, rIdx)} className="absolute top-1/2 -left-10 -translate-y-1/2 bg-red-500 text-white rounded p-0.5 opacity-0 group-hover/tr:opacity-100 z-10 scale-75 hover:scale-100 transition-all shadow-sm" title="행 삭제"><X className="w-3 h-3"/></button>
                                                    )}
                                                </td>
                                            )})}
                                              </tr>
                                            ));
                                        })()}
                                    </tbody>
                            </table>
                          </div>
                        ) : (
                          <div className="p-2 h-44 w-full">
                            <ResponsiveContainer width="100%" height="100%">
                              {(() => {
                                const headers = item.tableData[0];
                                const data = item.tableData.slice(1).map(row => {
                                  const obj: any = { name: row[0] };
                                  for(let i=1; i<row.length; i++) {
                                    obj[headers[i] || `col${i}`] = Number(row[i]) || 0;
                                  }
                                  return obj;
                                });
                                const selectedColIndices = item.chartColumnSelection 
                                    ? item.chartColumnSelection.map((sel, idx) => sel && idx > 0 ? idx : -1).filter(idx => idx !== -1)
                                    : headers.map((_, idx) => idx > 0 ? idx : -1).filter(idx => idx !== -1);
                                if (selectedColIndices.length === 0 && headers.length > 1) selectedColIndices.push(1);
                                const keys = selectedColIndices.map(idx => headers[idx] || `col${idx}`);
                                const colors = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6'];
                                
                                if (item.chartType === 'bar') {
                                  return (
                                    <BarChart data={data} margin={{ top: 5, right: 5, left: -20, bottom: 0 }}>
                                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                                      <XAxis dataKey="name" tick={{fontSize: 10}} stroke="#94a3b8" />
                                      <YAxis tick={{fontSize: 10}} stroke="#94a3b8" />
                                      <Tooltip contentStyle={{ fontSize: 10, borderRadius: 4, padding: 4 }} />
                                      <Legend wrapperStyle={{ fontSize: 10 }} />
                                      {keys.map((k, i) => <Bar key={k} dataKey={k} fill={colors[i % colors.length]} radius={[2,2,0,0]} />)}
                                    </BarChart>
                                  );
                                } else if (item.chartType === 'line') {
                                  return (
                                    <LineChart data={data} margin={{ top: 5, right: 5, left: -20, bottom: 0 }}>
                                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                                      <XAxis dataKey="name" tick={{fontSize: 10}} stroke="#94a3b8" />
                                      <YAxis tick={{fontSize: 10}} stroke="#94a3b8" />
                                      <Tooltip contentStyle={{ fontSize: 10, borderRadius: 4, padding: 4 }} />
                                      <Legend wrapperStyle={{ fontSize: 10 }} />
                                      {keys.map((k, i) => <Line key={k} type="monotone" dataKey={k} stroke={colors[i % colors.length]} strokeWidth={2} />)}
                                    </LineChart>
                                  );
                                } else if (item.chartType === 'pie') {
                                  const pieData = data.map(d => ({ name: d.name, value: d[keys[0]] || 0 }));
                                  return (
                                    <PieChart margin={{ top: 0, right: 0, left: 0, bottom: 0 }}>
                                      <Pie data={pieData} cx="50%" cy="50%" outerRadius={60} fill="#8884d8" dataKey="value" label={({name, percent}) => `${name} ${(percent * 100).toFixed(0)}%`} labelLine={false} style={{fontSize: 10}}>
                                        {pieData.map((entry, index) => <Cell key={`cell-${index}`} fill={colors[index % colors.length]} />)}
                                      </Pie>
                                      <Tooltip contentStyle={{ fontSize: 10, borderRadius: 4, padding: 4 }} />
                                    </PieChart>
                                  );
                                }
                                return null;
                              })()}
                            </ResponsiveContainer>
                          </div>
                        )}
                      </div>
                    </div>
                  )}
                </div>
              );
            })()})})()}
            
            <div className="flex gap-2 mt-4 pt-4 border-t border-slate-100">
              <button
                onClick={() => addNewItem(-1, 1)}
                className="flex-1 flex items-center justify-center gap-2 border-2 border-dashed border-slate-200 hover:border-blue-400 hover:bg-blue-50 text-slate-500 hover:text-blue-600 p-2.5 rounded-lg text-sm font-medium transition-all"
              >
                <Plus className="w-4 h-4" /> 세부 항목(L1) 추가
              </button>
              <button
                onClick={() => {
                  const newItem = { id: nextId, text: "", level: 0, isFixed: false };
                  const newChild = { id: nextId + 1, text: "", level: 1 };
                  setNextId(prev => prev + 2);
                  setReportData(data => [...data, newItem, newChild]);
                  setTimeout(() => { document.getElementById(`input-${newItem.id}`)?.focus(); }, 10);
                }}
                className="flex-1 flex items-center justify-center gap-2 bg-slate-100 hover:bg-slate-200 border border-slate-200 text-slate-700 p-2.5 rounded-lg text-sm font-medium transition-colors"
              >
                <Plus className="w-4 h-4" /> 대항목(L0) 추가
              </button>
              <button
                onClick={() => docUploadRef.current?.click()}
                className="flex items-center justify-center gap-1.5 bg-blue-50 hover:bg-blue-100 border border-blue-200 text-blue-700 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors"
                title="워드(.docx) 또는 한글(.hwpx) 파일 내용 가져오기"
              >
                <Upload className="w-4 h-4" /> 문서
              </button>
            </div>
          </div>
          
          <div className="pt-4 border-t border-slate-200 mt-auto shrink-0 space-y-2">
            <div className="flex items-center justify-center gap-3 text-[10px] font-semibold text-slate-400 select-none pb-2 border-b border-slate-100 mb-1">
              <span><kbd className="font-mono bg-slate-50 border border-slate-200 px-1 py-0.2 rounded text-slate-500 mr-0.5 shadow-sm">Tab</kbd>들여쓰기</span>
              <span className="text-slate-200">|</span>
              <span><kbd className="font-mono bg-slate-50 border border-slate-200 px-1 py-0.2 rounded text-slate-500 mr-0.5 shadow-sm">Shift+Tab</kbd>내어쓰기</span>
              <span className="text-slate-200">|</span>
              <span><kbd className="font-mono bg-slate-50 border border-slate-200 px-1 py-0.2 rounded text-slate-500 mr-0.5 shadow-sm">Enter</kbd>항목 추가</span>
              <span className="text-slate-200">|</span>
              <span><kbd className="font-mono bg-slate-50 border border-slate-200 px-1 py-0.2 rounded text-slate-500 mr-0.5 shadow-sm">↑/↓</kbd>이동</span>
            </div>
            <div className="flex items-center justify-between px-1 pb-1">
              <div className="flex flex-col gap-0.5">
                <span className="text-xs text-slate-500 font-medium flex items-center">
                  {isSaving ? <RefreshCw className="w-3 h-3 animate-spin mr-1.5" /> : <Save className="w-3 h-3 mr-1.5 text-blue-500" />}
                  {isSaving ? '저장 중...' : lastSaved ? `${lastSaved} 자동 저장됨` : '자동 저장 대기 중'}
                </span>
                {driveSaveResult && (
                  <span className={`text-[11px] font-semibold flex items-center gap-1 ${driveSaveResult === 'ok' ? 'text-emerald-600' : driveSaveResult === 'saving' ? 'text-blue-500' : driveSaveResult === 'skipped' ? 'text-slate-400' : 'text-red-500'}`}>
                    {driveSaveResult === 'ok' && `☁️ 구글 드라이브 저장 완료${driveSavedAt ? ` · ${driveSavedAt}` : ''}`}
                    {driveSaveResult === 'saving' && '☁️ 드라이브 저장 중...'}
                    {driveSaveResult === 'skipped' && '⚠️ 드라이브 미연결 — /api/google-auth 에서 인증 필요'}
                    {driveSaveResult === 'error' && '⚠️ 드라이브 저장 실패 — 관리자 문의'}
                  </span>
                )}
              </div>
              <button 
                onClick={handleReset}
                className="text-xs text-red-500 hover:text-red-700 font-medium flex items-center gap-1 transition-colors px-2 py-1 rounded hover:bg-red-50"
              >
                <RefreshCw className="w-3 h-3" /> 초기화 (새 보고서 시작)
              </button>
            </div>
            <div className="flex gap-2">
              {activeTab === 'notice_write' ? (
                <button 
                  onClick={handlePublishNotice}
                  className="w-full flex items-center justify-center gap-2 p-3 rounded-lg font-bold shadow-sm transition-colors bg-emerald-600 hover:bg-emerald-700 text-white"
                >
                  <Check className="w-5 h-5" />
                  공지사항 작성 및 등록 완료
                </button>
              ) : (
                <>
                  <button 
                    onClick={() => handleSave(true)}
                    disabled={isSaving || status === 'submitted'}
                    className={`flex-1 flex items-center justify-center gap-2 p-3 rounded-lg font-bold transition-colors disabled:opacity-70 ${status === 'submitted' ? 'bg-emerald-100 text-emerald-800 border border-emerald-300' : 'bg-slate-100 hover:bg-slate-200 text-slate-700'}`}
                  >
                    <Check className="w-4 h-4" />
                    {status === 'submitted' ? '제출 완료' : '제출 확정'}
                  </button>
                  {status === 'submitted' && (
                    <button 
                      onClick={() => { setStatus('draft'); handleSave(false); }}
                      className="flex-1 flex items-center justify-center gap-2 p-3 rounded-lg font-bold transition-colors bg-amber-100 hover:bg-amber-200 text-amber-800 border border-amber-300"
                    >
                      <X className="w-4 h-4" /> 제출 취소
                    </button>
                  )}
                </>
              )}
            </div>
            <button
              onClick={() => {
                if (isDownloadUnlocked) { checkWithAI(); return; }
                const pwd = prompt('비밀번호를 입력하세요:');
                if (pwd === 'skmt0909!') {
                  setIsDownloadUnlocked(true);
                  sessionStorage.setItem('download_unlocked', '1');
                  checkWithAI();
                } else if (pwd === 'samu') {
                  exportToWord();
                } else if (pwd !== null) {
                  alert('비밀번호가 일치하지 않습니다.');
                }
              }}
              className="w-full flex items-center justify-center gap-2 bg-blue-600 hover:bg-blue-700 text-white p-3 rounded-lg font-medium shadow-sm transition-colors"
            >
              <Bot className="w-5 h-5" />
              AI 문맥 검토 및 내보내기
            </button>
          </div>
        </div>

        {/* Preview Panel */}
        <div className={`bg-[#fafbfc] p-5 sm:p-8 rounded-xl shadow-sm border border-slate-200 flex-col h-[calc(100vh-8rem)] xl:h-[calc(100vh-3rem)] overflow-y-auto ${mobileView === 'preview' ? 'flex' : 'hidden xl:flex'}`}>
          <div className="mb-2 flex items-center justify-end">
            <span className="text-[11px] text-slate-400 font-medium bg-slate-100 px-2 py-1 rounded-md flex items-center gap-1">
              <span>✏️</span> 좌측 에디터 또는 이 미리보기 창에서 직접 편집 가능
            </span>
          </div>
          <div className="mb-1 font-serif">
            <div className="border-t-2 border-b-[3px] border-[#4eaee7] py-3 mb-2">
              <div className="text-3xl font-black text-center text-slate-800 tracking-tight">
                {activeTab === 'notice_write' ? '공지사항 미리보기' : `<${getDisplayParish(parish)}> 주간업무보고`}
              </div>
            </div>
            <div className="text-xl font-black text-blue-700 mb-1 drop-shadow-sm">
              {activeTab === 'notice_write' ? (noticeTitle || '제목을 입력해주세요') : `${(PARISH_CHURCH_MAP[parish] || []).indexOf(church) + 1}. ${getDisplayChurch(church)}`}
            </div>
          </div>
          <div
            className="flex-1 font-serif text-slate-900"
            onPaste={e => {
              const imageItem = Array.from(e.clipboardData.items).find(i => i.type.startsWith('image/'));
              if (imageItem) {
                e.preventDefault();
                const file = imageItem.getAsFile();
                if (!file) return;
                const reader = new FileReader();
                reader.onload = ev => {
                  const dataUrl = ev.target?.result as string;
                  const img = new Image();
                  img.onload = () => {
                    setReportData(prev => {
                      const newId = Math.max(0, ...prev.map(d => d.id)) + 1;
                      setNextId(newId + 1);
                      return [...prev, { id: newId, text: '', level: 1, image: dataUrl, imageWidth: img.naturalWidth || 400, imageHeight: img.naturalHeight || 300 }];
                    });
                  };
                  img.src = dataUrl;
                };
                reader.readAsDataURL(file);
                return;
              }
              const html = e.clipboardData.getData('text/html');
              if (html) {
                const parsed = parseHtmlTable(html);
                if (parsed?.tableData?.length) {
                  e.preventDefault();
                  setReportData(prev => {
                    const newId = Math.max(0, ...prev.map(d => d.id)) + 1;
                    setNextId(newId + 1);
                    return [...prev, { id: newId, text: '', level: 1, tableData: parsed.tableData, tableSpans: parsed.tableSpans, tableHighlights: parsed.tableData.map((r: any[]) => r.map(() => false)), chartType: 'none' as const }];
                  });
                }
              }
            }}
          >
            {renderPreviewLines()}
          </div>
        </div>
      </div>
      </div>
      )}

      {/* Admin Consolidated Console Tab View */}
      {activeTab === 'admin_console' && (
        <div className="w-full max-w-full px-1 sm:px-4 lg:px-8 mx-auto flex flex-col flex-1 min-h-0 overflow-hidden">
          <div className="grid grid-cols-1 xl:grid-cols-[40%_60%] gap-4 lg:gap-6 flex-1 min-h-0">
            {/* Left Panel: Submission Status Grid */}
            <div className="bg-white p-4 sm:p-6 rounded-xl shadow-sm border border-slate-200 flex flex-col h-[calc(100vh-12rem)] xl:h-[calc(100vh-8rem)]">
              <div className="flex items-center justify-between pb-4 border-b border-slate-200 mb-4">
                <div>
                  <h2 className="text-lg font-black text-slate-800 flex items-center gap-2">
                    <Settings className="w-5 h-5 text-purple-600 animate-spin-slow" />
                    전체 교구 제출현황 및 제어
                  </h2>
                  <p className="text-xs text-slate-500 mt-0.5">전국 교구와 협회 부서의 작성 현황을 확인하고 제어합니다.</p>
                </div>
                <button 
                  onClick={async () => {
                    const pwd = prompt("전체 교구의 데이터를 초기화하려면 비밀번호를 입력해 주세요:");
                    if (pwd === "skmt0909!") {
                      if (window.confirm("정말로 모든 교구와 협회의 주간보고 데이터를 초기화하시겠습니까? (이 작업은 되돌릴 수 없습니다)")) {
                        const defaultData = { data: DEFAULT_REPORT, status: 'draft', lastSaved: null };
                        for (const p of Object.keys(PARISH_CHURCH_MAP)) {
                          for (const c of PARISH_CHURCH_MAP[p]) {
                            const key = `report_${p}_${c}`;
                            localStorage.setItem(key, JSON.stringify(defaultData));
                            if (supabase) {
                              try {
                                await supabase.from('reports').upsert({ id: `${p}_${c}`, ...defaultData, updated_at: new Date().toISOString() });
                              } catch(e){}
                            }
                          }
                        }
                        loadAllReportsStatus();
                        alert("전 교구 및 협회 데이터가 전체 초기화되었습니다.");
                      }
                    } else if (pwd !== null) {
                      alert("비밀번호가 일치하지 않습니다.");
                    }
                  }}
                  className="px-2.5 py-1.5 bg-red-50 hover:bg-red-100 border border-red-200 text-red-600 rounded-lg text-xs font-bold transition-colors"
                >
                  전체 데이터 초기화
                </button>
              </div>

              {/* Parish Selector Tabs inside left panel */}
              <div className="flex gap-1.5 overflow-x-auto pb-2 scrollbar-hide snap-x whitespace-nowrap mb-4">
                <button 
                  onClick={() => setAdminActiveParish('전체')} 
                  className={`px-3 py-1.5 rounded-full text-xs font-bold transition-all ${adminActiveParish === '전체' ? 'bg-purple-600 text-white shadow-sm' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'}`}
                >
                  전체보기
                </button>
                {Object.keys(PARISH_CHURCH_MAP).map(p => (
                  <button 
                    key={p}
                    onClick={() => setAdminActiveParish(p)} 
                    className={`px-3 py-1.5 rounded-full text-xs font-bold transition-all ${adminActiveParish === p ? 'bg-purple-600 text-white shadow-sm' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'}`}
                  >
                    {getDisplayParish(p)}
                  </button>
                ))}
              </div>

              {/* Church submission list */}
              <div className="flex-1 overflow-y-auto pr-1 space-y-4">
                {Object.keys(PARISH_CHURCH_MAP)
                  .filter(p => adminActiveParish === '전체' || adminActiveParish === p)
                  .map(p => {
                    const churches = PARISH_CHURCH_MAP[p];
                    const submittedCount = churches.filter(c => adminReportStatusMap[`${p}_${c}`] === 'submitted').length;
                    const draftCount = churches.filter(c => adminReportStatusMap[`${p}_${c}`] === 'draft').length;
                    
                    return (
                      <div key={p} className="bg-slate-50 border border-slate-200 rounded-xl p-4 transition-all hover:shadow-sm">
                        <div className="flex items-center justify-between border-b border-slate-200 pb-2 mb-3">
                          <h3 className="font-extrabold text-sm text-slate-800 flex items-center gap-1.5">
                            <span className="w-2 h-2 rounded-full bg-purple-500"></span>
                            {getDisplayParish(p)}
                          </h3>
                          <div className="flex gap-2 text-xs font-semibold text-slate-500">
                            <span className="text-emerald-600">제출: {submittedCount}</span>
                            <span className="text-amber-600">작성중: {draftCount}</span>
                            <span className="text-slate-400 font-medium">미작성: {churches.length - submittedCount - draftCount}</span>
                          </div>
                        </div>

                        <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                          {churches.map(c => {
                            const currentStatus = adminReportStatusMap[`${p}_${c}`] || 'empty';
                            const statusConfig = {
                              submitted: { bg: 'bg-emerald-50 text-emerald-700 border-emerald-200', label: '제출 완료', icon: '✅' },
                              draft: { bg: 'bg-amber-50 text-amber-700 border-amber-200', label: '작성 중', icon: '📝' },
                              empty: { bg: 'bg-slate-100 text-slate-400 border-slate-200', label: '미작성', icon: '⚪' }
                            }[currentStatus];

                            return (
                              <div 
                                key={c} 
                                onClick={() => {
                                  if (p === '협회') {
                                    setActiveTab('association');
                                    setParish('협회');
                                    setChurch(c);
                                  } else {
                                    setActiveTab('report');
                                    setParish(p);
                                    setChurch(c);
                                  }
                                }}
                                className={`px-2.5 py-2 border rounded-lg flex flex-col justify-between h-14 cursor-pointer hover:shadow-sm hover:scale-[1.02] active:scale-[0.98] transition-all ${statusConfig.bg}`}
                              >
                                <span className="text-xs font-bold truncate">{getDisplayChurch(c)}</span>
                                <span className="text-[10px] font-black flex items-center gap-1 mt-0.5">
                                  <span>{statusConfig.icon}</span>
                                  {statusConfig.label}
                                </span>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    );
                  })}
              </div>
            </div>

            {/* Right Panel: Batch AI Review & Word compilation */}
            <div className="bg-white p-4 sm:p-6 rounded-xl shadow-sm border border-slate-200 flex flex-col h-[calc(100vh-12rem)] xl:h-[calc(100vh-8rem)]">
              <div className="pb-4 border-b border-slate-200 mb-6 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
                <div>
                  <h2 className="text-lg font-black text-slate-800 flex items-center gap-2">
                    <Sparkles className="w-5 h-5 text-indigo-600 animate-pulse" />
                    AI 종합 취합 및 통합 검토 콘솔
                  </h2>
                  <p className="text-xs text-slate-500 mt-0.5">전국 보고서를 실시간으로 취합하여 일괄 AI 맞춤법 교정 및 마스터 워드 파일로 다운로드합니다.</p>
                </div>
              </div>

              {/* ── Google Drive · AI Studio Build 연동 ── */}
              <div className="mb-5 space-y-3">

                {/* 상태 배너 */}
                <div className={`flex items-center justify-between gap-3 rounded-xl px-4 py-3 border ${
                  driveStatus?.authenticated
                    ? 'bg-emerald-50 border-emerald-200'
                    : driveStatus?.configured
                    ? 'bg-amber-50 border-amber-200'
                    : 'bg-red-50 border-red-200'
                }`}>
                  <div className="flex items-center gap-2.5 min-w-0">
                    {driveStatusLoading ? (
                      <div className="w-3.5 h-3.5 rounded-full border-2 border-slate-300 border-t-slate-600 animate-spin shrink-0" />
                    ) : driveStatus?.authenticated ? (
                      <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse shrink-0" />
                    ) : driveStatus?.configured ? (
                      <span className="w-2.5 h-2.5 rounded-full bg-amber-400 shrink-0" />
                    ) : (
                      <span className="w-2.5 h-2.5 rounded-full bg-red-400 shrink-0" />
                    )}
                    <div className="min-w-0">
                      <p className={`text-xs font-black ${driveStatus?.authenticated ? 'text-emerald-700' : driveStatus?.configured ? 'text-amber-700' : 'text-red-700'}`}>
                        Google Drive:{' '}
                        {driveStatusLoading
                          ? '상태 확인 중...'
                          : driveStatus?.authenticated
                          ? '✅ 연결됨 — 보고서 제출 시 자동 업로드'
                          : driveStatus?.configured
                          ? '⚠️ Refresh Token 미등록 — 아래 3단계 완료 필요'
                          : '❌ 미연결 — 아래 설정 가이드를 따라주세요'}
                      </p>
                      {driveStatus?.appUrl && (
                        <p className="text-[10px] text-slate-500 mt-0.5 truncate">앱 URL: {driveStatus.appUrl}</p>
                      )}
                    </div>
                  </div>
                  <div className="flex items-center gap-2 shrink-0">
                    <button onClick={checkDriveStatus} className="p-1.5 rounded-lg hover:bg-white/60 text-slate-400 transition-colors" title="새로고침">
                      <RefreshCw className="w-3.5 h-3.5" />
                    </button>
                    <button
                      onClick={() => setShowAiStudioGuide(v => !v)}
                      className="px-3 py-1.5 bg-violet-600 hover:bg-violet-700 text-white text-[11px] font-bold rounded-lg transition-colors flex items-center gap-1.5"
                    >
                      <Settings className="w-3 h-3" />
                      {showAiStudioGuide ? '가이드 닫기' : 'Drive 연동 설정 가이드'}
                    </button>
                  </div>
                </div>

                {/* AI Studio Build 전용 설정 가이드 */}
                {showAiStudioGuide && (
                  <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-sm">
                    {/* 헤더 */}
                    <div className="bg-gradient-to-r from-violet-600 to-indigo-600 px-5 py-4 flex items-center justify-between">
                      <div>
                        <h3 className="text-sm font-black text-white flex items-center gap-2">
                          <Sparkles className="w-4 h-4" /> Google Drive 연동 + AI Studio Build 사용 가이드
                        </h3>
                        <p className="text-violet-200 text-[11px] mt-0.5">GitHub 연결된 AI Studio Build 앱 기준 설정 방법</p>
                      </div>
                      <button onClick={() => setShowAiStudioGuide(false)} className="p-1 rounded-full hover:bg-white/20 text-white/70 transition-colors">
                        <X className="w-4 h-4" />
                      </button>
                    </div>

                    <div className="p-4 space-y-4">

                      {/* PART A: Drive 연동 설정 (미연결일 때 강조) */}
                      <div>
                        <p className="text-[11px] font-black text-slate-400 uppercase tracking-widest mb-2">A. Google Drive 연동 설정 (최초 1회)</p>
                        <div className="space-y-2 text-xs">

                          {/* Step 1 */}
                          <div className="flex gap-3 bg-slate-50 border border-slate-200 rounded-lg p-3">
                            <div className="w-5 h-5 rounded-full bg-indigo-600 text-white text-[10px] font-black flex items-center justify-center shrink-0 mt-0.5">1</div>
                            <div className="min-w-0 flex-1">
                              <p className="font-black text-slate-800">Google Cloud Console → OAuth 클라이언트 등록</p>
                              <p className="text-slate-500 mt-1 leading-relaxed">
                                <a href="https://console.cloud.google.com/apis/credentials" target="_blank" rel="noopener noreferrer" className="text-indigo-600 underline font-bold">console.cloud.google.com</a>
                                {' '}→ Drive API 활성화 → 사용자 인증 정보 → OAuth 2.0 클라이언트 ID (웹 애플리케이션) 생성
                              </p>
                              <p className="text-slate-500 mt-1">승인된 리디렉션 URI에 아래 주소를 추가하세요:</p>
                              {driveStatus?.callbackUrl ? (
                                <div className="flex items-center gap-2 mt-1.5">
                                  <code className="flex-1 bg-indigo-50 border border-indigo-200 text-indigo-800 px-2.5 py-1.5 rounded-lg font-mono text-[11px] break-all">
                                    {driveStatus.callbackUrl}
                                  </code>
                                  <button
                                    onClick={() => copyToClipboard(driveStatus.callbackUrl!, 'callback')}
                                    className="shrink-0 px-2.5 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-[10px] font-bold transition-colors flex items-center gap-1"
                                  >
                                    {copiedField === 'callback' ? <><Check className="w-3 h-3"/>복사됨</> : <><Copy className="w-3 h-3"/>복사</>}
                                  </button>
                                </div>
                              ) : (
                                <code className="block mt-1.5 bg-slate-100 px-2.5 py-1.5 rounded text-[11px] text-slate-500">{window.location.origin}/api/google-auth/callback</code>
                              )}
                            </div>
                          </div>

                          {/* Step 2 */}
                          <div className="flex gap-3 bg-slate-50 border border-slate-200 rounded-lg p-3">
                            <div className="w-5 h-5 rounded-full bg-indigo-600 text-white text-[10px] font-black flex items-center justify-center shrink-0 mt-0.5">2</div>
                            <div className="min-w-0 flex-1">
                              <p className="font-black text-slate-800">AI Studio Build → Secrets 패널에 등록</p>
                              <p className="text-slate-500 mt-1 leading-relaxed">Google AI Studio → 앱 편집기 → 우측 상단 <strong className="text-slate-700">Secrets</strong> 탭 → + 추가</p>
                              <div className="mt-2 space-y-1.5">
                                {[
                                  { name: 'GOOGLE_CLIENT_ID', desc: 'OAuth 클라이언트 ID' },
                                  { name: 'GOOGLE_CLIENT_SECRET', desc: 'OAuth 클라이언트 보안 비밀번호' },
                                ].map(s => (
                                  <div key={s.name} className="flex items-center gap-2">
                                    <div className="flex-1 bg-slate-100 border border-slate-200 rounded px-2 py-1 font-mono text-[11px] text-slate-700">{s.name}</div>
                                    <span className="text-slate-400 text-[10px]">← {s.desc}</span>
                                  </div>
                                ))}
                              </div>
                              <p className="text-slate-400 mt-1.5 text-[10px]">등록 후 AI Studio에서 앱 재배포 (Save & Deploy)</p>
                            </div>
                          </div>

                          {/* Step 3 */}
                          <div className={`flex gap-3 rounded-lg p-3 border ${driveStatus?.configured && !driveStatus?.authenticated ? 'bg-violet-50 border-violet-300' : 'bg-slate-50 border-slate-200'}`}>
                            <div className="w-5 h-5 rounded-full bg-indigo-600 text-white text-[10px] font-black flex items-center justify-center shrink-0 mt-0.5">3</div>
                            <div className="min-w-0 flex-1">
                              <p className="font-black text-slate-800">구글 계정 인증 → Refresh Token 발급</p>
                              <p className="text-slate-500 mt-1 leading-relaxed">아래 버튼으로 구글 계정 허용 → 표시된 Refresh Token을 복사</p>
                              <a
                                href={driveStatus?.authUrl || `${getLocalServerUrl()}/api/google-auth`}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="mt-2 inline-flex items-center gap-1.5 px-3 py-2 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-lg text-[11px] transition-colors"
                              >
                                <Key className="w-3 h-3" /> 구글 계정 인증하기 →
                              </a>
                              <p className="text-slate-400 mt-1.5 text-[10px]">인증 완료 화면에서 GOOGLE_REFRESH_TOKEN 값을 복사해 Secrets에 추가 후 재배포</p>
                            </div>
                          </div>

                          {/* Step 4 (optional folder) */}
                          <div className="flex gap-3 bg-slate-50 border border-dashed border-slate-200 rounded-lg p-3">
                            <div className="w-5 h-5 rounded-full bg-slate-400 text-white text-[10px] font-black flex items-center justify-center shrink-0 mt-0.5">4</div>
                            <div className="min-w-0">
                              <p className="font-black text-slate-600">선택 — 업로드 대상 Drive 폴더 지정</p>
                              <p className="text-slate-400 mt-1 leading-relaxed text-[11px]">비워두면 내 드라이브에 <code className="bg-slate-100 px-1 rounded">주간보고_제출현황</code> 폴더가 자동 생성됩니다.<br/>특정 폴더를 지정하려면 Drive 폴더 URL 마지막 ID를 <code className="bg-slate-100 px-1 rounded">GOOGLE_DRIVE_FOLDER_ID</code> Secret으로 등록하세요.</p>
                            </div>
                          </div>

                        </div>
                      </div>

                      {/* PART B: 매주 AI Studio에서 보고서 검토하기 */}
                      <div>
                        <p className="text-[11px] font-black text-slate-400 uppercase tracking-widest mb-2">B. 매주 AI Studio에서 보고서 검토 · 다운로드</p>
                        <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 text-xs">
                          {[
                            { icon: '📤', title: '보고서 제출', desc: '각 교구·협회가 제출하면 Drive에 자동 업로드됩니다' },
                            { icon: '🤖', title: 'AI Studio에서 열기', desc: 'AI Studio 채팅 → 파일 추가 → Drive → 주간보고_제출현황 폴더 선택' },
                            { icon: '⬇️', title: 'AI 검토 후 다운로드', desc: 'AI가 서식 그대로 교정 → 워드 파일로 다운로드' },
                          ].map(({ icon, title, desc }) => (
                            <div key={title} className="bg-slate-50 border border-slate-200 rounded-lg p-3 text-center">
                              <p className="text-2xl mb-1">{icon}</p>
                              <p className="font-black text-slate-800 mb-1">{title}</p>
                              <p className="text-slate-500 leading-relaxed text-[11px]">{desc}</p>
                            </div>
                          ))}
                        </div>

                        {/* 바로가기 버튼 2개 */}
                        <div className="grid grid-cols-2 gap-2 mt-2">
                          <a
                            href="https://aistudio.google.com/app/prompts"
                            target="_blank"
                            rel="noopener noreferrer"
                            className="flex items-center justify-center gap-2 px-4 py-2.5 bg-gradient-to-r from-blue-600 to-violet-600 hover:from-blue-700 hover:to-violet-700 text-white font-bold rounded-lg text-xs transition-all"
                          >
                            <Sparkles className="w-3.5 h-3.5" /> Google AI Studio 열기 →
                          </a>
                          <a
                            href="https://drive.google.com"
                            target="_blank"
                            rel="noopener noreferrer"
                            className="flex items-center justify-center gap-2 px-4 py-2.5 bg-white border border-slate-200 hover:bg-slate-50 text-slate-700 font-bold rounded-lg text-xs transition-all"
                          >
                            <Folder className="w-3.5 h-3.5 text-emerald-600" /> Google Drive 열기 →
                          </a>
                        </div>
                      </div>

                    </div>
                  </div>
                )}

              </div>

              {/* Action Buttons Container */}
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mb-6">
                <button
                  disabled={isAdminCheckingAI}
                  onClick={startAdminAiReview}
                  className="relative overflow-hidden group bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-700 hover:to-indigo-700 text-white px-5 py-4 rounded-xl font-bold flex flex-col items-center justify-center gap-1 shadow-md transition-all active:scale-[0.98] disabled:opacity-50"
                >
                  <div className="absolute inset-0 bg-white/10 opacity-0 group-hover:opacity-100 transition-opacity" />
                  <div className="flex items-center gap-2">
                    <Bot className={`w-5 h-5 ${isAdminCheckingAI ? 'animate-spin' : 'animate-bounce'}`} />
                    <span className="text-base">전체 취합 AI 일괄 검토 시작</span>
                  </div>
                  <span className="text-[10px] text-purple-200 font-medium font-sans text-center">실시간 데이터 100% 취합 + AI 문맥·오타 완벽 교정</span>
                </button>

                <button
                  disabled={isAdminCheckingAI}
                  onClick={exportMasterToWord}
                  className="relative overflow-hidden group bg-gradient-to-r from-blue-600 to-sky-600 hover:from-blue-700 hover:to-sky-700 text-white px-5 py-4 rounded-xl font-bold flex flex-col items-center justify-center gap-1 shadow-md transition-all active:scale-[0.98] disabled:opacity-50"
                >
                  <div className="absolute inset-0 bg-white/10 opacity-0 group-hover:opacity-100 transition-opacity" />
                  <div className="flex items-center gap-2">
                    <Download className="w-5 h-5" />
                    <span className="text-base">통합 마스터 워드(.docx) 다운로드</span>
                  </div>
                  <span className="text-[10px] text-blue-200 font-medium font-sans text-center">전국 모든 교구 + 협회 업무보고서 A4 한 권으로 즉시 출력</span>
                </button>

                <button
                  onClick={async () => {
                    try {
                      const serverUrl = getLocalServerUrl();
                      const res = await fetch(`${serverUrl}/api/open-folder`, { method: 'POST' });
                      if (res.ok) {
                        const data = await res.json();
                        alert(`📂 구글 드라이브 동기화 폴더가 성공적으로 열렸습니다!\n경로: ${data.path}`);
                      } else {
                        alert("로컬 서버가 실행되지 않았거나 폴더를 열 수 없습니다.\n로컬 서버 구동 상태를 확인해 주세요.");
                      }
                    } catch (e) {
                      alert("로컬 서버 연결에 실패했습니다.\n로컬 서버를 먼저 실행해 주세요.");
                    }
                  }}
                  className="relative overflow-hidden group bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 text-white px-5 py-4 rounded-xl font-bold flex flex-col items-center justify-center gap-1 shadow-md transition-all active:scale-[0.98]"
                >
                  <div className="absolute inset-0 bg-white/10 opacity-0 group-hover:opacity-100 transition-opacity" />
                  <div className="flex items-center gap-2">
                    <Folder className="w-5 h-5" />
                    <span className="text-base">구글 드라이브 폴더 열기</span>
                  </div>
                  <span className="text-[10px] text-emerald-200 font-medium font-sans text-center">실시간 개별 저장된 전국 텍스트(.txt) 보고서 즉시 확인</span>
                </button>
              </div>

              {/* Live Loading/Progress indicator */}
              {adminCompilationProgress && (
                <div className="flex-1 flex flex-col items-center justify-center p-6 bg-slate-50 border border-dashed border-slate-200 rounded-xl">
                  <div className="w-12 h-12 rounded-full border-4 border-purple-200 border-t-purple-600 animate-spin mb-4" />
                  <p className="text-sm font-bold text-purple-700 animate-pulse">{adminCompilationProgress}</p>
                  <p className="text-xs text-slate-400 mt-2">이 작업은 취합되는 보고서 수에 따라 최대 30초 정도 소요될 수 있습니다.</p>
                </div>
              )}

              {/* Placeholder when idle */}
              {!adminCompilationProgress && !adminAiCorrections && (
                <div className="flex-1 flex flex-col items-center justify-center text-center p-6 bg-slate-50 border border-dashed border-slate-200 rounded-xl">
                  <Sparkles className="w-16 h-16 text-indigo-300 opacity-40 mb-4 animate-pulse" />
                  <h3 className="font-extrabold text-slate-700 text-base">취합 및 AI 일괄 검토 대기 중</h3>
                  <p className="text-xs text-slate-400 mt-2 max-w-sm leading-relaxed">
                    상단의 <strong>'전체 취합 AI 일괄 검토 시작'</strong> 버튼을 클릭하여 전국에서 수집된 보고서들을 실시간으로 수집하고 AI 문장 교정을 시작하세요.
                  </p>
                </div>
              )}

              {/* Suggestions list */}
              {!adminCompilationProgress && adminAiCorrections && (
                <div className="flex-1 flex flex-col min-h-0">
                  <div className="flex items-center justify-between mb-3">
                    <h3 className="font-extrabold text-sm text-slate-800 flex items-center gap-1.5">
                      <Sparkles className="w-4 h-4 text-purple-600" />
                      AI 교정 제안 목록 ({adminAiCorrections.length}건 발견)
                    </h3>
                    <span className="text-xs text-indigo-600 font-bold bg-indigo-50 px-2.5 py-1.5 rounded-full">실시간 완벽 편집 지원</span>
                  </div>

                  <div className="flex-1 overflow-y-auto space-y-3 pr-1 min-h-0 pb-4">
                    {adminAiCorrections.length === 0 ? (
                      <div className="text-center py-12 text-slate-400 text-sm font-bold">오타나 수정이 필요한 어색한 항목이 전혀 발견되지 않았습니다. 완벽합니다! 🎉</div>
                    ) : (
                      adminAiCorrections.map((c: any, index: number) => {
                        const key = `${c.parish}_${c.church}_${c.id}`;
                        const isChecked = !!adminSelectedCorrections[key];

                        return (
                          <div 
                            key={index} 
                            onClick={() => toggleAdminCorrectionSelected(c.parish, c.church, c.id)}
                            className={`border rounded-xl p-4 transition-all cursor-pointer hover:shadow-md select-none ${isChecked ? 'bg-indigo-50/50 border-indigo-200 shadow-sm' : 'bg-white border-slate-200 opacity-60'}`}
                          >
                            {/* Correction Header info */}
                            <div className="flex items-center justify-between mb-3 border-b border-slate-100 pb-2">
                              <div className="flex items-center gap-2">
                                <input 
                                  type="checkbox" 
                                  checked={isChecked} 
                                  onChange={() => {}} 
                                  className="w-4 h-4 rounded text-indigo-600 focus:ring-indigo-400 cursor-pointer"
                                />
                                <span className="text-xs font-black text-slate-800 bg-slate-100 px-2 py-1 rounded">
                                  {getDisplayParish(c.parish)} &gt; {getDisplayChurch(c.church)}
                                </span>
                              </div>
                              {c.reason && (
                                <span className="text-[10px] font-black text-indigo-700 bg-indigo-50 border border-indigo-100 px-2 py-0.5 rounded-full">
                                  💡 {c.reason}
                                </span>
                              )}
                            </div>

                            {/* Side by side comparison */}
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs leading-relaxed">
                              {/* Original */}
                              <div className="bg-red-50/50 border border-red-100 p-3 rounded-lg flex flex-col gap-1">
                                <span className="text-[10px] font-black text-red-600 select-none">수정 전 원본 내용</span>
                                <span className="text-slate-700 font-medium font-sans break-all line-through">{c.original}</span>
                              </div>
                              {/* Corrected */}
                              <div className="bg-emerald-50/50 border border-emerald-100 p-3 rounded-lg flex flex-col gap-1">
                                <span className="text-[10px] font-black text-emerald-600 select-none">AI 추천 교정 내용</span>
                                <span className="text-slate-800 font-extrabold font-sans break-all">{c.corrected}</span>
                              </div>
                            </div>
                          </div>
                        );
                      })
                    )}
                  </div>

                  {/* Actions footer bar inside right panel */}
                  {adminAiCorrections.length > 0 && (
                    <div className="pt-4 border-t border-slate-200 mt-2 flex items-center justify-end gap-3 bg-white">
                      <button
                        onClick={() => {
                          setAdminAiCorrections(null);
                        }}
                        className="px-4 py-2.5 text-slate-500 hover:bg-slate-100 rounded-lg text-xs font-bold transition-all"
                      >
                        제안 닫기
                      </button>
                      <button
                        onClick={applySelectedAdminCorrections}
                        className="px-5 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-xs font-extrabold shadow-md shadow-indigo-100 transition-all flex items-center gap-1.5"
                      >
                        <Check className="w-4 h-4" />
                        선택된 {Object.values(adminSelectedCorrections).filter(Boolean).length}개 제안 실시간 반영 및 데이터 적용
                      </button>
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* Save Loading Overlay */}
      {isSaving && (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm flex items-center justify-center z-[60]">
          <div className="bg-white rounded-2xl shadow-2xl px-10 py-8 flex flex-col items-center gap-4">
            <RefreshCw className="w-10 h-10 text-blue-500 animate-spin" />
            <p className="text-slate-800 font-bold text-lg">저장 중...</p>
            <p className="text-slate-400 text-sm">잠시만 기다려 주세요</p>
          </div>
        </div>
      )}

      {/* AI Review & Export Modal */}
      {showAiModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl flex flex-col max-h-[90vh] overflow-hidden">
            <div className="bg-gradient-to-r from-blue-600 to-indigo-600 p-6 text-white flex items-center justify-between">
              <div className="flex items-center gap-3">
                <Bot className="w-7 h-7 text-blue-100" />
                <div>
                  <h3 className="text-xl font-bold">AI 편집 및 워드 생성</h3>
                  <p className="text-blue-100 text-sm mt-1">작성된 내용을 분석하여 오타 및 양식을 교정합니다.</p>
                </div>
              </div>
              <button disabled={isCheckingAI} onClick={() => setShowAiModal(false)} className="p-2 hover:bg-white/10 rounded-full text-white/80 transition-colors">
                <X className="w-6 h-6" />
              </button>
            </div>
            
            <div className="p-6 overflow-y-auto flex-1 bg-slate-50">
              {isCheckingAI ? (
                <div className="flex flex-col items-center justify-center py-16 text-slate-500">
                  <div className="w-12 h-12 border-4 border-blue-200 border-t-blue-600 rounded-full animate-spin mb-4"></div>
                  <p className="text-lg font-medium text-slate-700">AI가 보고서를 꼼꼼히 읽고 있습니다...</p>
                  <p className="text-sm mt-2">맞춤법, 띄어쓰기, 보고서 체재를 확인 중입니다.</p>
                </div>
              ) : aiCorrections ? (
                <div>
                  {aiCorrections.length > 0 ? (
                    <div className="space-y-4">
                      <div className="flex items-center gap-2 text-indigo-700 mb-4 font-medium bg-indigo-50 p-3 rounded-lg border border-indigo-100">
                        <AlertCircle className="w-5 h-5" /> 
                        총 {aiCorrections.length}건의 수정 제안이 있습니다.
                      </div>
                      {Object.entries(
                        aiCorrections.reduce((acc, corr) => {
                          if (!acc[corr.church]) acc[corr.church] = [];
                          acc[corr.church].push(corr);
                          return acc;
                        }, {} as Record<string, any[]>)
                      ).map(([churchName, corrs]) => (
                        <div key={churchName} className="mb-6">
                          <h4 className="text-lg font-bold text-slate-800 mb-3 flex items-center gap-2">
                            <BookOpen className="w-5 h-5 text-blue-600" /> {churchName}
                          </h4>
                          <div className="space-y-4 pl-2 border-l-2 border-blue-200">
                            {(corrs as any[]).map((corr: any) => (
                              <div key={`${corr.church}-${corr.id}`} className="bg-white border text-sm border-slate-200 shadow-sm rounded-lg overflow-hidden flex flex-col">
                                <div className="p-3 border-b border-slate-100 flex gap-4">
                                  <div className="flex-1">
                                    <span className="text-xs font-bold text-red-500 bg-red-50 px-2 py-0.5 rounded mr-2">원본</span>
                                    <span className="text-slate-600 line-through decoration-red-300">{corr.original}</span>
                                  </div>
                                  <div className="flex-1">
                                    <span className="text-xs font-bold text-emerald-600 bg-emerald-50 px-2 py-0.5 rounded mr-2">수정안</span>
                                    <span className="text-slate-900 font-medium">{corr.corrected}</span>
                                    {corr.level !== undefined && <span className="ml-2 text-xs font-bold text-purple-600 bg-purple-50 px-1.5 py-0.5 rounded">L{corr.level}</span>}
                                  </div>
                                </div>
                                <div className="p-3 bg-slate-50 flex items-center justify-between gap-4">
                                  <p className="text-slate-500 text-xs leading-relaxed"><strong className="text-slate-700">이유:</strong> {corr.reason}</p>
                                  <button 
                                    onClick={() => applyCorrection(corr.church, corr.id, corr.corrected, corr.level)}
                                    className="px-3 py-1.5 bg-white border border-slate-300 hover:border-indigo-400 hover:text-indigo-600 text-slate-600 rounded text-xs font-medium transition-colors shrink-0 whitespace-nowrap"
                                  >
                                    이 항목만 적용
                                  </button>
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <div className="flex flex-col items-center justify-center py-16 text-center">
                      <div className="w-16 h-16 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center mb-4">
                        <Check className="w-8 h-8" />
                      </div>
                      <h4 className="text-xl font-bold text-slate-800 mb-2">완벽합니다!</h4>
                      <p className="text-slate-500">발견된 오타나 양식 불일치가 없습니다. 바로 문서를 생성할 수 있습니다.</p>
                    </div>
                  )}
                </div>
              ) : null}
            </div>

            {!isCheckingAI && aiCorrections && (
              <div className="p-5 border-t border-slate-200 bg-white flex justify-between items-center gap-3">
                <button 
                  onClick={() => setShowAiModal(false)}
                  className="px-5 py-2.5 text-slate-600 hover:bg-slate-100 rounded-lg font-medium transition-colors"
                >
                  취소
                </button>
                <div className="flex gap-3">
                  {aiCorrections.length > 0 && (
                     <button 
                       onClick={applyAllCorrections}
                       className="px-5 py-2.5 bg-indigo-50 text-indigo-700 hover:bg-indigo-100 rounded-lg font-bold transition-colors"
                     >
                       모두 적용하기
                     </button>
                  )}
                  <button 
                    onClick={() => {
                      exportToWord();
                      setShowAiModal(false);
                    }}
                    className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-bold flex items-center gap-2 shadow-md transition-colors"
                  >
                    <Download className="w-5 h-5" />
                    워드로 다운로드
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {/* Crop Modal */}
      {cropModalOpen && (
        <div className="fixed inset-0 bg-black/60 flex items-center justify-center p-4 z-[99]">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-2xl flex flex-col max-h-[90vh]">
            <div className="p-4 border-b border-slate-200 flex items-center justify-between">
              <h3 className="text-lg font-bold">이미지 자르기 및 크기 조절</h3>
              <button onClick={() => setCropModalOpen(false)} className="p-2 hover:bg-slate-100 rounded-md text-slate-500">
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="p-4 overflow-auto flex-1 flex justify-center bg-slate-50">
              <ReactCrop
                crop={crop}
                onChange={(_, percentCrop) => setCrop(percentCrop)}
                onComplete={(c) => setCompletedCrop(c)}
              >
                <img
                  ref={imgRef}
                  src={cropImageSrc}
                  alt="Crop area"
                  className="max-w-full"
                />
              </ReactCrop>
            </div>
            <div className="p-4 border-t border-slate-200 flex justify-end gap-2">
              <button
                onClick={() => setCropModalOpen(false)}
                className="px-4 py-2 border border-slate-300 rounded-md text-slate-700 hover:bg-slate-50 font-medium"
              >
                취소
              </button>
              <button
                onClick={applyCrop}
                className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 font-medium flex items-center gap-2"
              >
                <CropIcon className="w-4 h-4" /> 영역 저장
              </button>
            </div>
          </div>
        </div>
      )}

      {/* JSON Export Modal */}
      {showJsonModal && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-3xl flex flex-col max-h-[90vh]">
            <div className="p-5 border-b border-slate-200 flex items-center justify-between">
              <div>
                <h3 className="text-lg font-bold">데이터 구조 내보내기</h3>
                <p className="text-sm text-slate-500 mt-1">현재 선택된 <strong>{parish}</strong> 소속 전체 교회의 데이터를 추출합니다.</p>
              </div>
              <button onClick={() => setShowJsonModal(false)} className="p-2 hover:bg-slate-100 rounded-md text-slate-500">
                <X className="w-5 h-5" />
              </button>
            </div>
            
            <div className="p-5 overflow-y-auto flex-1">
              <div className="flex gap-2 mb-4">
                <button 
                  onClick={() => setJsonFormat('flat')}
                  className={`px-4 py-2 rounded-md text-sm font-medium transition-colors ${jsonFormat === 'flat' ? 'bg-slate-800 text-white' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'}`}
                >
                  Flat Array (현재 상태)
                </button>
                <button 
                  onClick={() => setJsonFormat('tree')}
                  className={`px-4 py-2 rounded-md text-sm font-medium transition-colors ${jsonFormat === 'tree' ? 'bg-slate-800 text-white' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'}`}
                >
                  Tree Structure (계층형)
                </button>
              </div>

              <pre className="bg-slate-900 text-slate-50 p-4 rounded-lg overflow-x-auto text-sm leading-relaxed max-h-[50vh]">
                {JSON.stringify(getParishData(), null, 2)}
              </pre>
            </div>

            <div className="p-5 border-t border-slate-200 flex justify-end gap-3">
              <button 
                onClick={() => setShowJsonModal(false)}
                className="px-4 py-2 text-slate-600 hover:bg-slate-100 rounded-md font-medium"
              >
                닫기
              </button>
              <button 
                onClick={handleCopyJson}
                className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-md font-medium flex items-center gap-2"
              >
                {copied ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
                {copied ? '복사됨!' : 'JSON 복사하기'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Guide Modal */}
      {showGuideModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-2xl flex flex-col max-h-[85vh]">
            <div className="flex items-center justify-between p-5 border-b border-slate-200 bg-slate-50">
              <h2 className="text-xl font-bold flex items-center gap-2 text-slate-800">
                <BookOpen className="w-5 h-5 text-blue-600" /> 주간업무보고 100% 활용하기
              </h2>
              <button onClick={() => setShowGuideModal(false)} className="text-slate-400 hover:bg-slate-200 hover:text-slate-600 rounded-full p-2 transition-colors">
                <X className="w-5 h-5" />
              </button>
            </div>
            
            <div className="p-6 overflow-y-auto space-y-6 text-slate-700 bg-white">
              
              <section>
                <h3 className="font-bold text-lg text-slate-900 border-b border-slate-200 pb-2 mb-3 flex items-center gap-2">
                  <span className="bg-blue-100 text-blue-700 w-6 h-6 rounded-full flex items-center justify-center text-sm font-bold">1</span>
                  다양한 단축키로 빠른 작성
                </h3>
                <ul className="list-disc list-inside space-y-2 text-sm ml-2">
                  <li><kbd className="px-1.5 py-0.5 bg-slate-100 border border-slate-300 rounded font-mono text-xs">Enter</kbd> : 같은 수준의 새 항목을 아래에 추가합니다.</li>
                  <li><kbd className="px-1.5 py-0.5 bg-slate-100 border border-slate-300 rounded font-mono text-xs">Tab</kbd> / <kbd className="px-1.5 py-0.5 bg-slate-100 border border-slate-300 rounded font-mono text-xs">Shift + Tab</kbd> : 항목의 수준(Level)을 내리거나 올립니다.</li>
                  <li><kbd className="px-1.5 py-0.5 bg-slate-100 border border-slate-300 rounded font-mono text-xs">상/하 방향키</kbd> : 위/아래 항목으로 빠르게 이동합니다.</li>
                </ul>
              </section>

              <section>
                <h3 className="font-bold text-lg text-slate-900 border-b border-slate-200 pb-2 mb-3 flex items-center gap-2">
                  <span className="bg-emerald-100 text-emerald-700 w-6 h-6 rounded-full flex items-center justify-center text-sm font-bold">2</span>
                  표(Table)와 차트 기능
                </h3>
                <div className="bg-slate-50 p-4 rounded-lg text-sm space-y-3">
                  <p><strong>편리한 붙여넣기:</strong> 엑셀이나 워드에 작성된 표를 복사하여 빈 항목에 붙여넣기 해보세요. <strong>자동으로 표가 생성</strong>됩니다.</p>
                  <p><strong>수동 생성:</strong> 각 항목 우측 메뉴바에서 <TableIcon className="w-4 h-4 inline text-indigo-500" /> 버튼을 눌러 표를 직접 삽입할 수도 있습니다.</p>
                  <p><strong>강조 효과:</strong> 표의 각 셀이나 행 좌측에 마우스를 올리면 나타나는 <Highlighter className="w-3.5 h-3.5 inline text-blue-500" /> 버튼을 눌러 파란색 볼드체로 강조하세요.</p>
                  <p><strong>차트 변환:</strong> 표 데이터를 바탕으로 '막대/선/원형 차트'로 즉시 전환할 수 있습니다.</p>
                </div>
              </section>

              <section>
                <h3 className="font-bold text-lg text-slate-900 border-b border-slate-200 pb-2 mb-3 flex items-center gap-2">
                  <span className="bg-amber-100 text-amber-700 w-6 h-6 rounded-full flex items-center justify-center text-sm font-bold">3</span>
                  복잡한 자료는 '사진 첨부'를 권장합니다.
                </h3>
                <div className="flex items-start gap-3 bg-amber-50 p-4 rounded-lg border border-amber-100">
                  <AlertCircle className="w-5 h-5 text-amber-500 shrink-0 mt-0.5" />
                  <div className="text-sm text-slate-800 leading-relaxed">
                    시스템 내에서 직접 구현하기 까다로운 <strong>복잡한 디자인의 표, 외부 그래프, 특수 기호</strong> 등은 무리해서 옮기지 마세요.<br /><br />
                    해당 자료를 캡처하여 사진으로 저장한 뒤, 항목 우측의 <ImageIcon className="w-4 h-4 inline text-emerald-600" /> 아이콘을 클릭해 <strong>사진 형태로 첨부</strong>하시는 것이 훨씬 깔끔하고 편리합니다.
                  </div>
                </div>
              </section>

              <section>
                <h3 className="font-bold text-lg text-slate-900 border-b border-slate-200 pb-2 mb-3 flex items-center gap-2">
                  <span className="bg-purple-100 text-purple-700 w-6 h-6 rounded-full flex items-center justify-center text-sm font-bold">4</span>
                  AI 정리 및 워드 내보내기
                </h3>
                <p className="text-sm mb-2">하단의 <span className="font-semibold text-purple-700"><Bot className="w-4 h-4 inline mr-1" />AI 문맥 검토</span> 버튼을 눌러 오탈자를 점검하고, 우측 하단 <span className="font-semibold text-blue-700"><Download className="w-4 h-4 inline mr-1"/>Word로 다운로드</span> 버튼으로 깨끗하게 <code>.docx</code> 파일로 출력하세요.</p>
              </section>

              <section className="bg-indigo-50 border border-indigo-100 p-4 rounded-lg">
                <h3 className="font-bold text-base text-indigo-900 pb-2 mb-3 border-b border-indigo-200 flex items-center gap-2">
                  <Check className="w-4 h-4 text-indigo-700" /> 주간보고 작성 표준 예시 (개조식)
                </h3>
                <div className="space-y-2 text-xs text-indigo-950 leading-relaxed font-medium">
                  <p>1. 00교구 000선교회 집회 <span className="text-[10px] bg-indigo-200 text-indigo-800 px-1.5 py-0.5 rounded font-bold ml-2">1단계 (L0)</span></p>
                  <p className="ml-5">1) 일시: 2026년 00월 00일 0요일 오전 10시 <span className="text-[10px] bg-indigo-200 text-indigo-800 px-1.5 py-0.5 rounded font-bold ml-2">2단계 (L1)</span></p>
                  <p className="ml-5">2) 장소: 00교회 대성전</p>
                  <p className="ml-5">3) 참석인원: 60명 (선교사 50명, 교구 공직자 6명, 청년스텝 4명)</p>
                  <p className="ml-5">4) 내용: 000 교구장 환영사, 특강, 000 원장 격려사, 화동 프로그램</p>
                  <div className="ml-5 flex items-center gap-2">
                    5) 대표사진: <span className="text-indigo-700 text-xs flex items-center gap-1"><ImageIcon className="w-3.5 h-3.5"/> (우측 사진 아이콘을 클릭하여 첨부)</span>
                  </div>
                </div>
              </section>

            </div>
          </div>
        </div>
      )}


      {tableContextMenu && (
        <div 
          className="fixed z-50 bg-white border border-slate-200 shadow-xl rounded py-1 w-36 text-sm"
          style={{ left: tableContextMenu.x, top: tableContextMenu.y }}
        >
          <div className="px-2 py-1.5 flex justify-between">
            <button onClick={() => setCellAlignment(tableContextMenu.id, tableContextMenu.r, tableContextMenu.c, 'left')} className="p-1.5 hover:bg-slate-100 rounded text-slate-600" title="왼쪽 맞춤"><AlignLeft className="w-4 h-4" /></button>
            <button onClick={() => setCellAlignment(tableContextMenu.id, tableContextMenu.r, tableContextMenu.c, 'center')} className="p-1.5 hover:bg-slate-100 rounded text-slate-600" title="가운데 맞춤"><AlignCenter className="w-4 h-4" /></button>
            <button onClick={() => setCellAlignment(tableContextMenu.id, tableContextMenu.r, tableContextMenu.c, 'right')} className="p-1.5 hover:bg-slate-100 rounded text-slate-600" title="오른쪽 맞춤"><AlignRight className="w-4 h-4" /></button>
          </div>
          <div className="h-px bg-slate-200 my-1"></div>
          {tableSelection?.start?.r !== tableSelection?.end?.r || tableSelection?.start?.c !== tableSelection?.end?.c ? (
            <button className="w-full text-left px-3 py-1.5 hover:bg-slate-100 font-medium text-blue-600" onClick={() => handleMergeSelection(tableContextMenu.id)}>선택 병합</button>
          ) : (
            <div className="px-3 py-1.5 text-slate-400 text-xs text-center border-b border-slate-100">병합할 셀들을<br/>드래그하세요</div>
          )}
          <button className="w-full text-left px-3 py-1.5 hover:bg-slate-100" onClick={() => unmergeCell(tableContextMenu.id, tableContextMenu.r, tableContextMenu.c)}>병합 해제</button>
          <div className="h-px bg-slate-200 my-1"></div>
          <button className="w-full text-left px-3 py-1.5 hover:bg-slate-100" onClick={() => addTableRow(tableContextMenu.id)}>행 추가</button>
          <button className="w-full text-left px-3 py-1.5 hover:bg-slate-100 text-red-600" onClick={() => removeTableRow(tableContextMenu.id, tableContextMenu.r)}>행 삭제</button>
          <div className="h-px bg-slate-200 my-1"></div>
          <button className="w-full text-left px-3 py-1.5 hover:bg-slate-100" onClick={() => addTableCol(tableContextMenu.id)}>열 추가</button>
          <button className="w-full text-left px-3 py-1.5 hover:bg-slate-100 text-red-600" onClick={() => removeTableCol(tableContextMenu.id, tableContextMenu.c)}>열 삭제</button>
        </div>
      )}

    </div>
  );
}
