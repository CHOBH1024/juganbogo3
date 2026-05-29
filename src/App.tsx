import React, { useState, useEffect, useRef, useCallback } from 'react';
import { isBrowserAIReady, initBrowserAI, runBrowserAI } from './browserAI';
import { Plus, X, ArrowLeft, ArrowRight, ArrowUp, ArrowDown, FileJson, Copy, Check, Save, Download, Bot, Clock, AlertCircle, RefreshCw, Image as ImageIcon, Crop as CropIcon, Table as TableIcon, BarChart2, Trash2, Highlighter, BookOpen, AlignLeft, AlignCenter, AlignRight, Settings, Key, Bell, Upload, FileText, Sparkles, Folder, User, CheckCircle, Info, AlertTriangle } from 'lucide-react';
import { GoogleGenAI, Type } from '@google/genai';
import { Document, Packer, Paragraph, TextRun, HeadingLevel, ImageRun, Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign } from "docx";
import TextareaAutosize from 'react-textarea-autosize';
import ReactCrop, { Crop, PixelCrop, centerCrop, makeAspectCrop } from 'react-image-crop';
import 'react-image-crop/dist/ReactCrop.css';
import { BarChart, Bar, LineChart, Line, PieChart, Pie, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, Cell } from 'recharts';
import { parseHtmlTable } from './lib/tableParser';
import { supabase } from './lib/supabase';
import RoleSelection, { Role } from './RoleSelection';
const getLocalServerUrl = () => {
  const customUrl = localStorage.getItem('LOCAL_SERVER_URL');
  if (customUrl) return customUrl;
  return 'http://localhost:5000';
};

// ngrok 터널 경유 시 브라우저 경고창 우회 헤더 자동 추가
const localFetch = (url: string, options: RequestInit = {}): Promise<Response> => {
  const serverUrl = getLocalServerUrl();
  const isNgrok = serverUrl.includes('ngrok');
  if (isNgrok) {
    options = {
      ...options,
      headers: {
        'ngrok-skip-browser-warning': 'true',
        ...(options.headers || {}),
      },
    };
  }
  return fetch(url, options);
};

// ── 토스트 알림 시스템 ──────────────────────────────────────────
type ToastType = 'success' | 'error' | 'info' | 'warning';
interface Toast { id: number; message: string; type: ToastType; }
let _toastId = 0;
let _toastDispatch: ((t: Toast) => void) | null = null;
const toast = {
  success: (msg: string) => _toastDispatch?.({ id: ++_toastId, message: msg, type: 'success' }),
  error: (msg: string) => _toastDispatch?.({ id: ++_toastId, message: msg, type: 'error' }),
  info: (msg: string) => _toastDispatch?.({ id: ++_toastId, message: msg, type: 'info' }),
  warning: (msg: string) => _toastDispatch?.({ id: ++_toastId, message: msg, type: 'warning' }),
};

// ── RLS 우회를 위한 Storage 기반 JSON DB 헬퍼 (로컬 및 클라우드 동시 지원)
const fetchDbData = async (id: string) => {
  const isLocal = localStorage.getItem('IS_LOCAL_MODE') === 'true';
  if (isLocal) {
    try {
      const serverUrl = getLocalServerUrl();
      const res = await localFetch(`${serverUrl}/api/load-data/${id}`);
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
      await localFetch(`${serverUrl}/api/save-data`, {
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

  // 토스트 알림 상태
  const [toasts, setToasts] = useState<Toast[]>([]);
  useEffect(() => {
    _toastDispatch = (t: Toast) => {
      setToasts(prev => [...prev.slice(-4), t]);
      setTimeout(() => setToasts(prev => prev.filter(x => x.id !== t.id)), 3500);
    };
    return () => { _toastDispatch = null; };
  }, []);

  const [role, setRole] = useState<Role>(null);
  const [isLocalMode, setIsLocalMode] = useState(() => localStorage.getItem('IS_LOCAL_MODE') === 'true');
  const [isCloudAiAvailable, setIsCloudAiAvailable] = useState(false);
  const [isBrowserAiReady, setIsBrowserAiReady] = useState(false);
  const [isBrowserAiLoading, setIsBrowserAiLoading] = useState(false);
  const [browserAiProgress, setBrowserAiProgress] = useState(0);

  // 로컬 서버 자동 감지 (마운트 시 ping — 실패 시 isLocalMode 리셋)
  useEffect(() => {
    const detect = async () => {
      try {
        const res = await localFetch(`${getLocalServerUrl()}/api/ping`, { signal: AbortSignal.timeout(2000) });
        if (res.ok) {
          localStorage.setItem('IS_LOCAL_MODE', 'true');
          setIsLocalMode(true);
          return;
        }
      } catch {}
      // 서버 응답 없으면 로컬 모드 해제
      localStorage.removeItem('IS_LOCAL_MODE');
      setIsLocalMode(false);
    };
    detect();
  }, []);

  // Vercel 클라우드 AI 가용 여부 확인
  useEffect(() => {
    fetch('/api/ai-review')
      .then(r => r.json())
      .then(d => { if (d.available) setIsCloudAiAvailable(true); })
      .catch(() => {});
  }, []);
  const [parish, setParish] = useState(() => localStorage.getItem('APP_PARISH') || "천원특별");
  const [church, setChurch] = useState(() => localStorage.getItem('APP_CHURCH') || PARISH_CHURCH_MAP["천원특별"][0]);
  
  const handleSelectRole = (selectedRole: Role, data?: any) => {
    localStorage.setItem('APP_ROLE', selectedRole || '');
    setRole(selectedRole);
    if (data?.parish) {
      localStorage.setItem('APP_PARISH', data.parish);
      setParish(data.parish);
    }
    if (data?.church) {
      localStorage.setItem('APP_CHURCH', data.church);
      setChurch(data.church);
    }
    if (selectedRole === 'admin') {
      setActiveTab('admin_console');
    }
    if (selectedRole === 'manager' && data?.parish) {
      setAdminConsoleParish(data.parish);
    }
  };
  
  const [reportData, setReportData] = useState<ReportItem[]>(DEFAULT_REPORT);
  const [lastSaved, setLastSaved] = useState<string | null>(null);
  const [tableSelection, setTableSelection] = useState<{ id: number, start: { r: number, c: number }, end: { r: number, c: number }, isDragging: boolean } | null>(null);
  const [tableContextMenu, setTableContextMenu] = useState<{ x: number, y: number, id: number, r: number, c: number } | null>(null);
  
  const [nextId, setNextId] = useState(5);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [showGuideModal, setShowGuideModal] = useState(false);
  const [showAiPasteModal, setShowAiPasteModal] = useState(false);
  const [isAiProcessing, setIsAiProcessing] = useState(false);
  const aiPasteRef = useRef<HTMLDivElement>(null);
  const [copied, setCopied] = useState(false);
  const [appConfig, setAppConfig] = useState<{solarDate: string, heavenlyDate: string, deadline?: string} | null>(null);
  const [jsonFormat, setJsonFormat] = useState<'flat' | 'tree'>('flat');

  const [isSaving, setIsSaving] = useState(false);
  const [showSubmitCelebration, setShowSubmitCelebration] = useState(false);
  const [driveSaveResult, setDriveSaveResult] = useState<'ok' | 'skipped' | 'error' | 'saving' | null>(null);
  const [driveSavedAt, setDriveSavedAt] = useState<string | null>(null);
  const [quickEntryMode, setQuickEntryMode] = useState(false);
  const [quickEntryText, setQuickEntryText] = useState('');
  const [simpleMode, setSimpleMode] = useState(false);

  // 로컬 Claude AI 검토 상태
  const [isCheckingAI, setIsCheckingAI] = useState(false);
  const [aiCorrections, setAiCorrections] = useState<{id: number; original: string; corrected: string; reason: string}[] | null>(null);

  const [cropModalOpen, setCropModalOpen] = useState(false);
  const [cropImageSrc, setCropImageSrc] = useState<string>('');
  const [cropItemId, setCropItemId] = useState<number | null>(null);
  const [crop, setCrop] = useState<Crop>();
  const [completedCrop, setCompletedCrop] = useState<PixelCrop>();
  const imgRef = useRef<HTMLImageElement>(null);
  const isLoadingDataRef = useRef(false);
  const [isLoadingData, setIsLoadingData] = useState(false);
  const docUploadRef = useRef<HTMLInputElement>(null);

  const [status, setStatus] = useState<'draft' | 'submitted'>('draft');
  const [parishStats, setParishStats] = useState<Record<string, 'empty' | 'draft' | 'submitted'>>({});

  const [activeTab, setActiveTab] = useState<'report' | 'notice_write' | 'notice' | 'admin_console'>('report');
  const [noticeTitle, setNoticeTitle] = useState('');
  const [noticeCategory, setNoticeCategory] = useState('공지');
  const [noticePdfUrl, setNoticePdfUrl] = useState<string | null>(null);
  const [mobileView, setMobileView] = useState<'editor' | 'preview'>('editor');
  const [noticeSearchQuery, setNoticeSearchQuery] = useState('');
  const [noticeFilterCategory, setNoticeFilterCategory] = useState<string>('전체');

  const handleNoticeWriteTab = () => {
    if (activeTab === 'notice_write') return;
    withAdminBypass('공지 작성 모드', (pwd) => pwd === 'chongmu2027', () => {
      setActiveTab('notice_write');
      setReportData([]);
      setNoticeTitle('');
      setNoticePdfUrl(null);
    });
  };

  const [notices, setNotices] = useState<any[]>([]);
  const [readNoticeIds, setReadNoticeIds] = useState<Set<string>>(() => {
    try { return new Set(JSON.parse(localStorage.getItem('read_notice_ids') || '[]')); } catch { return new Set(); }
  });
  const [activeNotice, setActiveNotice] = useState<any | null>(null);
  const [isUploadingNotice, setIsUploadingNotice] = useState(false);

  const markNoticeRead = (id: string) => {
    setReadNoticeIds(prev => {
      const next = new Set(prev);
      next.add(id);
      localStorage.setItem('read_notice_ids', JSON.stringify([...next]));
      return next;
    });
  };

  // 관리자 통합 취합 콘솔용 상태
  const [isAdminCheckingAI, setIsAdminCheckingAI] = useState(false);
  const [adminAiCorrections, setAdminAiCorrections] = useState<any[] | null>(null);
  const [adminReportStatusMap, setAdminReportStatusMap] = useState<Record<string, 'empty' | 'draft' | 'submitted'>>({});
  const [adminReportTimestampMap, setAdminReportTimestampMap] = useState<Record<string, string | null>>({});
  const [adminExpandedParish, setAdminExpandedParish] = useState<string | null>(null);
  const [adminActiveParish, setAdminActiveParish] = useState<string>('전체');
  const [adminConsoleParish, setAdminConsoleParish] = useState<string>(() => {
    const savedRole = localStorage.getItem('APP_ROLE');
    if (savedRole === 'manager') return localStorage.getItem('APP_PARISH') || Object.keys(PARISH_CHURCH_MAP)[0];
    return Object.keys(PARISH_CHURCH_MAP)[0];
  });
  const [adminLastRefreshed, setAdminLastRefreshed] = useState<Date | null>(null);
  const [adminAutoRefresh, setAdminAutoRefresh] = useState(false);
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
  const [showResetModal, setShowResetModal] = useState(false);
  const [resetSelectedParishes, setResetSelectedParishes] = useState<string[]>([]);
  const [resetSelectedChurches, setResetSelectedChurches] = useState<{ [k: string]: string[] }>({});
  const [resetMode, setResetMode] = useState<'quick' | 'custom'>('quick');

  // 내 교회 변경 모달 (church 역할 전용)
  const [showChurchChangeModal, setShowChurchChangeModal] = useState(false);
  const [churchChangeParish, setChurchChangeParish] = useState(parish);
  const [churchChangeChurch, setChurchChangeChurch] = useState(church);

  // 보고서 섹션 접기/펼치기 (대항목 ID → collapsed)
  const [collapsedSections, setCollapsedSections] = useState<Set<number>>(new Set());
  const toggleSection = (id: number) => setCollapsedSections(prev => {
    const next = new Set(prev);
    if (next.has(id)) next.delete(id); else next.add(id);
    return next;
  });

  // 교회 보고서 미리보기 모달 (admin console)
  const [previewChurchModal, setPreviewChurchModal] = useState<{
    parish: string;
    church: string;
    status: 'empty' | 'draft' | 'submitted';
    data: any[];
    timestamp: string | null;
  } | null>(null);

  const openChurchPreview = async (p: string, c: string) => {
    const key = `report_${p}_${c}`;
    const st = (adminReportStatusMap[`${p}_${c}`] || 'empty') as 'empty' | 'draft' | 'submitted';
    const ts = adminReportTimestampMap[`${p}_${c}`] || null;
    let data: any[] = [];
    try {
      const sess = sessionStorage.getItem(key);
      const local = localStorage.getItem(key);
      const raw = sess || local;
      if (raw) {
        const parsed = JSON.parse(raw);
        data = parsed.data || [];
      } else {
        const db = await fetchDbData(`${p}_${c}`);
        data = db?.data || [];
      }
    } catch {}
    setPreviewChurchModal({ parish: p, church: c, status: st, data, timestamp: ts });
  };

  // 새 주간보고 시작 모달 (admin 전용)
  const [showNewWeekModal, setShowNewWeekModal] = useState(false);
  const [newWeekPassword, setNewWeekPassword] = useState('');
  const [newWeekSolarDate, setNewWeekSolarDate] = useState('');
  const [newWeekHeavenlyDate, setNewWeekHeavenlyDate] = useState('');
  const [newWeekDeadline, setNewWeekDeadline] = useState('');
  const [newWeekPasswordError, setNewWeekPasswordError] = useState('');
  const [newWeekIsResetting, setNewWeekIsResetting] = useState(false);

  // 범용 비밀번호 모달
  const [showPasswordModal, setShowPasswordModal] = useState(false);
  const [passwordModalTitle, setPasswordModalTitle] = useState('');
  const [passwordModalInput, setPasswordModalInput] = useState('');
  const [passwordModalError, setPasswordModalError] = useState('');
  const [passwordModalCallback, setPasswordModalCallback] = useState<((pwd: string) => boolean) | null>(null);
  const [passwordModalOnSuccess, setPasswordModalOnSuccess] = useState<((pwd: string) => void) | null>(null);

  const openPasswordModal = (title: string, checkFn: (pwd: string) => boolean, onSuccess: (pwd: string) => void) => {
    setPasswordModalTitle(title);
    setPasswordModalInput('');
    setPasswordModalError('');
    setPasswordModalCallback(() => checkFn);
    setPasswordModalOnSuccess(() => onSuccess);
    setShowPasswordModal(true);
  };

  // 관리자 모드에서는 비밀번호 모달 없이 바로 실행
  const withAdminBypass = (title: string, checkFn: (pwd: string) => boolean, onSuccess: (pwd: string) => void) => {
    if (role === 'admin') { onSuccess('admin'); return; }
    openPasswordModal(title, checkFn, onSuccess);
  };

  const submitPasswordModal = () => {
    if (passwordModalCallback && passwordModalCallback(passwordModalInput)) {
      const enteredPwd = passwordModalInput;
      setShowPasswordModal(false);
      passwordModalOnSuccess?.(enteredPwd);
    } else {
      setPasswordModalError('비밀번호가 일치하지 않습니다.');
    }
  };

  const getRelativeTime = (dateStr: string | null): string => {
    if (!dateStr) return '';
    try {
      // dateStr can be "2026. 5. 27. 오전 3:25:00" (ko-KR locale) or ISO
      const date = new Date(dateStr.replace(/(\d{4})\. (\d+)\. (\d+)\. (오전|오후) (\d+):(\d+):(\d+)/, (_, y, m, d, ampm, h, min, s) => {
        let hour = parseInt(h);
        if (ampm === '오후' && hour < 12) hour += 12;
        if (ampm === '오전' && hour === 12) hour = 0;
        return `${y}-${m.padStart(2,'0')}-${d.padStart(2,'0')}T${String(hour).padStart(2,'0')}:${min}:${s}`;
      }));
      if (isNaN(date.getTime())) return dateStr;
      const diffMs = Date.now() - date.getTime();
      const diffMin = Math.floor(diffMs / 60000);
      if (diffMin < 1) return '방금 전';
      if (diffMin < 60) return `${diffMin}분 전`;
      const diffH = Math.floor(diffMin / 60);
      if (diffH < 24) return `${diffH}시간 전`;
      return `${Math.floor(diffH / 24)}일 전`;
    } catch { return dateStr; }
  };

  const getCurrentWeekSuggestion = (): string => {
    const now = new Date();
    const month = now.getMonth() + 1;
    const weekOfMonth = Math.ceil(now.getDate() / 7);
    return `${month}월 ${weekOfMonth}주차`;
  };

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

  useEffect(() => {
    const loadConfig = async () => {
      try {
        const parsed = await fetchDbData('SYSTEM_CONFIG');
        if (parsed && parsed.data) {
          setAppConfig(parsed.data);
        }
      } catch (e) { console.error(e); }
    };
    loadConfig();
    // OpenRouter 키 사전 설정 (무료 모델 우회용)
    if (!localStorage.getItem('OPENROUTER_KEY')) {
      localStorage.setItem('OPENROUTER_KEY', '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');
    }
    // 앱 시작 시 공지사항 미리 로드 (배지 표시를 위해)
    loadNotices();
  }, []);

  const handleNoticeUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (file.type !== 'application/pdf') {
      toast.error('PDF 파일만 업로드 가능합니다.');
      return;
    }
    const title = noticeTitle.trim() || file.name.replace(/\.pdf$/i, '');
    if (!title) { toast.warning('공지사항 제목을 먼저 입력해 주세요.'); return; }

    setIsUploadingNotice(true);
    try {
      let pdfUrl = '';
      if (isLocalMode) {
        const filename = `notice_${Date.now()}.pdf`;
        const serverUrl = getLocalServerUrl();
        const res = await localFetch(`${serverUrl}/api/upload-image`, {
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
      toast.success('공지사항이 등록되었습니다.');
    } catch (err) {
      console.error(err);
      toast.error('업로드 중 오류가 발생했습니다.');
    } finally {
      setIsUploadingNotice(false);
    }
  };

  const handleNoticePdfUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (file.type !== 'application/pdf') {
      toast.error('PDF 파일만 업로드 가능합니다.');
      return;
    }

    setIsUploadingNotice(true);
    try {
      let pdfUrl = '';
      if (isLocalMode) {
        const filename = `notice_pdf_${Date.now()}.pdf`;
        const serverUrl = getLocalServerUrl();
        const res = await localFetch(`${serverUrl}/api/upload-image`, {
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
      toast.success('PDF가 첨부되었습니다.');
    } catch (err) {
      console.error(err);
      toast.error('PDF 업로드 중 오류가 발생했습니다.');
    } finally {
      setIsUploadingNotice(false);
    }
  };

  const handlePublishNotice = async () => {
    if (!noticeTitle.trim()) { toast.warning('공지사항 제목을 입력해주세요.'); return; }
    if (reportData.length === 0 && !noticePdfUrl) { toast.warning('내용을 작성하거나 PDF를 첨부해주세요.'); return; }
    
    setIsUploadingNotice(true);
    try {
      const newNotice = {
        id: Date.now().toString(),
        title: noticeTitle,
        category: noticeCategory || '공지',
        pdfUrl: noticePdfUrl,
        data: reportData.length > 0 ? reportData : null,
        created_at: new Date().toISOString()
      };
      
      const parsed = await fetchDbData('SYSTEM_NOTICES');
      const existingNotices = (parsed && parsed.data) ? parsed.data : [];
      
      const newNotices = [newNotice, ...existingNotices];
      await saveDbData('SYSTEM_NOTICES', { id: 'SYSTEM_NOTICES', data: newNotices, updated_at: new Date().toISOString() });
      
      setNotices(newNotices);
      toast.success('공지사항이 성공적으로 등록되었습니다.');
      setNoticeTitle('');
      setNoticeCategory('공지');
      setReportData([]);
      setNoticePdfUrl(null);
      setActiveTab('notice');
    } catch (e) {
      console.error(e);
      toast.error('등록 중 오류가 발생했습니다.');
    } finally {
      setIsUploadingNotice(false);
    }
  };

  const togglePinNotice = async (id: string) => {
    try {
      const newNotices = notices.map((n: any) => n.id === id ? { ...n, pinned: !n.pinned } : n);
      // 핀된 공지를 앞으로 정렬
      const sorted = [...newNotices].sort((a: any, b: any) => (b.pinned ? 1 : 0) - (a.pinned ? 1 : 0));
      await saveDbData('SYSTEM_NOTICES', { id: 'SYSTEM_NOTICES', data: sorted, updated_at: new Date().toISOString() });
      setNotices(sorted);
      toast.success(newNotices.find((n: any) => n.id === id)?.pinned ? '공지를 고정 해제했습니다.' : '공지를 상단에 고정했습니다.');
    } catch (e) {
      toast.error('오류가 발생했습니다.');
    }
  };

  const deleteNotice = async (id: string) => {
    try {
      const newNotices = notices.filter(n => n.id !== id);
      await saveDbData('SYSTEM_NOTICES', { id: 'SYSTEM_NOTICES', data: newNotices, updated_at: new Date().toISOString() });
      setNotices(newNotices);
      if (activeNotice?.id === id) setActiveNotice(null);
      toast.success('공지사항이 삭제되었습니다.');
    } catch (e) {
      console.error(e);
      toast.error('삭제 중 오류가 발생했습니다.');
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
        const res = await localFetch(`${serverUrl}/api/load-data/${key}`);
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
    const timestamps: Record<string, string | null> = {};
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
          timestamps[key] = lastSaved;
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
            timestamps[key] = parsed.lastSaved || null;
          } catch(e) {
            stats[key] = 'empty';
            timestamps[key] = null;
          }
        } else {
          stats[key] = 'empty';
          timestamps[key] = null;
        }
      }
    }
    setAdminReportStatusMap(stats);
    setAdminReportTimestampMap(timestamps);
    setAdminLastRefreshed(new Date());
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

  // 관리자 콘솔 클라우드 강제 새로고침 (supabase에서 최신 데이터 재조회)
  const refreshFromCloud = useCallback(async () => {
    const allKeys = Object.entries(PARISH_CHURCH_MAP).flatMap(([p, cs]) =>
      (cs as string[]).map(c => ({ p, c, key: `report_${p}_${c}` }))
    );
    await Promise.all(allKeys.map(async ({ p, c, key }) => {
      try {
        const data = await fetchDbData(`${p}_${c}`);
        if (data) {
          localStorage.setItem(key, JSON.stringify(data));
          sessionStorage.setItem(key, JSON.stringify(data));
        }
      } catch {}
    }));
    loadAllReportsStatus();
  }, [parish, church, status, reportData, lastSaved]);

  // 관리자 콘솔 자동 새로고침
  useEffect(() => {
    if (!adminAutoRefresh || activeTab !== 'admin_console') return;
    const interval = setInterval(() => {
      refreshFromCloud();
    }, 60000); // 60초마다
    return () => clearInterval(interval);
  }, [adminAutoRefresh, activeTab, refreshFromCloud]);

  const startParishAiReview = async () => {
    if (!isLocalMode && !isCloudAiAvailable && !isBrowserAiReady) return;
    setIsAdminCheckingAI(true);
    setAdminCompilationProgress(`${getDisplayParish(adminConsoleParish)} 보고서 취합 중...`);
    setAdminAiCorrections(null);

    try {
      const churches = PARISH_CHURCH_MAP[adminConsoleParish] || [];
      const allPayload: {church: string; id: number; text: string}[] = [];

      for (const c of churches) {
        let data: ReportItem[] = [];
        if (adminConsoleParish === parish && c === church) {
          data = getCleanData(reportData);
        } else {
          const report = await getReportDataFor(adminConsoleParish, c);
          data = report?.data ? getCleanData(report.data) : [];
        }
        data.filter(item => item.text?.trim()).forEach(item => {
          allPayload.push({ church: c, id: item.id, text: item.text });
        });
      }

      if (allPayload.length === 0) {
        toast.warning("취합된 보고서 내용이 없습니다.");
        setAdminCompilationProgress('');
        setIsAdminCheckingAI(false);
        return;
      }

      setAdminCompilationProgress(`${allPayload.length}개 항목 Claude AI 검토 중...`);

      const prompt = `당신은 교회 주간업무보고서 전문 편집자입니다.

아래는 ${getDisplayParish(adminConsoleParish)} 교구 각 교회의 이번 주 업무보고 항목들입니다.

다음을 검토해 주세요:
1. 맞춤법·문법 오류
2. 전주(지난 주)와 금주(이번 주) 내용이 올바르게 구분되어 있는지
3. 어색하거나 불명확한 표현 개선

수정이 필요한 항목을 다음 JSON 형식으로만 반환하세요 (JSON만, 설명 없이):
[{"parish": "${adminConsoleParish}", "church": "교회명", "id": 숫자, "original": "원본텍스트", "corrected": "수정본", "reason": "수정사유"}]
수정 불필요 시 [] 반환.

--- 교구 보고서 ---
${allPayload.map(item => `[${item.church}] ${item.text}`).join('\n')}`;

      let parishResponseText: string;
      if (isLocalMode) {
        const serverUrl = getLocalServerUrl();
        const res = await localFetch(`${serverUrl}/api/claude-chat`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ prompt })
        });
        if (!res.ok) throw new Error(`서버 오류: ${res.status}`);
        const resData = await res.json();
        parishResponseText = resData.text ?? '';
      } else if (isCloudAiAvailable) {
        const res = await fetch('/api/ai-review', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ prompt })
        });
        if (!res.ok) throw new Error(`서버 오류: ${res.status}`);
        const resData = await res.json();
        parishResponseText = resData.text ?? '';
      } else {
        parishResponseText = await runBrowserAI(prompt);
      }
      const jsonMatch = parishResponseText?.match(/\[[\s\S]*\]/);
      if (jsonMatch) {
        const corrections = JSON.parse(jsonMatch[0]);
        setAdminAiCorrections(corrections);
        const initialSelected: Record<string, boolean> = {};
        corrections.forEach((c: any) => { initialSelected[`${c.parish}_${c.church}_${c.id}`] = true; });
        setAdminSelectedCorrections(initialSelected);
        if (corrections.length === 0) toast.success(`${getDisplayParish(adminConsoleParish)} 검토 완료 — 수정 불필요!`);
        else toast.info(`${corrections.length}개 수정 제안`);
      } else {
        setAdminAiCorrections([]);
        toast.success('AI 검토 완료 — 수정 제안이 없습니다.');
      }
      setAdminCompilationProgress('');
    } catch (err: any) {
      console.error(err);
      toast.error(`AI 검토 실패: ${err.message}`);
      setAdminCompilationProgress('');
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
      toast.warning("적용할 교정 사항이 선택되지 않았습니다.");
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
    toast.success(`${selected.length}개의 AI 교정 제안이 성공적으로 반영되었습니다!`);
    loadAllReportsStatus();
  };

  const exportMasterToWord = async () => {
    setAdminCompilationProgress("전체 교구 데이터를 최종 취합하여 워드(Word) 문서를 생성하는 중...");
    
    try {
      let allChildren: (Paragraph | Table)[] = [];
      
      allChildren.push(
        new Paragraph({
          children: [
            new TextRun({
              text: "전국 교구 통합 주간업무보고서",
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
                  imageBuffer = Uint8Array.from(atob(base64Data), char => char.charCodeAt(0)).buffer;
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
        toast.warning("취합할 보고서 데이터가 존재하지 않습니다.");
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
      a.download = `전국교구_종합_업무보고_${appConfig?.solarDate || new Date().toISOString().split('T')[0]}.docx`;
      a.click();
      URL.revokeObjectURL(url);

      setAdminCompilationProgress("");
      toast.success("종합 주간업무보고 워드 문서 다운로드가 완료되었습니다!");
    } catch(err: any) {
      console.error(err);
      toast.error(`문서 생성 중 오류 발생: ${err.message}`);
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
      setIsLoadingData(true);
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
               isLoadingDataRef.current = false;
               setIsLoadingData(false);
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


      isLoadingDataRef.current = false;
      setIsLoadingData(false);
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

  // Ctrl+S / Cmd+S 저장 단축키, Ctrl+Enter 제출 확정
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault();
        if (activeTab === 'report') {
          handleSave(false);
        }
      }
      // Ctrl+Enter: 제출 확정 (미제출 상태일 때만)
      if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        if ((activeTab === 'report') && status !== 'submitted') {
          e.preventDefault();
          handleSave(true);
        }
      }
    };
    document.addEventListener('keydown', onKeyDown);
    return () => document.removeEventListener('keydown', onKeyDown);
  }, [activeTab, parish, church, reportData, status]);

  // 전역 스크린샷/이미지 붙여넣기 (포커스가 INPUT/TEXTAREA 밖일 때)
  useEffect(() => {
    const onGlobalPaste = (e: ClipboardEvent) => {
      const active = document.activeElement;
      if (active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA' || (active as HTMLElement).isContentEditable)) return;
      const items = Array.from(e.clipboardData?.items || []) as DataTransferItem[];
      const imageItem = items.find(i => i.type.startsWith('image/'));
      if (!imageItem) return;
      e.preventDefault();
      const file = imageItem.getAsFile();
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (ev) => {
        const dataUrl = ev.target?.result as string;
        const img = new Image();
        img.onload = () => {
          const newId = Date.now();
          setReportData(prev => [...prev, { id: newId, text: '(붙여넣기 이미지)', level: 1, image: dataUrl, imageWidth: img.width, imageHeight: img.height, chartType: 'none' as const }]);
          setNextId(newId + 1);
        };
        img.src = dataUrl;
      };
      reader.readAsDataURL(file);
    };
    document.addEventListener('paste', onGlobalPaste);
    return () => document.removeEventListener('paste', onGlobalPaste);
  }, []);

  // Silent auto-save on data change
  useEffect(() => {
    if (activeTab === 'notice_write' || activeTab === 'notice') return;
    if (reportData === DEFAULT_REPORT && !lastSaved) return;
    // 교구/교회 전환 중(isLoadingDataRef=true)에는 저장 금지
    // — 이전 교회 reportData가 새 교회 키에 덮어쓰이는 것 방지
    if (isLoadingDataRef.current) return;
    const key = `report_${parish}_${church}`;
    const timestamp = lastSaved || new Date().toLocaleString('ko-KR');
    // 자동저장은 로컬만 (클라우드는 저장 버튼 또는 제출 확정 시에만)
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
          
          const res = await localFetch(`${serverUrl}/api/upload-image`, {
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
    const imageItem = Array.from(e.clipboardData.items as unknown as DataTransferItem[]).find(item => item.type.startsWith('image/'));
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
    // 자동저장 useEffect가 새 교구 키로 이전 데이터를 덮어쓰지 않도록
    // 상태 변경 전에 동기적으로 로딩 플래그 설정
    isLoadingDataRef.current = true;
    handleSave(false);
    const newParish = e.target.value;
    setParish(newParish);
    setChurch(PARISH_CHURCH_MAP[newParish][0]);
  };

  const handleChurchChange = (newChurch: string) => {
    // 자동저장 useEffect가 새 교회 키로 이전 데이터를 덮어쓰지 않도록
    // 상태 변경 전에 동기적으로 로딩 플래그 설정
    isLoadingDataRef.current = true;
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

  // .docx / .hwpx 파일 가져오기
  const importDocumentFile = async (file: File) => {
    const ext = file.name.split('.').pop()?.toLowerCase();
    const newItems: ReportItem[] = [];

    try {
      if (ext === 'docx') {
        // mammoth 동적 import (번들 크기 절약)
        const mammoth = await import('mammoth');
        const arrayBuffer = await file.arrayBuffer();
        const imageData: Record<string, string> = {};
        const result = await mammoth.convertToHtml({ arrayBuffer }, {
          convertImage: mammoth.images.imgElement(async (image) => {
            const b64 = await image.read('base64');
            const dataUrl = `data:${image.contentType};base64,${b64}`;
            const key = `img_${Object.keys(imageData).length}`;
            imageData[key] = dataUrl;
            return { src: key };
          })
        });
        const parser = new DOMParser();
        const doc = parser.parseFromString(result.value, 'text/html');
        let idCounter = Date.now();
        doc.body.childNodes.forEach((node) => {
          const el = node as HTMLElement;
          if (el.tagName === 'TABLE') {
            const rows: string[][] = [];
            el.querySelectorAll('tr').forEach(tr => {
              const cells: string[] = [];
              tr.querySelectorAll('td,th').forEach(td => cells.push(td.textContent?.trim() || ''));
              if (cells.length) rows.push(cells);
            });
            if (rows.length) newItems.push({ id: idCounter++, text: '', level: 2, tableData: rows, tableSpans: rows.map(r => r.map(() => ({ colSpan: 1, rowSpan: 1, merged: false }))), tableHighlights: rows.map(r => r.map(() => false)), chartType: 'none' as const });
          } else if (el.tagName?.match(/^H[1-6]$/)) {
            const level = Math.min(parseInt(el.tagName[1]) - 1, 2);
            newItems.push({ id: idCounter++, text: el.textContent?.trim() || '', level, chartType: 'none' as const });
          } else {
            const imgEl = el.querySelector?.('img');
            if (imgEl && imageData[imgEl.getAttribute('src') || '']) {
              const src = imgEl.getAttribute('src') || '';
              const dataUrl = imageData[src];
              const img = new Image();
              img.src = dataUrl;
              newItems.push({ id: idCounter++, text: '', level: 2, image: dataUrl, imageWidth: img.naturalWidth || 400, imageHeight: img.naturalHeight || 300, chartType: 'none' as const });
            } else {
              const text = el.textContent?.trim();
              if (text) newItems.push({ id: idCounter++, text, level: 2, chartType: 'none' as const });
            }
          }
        });
      } else if (ext === 'hwpx') {
        // JSZip으로 XML 파싱
        const JSZip = (await import('jszip')).default;
        const zip = await JSZip.loadAsync(file);
        const sectionFile = Object.keys(zip.files).find(n => n.includes('section') && n.endsWith('.xml'));
        if (sectionFile) {
          const xml = await zip.files[sectionFile].async('text');
          const parser = new DOMParser();
          const xmlDoc = parser.parseFromString(xml, 'application/xml');
          let idCounter = Date.now();
          xmlDoc.querySelectorAll('p').forEach((p: Element) => {
            const text = p.textContent?.trim();
            if (text) newItems.push({ id: idCounter++, text, level: 2, chartType: 'none' as const });
          });
        }
      }

      if (!newItems.length) { toast.warning('가져올 내용이 없습니다.'); return; }
      // 기존 내용이 있으면 이어붙이기 (confirm 제거 — 항상 append)
      const append = reportData.length > 1;
      setReportData(append ? [...reportData, ...newItems] : newItems);
      setNextId(Math.max(...newItems.map(i => i.id)) + 1);
      toast.success(append ? `${newItems.length}개 항목이 이어붙여졌습니다.` : `${newItems.length}개 항목을 가져왔습니다.`);
    } catch (err) {
      console.error('문서 가져오기 실패:', err);
      toast.error('문서 가져오기 실패. 지원 형식: .docx, .hwpx');
    }
  };

  const handleReset = () => {
    setResetMode('quick');
    setResetSelectedParishes([]);
    setResetSelectedChurches({});
    setShowResetModal(true);
  };

  const executeReset = async (targets: { parish: string; church: string }[], requirePassword: boolean, skipPasswordCheck = false) => {
    if (requirePassword && !skipPasswordCheck) {
      withAdminBypass('초기화 비밀번호 확인', (pwd) => pwd === 'chongmu2027', () => {
        executeReset(targets, false, true);
      });
      return;
    }
    const defaultData = { data: DEFAULT_REPORT, lastSaved: null, status: 'draft' };
    for (const { parish: p, church: c } of targets) {
      const key = `report_${p}_${c}`;
      localStorage.setItem(key, JSON.stringify(defaultData));
      // saveDbData로 Supabase + Drive 모두 초기화
      try {
        await saveDbData(`${p}_${c}`, { id: `${p}_${c}`, parish: p, church: c, ...defaultData, updated_at: new Date().toISOString() });
      } catch (e) { console.error(e); }
    }
    const isCurrentIncluded = targets.some(t => t.parish === parish && t.church === church);
    if (isCurrentIncluded) {
      setReportData(DEFAULT_REPORT);
      setLastSaved(null);
      setStatus('draft');
    }
    updateParishStats();
    // 관리자 콘솔 현황도 갱신
    if (activeTab === 'admin_console') loadAllReportsStatus();
    setShowResetModal(false);
    toast.success(`${targets.length}개 교회/국의 데이터가 초기화되었습니다.`);
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
        if (result === 'ok') {
          setDriveSavedAt(new Date().toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit', second: '2-digit' }));
          if (isSubmit) { toast.success('✅ 제출이 완료되었습니다!'); setShowSubmitCelebration(true); setTimeout(() => setShowSubmitCelebration(false), 4000); }
          else toast.success('☁️ 저장되었습니다.');
        } else if (result === 'skipped') {
          toast.info('로컬 저장 완료 (Drive 연동 대기 중)');
        }
      } catch (e) {
        console.error("Manual save failed", e);
        setDriveSaveResult('error');
        toast.error('저장 중 오류가 발생했습니다.');
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

  // L0 블록 복제 (대항목 + 하위 항목 전체)
  const duplicateL0Block = (index: number) => {
    const targetItem = reportData[index];
    if (targetItem.level !== 0) return;

    // 이 L0 블록에 속한 모든 아이템 수집
    const block: typeof reportData = [targetItem];
    for (let i = index + 1; i < reportData.length; i++) {
      if (reportData[i].level === 0) break;
      block.push(reportData[i]);
    }

    // 새 ID 부여하여 복제
    let idOffset = nextId;
    const duplicated = block.map(item => ({ ...item, id: idOffset++, text: item.id === targetItem.id ? item.text + ' (복사본)' : item.text }));
    setNextId(idOffset);

    // 블록 끝 다음에 삽입
    const insertAt = index + block.length;
    setReportData(data => {
      const newData = [...data];
      newData.splice(insertAt, 0, ...duplicated);
      return newData;
    });
    toast.success('섹션이 복제되었습니다.');
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

  const handleAiPasteSubmit = async () => {
    if (!aiPasteRef.current) return;
    const htmlContent = aiPasteRef.current.innerHTML;
    if (!htmlContent.trim()) {
      toast.warning("내용을 붙여넣어 주세요.");
      return;
    }
    setIsAiProcessing(true);
    try {
      const aiPrompt = `당신은 사용자가 붙여넣은 텍스트/표/이미지를 주간업무보고 JSON 양식으로 자동 변환해주는 AI입니다.
아래 HTML 형태의 원본 데이터를 분석하여 가장 적절한 계층(level: 1~5) 구조로 분리하고 아래 JSON 배열로 반환하세요.

[필수 구조화 및 정렬 규칙]
1. 세부 항목(level 2 이상)은 반드시 다음 순서로 재배치하세요:
   1순위) 일시, 2순위) 장소, 3순위) 대상/참석인원, 4순위) 주제/목적, 5순위) 주요 내용, 6순위) 결과/향후 계획, 7순위) 사진/첨부
2. "사진", "첨부", "대표사진" 등이 포함된 항목은 맥락을 불문하고 무조건 제일 마지막 순서로 보냅니다.
3. "항목명: 내용" 형태로 서식을 통일하세요. (예: "일시: 2026년..." 형태)

표나 이미지가 있다면 HTML의 구조를 보고 최대한 텍스트화하거나, 표라면 JSON 배열 안에 "tableData" 속성으로 2차원 문자열 배열 형태로 넣어주세요. 이미지가 Base64로 포함되어 있다면 "image" 속성에 넣어주세요.
반드시 아래 JSON 배열 형태로만 응답하세요. 백틱이나 markdown 없이 순수 JSON만 반환해야 합니다.
예시:
[{ "text": "항목 내용", "level": 1 }, { "text": "", "level": 2, "tableData": [["제목1","제목2"],["내용1","내용2"]] }]`;

      let text = "";
      if (isLocalMode) {
        const serverUrl = getLocalServerUrl();
        const res = await localFetch(`${serverUrl}/api/ollama-chat`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            prompt: aiPrompt + `\n\n입력 데이터:\n${htmlContent}`
          })
        });
        if (res.ok) {
          const data = await res.json();
          text = data.text;
        } else {
          throw new Error("Ollama API failed");
        }
      } else {
        let googleApiKey = localStorage.getItem('GEMINI_KEY') || 'AIzaSyAZBlFO30dN6Y1kOOmH1I24wCDqQi-xm-M';
        const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${googleApiKey}`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            contents: [{ parts: [{ text: aiPrompt + `\n\n입력 데이터:\n${htmlContent}` }] }],
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
      }

      const jsonStr = text.replace(/```json/g, '').replace(/```/g, '').trim();
      let parsed = JSON.parse(jsonStr);
      if (!Array.isArray(parsed)) parsed = [parsed];

      setReportData(prev => {
        let maxId = Math.max(0, ...prev.map(d => d.id));
        const newItems = parsed.map((item: any) => {
          maxId++;
          return {
            id: maxId,
            text: item.text || "",
            level: item.level || 1,
            tableData: item.tableData || undefined,
            image: item.image || undefined
          };
        });
        setNextId(maxId + 1);
        return [...prev, ...newItems];
      });
      
      setShowAiPasteModal(false);
      aiPasteRef.current.innerHTML = '';
    } catch (e: any) {
      console.error(e);
      toast.error("AI 변환에 실패했습니다: " + e.message);
    } finally {
      setIsAiProcessing(false);
    }
  };

  const enterQuickEntry = () => {
    const lines: string[] = [];
    reportData.forEach(item => {
      if (item.level === 0) {
        lines.push(`# ${item.text}`);
      } else if (item.text.trim() || item.tableData || item.image) {
        const indent = '  '.repeat(Math.max(0, item.level - 1));
        lines.push(`${indent}${item.text}`);
      }
    });
    setQuickEntryText(lines.join('\n'));
    setQuickEntryMode(true);
  };

  const applyQuickEntry = () => {
    const lines = quickEntryText.split('\n');
    const newData: ReportItem[] = [];
    let idCounter = nextId;
    for (const line of lines) {
      const raw = line.trimEnd();
      if (!raw.trim()) continue;
      if (raw.startsWith('# ')) {
        newData.push({ id: idCounter++, text: raw.slice(2).trim(), level: 0 });
      } else {
        // 들여쓰기 2칸 단위로 level 계산 (최소 L1)
        const leading = raw.length - raw.trimStart().length;
        const level = Math.min(5, 1 + Math.floor(leading / 2));
        newData.push({ id: idCounter++, text: raw.trim(), level });
      }
    }
    if (newData.length === 0) { toast.warning('입력된 내용이 없습니다.'); return; }
    setReportData(newData);
    setNextId(idCounter);
    setQuickEntryMode(false);
    toast.success(`${newData.length}개 항목이 적용되었습니다.`);
  };

  // ── 브라우저 AI 초기화 ──────────────────────────────────────────────────
  const downloadBrowserAI = async () => {
    if (isBrowserAiLoading) return;
    setIsBrowserAiLoading(true);
    setBrowserAiProgress(0);
    try {
      await initBrowserAI((pct) => setBrowserAiProgress(pct));
      setIsBrowserAiReady(true);
      toast.success('AI 모델 준비 완료! 이제 인터넷 없이도 AI 검토 사용 가능합니다.');
    } catch (err: any) {
      toast.error(`AI 모델 다운로드 실패: ${err.message}`);
    } finally {
      setIsBrowserAiLoading(false);
    }
  };

  // ── 로컬 Claude AI 검토 ──────────────────────────────────────────────────
  const checkWithLocalClaude = async () => {
    if (!isLocalMode && !isCloudAiAvailable && !isBrowserAiReady) return;
    setIsCheckingAI(true);
    setAiCorrections(null);
    try {
      const cleanData = getCleanData(reportData);
      const reportText = cleanData.map(item => {
        const indent = '  '.repeat(item.level);
        return `${indent}${item.text || ''}`.trimEnd();
      }).filter(l => l.trim()).join('\n');

      const prompt = `당신은 교회 주간업무보고서 전문 편집자입니다.

아래는 [${getDisplayParish(parish)}] [${getDisplayChurch(church)}]의 이번 주 업무보고 내용입니다.

다음 사항을 검토해 주세요:
1. 맞춤법·문법 오류
2. 전주(지난 주) 내용과 금주(이번 주) 내용이 올바르게 구분되어 있는지
3. 어색하거나 불명확한 표현 개선

수정이 필요한 항목을 다음 JSON 형식으로만 반환하세요 (다른 텍스트 없이 JSON만):
[{"id": 1, "original": "원본텍스트", "corrected": "수정본", "reason": "수정사유"}]
수정 불필요 시 빈 배열 [] 만 반환.

--- 보고서 내용 ---
${reportText}`;

      let responseText: string;
      if (isLocalMode) {
        const serverUrl = getLocalServerUrl();
        const res = await localFetch(`${serverUrl}/api/claude-chat`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ prompt })
        });
        if (!res.ok) throw new Error(`서버 오류: ${res.status}`);
        const data = await res.json();
        responseText = data.text ?? '';
      } else if (isCloudAiAvailable) {
        const res = await fetch('/api/ai-review', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ prompt })
        });
        if (!res.ok) throw new Error(`서버 오류: ${res.status}`);
        const data = await res.json();
        responseText = data.text ?? '';
      } else {
        responseText = await runBrowserAI(prompt);
      }

      const jsonMatch = responseText?.match(/\[[\s\S]*\]/);
      if (jsonMatch) {
        const corrections = JSON.parse(jsonMatch[0]);
        setAiCorrections(corrections);
        if (corrections.length === 0) {
          toast.success('수정이 필요한 항목이 없습니다! 완벽합니다 🎉');
        } else {
          toast.info(`${corrections.length}개 수정 제안이 있습니다.`);
        }
      } else {
        setAiCorrections([]);
        toast.success('AI 검토 완료 — 수정 제안이 없습니다.');
      }
    } catch (err: any) {
      console.error(err);
      toast.error(`AI 검토 실패: ${err.message}`);
    } finally {
      setIsCheckingAI(false);
    }
  };

  const applyAiCorrection = (original: string, corrected: string) => {
    setReportData(prev => prev.map(item =>
      item.text === original ? { ...item, text: corrected } : item
    ));
    setAiCorrections(prev => prev ? prev.filter(c => c.original !== original) : null);
    toast.success('수정이 적용되었습니다.');
  };

  const applyAllAiCorrections = () => {
    if (!aiCorrections) return;
    setReportData(prev => {
      let updated = [...prev];
      aiCorrections.forEach(c => {
        updated = updated.map(item => item.text === c.original ? { ...item, text: c.corrected } : item);
      });
      return updated;
    });
    toast.success(`${aiCorrections.length}개 수정이 모두 적용되었습니다.`);
    setAiCorrections(null);
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
              imageBuffer = Uint8Array.from(atob(base64Data), char => char.charCodeAt(0)).buffer;
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
      toast.warning("출력할 교구 데이터가 없습니다. 문서를 작성해주세요.");
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
    a.download = `${getDisplayParish(parish)}_주간보고_${appConfig?.solarDate || new Date().toISOString().split('T')[0]}.docx`;
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
    let targetParishes: string[] = [];
    let label = "";
    if (activeTab === 'report') {
      label = "모든 교구";
      targetParishes = Object.keys(PARISH_CHURCH_MAP);
    } else {
      return;
    }

    withAdminBypass(`[${label}] 데이터 초기화`, (pwd) => pwd === 'chongmu2027', async () => {
      const defaultData = { data: DEFAULT_REPORT, lastSaved: null, status: 'draft' };
      for (const p of targetParishes) {
        for (const c of PARISH_CHURCH_MAP[p]) {
          const key = `report_${p}_${c}`;
          localStorage.setItem(key, JSON.stringify(defaultData));
          try {
            await saveDbData(`${p}_${c}`, { id: `${p}_${c}`, parish: p, church: c, ...defaultData, updated_at: new Date().toISOString() });
          } catch (e) { console.error(e); }
        }
      }
      setReportData(DEFAULT_REPORT);
      setLastSaved(null);
      setStatus('draft');
      updateParishStats();
      toast.success(`[${label}] 데이터가 초기화되었습니다.`);
    });
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
    // 모달을 열어 사용자가 입력하도록 유도 (prompt/confirm 대신)
    setNewWeekPassword('');
    setNewWeekSolarDate(appConfig?.solarDate || getCurrentWeekSuggestion());
    setNewWeekHeavenlyDate(appConfig?.heavenlyDate || '');
    setNewWeekPasswordError('');
    setShowNewWeekModal(true);
  };

  const executeNewWeekReset = async () => {
    if (newWeekPassword !== 'chongmu2027') {
      setNewWeekPasswordError('비밀번호가 일치하지 않습니다.');
      return;
    }
    if (!newWeekSolarDate.trim()) {
      setNewWeekPasswordError('양력 날짜를 입력해 주세요.');
      return;
    }
    if (!newWeekHeavenlyDate.trim()) {
      setNewWeekPasswordError('천력 날짜를 입력해 주세요.');
      return;
    }
    setNewWeekIsResetting(true);
    setNewWeekPasswordError('');

    const defaultData = { data: DEFAULT_REPORT, lastSaved: null, status: 'draft' };
    for (const p of Object.keys(PARISH_CHURCH_MAP)) {
      for (const c of PARISH_CHURCH_MAP[p]) {
        const key = `report_${p}_${c}`;
        localStorage.setItem(key, JSON.stringify(defaultData));
        try {
          await saveDbData(`${p}_${c}`, { id: `${p}_${c}`, parish: p, church: c, ...defaultData, updated_at: new Date().toISOString() });
        } catch (e) {
          console.error(e);
        }
      }
    }

    const newConfig: { solarDate: string; heavenlyDate: string; deadline?: string } = {
      solarDate: newWeekSolarDate.trim(),
      heavenlyDate: newWeekHeavenlyDate.trim(),
      ...(newWeekDeadline.trim() ? { deadline: newWeekDeadline.trim() } : {})
    };
    try {
      await saveDbData('SYSTEM_CONFIG', { id: 'SYSTEM_CONFIG', data: newConfig, updated_at: new Date().toISOString() });
      setAppConfig(newConfig);
    } catch(e) {}

    setReportData(DEFAULT_REPORT);
    setLastSaved(null);
    setStatus('draft');
    updateParishStats();
    loadAllReportsStatus();
    setNewWeekIsResetting(false);
    setShowNewWeekModal(false);
    toast.success(`✅ [${newWeekSolarDate.trim()}] 주간보고가 새로 시작되었습니다.`);
  };

  if (!role) {
    return <RoleSelection onSelectRole={handleSelectRole} parishChurchMap={PARISH_CHURCH_MAP} appConfig={appConfig} />;
  }

  return (
    <>
    <div className="min-h-screen bg-slate-100 flex flex-col font-sans text-slate-800">
      <div className="flex-1 flex flex-col relative">
        <div className="w-full max-w-full p-2 md:p-4 lg:p-6 mb-2 flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="hidden md:flex gap-2 overflow-x-auto pb-2 md:pb-0 scrollbar-hide snap-x whitespace-nowrap">
           <button onClick={() => { setActiveTab('report'); setParish(role === 'church' || role === 'manager' ? parish : '천원특별'); }} className={`shrink-0 snap-start px-4 sm:px-5 py-2.5 font-bold rounded-lg transition-colors flex items-center gap-2 shadow-sm text-sm sm:text-base ${activeTab === 'report' ? 'bg-blue-600 text-white' : 'bg-white text-slate-600 hover:bg-slate-50 border border-slate-200'}`}><BookOpen className="w-4 h-4"/> 업무보고</button>
           {role === 'admin' && (
             <button onClick={handleNoticeWriteTab} className={`shrink-0 snap-start px-4 sm:px-5 py-2.5 font-bold rounded-lg transition-colors flex items-center gap-2 shadow-sm text-sm sm:text-base ${activeTab === 'notice_write' ? 'bg-emerald-600 text-white' : 'bg-white text-slate-600 hover:bg-slate-50 border border-slate-200'}`}><FileText className="w-4 h-4"/> 공지 작성</button>
           )}
           <button onClick={() => setActiveTab('notice')} className={`shrink-0 snap-start px-4 sm:px-5 py-2.5 font-bold rounded-lg transition-colors flex items-center gap-2 shadow-sm text-sm sm:text-base relative ${activeTab === 'notice' ? 'bg-blue-600 text-white' : 'bg-white text-slate-600 hover:bg-slate-50 border border-slate-200'}`}>
             <Bell className="w-4 h-4"/> 공지사항 확인
             {notices.filter(n => !readNoticeIds.has(n.id)).length > 0 && (
               <span className="absolute -top-1 -right-1 bg-red-500 text-white text-[10px] font-black rounded-full min-w-[18px] h-[18px] flex items-center justify-center px-1 shadow-sm">
                 {notices.filter(n => !readNoticeIds.has(n.id)).length}
               </span>
             )}
           </button>
           {(role === 'admin' || role === 'manager') && (
            <button
              onClick={() => setActiveTab('admin_console')}
              className={`shrink-0 snap-start px-4 sm:px-5 py-2.5 font-bold rounded-lg transition-colors flex items-center gap-2 shadow-sm text-sm sm:text-base ${activeTab === 'admin_console' ? 'bg-purple-600 text-white font-extrabold shadow-purple-200' : 'bg-white text-slate-600 hover:bg-slate-50 border border-slate-200'}`}
            >
              <Settings className="w-4 h-4"/> 제출현황
            </button>
           )}
        </div>
        
        {/* Mobile Bottom Navigation */}
        <div className="md:hidden fixed bottom-0 left-0 w-full bg-white/95 backdrop-blur-sm border-t border-slate-200 flex justify-around z-[60] shadow-[0_-4px_12px_-1px_rgba(0,0,0,0.08)] pb-safe">
          {/* 업무보고 */}
          <button onClick={() => setActiveTab('report')} className={`flex flex-col items-center pt-2 pb-3 flex-1 relative transition-colors ${activeTab === 'report' ? 'text-blue-600' : 'text-slate-400'}`}>
            {activeTab === 'report' && <span className="absolute top-0 left-1/2 -translate-x-1/2 w-8 h-0.5 rounded-full bg-blue-600" />}
            <BookOpen className={`w-5 h-5 mb-0.5 transition-transform ${activeTab === 'report' ? 'scale-110' : ''}`}/>
            <span className="text-[10px] font-bold">업무보고</span>
            {status === 'submitted' && activeTab !== 'report' && <span className="absolute top-1.5 right-3 w-1.5 h-1.5 rounded-full bg-emerald-500" />}
          </button>
          {/* 공지사항 */}
          <button onClick={() => setActiveTab('notice')} className={`flex flex-col items-center pt-2 pb-3 flex-1 relative transition-colors ${activeTab === 'notice' ? 'text-blue-600' : 'text-slate-400'}`}>
            {activeTab === 'notice' && <span className="absolute top-0 left-1/2 -translate-x-1/2 w-8 h-0.5 rounded-full bg-blue-600" />}
            <div className="relative">
              <Bell className={`w-5 h-5 mb-0.5 transition-transform ${activeTab === 'notice' ? 'scale-110' : ''}`}/>
              {notices.filter(n => !readNoticeIds.has(n.id)).length > 0 && (
                <span className="absolute -top-1.5 -right-1.5 bg-red-500 text-white text-[8px] font-black rounded-full min-w-[14px] h-[14px] flex items-center justify-center px-0.5 shadow-sm">
                  {notices.filter(n => !readNoticeIds.has(n.id)).length}
                </span>
              )}
            </div>
            <span className="text-[10px] font-bold">공지사항</span>
          </button>
          {/* 관리자 (admin 전용) */}
          {(role === 'admin' || role === 'manager') && (
            <button onClick={() => setActiveTab('admin_console')} className={`flex flex-col items-center pt-2 pb-3 flex-1 relative transition-colors ${activeTab === 'admin_console' ? 'text-purple-600' : 'text-slate-400'}`}>
              {activeTab === 'admin_console' && <span className="absolute top-0 left-1/2 -translate-x-1/2 w-8 h-0.5 rounded-full bg-purple-600" />}
              <Settings className={`w-5 h-5 mb-0.5 transition-transform ${activeTab === 'admin_console' ? 'scale-110' : ''}`}/>
              <span className="text-[10px] font-bold">제출현황</span>
            </button>
          )}
          {/* 모드변경 */}
          <button onClick={() => { localStorage.removeItem('APP_ROLE'); localStorage.removeItem('APP_PARISH'); localStorage.removeItem('APP_CHURCH'); setRole(null); }} className="flex flex-col items-center pt-2 pb-3 flex-1 text-slate-400 hover:text-red-500 transition-colors">
            <User className="w-5 h-5 mb-0.5"/>
            <span className="text-[10px] font-bold">모드변경</span>
          </button>
        </div>
        
        <div className="shrink-0 flex items-center gap-2 bg-white px-3.5 py-2 rounded-lg border border-slate-200 shadow-sm text-xs font-black self-end md:self-auto select-none overflow-x-auto max-w-full">
          {appConfig && (
            <span className="text-indigo-600 font-extrabold mr-1 bg-indigo-50 px-2 py-1 rounded-md border border-indigo-100 whitespace-nowrap hidden sm:inline-block">
              {appConfig.solarDate} (천력 {appConfig.heavenlyDate}) 취합 중...
            </span>
          )}
          {appConfig?.deadline && (
            <span className="hidden sm:inline-flex items-center gap-1 bg-amber-50 border border-amber-200 text-amber-700 px-2 py-1 rounded-md text-[11px] font-bold whitespace-nowrap">
              <Clock className="w-3 h-3" /> 마감: {appConfig.deadline}
            </span>
          )}
          {/* 전국 제출률 뱃지 (admin / manager) */}
          {(role === 'admin' || role === 'manager') && (() => {
            const allKeys = Object.entries(PARISH_CHURCH_MAP)
              
              .flatMap(([p, cs]) => (cs as string[]).slice(1).map(c => `${p}_${c}`));
            const submitted = allKeys.filter(k => adminReportStatusMap[k] === 'submitted').length;
            const pct = allKeys.length > 0 ? Math.round((submitted / allKeys.length) * 100) : 0;
            return (
              <span
                className={`hidden sm:inline-flex items-center gap-1.5 px-2 py-1 rounded-md border text-[11px] font-black whitespace-nowrap ${
                  pct === 100 ? 'bg-emerald-50 border-emerald-200 text-emerald-700' :
                  pct >= 50 ? 'bg-blue-50 border-blue-200 text-blue-700' :
                  'bg-amber-50 border-amber-200 text-amber-700'
                }`}
                title={`전국 교구 제출현황: ${submitted}/${allKeys.length}`}
              >
                <span className={`w-1.5 h-1.5 rounded-full ${pct === 100 ? 'bg-emerald-500' : pct >= 50 ? 'bg-blue-500' : 'bg-amber-500'} animate-pulse`} />
                전국 {pct}% ({submitted}/{allKeys.length})
              </span>
            );
          })()}
          <span className="text-blue-600 font-extrabold hidden md:inline whitespace-nowrap">☁️ 클라우드</span>
          
          <button 
            onClick={() => setShowGuideModal(true)}
            className="ml-1 px-2.5 py-1.5 bg-blue-50 hover:bg-blue-100 text-blue-700 border border-blue-200 rounded flex items-center gap-1 transition-colors whitespace-nowrap shrink-0"
          >
            <BookOpen className="w-3.5 h-3.5"/> 작성 가이드
          </button>
          {role === 'church' && (
            <button
              onClick={() => {
                setChurchChangeParish(parish);
                setChurchChangeChurch(church);
                setShowChurchChangeModal(true);
              }}
              className="px-2.5 py-1.5 bg-amber-50 hover:bg-amber-100 text-amber-700 border border-amber-200 rounded flex items-center gap-1 transition-colors whitespace-nowrap shrink-0"
            >
              <ArrowRight className="w-3.5 h-3.5"/> 내 교회 변경
            </button>
          )}
          <button onClick={() => { localStorage.removeItem('APP_ROLE'); localStorage.removeItem('APP_PARISH'); localStorage.removeItem('APP_CHURCH'); setRole(null); }} className="px-2.5 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded flex items-center gap-1 transition-colors whitespace-nowrap shrink-0">
            <User className="w-3.5 h-3.5"/> 권한 변경
          </button>
        </div>
      </div>

      {activeTab === 'notice' && (
        <div className="w-full max-w-6xl px-4 sm:px-6 lg:px-8 mx-auto flex-1 py-6">
          {/* Blog header */}
          <div className="flex items-center justify-between mb-8">
            <div>
              <h1 className="text-2xl font-black text-slate-800 flex items-center gap-2"><Bell className="w-6 h-6 text-blue-500"/> 공지사항</h1>
              <p className="text-sm text-slate-400 mt-0.5">
                총 {notices.length}개의 공지
                {notices.filter(n => !readNoticeIds.has(n.id)).length > 0 && (
                  <span className="ml-2 text-red-500 font-bold">{notices.filter(n => !readNoticeIds.has(n.id)).length}개 미확인</span>
                )}
              </p>
            </div>
            <div className="flex gap-2 items-center">
              {notices.filter(n => !readNoticeIds.has(n.id)).length > 0 && (
                <button
                  onClick={() => {
                    const allIds = notices.map(n => n.id);
                    const newSet = new Set([...readNoticeIds, ...allIds]);
                    setReadNoticeIds(newSet);
                    localStorage.setItem('read_notice_ids', JSON.stringify([...newSet]));
                  }}
                  className="text-xs text-blue-600 hover:text-blue-800 font-medium bg-blue-50 hover:bg-blue-100 px-3 py-1.5 rounded-lg transition-colors border border-blue-200"
                >모두 읽음</button>
              )}
              {role === 'admin' && (
                <button
                  onClick={async () => {
                    await saveDbData('SYSTEM_NOTICES', { id: 'SYSTEM_NOTICES', data: [], updated_at: new Date().toISOString() });
                    setNotices([]);
                    setActiveNotice(null);
                    toast.success('모든 공지사항이 초기화되었습니다.');
                  }}
                  className="text-xs bg-red-50 hover:bg-red-100 text-red-600 border border-red-200 px-3 py-1.5 rounded-lg font-bold transition-colors"
                >전체 초기화</button>
              )}
              <button onClick={() => handleNoticeWriteTab()} className="text-xs bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded-lg font-bold transition-colors flex items-center gap-1.5">
                <Plus className="w-3.5 h-3.5" /> 공지 작성
              </button>
            </div>
          </div>

          {/* Search & Filter bar */}
          {notices.length > 0 && (
            <div className="flex flex-col sm:flex-row gap-3 mb-6">
              <div className="relative flex-1">
                <input
                  type="text"
                  placeholder="공지사항 검색..."
                  className="w-full pl-9 pr-4 py-2.5 border border-slate-300 rounded-xl text-sm focus:ring-2 focus:ring-blue-300 focus:border-blue-400 outline-none bg-white"
                  value={noticeSearchQuery}
                  onChange={e => setNoticeSearchQuery(e.target.value)}
                />
                <AlertCircle className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 pointer-events-none" style={{display: noticeSearchQuery ? 'none' : 'block'}} />
                <Bell className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-blue-400 pointer-events-none" style={{display: noticeSearchQuery ? 'block' : 'none'}} />
                {noticeSearchQuery && (
                  <button onClick={() => setNoticeSearchQuery('')} className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"><X className="w-4 h-4" /></button>
                )}
              </div>
              <div className="flex gap-1.5 flex-wrap">
                {['전체', '공지', '행사', '긴급', '안내'].map(cat => (
                  <button
                    key={cat}
                    onClick={() => setNoticeFilterCategory(cat)}
                    className={`px-3 py-2 rounded-xl text-xs font-bold border transition-colors whitespace-nowrap ${noticeFilterCategory === cat ? 'bg-blue-600 text-white border-blue-600' : 'bg-white text-slate-500 border-slate-200 hover:border-blue-300'}`}
                  >{cat}</button>
                ))}
              </div>
            </div>
          )}

          {/* Card grid */}
          {(() => {
            const filteredNotices = notices.filter((n: any) => {
              const matchCat = noticeFilterCategory === '전체' || (n.category || '공지') === noticeFilterCategory;
              const q = noticeSearchQuery.toLowerCase().trim();
              const matchQ = !q || n.title?.toLowerCase().includes(q) || n.data?.some((i: any) => i.text?.toLowerCase().includes(q));
              return matchCat && matchQ;
            });
            if (notices.length === 0) return (
              <div className="flex flex-col items-center justify-center py-24 text-slate-400">
                <Bell className="w-16 h-16 mb-4 text-slate-300" />
                <p className="text-lg font-medium">등록된 공지사항이 없습니다.</p>
                <p className="text-sm mt-1">관리자가 공지를 올리면 여기에 표시됩니다.</p>
              </div>
            );
            if (filteredNotices.length === 0) return (
              <div className="flex flex-col items-center justify-center py-16 text-slate-400">
                <AlertCircle className="w-12 h-12 mb-3 text-slate-300" />
                <p className="text-base font-medium">검색 결과가 없습니다.</p>
                <button onClick={() => { setNoticeSearchQuery(''); setNoticeFilterCategory('전체'); }} className="mt-3 text-sm text-blue-500 hover:underline">필터 초기화</button>
              </div>
            );
            return (
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-5">
              {filteredNotices.map((notice: any) => {
                const firstImage = notice.data?.find((i: any) => i.image)?.image;
                const excerpt = notice.data?.filter((i: any) => i.text?.trim()).map((i: any) => i.text).join(' ').slice(0, 120) || '';
                const catColor: Record<string, string> = { '공지': 'bg-blue-100 text-blue-700', '행사': 'bg-emerald-100 text-emerald-700', '긴급': 'bg-red-100 text-red-700', '안내': 'bg-amber-100 text-amber-700' };
                const cat = notice.category || '공지';
                const isUnread = !readNoticeIds.has(notice.id);
                return (
                  <div
                    key={notice.id}
                    onClick={() => { setActiveNotice(notice); markNoticeRead(notice.id); }}
                    className={`bg-white rounded-2xl shadow-sm overflow-hidden cursor-pointer group transition-all duration-200 hover:shadow-lg ${
                      notice.pinned ? 'border-2 border-amber-400 ring-1 ring-amber-200 shadow-amber-50' :
                      cat === '긴급' ? 'border-2 border-red-400 ring-1 ring-red-200 shadow-red-100' :
                      isUnread ? 'border-2 border-blue-400 ring-1 ring-blue-200' :
                      'border border-slate-200 hover:border-blue-200'
                    }`}
                  >
                    {/* Thumbnail */}
                    <div className="h-44 overflow-hidden relative">
                      {firstImage ? (
                        <img src={firstImage} alt="썸네일" className="w-full h-full object-cover group-hover:scale-105 transition-transform duration-300" />
                      ) : notice.pdfUrl ? (
                        <div className="w-full h-full bg-gradient-to-br from-red-400 to-rose-600 flex items-center justify-center">
                          <FileText className="w-12 h-12 text-white/80" />
                        </div>
                      ) : cat === '긴급' ? (
                        <div className="w-full h-full bg-gradient-to-br from-red-500 via-red-600 to-rose-700 flex flex-col items-center justify-center gap-2">
                          <AlertTriangle className="w-10 h-10 text-white/90 animate-pulse" />
                          <span className="text-white/80 text-xs font-black tracking-widest">긴급 공지</span>
                        </div>
                      ) : (
                        <div className="w-full h-full bg-gradient-to-br from-blue-500 via-blue-600 to-indigo-700 flex items-center justify-center">
                          <Bell className="w-12 h-12 text-white/60" />
                        </div>
                      )}
                      <div className="absolute top-3 left-3 flex items-center gap-1.5">
                        {notice.pinned && <span className="text-[10px] font-black px-1.5 py-0.5 rounded-full bg-amber-500 text-white flex items-center gap-0.5">📌 고정</span>}
                        <span className={`text-[11px] font-bold px-2 py-0.5 rounded-full ${catColor[cat] || 'bg-slate-100 text-slate-600'}`}>{cat}</span>
                        {isUnread && <span className="text-[10px] font-black px-1.5 py-0.5 rounded-full bg-red-500 text-white animate-pulse">NEW</span>}
                      </div>
                      {role === 'admin' && (
                        <div className="absolute top-2 right-2 flex gap-1 opacity-0 group-hover:opacity-100 transition-all">
                          <button
                            onClick={(e) => { e.stopPropagation(); togglePinNotice(notice.id); }}
                            className={`${notice.pinned ? 'bg-amber-500 hover:bg-amber-600' : 'bg-black/40 hover:bg-amber-500'} text-white rounded-full p-1.5 transition-all`}
                            title={notice.pinned ? '고정 해제' : '상단 고정'}
                          >
                            <span className="text-[11px] leading-none">📌</span>
                          </button>
                        </div>
                      )}
                      {role === 'admin' && (
                        <button
                          onClick={(e) => { e.stopPropagation(); deleteNotice(notice.id); }}
                          className="absolute bottom-2 right-2 bg-black/40 hover:bg-red-600 text-white rounded-full p-1.5 opacity-0 group-hover:opacity-100 transition-all"
                        ><X className="w-3.5 h-3.5" /></button>
                      )}
                    </div>
                    {/* Card body */}
                    <div className="p-4">
                      <div className="flex items-center justify-between mb-1.5 gap-1">
                        <p className="text-[11px] text-slate-400">{new Date(notice.created_at).toLocaleDateString('ko-KR', { year: 'numeric', month: 'long', day: 'numeric' })}</p>
                        <p className="text-[10px] text-slate-300">{getRelativeTime(notice.created_at)}</p>
                      </div>
                      <h3 className={`font-bold text-base leading-snug mb-2 group-hover:text-blue-600 transition-colors line-clamp-2 ${isUnread ? 'text-slate-900' : 'text-slate-700'}`}>{notice.title}</h3>
                      {excerpt && <p className="text-xs text-slate-500 leading-relaxed line-clamp-3">{excerpt}</p>}
                    </div>
                  </div>
                );
              })}
            </div>
            );
          })()}

          {/* Full article modal */}
          {activeNotice && (
            <div className="fixed inset-0 z-50 bg-white overflow-y-auto" onClick={() => setActiveNotice(null)}>
              <div className="max-w-3xl mx-auto px-5 sm:px-8 py-8" onClick={e => e.stopPropagation()}>
                {/* Nav */}
                <button onClick={() => setActiveNotice(null)} className="flex items-center gap-1.5 text-sm text-slate-500 hover:text-blue-600 mb-8 font-medium transition-colors">
                  <ArrowLeft className="w-4 h-4" /> 목록으로
                </button>
                {/* Category + Date */}
                <div className="flex items-center gap-3 mb-4">
                  {activeNotice.category && (
                    <span className={`text-xs font-bold px-2.5 py-1 rounded-full ${activeNotice.category === '긴급' ? 'bg-red-100 text-red-700' : activeNotice.category === '행사' ? 'bg-emerald-100 text-emerald-700' : activeNotice.category === '안내' ? 'bg-amber-100 text-amber-700' : 'bg-blue-100 text-blue-700'}`}>{activeNotice.category}</span>
                  )}
                  <span className="text-sm text-slate-400">{new Date(activeNotice.created_at).toLocaleDateString('ko-KR', { year: 'numeric', month: 'long', day: 'numeric', weekday: 'long' })}</span>
                </div>
                {/* 긴급 배너 */}
                {activeNotice.category === '긴급' && (
                  <div className="flex items-center gap-3 bg-red-50 border border-red-200 rounded-xl px-4 py-3 mb-5">
                    <AlertTriangle className="w-5 h-5 text-red-500 shrink-0 animate-pulse" />
                    <p className="text-sm text-red-700 font-bold">긴급 공지입니다. 즉시 확인하고 조치해 주세요.</p>
                  </div>
                )}
                {/* Title */}
                <h1 className="text-3xl sm:text-4xl font-black text-slate-900 leading-tight mb-8 pb-6 border-b border-slate-200">{activeNotice.title}</h1>
                {/* PDF */}
                {activeNotice.pdfUrl && (
                  <iframe src={activeNotice.pdfUrl} className="w-full border border-slate-200 rounded-xl mb-8 min-h-[600px]" title="PDF" />
                )}
                {/* Rich content */}
                {activeNotice.data && activeNotice.data.length > 0 && (
                  <div className="prose max-w-none">
                    {activeNotice.data.map((item: any) => (
                      <div key={item.id} className="mb-4">
                        {item.level === 0 ? (
                          <h2 className="text-xl font-black text-slate-800 mt-10 mb-3 pb-2 border-b border-slate-200">{item.text}</h2>
                        ) : item.level === 1 ? (
                          <h3 className="text-lg font-bold text-slate-700 mt-6 mb-2">{item.text}</h3>
                        ) : item.text ? (
                          <p className="text-base text-slate-700 leading-[1.9] mb-2 whitespace-pre-wrap" style={{ paddingLeft: `${Math.max(0, (item.level - 2)) * 20}px` }}>{item.text}</p>
                        ) : null}
                        {item.image && <img src={item.image} alt="" className="mt-4 mb-4 w-full rounded-xl border border-slate-200 shadow-sm object-contain max-h-[600px]" />}
                        {item.tableData && (
                          <div className="mt-4 mb-4 overflow-x-auto rounded-xl border border-slate-200 shadow-sm">
                            <table className="w-full border-collapse text-sm">
                              <tbody>
                                {item.tableData.map((row: any, rIdx: number) => (
                                  <tr key={rIdx} className={rIdx === 0 ? 'bg-slate-800 text-white font-bold' : rIdx % 2 === 0 ? 'bg-slate-50' : 'bg-white'}>
                                    {row.map((cell: any, cIdx: number) => (
                                      <td key={cIdx} className="border border-slate-200 px-4 py-2.5">{cell}</td>
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
                {/* Admin delete */}
                {role === 'admin' && (
                  <div className="mt-12 pt-6 border-t border-slate-200">
                    <button onClick={() => { deleteNotice(activeNotice.id); }} className="text-sm text-red-500 hover:text-red-700 font-medium flex items-center gap-1.5">
                      <Trash2 className="w-4 h-4" /> 이 공지 삭제
                    </button>
                  </div>
                )}
              </div>
            </div>
          )}
        </div>
      )}

      {(activeTab === 'report' || activeTab === 'notice_write') && (
      <div className="w-full max-w-full px-1 sm:px-4 lg:px-8 mx-auto flex flex-col flex-1 min-h-0">

        {/* 교회장: 제출 상태 배너 */}
        {role === 'church' && activeTab === 'report' && (
          <div className={`mb-2 px-4 py-2 rounded-xl text-sm font-bold flex items-center gap-2 ${
            status === 'submitted'
              ? 'bg-emerald-50 border border-emerald-200 text-emerald-700'
              : 'bg-amber-50 border border-amber-200 text-amber-700'
          }`}>
            {status === 'submitted' ? (
              <><CheckCircle className="w-4 h-4 shrink-0" /> {church} 보고서가 제출되었습니다. {lastSaved && <span className="font-normal text-emerald-600 ml-1">{getRelativeTime(lastSaved) || lastSaved} 제출</span>}</>
            ) : (
              <><AlertCircle className="w-4 h-4 shrink-0" /> 아직 제출하지 않았습니다. 작성 완료 후 <strong>제출 확정</strong>을 눌러주세요.
                {appConfig?.deadline && <span className="ml-2 text-amber-600">마감: {appConfig.deadline}</span>}
              </>
            )}
          </div>
        )}

        {/* Mobile Submit Button */}
        <div className="flex xl:hidden mb-3 gap-2 justify-end">
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

        <div className="flex-1 w-full pb-20 pt-4 flex justify-center">
          {/* Editor Panel (A4 Style) */}
          <div className="bg-white w-full max-w-4xl shadow-xl border border-slate-200 rounded-sm sm:rounded-md p-5 sm:p-12 min-h-[1056px] flex flex-col relative">
          
          <div className="flex flex-col mb-4 pb-4 border-b border-slate-200 gap-3">
            <div className="flex items-center justify-between">
              <h2 className="text-xl font-bold flex items-center gap-2">
                {activeTab === 'notice_write' ? '공지사항 작성 에디터' : '교구 보고서 작성'}
                {isLoadingData ? (
                  <span className="text-xs font-normal text-blue-500 flex items-center gap-1 bg-blue-50 px-2 py-1 rounded-full animate-pulse">
                    <RefreshCw className="w-3 h-3 animate-spin" /> 불러오는 중...
                  </span>
                ) : lastSaved ? (
                  <span className="text-xs font-normal text-slate-500 flex items-center gap-1 bg-slate-100 px-2 py-1 rounded-full" title={lastSaved}>
                    <Clock className="w-3 h-3" /> 저장: {getRelativeTime(lastSaved) || lastSaved}
                  </span>
                ) : null}
              </h2>
              <div className="flex gap-2 flex-wrap">
                {activeTab !== 'notice_write' && (
                  <button
                    onClick={() => { setSimpleMode(v => !v); if (quickEntryMode) setQuickEntryMode(false); }}
                    className={`flex items-center gap-1.5 text-sm px-3 py-1.5 rounded-md font-bold transition-colors shadow-sm border ${simpleMode ? 'bg-emerald-500 border-emerald-500 text-white' : 'bg-emerald-50 border-emerald-200 text-emerald-700 hover:bg-emerald-100'}`}
                    title="모바일 간편 입력 — 글+사진 카드 방식"
                  >
                    <ImageIcon className="w-4 h-4" />
                    {simpleMode ? '일반 편집기' : '간편 입력'}
                  </button>
                )}
                {activeTab !== 'notice_write' && !simpleMode && (
                  <button
                    onClick={() => quickEntryMode ? setQuickEntryMode(false) : enterQuickEntry()}
                    className={`flex items-center gap-1.5 text-sm px-3 py-1.5 rounded-md font-bold transition-colors shadow-sm border ${quickEntryMode ? 'bg-amber-500 border-amber-500 text-white' : 'bg-amber-50 border-amber-200 text-amber-700 hover:bg-amber-100'}`}
                    title="카카오톡 메시지 붙여넣기 입력"
                  >
                    <AlignLeft className="w-4 h-4" />
                    {quickEntryMode ? '편집기 복귀' : '붙여넣기'}
                  </button>
                )}
                {/* AI 검토 — 로컬 모드 + 사무장/관리자 전용 */}
                {(isLocalMode || isCloudAiAvailable || isBrowserAiReady) && activeTab !== 'notice_write' && (role === 'manager' || role === 'admin') && (
                  <button
                    onClick={checkWithLocalClaude}
                    disabled={isCheckingAI}
                    className={`flex items-center gap-1.5 text-sm px-3 py-1.5 rounded-md font-bold transition-colors shadow-sm border ${isCheckingAI ? 'bg-purple-100 border-purple-300 text-purple-500 cursor-wait' : 'bg-purple-50 border-purple-200 text-purple-700 hover:bg-purple-100'}`}
                    title="Claude AI로 맞춤법·전주금주 구분 검토"
                  >
                    <Bot className={`w-4 h-4 ${isCheckingAI ? 'animate-spin' : ''}`} />
                    {isCheckingAI ? '검토 중...' : 'AI 검토'}
                  </button>
                )}
                {/* 보고서 전체 텍스트 복사 */}
                {activeTab !== 'notice_write' && (
                  <button
                    onClick={() => {
                      const lines: string[] = [];
                      let l0Cnt = 0;
                      const counters = [0,0,0,0,0,0];
                      reportData.forEach(item => {
                        if (!item.text?.trim() && !item.tableData) return;
                        if (item.level === 0) {
                          l0Cnt++; counters[0]=l0Cnt; counters[1]=0; counters[2]=0;
                          lines.push(`\n${toRoman(l0Cnt)}. ${item.text}`);
                        } else if (item.level === 1) {
                          counters[1]++;
                          lines.push(`  ${counters[1]}. ${item.text}`);
                        } else if (item.level === 2) {
                          counters[2]++;
                          lines.push(`    ${counters[2]}) ${item.text}`);
                        } else {
                          lines.push(`      - ${item.text}`);
                        }
                        if (item.tableData) {
                          item.tableData.forEach((row: string[]) => lines.push(`    | ${row.join(' | ')} |`));
                        }
                      });
                      navigator.clipboard.writeText(lines.join('\n').trim());
                      toast.success('보고서 전체 텍스트가 복사되었습니다.');
                    }}
                    className="flex items-center gap-1.5 text-sm bg-blue-50 border border-blue-200 hover:bg-blue-100 text-blue-700 px-3 py-1.5 rounded-md transition-colors shadow-sm"
                    title="보고서 전체 내용을 텍스트로 복사"
                  >
                    <Copy className="w-4 h-4" />
                    전체 복사
                  </button>
                )}
                <button
                  onClick={() => setShowJsonModal(true)}
                  className="flex items-center gap-2 text-sm bg-slate-100 border border-slate-200 hover:bg-slate-200 text-slate-700 px-3 py-1.5 rounded-md transition-colors shadow-sm"
                  title="데이터 구조 확인"
                >
                  <FileJson className="w-4 h-4" />
                  추출
                </button>
                <button
                  onClick={() => { setResetMode('quick'); setResetSelectedParishes([]); setResetSelectedChurches({}); setShowResetModal(true); }}
                  className="flex items-center gap-1.5 text-sm bg-red-50 border border-red-200 hover:bg-red-100 text-red-700 font-bold px-3 py-1.5 rounded-md transition-colors shadow-sm"
                  title="데이터 초기화 옵션"
                >
                  <Trash2 className="w-4 h-4" />
                  전체 초기화
                </button>
              </div>
            </div>
            
            <div className="flex flex-col gap-3">
              <div className="flex gap-3">
                {activeTab === 'notice_write' ? (
                  <div className="flex-1 flex flex-col gap-2">
                    <input
                      type="text"
                      value={noticeTitle}
                      onChange={e => setNoticeTitle(e.target.value)}
                      className="w-full px-0 py-1 border-0 border-b-2 border-slate-200 focus:border-blue-500 text-2xl font-black text-slate-800 bg-transparent outline-none placeholder:text-slate-300 transition-colors"
                      placeholder="공지 제목을 입력하세요..."
                    />
                    <div className="flex gap-2 items-center flex-wrap">
                      <div className="flex gap-1.5">
                        {['공지', '행사', '긴급', '안내'].map(cat => (
                          <button
                            key={cat}
                            onClick={() => setNoticeCategory(cat)}
                            className={`text-xs font-bold px-2.5 py-1 rounded-full border transition-colors ${noticeCategory === cat ? 'bg-blue-600 text-white border-blue-600' : 'bg-white text-slate-500 border-slate-200 hover:border-blue-300'}`}
                          >{cat}</button>
                        ))}
                      </div>
                      <div className="h-4 w-px bg-slate-200" />
                      <label className="text-xs bg-slate-50 hover:bg-slate-100 text-slate-600 border border-slate-200 px-2.5 py-1 rounded-full cursor-pointer font-medium flex items-center gap-1 transition-colors shrink-0">
                        {isUploadingNotice ? <RefreshCw className="w-3 h-3 animate-spin" /> : <Upload className="w-3 h-3" />}
                        {isUploadingNotice ? 'PDF 첨부 중...' : 'PDF 첨부'}
                        <input type="file" accept="application/pdf" className="hidden" onChange={handleNoticePdfUpload} disabled={isUploadingNotice} />
                      </label>
                      {noticePdfUrl && <span className="text-xs text-blue-600 font-semibold flex items-center gap-1"><Check className="w-3 h-3" /> PDF 첨부됨</span>}
                    </div>
                  </div>
                ) : (
                  <>
                    {(
                      <div className="flex-1">
                        <label htmlFor="parishSelect" className="block text-xs font-bold text-slate-500 mb-1">교구</label>
                        <select 
                          id="parishSelect" 
                          value={parish}
                          onChange={handleParishChange}
                          disabled={role === 'manager' || role === 'church'}
                          className="w-full px-3 py-2 border border-slate-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm font-medium disabled:bg-slate-100 disabled:text-slate-500"
                        >
                          {Object.keys(PARISH_CHURCH_MAP).map(p => (
                            <option key={p} value={p}>{getDisplayParish(p)}</option>
                          ))}
                        </select>
                      </div>
                    )}
                    <div className="flex-1">
                      <label htmlFor="churchSelect" className="block text-xs font-bold text-slate-500 mb-1">교회</label>
                      <select 
                        id="churchSelect" 
                        value={church}
                        onChange={(e) => handleChurchChange(e.target.value)}
                        disabled={role === 'church'}
                        className="w-full px-3 py-2 border border-slate-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm font-medium disabled:bg-slate-100 disabled:text-slate-500"
                      >
                        {(PARISH_CHURCH_MAP[parish] || []).map(c => (
                          <option key={c} value={c}>{getDisplayChurch(c)}</option>
                        ))}
                      </select>
                    </div>
                  </>
                )}
              </div>
              {role !== 'church' && (
                <div className="bg-slate-50 border border-slate-200 rounded-lg p-3">
                  <div className="flex justify-between items-center mb-2">
                    <div className="flex items-center gap-2">
                      <span className="text-xs font-bold text-slate-500">교회별 제출 현황</span>
                      {(() => {
                        const churches = PARISH_CHURCH_MAP[parish] || [];
                        const submitted = churches.filter(c => parishStats[c] === 'submitted').length;
                        const total = churches.length;
                        return (
                          <span className={`text-xs font-black px-2 py-0.5 rounded-full ${submitted === total ? 'bg-emerald-100 text-emerald-700' : submitted > 0 ? 'bg-blue-100 text-blue-700' : 'bg-slate-200 text-slate-500'}`}>
                            {submitted}/{total} 제출
                          </span>
                        );
                      })()}
                    </div>
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
                  {/* 미제출 알림 문자 복사 */}
                  {(() => {
                    const notSubmitted = PARISH_CHURCH_MAP[parish].filter(c => parishStats[c] !== 'submitted');
                    if (notSubmitted.length === 0) return (
                      <div className="mt-2 text-xs text-emerald-600 font-bold flex items-center gap-1.5">
                        <CheckCircle className="w-3.5 h-3.5" /> 모든 교회 제출 완료! 🎉
                      </div>
                    );
                    return (
                      <button
                        onClick={() => {
                          const dateStr = appConfig?.solarDate || '이번 주';
                          const msg = `[${parish} 주간보고 알림]\n${dateStr} 주간보고 미제출 교회입니다.\n미제출: ${notSubmitted.join(', ')}\n빠른 제출 부탁드립니다.`;
                          navigator.clipboard.writeText(msg);
                          toast.success('미제출 알림 문자가 복사되었습니다.');
                        }}
                        className="mt-2 w-full flex items-center justify-center gap-1.5 px-3 py-2 bg-amber-50 hover:bg-amber-100 border border-amber-200 text-amber-700 rounded-lg text-xs font-bold transition-colors"
                      >
                        <Copy className="w-3.5 h-3.5" /> 미제출({notSubmitted.length}개) 알림 문자 복사
                      </button>
                    );
                  })()}
                  {/* 교구 Word 내보내기 (사무장용) */}
                  {role === 'manager' && (
                    <button
                      onClick={exportToWord}
                      className="mt-1.5 w-full flex items-center justify-center gap-1.5 px-3 py-2 bg-blue-50 hover:bg-blue-100 border border-blue-200 text-blue-700 rounded-lg text-xs font-bold transition-colors"
                    >
                      <Download className="w-3.5 h-3.5" /> {getDisplayParish(parish)} 교구 Word 내보내기
                    </button>
                  )}
                </div>
              )}
            </div>
          </div>



          {/* AI 교정 결과 패널 */}
          {aiCorrections && aiCorrections.length > 0 && (
            <div className="mb-4 bg-purple-50 border border-purple-200 rounded-xl overflow-hidden">
              <div className="flex items-center justify-between px-4 py-2.5 bg-purple-100 border-b border-purple-200">
                <span className="text-sm font-black text-purple-800 flex items-center gap-1.5">
                  <Bot className="w-4 h-4" /> AI 수정 제안 ({aiCorrections.length}건)
                </span>
                <div className="flex gap-2">
                  <button
                    onClick={applyAllAiCorrections}
                    className="px-3 py-1 bg-purple-600 hover:bg-purple-700 text-white text-xs font-bold rounded-lg transition-colors"
                  >
                    전체 적용
                  </button>
                  <button
                    onClick={() => setAiCorrections(null)}
                    className="px-3 py-1 bg-white hover:bg-slate-50 text-slate-500 text-xs font-bold rounded-lg border border-slate-200 transition-colors"
                  >
                    닫기
                  </button>
                </div>
              </div>
              <div className="divide-y divide-purple-100 max-h-72 overflow-y-auto">
                {aiCorrections.map((c, i) => (
                  <div key={i} className="flex items-start gap-3 px-4 py-3">
                    <div className="flex-1 min-w-0">
                      <div className="flex gap-2 text-xs mb-1">
                        <span className="text-red-600 line-through break-all">{c.original}</span>
                        <span className="text-slate-400">→</span>
                        <span className="text-emerald-700 font-bold break-all">{c.corrected}</span>
                      </div>
                      {c.reason && <p className="text-[10px] text-purple-600">{c.reason}</p>}
                    </div>
                    <button
                      onClick={() => applyAiCorrection(c.original, c.corrected)}
                      className="shrink-0 px-2.5 py-1 bg-emerald-500 hover:bg-emerald-600 text-white text-xs font-bold rounded-lg transition-colors"
                    >
                      적용
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}
          {aiCorrections && aiCorrections.length === 0 && (
            <div className="mb-4 flex items-center gap-2 px-4 py-3 bg-emerald-50 border border-emerald-200 rounded-xl text-sm text-emerald-700 font-bold">
              <Check className="w-4 h-4" /> AI 검토 완료 — 수정이 필요한 항목이 없습니다!
              <button onClick={() => setAiCorrections(null)} className="ml-auto text-xs text-slate-400 hover:text-slate-600"><X className="w-3.5 h-3.5" /></button>
            </div>
          )}

          {/* 줄별 빠른 입력 모드 */}
          {quickEntryMode && activeTab !== 'notice_write' && (
            <div className="mb-4 flex flex-col gap-3">
              <div className="bg-amber-50 border border-amber-200 rounded-lg px-4 py-3 text-xs text-amber-800 leading-relaxed">
                <strong>붙여넣기 입력</strong> — 카카오톡 메시지를 그대로 붙여넣기 하세요! <code className="bg-amber-100 px-1 rounded"># 섹션제목</code>으로 항목 구분, 나머지 줄은 세부 내용으로 들어갑니다. 나중에 AI로 정리 가능합니다.
              </div>
              <textarea
                value={quickEntryText}
                onChange={e => setQuickEntryText(e.target.value)}
                className="w-full min-h-[300px] sm:min-h-[400px] border border-slate-300 rounded-lg px-4 py-3 text-base leading-relaxed focus:ring-2 focus:ring-amber-400 focus:border-amber-400 outline-none resize-y"
                placeholder={`카카오톡 메시지 복사 후 여기에 붙여넣기\n\n예시:\n# 전주 결과보고\n행사 진행 완료\n참석 인원 50명\n\n# 금주 계획\n수요일 모임 준비`}
                autoFocus
              />
              <div className="flex gap-2">
                <button
                  onClick={applyQuickEntry}
                  className="flex-1 bg-amber-500 hover:bg-amber-600 text-white font-bold py-2.5 rounded-lg flex items-center justify-center gap-2 transition-colors"
                >
                  <Check className="w-4 h-4" /> 적용하기
                </button>
                <button
                  onClick={() => setQuickEntryMode(false)}
                  className="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-600 font-bold rounded-lg transition-colors"
                >
                  취소
                </button>
              </div>
            </div>
          )}

          {/* 간편 입력 모드 — 글+사진 카드 방식 */}
          {simpleMode && activeTab !== 'notice_write' && (
            <div className="mb-4 space-y-3">
              <div className="bg-emerald-50 border border-emerald-200 rounded-xl px-4 py-3 text-xs text-emerald-800 leading-relaxed flex items-start gap-2">
                <ImageIcon className="w-4 h-4 shrink-0 mt-0.5 text-emerald-600" />
                <span><strong>간편 입력 모드</strong> — 각 항목에 글을 쓰고 사진을 첨부하세요. 카카오톡 내용 복사 후 그대로 붙여넣기 가능합니다.</span>
              </div>
              {reportData.filter(item => item.level === 0).map((item, idx) => (
                <div key={item.id} className="bg-white border-2 border-slate-200 rounded-2xl overflow-hidden shadow-sm">
                  <div className="bg-slate-50 px-4 py-2.5 flex items-center justify-between border-b border-slate-200">
                    <span className="text-xs font-bold text-slate-500">항목 {idx + 1}</span>
                    <button onClick={() => removeItem(item.id)} className="p-1 text-red-400 hover:text-red-600 rounded transition-colors">
                      <X className="w-4 h-4" />
                    </button>
                  </div>
                  <div className="p-4">
                    <TextareaAutosize
                      value={item.text || ''}
                      onChange={e => updateText(item.id, e.target.value)}
                      className="w-full text-base leading-relaxed outline-none resize-none font-medium placeholder:text-slate-300"
                      placeholder="내용을 입력하거나 카카오톡 메시지를 붙여넣기 하세요..."
                      minRows={3}
                    />
                  </div>
                  {item.image && (
                    <div className="px-4 pb-3">
                      <div className="relative inline-block w-full">
                        <img src={item.image} alt="첨부" className="w-full max-h-64 object-contain rounded-xl border border-slate-200" />
                        <button
                          onClick={() => removeImage(item.id)}
                          className="absolute top-2 right-2 bg-red-500 text-white rounded-full p-1.5 shadow-md"
                        >
                          <X className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>
                  )}
                  <div className="px-4 pb-4">
                    <label className="flex items-center justify-center gap-2 w-full py-3.5 rounded-xl border-2 border-dashed border-slate-200 hover:border-blue-400 hover:bg-blue-50 cursor-pointer transition-colors active:scale-[0.98]">
                      <ImageIcon className="w-5 h-5 text-slate-400" />
                      <span className="text-sm text-slate-500 font-semibold">{item.image ? '사진 변경' : '📷 사진 첨부'}</span>
                      <input type="file" accept="image/*" capture="environment" className="hidden" onChange={e => handleImageUpload(e, item.id)} />
                    </label>
                  </div>
                </div>
              ))}
              <button
                onClick={() => {
                  const newId = Date.now();
                  setReportData(prev => [...prev, { id: newId, text: '', level: 0 }]);
                }}
                className="w-full py-4 rounded-2xl border-2 border-dashed border-emerald-300 hover:border-emerald-500 hover:bg-emerald-50 text-emerald-600 font-bold flex items-center justify-center gap-2 transition-colors active:scale-[0.98] text-sm"
              >
                <Plus className="w-5 h-5" /> 새 항목 추가
              </button>
            </div>
          )}

          <div className={`flex-1 pr-2 space-y-2 pb-4 ${quickEntryMode || simpleMode ? 'hidden' : ''}`}>
            {(() => {
              const editorCounters = [0, 0, 0, 0, 0, 0];
              let currentL0Id: number | null = null;
              return reportData.map((item, index) => {
                if (item.level === 0) {
                  currentL0Id = item.id;
                  editorCounters[0]++; editorCounters[1]=0; editorCounters[2]=0; editorCounters[3]=0; editorCounters[4]=0; editorCounters[5]=0;
                  const isCollapsed = collapsedSections.has(item.id);
                  return (
                    <div key={item.id} className="flex flex-col gap-2 py-3 mt-4 first:mt-0 group">
                      <div className="font-bold text-lg text-blue-800 border-b-2 border-blue-100 w-full pb-1 flex items-center gap-2">
                        <button
                          onClick={() => toggleSection(item.id)}
                          className="shrink-0 text-blue-400 hover:text-blue-600 transition-colors"
                          title={isCollapsed ? '섹션 펼치기' : '섹션 접기'}
                        >
                          <ArrowRight className={`w-4 h-4 transition-transform duration-200 ${isCollapsed ? '' : 'rotate-90'}`} />
                        </button>
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
                          <button onClick={() => duplicateL0Block(index)} className="p-1 text-emerald-400 hover:bg-emerald-50 hover:text-emerald-600 rounded" title="이 섹션 복제">
                            <Copy className="w-4 h-4" />
                          </button>
                          <button onClick={() => removeItem(item.id)} className="p-1 text-red-300 hover:bg-red-50 hover:text-red-500 rounded-md shrink-0 transition-colors" title="삭제">
                            <X className="w-4 h-4" />
                          </button>
                        </div>
                      </div>

                      {/* 접힌 상태 표시 */}
                      {isCollapsed && (
                        <div
                          onClick={() => toggleSection(item.id)}
                          className="cursor-pointer flex items-center gap-2 px-3 py-2 bg-blue-50 border border-dashed border-blue-200 rounded-lg text-xs text-blue-500 hover:bg-blue-100 transition-colors"
                        >
                          <ArrowRight className="w-3.5 h-3.5" />
                          <span className="font-medium">
                            {(() => {
                              // 이 L0 섹션에 속한 자식 아이템 수 계산
                              let count = 0;
                              for (let i = index + 1; i < reportData.length; i++) {
                                if (reportData[i].level === 0) break;
                                if (reportData[i].text?.trim()) count++;
                              }
                              return `${count}개 항목 숨김 — 클릭하여 펼치기`;
                            })()}
                          </span>
                        </div>
                      )}

                      {!isCollapsed && item.image && (
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

                      {!isCollapsed && item.tableData && (
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

                // 접힌 섹션의 자식 아이템은 렌더링 생략
                if (currentL0Id !== null && collapsedSections.has(currentL0Id)) {
                  return null;
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
            
            {/* 빠른 대항목 템플릿 */}
            <div className="mt-3 pt-3 border-t border-slate-100">
              <p className="text-[10px] text-slate-400 font-semibold mb-2 uppercase tracking-wide">빠른 대항목 추가</p>
              <div className="flex flex-wrap gap-1.5">
                {['전주 결과보고', '금주 계획 및 보고', '특별활동', '교육 및 훈련', '봉사 활동', '전도 및 선교', '기타'].map(title => (
                  <button
                    key={title}
                    onClick={() => {
                      const newItem = { id: nextId, text: title, level: 0, isFixed: false };
                      const newChild = { id: nextId + 1, text: "", level: 1 };
                      setNextId(prev => prev + 2);
                      setReportData(data => [...data, newItem, newChild]);
                      setTimeout(() => { document.getElementById(`input-${newChild.id}`)?.focus(); }, 10);
                    }}
                    className="px-2.5 py-1 bg-blue-50 hover:bg-blue-100 border border-blue-200 text-blue-700 rounded-lg text-xs font-medium transition-colors flex items-center gap-1"
                  >
                    <Plus className="w-3 h-3" />{title}
                  </button>
                ))}
              </div>
            </div>

            <div className="flex gap-2 mt-3 pt-3 border-t border-slate-100">
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
              <input
                ref={docUploadRef}
                type="file"
                accept=".docx,.hwpx"
                className="hidden"
                onChange={(e) => { const f = e.target.files?.[0]; if (f) { importDocumentFile(f); e.target.value = ''; } }}
              />
              <button
                onClick={() => setShowAiPasteModal(true)}
                className="flex items-center justify-center gap-1.5 bg-violet-50 hover:bg-violet-100 border border-violet-200 text-violet-700 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors"
                title="AI가 복사한 내용을 자동으로 양식에 맞게 변환"
              >
                <Sparkles className="w-4 h-4" /> AI 변환
              </button>
            </div>
          </div>
          
          <div className="pt-4 border-t border-slate-200 mt-auto shrink-0 space-y-2">
            {/* 보고서 완성도 */}
            {(() => {
              const cleanItems = getCleanData(reportData);
              const totalItems = cleanItems.length;
              const l0Count = cleanItems.filter(i => i.level === 1).length; // 대항목 수
              const filledItems = cleanItems.filter(i => i.text && i.text.trim().length > 2).length;
              const totalChars = cleanItems.reduce((acc, i) => acc + (i.text?.length || 0), 0);
              if (totalItems === 0) return null;
              const pct = Math.round((filledItems / totalItems) * 100);
              return (
                <div className="px-1 pb-2">
                  <div className="flex items-center justify-between text-[10px] text-slate-500 mb-1">
                    <span className="font-semibold">보고서 완성도</span>
                    <span className={`font-black ${pct >= 80 ? 'text-emerald-600' : pct >= 50 ? 'text-blue-600' : 'text-slate-400'}`}>{pct}% · {l0Count}개 대항목</span>
                  </div>
                  <div className="w-full bg-slate-100 rounded-full h-1.5 overflow-hidden">
                    <div
                      className={`h-full rounded-full transition-all duration-500 ${pct >= 80 ? 'bg-emerald-500' : pct >= 50 ? 'bg-blue-500' : 'bg-amber-400'}`}
                      style={{ width: `${pct}%` }}
                    />
                  </div>
                  <div className="flex items-center justify-between mt-1 text-[10px] text-slate-400">
                    <span>항목 {filledItems}/{totalItems} 작성</span>
                    <span>{totalChars.toLocaleString()}자</span>
                  </div>
                </div>
              );
            })()}
            <div className="flex items-center justify-center gap-2 text-[10px] font-semibold text-slate-400 select-none pb-2 border-b border-slate-100 mb-1 flex-wrap">
              <span><kbd className="font-mono bg-slate-50 border border-slate-200 px-1 py-0.2 rounded text-slate-500 mr-0.5 shadow-sm">Tab</kbd>들여쓰기</span>
              <span className="text-slate-200">|</span>
              <span><kbd className="font-mono bg-slate-50 border border-slate-200 px-1 py-0.2 rounded text-slate-500 mr-0.5 shadow-sm">Shift+Tab</kbd>내어쓰기</span>
              <span className="text-slate-200">|</span>
              <span><kbd className="font-mono bg-slate-50 border border-slate-200 px-1 py-0.2 rounded text-slate-500 mr-0.5 shadow-sm">Enter</kbd>항목 추가</span>
              <span className="text-slate-200">|</span>
              <span><kbd className="font-mono bg-slate-50 border border-slate-200 px-1 py-0.2 rounded text-slate-500 mr-0.5 shadow-sm">Ctrl+S</kbd>저장</span>
              <span className="text-slate-200">|</span>
              <span><kbd className="font-mono bg-slate-50 border border-slate-200 px-1 py-0.2 rounded text-slate-500 mr-0.5 shadow-sm">Ctrl+↵</kbd>제출</span>
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
                    className={`flex-1 flex items-center justify-center gap-2 p-3 rounded-lg font-bold transition-all disabled:opacity-70 ${
                      status === 'submitted'
                        ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                        : 'bg-gradient-to-r from-emerald-500 to-teal-500 hover:from-emerald-600 hover:to-teal-600 text-white shadow-md hover:shadow-lg active:scale-[0.98]'
                    }`}
                  >
                    <Check className="w-4 h-4" />
                    {status === 'submitted' ? '✅ 제출 완료' : '제출 확정'}
                  </button>
                  {status === 'submitted' && (
                    <button 
                      onClick={() => { setStatus('draft'); handleSave(false); }}
                      className="flex-1 flex items-center justify-center gap-2 p-3 rounded-lg font-bold transition-colors bg-amber-100 hover:bg-amber-200 text-amber-800 border border-amber-300"
                    >
                      <Save className="w-4 h-4" /> 저장
                    </button>
                  )}
                </>
              )}
            </div>
            {(role === 'admin' || role === 'manager') && (
            <button
              onClick={exportToWord}
              className="w-full flex items-center justify-center gap-2 bg-blue-600 hover:bg-blue-700 text-white p-3 rounded-lg font-medium shadow-sm transition-colors"
            >
              <Download className="w-5 h-5" />
              Word 내보내기
            </button>
            )}
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
              <div className="flex items-start justify-between pb-4 border-b border-slate-200 mb-4 gap-2">
                <div className="min-w-0">
                  <h2 className="text-lg font-black text-slate-800 flex items-center gap-2">
                    <Settings className="w-5 h-5 text-purple-600 animate-spin-slow" />
                    전체 교구 제출현황 및 제어
                  </h2>
                  <div className="flex items-center gap-2 mt-0.5 flex-wrap">
                    <p className="text-xs text-slate-500">전국 교구의 작성 현황</p>
                    {adminLastRefreshed && (
                      <span className="text-[10px] text-slate-400" title={adminLastRefreshed.toLocaleTimeString('ko-KR')}>
                        갱신: {adminLastRefreshed.toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' })}
                      </span>
                    )}
                  </div>
                </div>
                <div className="flex gap-1.5 shrink-0 flex-wrap justify-end">
                  <button
                    onClick={() => refreshFromCloud()}
                    className="px-2 py-1.5 bg-slate-50 hover:bg-slate-100 border border-slate-200 text-slate-600 rounded-lg text-xs font-bold transition-colors flex items-center gap-1"
                    title="클라우드에서 최신 데이터 새로고침"
                  >
                    <RefreshCw className="w-3 h-3" />
                  </button>
                  <button
                    onClick={() => setAdminAutoRefresh(v => !v)}
                    className={`px-2 py-1.5 border rounded-lg text-xs font-bold transition-colors flex items-center gap-1 ${adminAutoRefresh ? 'bg-emerald-50 border-emerald-300 text-emerald-700' : 'bg-slate-50 border-slate-200 text-slate-500 hover:bg-slate-100'}`}
                    title={adminAutoRefresh ? '자동 갱신 켜짐 (60초마다)' : '자동 갱신 꺼짐'}
                  >
                    <Clock className="w-3 h-3" /> {adminAutoRefresh ? '자동 ON' : '자동 OFF'}
                  </button>
                  <button
                    onClick={() => {
                      setNewWeekPassword('');
                      setNewWeekSolarDate(appConfig?.solarDate || getCurrentWeekSuggestion());
                      setNewWeekHeavenlyDate(appConfig?.heavenlyDate || '');
                      setNewWeekPasswordError('');
                      setShowNewWeekModal(true);
                    }}
                    className="px-2.5 py-1.5 bg-red-50 hover:bg-red-100 border border-red-200 text-red-600 rounded-lg text-xs font-bold transition-colors flex items-center gap-1"
                  >
                    <RefreshCw className="w-3 h-3" /> 새 주간보고
                  </button>
                </div>
              </div>

              {/* 전국 종합 현황 요약 */}
              {(() => {
                const allChurches = Object.entries(PARISH_CHURCH_MAP)
                  
                  .flatMap(([p, cs]) => cs.slice(1).map(c => `${p}_${c}`));
                const totalSubmitted = allChurches.filter(k => adminReportStatusMap[k] === 'submitted').length;
                const totalDraft = allChurches.filter(k => adminReportStatusMap[k] === 'draft').length;
                const totalEmpty = allChurches.length - totalSubmitted - totalDraft;
                const pct = allChurches.length > 0 ? Math.round((totalSubmitted / allChurches.length) * 100) : 0;
                return (
                  <div className="mb-3 bg-gradient-to-r from-purple-50 to-indigo-50 border border-purple-200 rounded-xl p-3">
                    <div className="flex items-center justify-between mb-1.5">
                      <span className="text-xs font-bold text-purple-800">전국 교구 제출 현황</span>
                      <span className={`text-sm font-black ${pct === 100 ? 'text-emerald-600' : 'text-purple-700'}`}>{pct}%</span>
                    </div>
                    <div className="w-full bg-white/60 rounded-full h-2 overflow-hidden mb-2">
                      <div className="h-full bg-gradient-to-r from-purple-500 to-indigo-500 rounded-full transition-all duration-700" style={{ width: `${pct}%` }} />
                    </div>
                    <div className="flex gap-3 text-[10px] font-semibold items-center">
                      <span className="text-emerald-600">✅ 제출 {totalSubmitted}</span>
                      <span className="text-blue-600">📝 작성중 {totalDraft}</span>
                      <span className="text-slate-400">⬜ 미작성 {totalEmpty}</span>
                      <span className="text-slate-500">합계 {allChurches.length}개</span>
                      <button
                        onClick={() => {
                          const lines = [`[${appConfig?.solarDate || '이번 주'} 전국 교구 제출현황]`, `제출률: ${pct}% (${totalSubmitted}/${allChurches.length}개)`, ''];
                          Object.entries(PARISH_CHURCH_MAP)
                            
                            .forEach(([p, cs]) => {
                              const tcs = (cs as string[]).slice(1);
                              const submitted = tcs.filter(c => adminReportStatusMap[`${p}_${c}`] === 'submitted').length;
                              const notSubmitted = tcs.filter(c => adminReportStatusMap[`${p}_${c}`] !== 'submitted');
                              lines.push(`• ${getDisplayParish(p)}: ${submitted}/${tcs.length}${notSubmitted.length > 0 ? ` (미제출: ${notSubmitted.map(c => getDisplayChurch(c)).join(', ')})` : ' ✅'}`);
                            });
                          navigator.clipboard.writeText(lines.join('\n'));
                          toast.success('전국 현황 요약이 복사되었습니다.');
                        }}
                        className="ml-auto flex items-center gap-1 px-2 py-1 bg-white/60 hover:bg-white border border-purple-200 text-purple-700 rounded-md text-[10px] font-bold transition-colors"
                        title="전국 현황 요약 텍스트 복사"
                      >
                        <Copy className="w-2.5 h-2.5" /> 현황 복사
                      </button>
                    </div>
                  </div>
                );
              })()}

              {/* Parish Dashboard (Accordion) — 미제출순 정렬 */}
              <div className="flex-1 overflow-y-auto pr-2 space-y-2">
                {Object.keys(PARISH_CHURCH_MAP)
                  .map(p => {
                    const churches = PARISH_CHURCH_MAP[p];
                    const targetChurches = churches.slice(1);
                    const submittedCount = targetChurches.filter(c => adminReportStatusMap[`${p}_${c}`] === 'submitted').length;
                    const completionRate = targetChurches.length > 0 ? Math.round((submittedCount / targetChurches.length) * 100) : 0;
                    return { p, targetChurches, submittedCount, completionRate };
                  })
                  .filter(({ targetChurches }) => targetChurches.length > 0)
                  .sort((a, b) => a.completionRate - b.completionRate)
                  .map(({ p, targetChurches, submittedCount, completionRate }) => {
                  const draftCount = targetChurches.filter(c => adminReportStatusMap[`${p}_${c}`] === 'draft').length;
                  const isExpanded = adminExpandedParish === p;
                  const isComplete = submittedCount === targetChurches.length;

                  return (
                    <div key={p} className={`border rounded-xl overflow-hidden transition-all ${isComplete ? 'border-emerald-200 bg-emerald-50/40' : isExpanded ? 'border-purple-300' : 'border-slate-200 bg-white'}`}>
                      {/* Header row */}
                      <div
                        onClick={() => { setAdminExpandedParish(isExpanded ? null : p); if (!isExpanded) setAdminConsoleParish(p); }}
                        className="flex items-center gap-3 p-3 cursor-pointer hover:bg-slate-50 transition-colors"
                      >
                        <div className={`w-2 h-2 rounded-full shrink-0 ${isComplete ? 'bg-emerald-500' : draftCount > 0 ? 'bg-blue-400' : 'bg-slate-300'}`} />
                        <span className="font-extrabold text-sm text-slate-800 flex-1">{getDisplayParish(p)}</span>
                        <div className="flex items-center gap-2 text-xs">
                          <span className={`font-black ${isComplete ? 'text-emerald-600' : 'text-slate-500'}`}>{submittedCount}/{targetChurches.length}</span>
                          <span className={`font-black text-sm ${isComplete ? 'text-emerald-600' : completionRate > 0 ? 'text-purple-600' : 'text-slate-400'}`}>{completionRate}%</span>
                        </div>
                        <ArrowRight className={`w-3.5 h-3.5 text-slate-400 shrink-0 transition-transform ${isExpanded ? 'rotate-90' : ''}`} />
                      </div>
                      {/* Progress bar */}
                      <div className="h-1.5 bg-slate-100 mx-3 mb-2 rounded-full overflow-hidden">
                        <div
                          className={`h-full rounded-full transition-all duration-700 ease-out ${isComplete ? 'bg-emerald-500' : 'bg-purple-500'}`}
                          style={{ width: `${completionRate}%` }}
                        />
                      </div>
                      {/* Expanded church list */}
                      {isExpanded && (
                        <div className="border-t border-slate-100 px-3 pb-3 pt-2">
                          {/* 교구 빠른 액션 */}
                          <div className="flex gap-1.5 mb-2">
                            <button
                              onClick={() => {
                                setActiveTab('report');
                                setParish(p);
                                setChurch(PARISH_CHURCH_MAP[p][0]);
                                setTimeout(() => exportToWord(), 500);
                              }}
                              className="flex items-center gap-1 px-2 py-1 bg-blue-50 hover:bg-blue-100 border border-blue-200 text-blue-700 rounded text-[10px] font-bold transition-colors"
                            >
                              <Download className="w-3 h-3" /> Word 내보내기
                            </button>
                            <button
                              onClick={() => {
                                const notSub = targetChurches.filter(c => adminReportStatusMap[`${p}_${c}`] !== 'submitted').map(c => getDisplayChurch(c));
                                if (notSub.length === 0) { toast.success('모든 교회가 제출 완료되었습니다!'); return; }
                                const msg = `[${getDisplayParish(p)} 주간보고 알림]\n${appConfig?.solarDate || '이번 주'} 미제출 교회:\n${notSub.map(c => `• ${c}`).join('\n')}\n\n빠른 제출 부탁드립니다.`;
                                navigator.clipboard.writeText(msg);
                                toast.success('교구 미제출 알림이 복사되었습니다.');
                              }}
                              className="flex items-center gap-1 px-2 py-1 bg-amber-50 hover:bg-amber-100 border border-amber-200 text-amber-700 rounded text-[10px] font-bold transition-colors"
                            >
                              <Copy className="w-3 h-3" /> 미제출 알림
                            </button>
                          </div>
                          <div className="grid grid-cols-2 gap-1.5">
                            {targetChurches.map(c => {
                              const st = adminReportStatusMap[`${p}_${c}`] || 'empty';
                              const ts = adminReportTimestampMap[`${p}_${c}`];
                              return (
                                <div
                                  key={c}
                                  className={`relative rounded-lg text-xs border group ${
                                    st === 'submitted' ? 'bg-emerald-50 border-emerald-200' :
                                    st === 'draft' ? 'bg-blue-50 border-blue-200' :
                                    'bg-slate-50 border-slate-200'
                                  }`}
                                >
                                  <button
                                    onClick={() => {
                                      setActiveTab('report'); setParish(p); setChurch(c);
                                    }}
                                    className="text-left px-2.5 py-2 w-full"
                                  >
                                    <div className={`font-bold truncate ${st === 'submitted' ? 'text-emerald-700' : st === 'draft' ? 'text-blue-700' : 'text-slate-500'}`}>{c}</div>
                                    <div className="text-[10px] text-slate-400 mt-0.5 flex items-center gap-1">
                                      {st === 'submitted' ? <><Check className="w-2.5 h-2.5 text-emerald-500" />제출완료</> :
                                       st === 'draft' ? <><Clock className="w-2.5 h-2.5 text-blue-400" />작성중</> :
                                       <span className="text-slate-300">미작성</span>}
                                    </div>
                                    {ts && <div className="text-[9px] text-slate-400 truncate mt-0.5" title={ts}>{getRelativeTime(ts) || ts}</div>}
                                  </button>
                                  {/* 미리보기 버튼 */}
                                  <button
                                    onClick={(e) => { e.stopPropagation(); openChurchPreview(p, c); }}
                                    className="absolute bottom-1 right-1 hidden group-hover:flex w-5 h-5 items-center justify-center bg-blue-50 hover:bg-blue-100 border border-blue-200 text-blue-400 hover:text-blue-600 rounded transition-colors"
                                    title="보고서 내용 미리보기"
                                  >
                                    <BookOpen className="w-3 h-3" />
                                  </button>
                                  {/* 초기화 버튼 — 관리자: 비밀번호 필요 / 사무장: 자기 교구 내 자유롭게 삭제 */}
                                  {(role === 'admin' || (role === 'manager' && p === parish)) && (
                                    <button
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        const doReset = async () => {
                                          const defaultData = { data: DEFAULT_REPORT, lastSaved: null, status: 'draft' };
                                          const key = `report_${p}_${c}`;
                                          localStorage.setItem(key, JSON.stringify(defaultData));
                                          sessionStorage.removeItem(key);
                                          try {
                                            await saveDbData(`${p}_${c}`, { id: `${p}_${c}`, parish: p, church: c, ...defaultData, updated_at: new Date().toISOString() });
                                          } catch {}
                                          setAdminReportStatusMap(prev => ({ ...prev, [`${p}_${c}`]: 'empty' }));
                                          setAdminReportTimestampMap(prev => ({ ...prev, [`${p}_${c}`]: null }));
                                          toast.success(`[${c}] 보고서가 초기화되었습니다.`);
                                        };
                                        if (role === 'manager') {
                                          // 사무장: 확인창만
                                          if (window.confirm(`[${c}] 보고서를 초기화하시겠습니까?`)) doReset();
                                        } else {
                                          withAdminBypass(`[${c}] 초기화`, (pwd) => pwd === 'chongmu2027', doReset);
                                        }
                                      }}
                                      className="absolute top-1 right-1 hidden group-hover:flex w-5 h-5 items-center justify-center bg-red-50 hover:bg-red-100 border border-red-200 text-red-400 hover:text-red-600 rounded transition-colors"
                                      title="이 교회 보고서 초기화"
                                    >
                                      <X className="w-3 h-3" />
                                    </button>
                                  )}
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
              {/* 미제출 교회별 알림 문자 복사 */}
              {(() => {
                const parishMissing = Object.entries(PARISH_CHURCH_MAP)
                  .map(([p, cs]) => {
                    const missing = (cs as string[]).slice(1).filter(c => adminReportStatusMap[`${p}_${c}`] !== 'submitted');
                    return { p, missing };
                  })
                  .filter(({ missing }) => missing.length > 0);
                const totalMissing = parishMissing.reduce((sum, { missing }) => sum + missing.length, 0);
                if (parishMissing.length === 0) return (
                  <div className="mt-3 pt-3 border-t border-slate-100 text-xs text-emerald-600 font-bold flex items-center gap-1.5">
                    <CheckCircle className="w-3.5 h-3.5" /> 전국 모든 교구 제출 완료! 🎉
                  </div>
                );
                return (
                  <button
                    onClick={() => {
                      const dateStr = appConfig?.solarDate || '이번 주';
                      const lines = [`[전국 주간보고 알림]`, `${dateStr} 미제출 교회 현황`, ''];
                      parishMissing.forEach(({ p, missing }) => {
                        lines.push(`◎ ${getDisplayParish(p)} (${missing.length}개 미제출)`);
                        missing.forEach(c => lines.push(`  • ${getDisplayChurch(c)}`));
                      });
                      lines.push('', '빠른 제출 독려 부탁드립니다.');
                      navigator.clipboard.writeText(lines.join('\n'));
                      toast.success('미제출 교회별 알림 문자가 복사되었습니다.');
                    }}
                    className="mt-3 w-full flex items-center justify-center gap-1.5 px-3 py-2 bg-amber-50 hover:bg-amber-100 border border-amber-200 text-amber-700 rounded-lg text-xs font-bold transition-colors"
                  >
                    <Copy className="w-3.5 h-3.5" /> 미제출 {parishMissing.length}개 교구 / {totalMissing}개 교회 알림 복사
                  </button>
                );
              })()}
            </div>

            {/* Right Panel: Parish Compilation Console */}
            <div className="bg-white p-4 sm:p-6 rounded-xl shadow-sm border border-slate-200 flex flex-col h-[calc(100vh-12rem)] xl:h-[calc(100vh-8rem)]">
              {/* Header */}
              <div className="pb-4 border-b border-slate-200 mb-4 flex items-start justify-between gap-2">
                <div className="flex-1 min-w-0">
                  <h2 className="text-lg font-black text-slate-800 flex items-center gap-2">
                    <Download className="w-5 h-5 text-blue-600" />
                    교구별 취합 콘솔
                  </h2>
                  <p className="text-xs text-slate-500 mt-0.5">교구 선택 후 Word 내보내기 및 미제출 알림을 관리하세요</p>
                </div>
              </div>

              {/* Parish selector — 관리자만 변경 가능, 사무장은 자기 교구 고정 */}
              <div className="mb-4">
                <label className="block text-xs font-bold text-slate-500 mb-1.5">취합할 교구</label>
                {role === 'manager' ? (
                  <div className="w-full px-3 py-2.5 border border-slate-200 rounded-lg bg-slate-50 text-sm font-bold text-slate-700">
                    {getDisplayParish(adminConsoleParish)} <span className="text-xs text-slate-400 font-normal">(내 담당 교구)</span>
                  </div>
                ) : (
                  <select
                    value={adminConsoleParish}
                    onChange={(e) => setAdminConsoleParish(e.target.value)}
                    className="w-full px-3 py-2.5 border border-slate-300 rounded-lg bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm font-bold"
                  >
                    {Object.keys(PARISH_CHURCH_MAP).map(p => (
                      <option key={p} value={p}>{getDisplayParish(p)}</option>
                    ))}
                  </select>
                )}
              </div>

              {/* Parish stats */}
              {(() => {
                const targets = (PARISH_CHURCH_MAP[adminConsoleParish] || []).slice(1);
                const submitted = targets.filter(c => adminReportStatusMap[`${adminConsoleParish}_${c}`] === 'submitted').length;
                const draft = targets.filter(c => adminReportStatusMap[`${adminConsoleParish}_${c}`] === 'draft').length;
                const empty = targets.length - submitted - draft;
                const pct = targets.length > 0 ? Math.round(submitted / targets.length * 100) : 0;
                return (
                  <div className="mb-4 bg-gradient-to-r from-blue-50 to-indigo-50 border border-blue-200 rounded-xl p-3.5">
                    <div className="flex items-center justify-between mb-2">
                      <span className="text-xs font-bold text-blue-800">{getDisplayParish(adminConsoleParish)} 제출 현황</span>
                      <span className={`text-sm font-black ${pct === 100 ? 'text-emerald-600' : 'text-blue-700'}`}>{pct}%</span>
                    </div>
                    <div className="w-full bg-white/60 rounded-full h-2 overflow-hidden mb-2.5">
                      <div className="h-full bg-gradient-to-r from-blue-500 to-indigo-500 rounded-full transition-all duration-700" style={{ width: `${pct}%` }} />
                    </div>
                    <div className="flex gap-3 text-[11px] font-semibold">
                      <span className="text-emerald-600">✅ {submitted}</span>
                      <span className="text-blue-600">📝 {draft}</span>
                      <span className="text-slate-400">⬜ {empty}</span>
                      <span className="text-slate-500 ml-auto">총 {targets.length}개</span>
                    </div>
                  </div>
                );
              })()}

              {/* Action buttons */}
              <div className="grid grid-cols-2 gap-2 mb-3">
                <button
                  onClick={() => {
                    setActiveTab('report');
                    setParish(adminConsoleParish);
                    setChurch(PARISH_CHURCH_MAP[adminConsoleParish][0]);
                    setTimeout(() => exportToWord(), 500);
                  }}
                  className="bg-gradient-to-r from-blue-600 to-sky-600 hover:from-blue-700 hover:to-sky-700 text-white px-3 py-3.5 rounded-xl font-bold flex flex-col items-center justify-center gap-1 shadow-sm transition-all active:scale-[0.98]"
                >
                  <Download className="w-4 h-4" />
                  <span className="text-xs">{getDisplayParish(adminConsoleParish)}</span>
                  <span className="text-[10px] text-blue-200">교구 Word 출력</span>
                </button>
                <button
                  onClick={exportMasterToWord}
                  disabled={!!adminCompilationProgress}
                  className="bg-gradient-to-r from-indigo-600 to-purple-600 hover:from-indigo-700 hover:to-purple-700 text-white px-3 py-3.5 rounded-xl font-bold flex flex-col items-center justify-center gap-1 shadow-sm transition-all active:scale-[0.98] disabled:opacity-60"
                >
                  <Download className="w-4 h-4" />
                  <span className="text-xs">전국 통합</span>
                  <span className="text-[10px] text-indigo-200">마스터 Word</span>
                </button>
              </div>

              {/* 삭제 버튼 */}
              <div className="flex gap-2 mb-3">
                {/* 사무장: 담당 교구 전체 삭제 */}
                {role === 'manager' && (
                  <button
                    onClick={() => {
                      const targets = (PARISH_CHURCH_MAP[adminConsoleParish] || []).slice(1);
                      if (window.confirm(`${getDisplayParish(adminConsoleParish)} 교구 전체(${targets.length}개 교회) 보고서를 초기화하시겠습니까?`)) {
                        Promise.all(targets.map(async c => {
                          const defaultData = { data: DEFAULT_REPORT, lastSaved: null, status: 'draft' };
                          const key = `report_${adminConsoleParish}_${c}`;
                          localStorage.setItem(key, JSON.stringify(defaultData));
                          sessionStorage.removeItem(key);
                          try { await saveDbData(`${adminConsoleParish}_${c}`, { id: `${adminConsoleParish}_${c}`, parish: adminConsoleParish, church: c, ...defaultData, updated_at: new Date().toISOString() }); } catch {}
                          setAdminReportStatusMap(prev => ({ ...prev, [`${adminConsoleParish}_${c}`]: 'empty' }));
                          setAdminReportTimestampMap(prev => ({ ...prev, [`${adminConsoleParish}_${c}`]: null }));
                        })).then(() => toast.success(`${getDisplayParish(adminConsoleParish)} 교구 전체 초기화 완료`));
                      }
                    }}
                    className="flex-1 flex items-center justify-center gap-1.5 px-3 py-2.5 rounded-xl border border-red-200 bg-red-50 hover:bg-red-100 text-red-700 font-bold text-sm transition-all"
                  >
                    <Trash2 className="w-4 h-4" /> 교구 전체 삭제
                  </button>
                )}
                {/* 관리자: 전체 삭제 */}
                {role === 'admin' && (
                  <button
                    onClick={() => {
                      withAdminBypass('전체 삭제 (비밀번호 확인)', (pwd) => pwd === 'chongmu2027', async () => {
                        const allTargets = Object.keys(PARISH_CHURCH_MAP).flatMap(p => PARISH_CHURCH_MAP[p].map(c => ({ parish: p, church: c })));
                        await Promise.all(allTargets.map(async ({ parish: p, church: c }) => {
                          const defaultData = { data: DEFAULT_REPORT, lastSaved: null, status: 'draft' };
                          localStorage.setItem(`report_${p}_${c}`, JSON.stringify(defaultData));
                          sessionStorage.removeItem(`report_${p}_${c}`);
                          try { await saveDbData(`${p}_${c}`, { id: `${p}_${c}`, parish: p, church: c, ...defaultData, updated_at: new Date().toISOString() }); } catch {}
                          setAdminReportStatusMap(prev => ({ ...prev, [`${p}_${c}`]: 'empty' }));
                          setAdminReportTimestampMap(prev => ({ ...prev, [`${p}_${c}`]: null }));
                        }));
                        toast.success(`전국 전체(${allTargets.length}개) 보고서 초기화 완료`);
                      });
                    }}
                    className="flex-1 flex items-center justify-center gap-1.5 px-3 py-2.5 rounded-xl border border-red-300 bg-red-50 hover:bg-red-100 text-red-700 font-bold text-sm transition-all"
                  >
                    <Trash2 className="w-4 h-4" /> 전체 삭제
                  </button>
                )}
              </div>

              {/* Parish AI Review — 로컬 or 클라우드 or 브라우저 AI */}
              {(isLocalMode || isCloudAiAvailable || isBrowserAiReady) && (
                <button
                  onClick={startParishAiReview}
                  disabled={isAdminCheckingAI || !!adminCompilationProgress}
                  className={`mb-3 w-full flex items-center justify-center gap-2 px-3 py-3 rounded-xl font-bold text-sm transition-all active:scale-[0.98] shadow-sm ${isAdminCheckingAI ? 'bg-purple-100 border border-purple-200 text-purple-500 cursor-wait' : 'bg-gradient-to-r from-purple-600 to-violet-600 hover:from-purple-700 hover:to-violet-700 text-white'}`}
                >
                  <Bot className={`w-4 h-4 ${isAdminCheckingAI ? 'animate-spin' : ''}`} />
                  {isAdminCheckingAI ? '검토 중...' : `${getDisplayParish(adminConsoleParish)} 교구 AI 일괄 검토`}
                </button>
              )}

              {/* Admin AI corrections panel */}
              {adminAiCorrections && adminAiCorrections.length > 0 && (
                <div className="mb-3 bg-purple-50 border border-purple-200 rounded-xl overflow-hidden">
                  <div className="flex items-center justify-between px-3 py-2 bg-purple-100 border-b border-purple-200">
                    <span className="text-xs font-black text-purple-800 flex items-center gap-1"><Bot className="w-3.5 h-3.5" /> AI 제안 {adminAiCorrections.length}건</span>
                    <div className="flex gap-1.5">
                      <button onClick={applySelectedAdminCorrections} className="px-2.5 py-1 bg-purple-600 hover:bg-purple-700 text-white text-[10px] font-bold rounded-lg">전체 적용</button>
                      <button onClick={() => setAdminAiCorrections(null)} className="px-2.5 py-1 bg-white border border-slate-200 text-slate-500 text-[10px] font-bold rounded-lg">닫기</button>
                    </div>
                  </div>
                  <div className="divide-y divide-purple-100 max-h-52 overflow-y-auto">
                    {adminAiCorrections.map((c: any, i: number) => (
                      <div key={i} className="px-3 py-2.5 flex items-start gap-2">
                        <div className="flex-1 min-w-0">
                          <p className="text-[10px] font-bold text-slate-500 mb-0.5">{c.church}</p>
                          <div className="flex gap-1.5 text-[11px] flex-wrap">
                            <span className="text-red-600 line-through break-all">{c.original}</span>
                            <span className="text-slate-400">→</span>
                            <span className="text-emerald-700 font-bold break-all">{c.corrected}</span>
                          </div>
                          {c.reason && <p className="text-[10px] text-purple-500 mt-0.5">{c.reason}</p>}
                        </div>
                        <button
                          onClick={() => {
                            const key = `${c.parish}_${c.church}_${c.id}`;
                            setAdminSelectedCorrections({ [key]: true });
                            applySelectedAdminCorrections();
                          }}
                          className="shrink-0 px-2 py-1 bg-emerald-500 hover:bg-emerald-600 text-white text-[10px] font-bold rounded-lg"
                        >적용</button>
                      </div>
                    ))}
                  </div>
                </div>
              )}
              {adminAiCorrections && adminAiCorrections.length === 0 && (
                <div className="mb-3 flex items-center gap-2 px-3 py-2 bg-emerald-50 border border-emerald-200 rounded-xl text-xs text-emerald-700 font-bold">
                  <Check className="w-3.5 h-3.5" /> 수정 필요 없음 — 완벽합니다!
                  <button onClick={() => setAdminAiCorrections(null)} className="ml-auto"><X className="w-3 h-3 text-slate-400" /></button>
                </div>
              )}

              {/* Missing notification */}
              {(() => {
                const targets = (PARISH_CHURCH_MAP[adminConsoleParish] || []).slice(1);
                const notSub = targets.filter(c => adminReportStatusMap[`${adminConsoleParish}_${c}`] !== 'submitted').map(c => getDisplayChurch(c));
                if (notSub.length === 0) return (
                  <div className="mb-3 text-xs text-emerald-600 font-bold flex items-center gap-1.5 px-3 py-2 bg-emerald-50 rounded-lg border border-emerald-200">
                    <CheckCircle className="w-3.5 h-3.5" /> 모든 교회 제출 완료! 🎉
                  </div>
                );
                return (
                  <button
                    onClick={() => {
                      const msg = `[${getDisplayParish(adminConsoleParish)} 주간보고 알림]\n${appConfig?.solarDate || '이번 주'} 미제출 교회:\n${notSub.map(c => `• ${c}`).join('\n')}\n\n빠른 제출 부탁드립니다.`;
                      navigator.clipboard.writeText(msg);
                      toast.success('미제출 알림 문자가 복사되었습니다.');
                    }}
                    className="mb-3 w-full flex items-center justify-center gap-1.5 px-3 py-2.5 bg-amber-50 hover:bg-amber-100 border border-amber-200 text-amber-700 rounded-lg text-xs font-bold transition-colors"
                  >
                    <Copy className="w-3.5 h-3.5" /> {getDisplayParish(adminConsoleParish)} 미제출({notSub.length}개) 알림 복사
                  </button>
                );
              })()}

              {/* Church list for selected parish */}
              <div className="flex-1 overflow-y-auto min-h-0">
                <p className="text-[11px] font-bold text-slate-500 mb-2">{getDisplayParish(adminConsoleParish)} 교회별 현황 (클릭 시 보고서 편집)</p>
                <div className="grid grid-cols-2 gap-1.5">
                  {(PARISH_CHURCH_MAP[adminConsoleParish] || []).slice(1).map(c => {
                    const st = adminReportStatusMap[`${adminConsoleParish}_${c}`] || 'empty';
                    const ts = adminReportTimestampMap[`${adminConsoleParish}_${c}`];
                    return (
                      <button
                        key={c}
                        onClick={() => { setActiveTab('report'); setParish(adminConsoleParish); setChurch(c); }}
                        className={`text-left p-2.5 rounded-lg border text-xs transition-all hover:shadow-sm ${
                          st === 'submitted' ? 'bg-emerald-50 border-emerald-200 hover:bg-emerald-100' :
                          st === 'draft' ? 'bg-blue-50 border-blue-200 hover:bg-blue-100' :
                          'bg-slate-50 border-slate-200 hover:bg-slate-100'
                        }`}
                      >
                        <div className={`font-bold truncate ${st === 'submitted' ? 'text-emerald-700' : st === 'draft' ? 'text-blue-700' : 'text-slate-400'}`}>{getDisplayChurch(c)}</div>
                        <div className="text-[10px] text-slate-400 mt-0.5">
                          {st === 'submitted' ? '✅ 제출' : st === 'draft' ? '📝 작성중' : '⬜ 미작성'}
                        </div>
                        {ts && <div className="text-[9px] text-slate-300 truncate mt-0.5">{getRelativeTime(ts) || ts}</div>}
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* Drive status + guide */}
              <div className="mt-3 pt-3 border-t border-slate-100 space-y-2">
                {/* ngrok / 외부 서버 URL 설정 */}
                <div className="flex items-center gap-2">
                  <input
                    type="text"
                    defaultValue={localStorage.getItem('LOCAL_SERVER_URL') || ''}
                    placeholder="https://xxxx.ngrok-free.app (외부 AI 서버 URL)"
                    className="flex-1 min-w-0 text-xs px-2.5 py-1.5 border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-400"
                    onBlur={e => {
                      const val = e.target.value.trim().replace(/\/$/, '');
                      if (val) {
                        localStorage.setItem('LOCAL_SERVER_URL', val);
                        localStorage.setItem('IS_LOCAL_MODE', 'true');
                        setIsLocalMode(true);
                        toast.success('서버 URL이 저장되었습니다.');
                      } else {
                        localStorage.removeItem('LOCAL_SERVER_URL');
                      }
                    }}
                  />
                  <span className={`text-[10px] font-bold shrink-0 ${(isLocalMode || isCloudAiAvailable || isBrowserAiReady) ? 'text-emerald-600' : 'text-slate-400'}`}>
                    {isLocalMode ? '🟢 로컬 AI' : isCloudAiAvailable ? '🟢 클라우드 AI' : isBrowserAiReady ? '🟢 브라우저 AI' : '⚪ AI 꺼짐'}
                  </span>
                </div>

                {/* 브라우저 AI 다운로드 — PC/API 없이 사용 */}
                {!isLocalMode && !isCloudAiAvailable && (
                  <div className="rounded-lg border border-violet-200 bg-violet-50 px-3 py-2">
                    {isBrowserAiReady ? (
                      <div className="flex items-center gap-2 text-xs text-violet-700 font-bold">
                        <Bot className="w-3.5 h-3.5" /> 브라우저 AI 사용 중 (오프라인 가능)
                      </div>
                    ) : isBrowserAiLoading ? (
                      <div className="space-y-1.5">
                        <div className="flex items-center gap-2 text-xs text-violet-700 font-bold">
                          <Bot className="w-3.5 h-3.5 animate-spin" /> AI 모델 다운로드 중... {browserAiProgress}%
                        </div>
                        <div className="w-full bg-violet-100 rounded-full h-1.5">
                          <div className="bg-violet-500 h-1.5 rounded-full transition-all" style={{ width: `${browserAiProgress}%` }} />
                        </div>
                        <div className="text-[10px] text-violet-500">최초 1회만 다운로드 (약 300MB). 이후 오프라인 사용 가능.</div>
                      </div>
                    ) : (
                      <div className="space-y-1.5">
                        <div className="text-xs text-violet-700 font-bold">브라우저 AI 설치 (PC·API 불필요)</div>
                        <div className="text-[10px] text-violet-500">최초 1회 약 300MB 다운로드. 이후 인터넷 없이도 AI 검토 가능.</div>
                        <button
                          onClick={downloadBrowserAI}
                          className="w-full flex items-center justify-center gap-1.5 text-xs px-3 py-1.5 rounded-md bg-violet-600 hover:bg-violet-700 text-white font-bold transition-colors"
                        >
                          <Bot className="w-3.5 h-3.5" /> AI 모델 다운로드
                        </button>
                      </div>
                    )}
                  </div>
                )}

                <div className={`flex items-center justify-between gap-2 rounded-lg px-3 py-2 border text-xs ${
                  driveStatus?.authenticated ? 'bg-emerald-50 border-emerald-200' :
                  driveStatus?.configured ? 'bg-amber-50 border-amber-200' : 'bg-red-50 border-red-200'
                }`}>
                  <span className={`font-bold truncate ${driveStatus?.authenticated ? 'text-emerald-700' : driveStatus?.configured ? 'text-amber-700' : 'text-red-700'}`}>
                    Drive: {driveStatusLoading ? '확인 중...' : driveStatus?.authenticated ? '✅ 연결됨' : driveStatus?.configured ? '⚠️ 토큰 필요' : '❌ 미연결'}
                  </span>
                  <div className="flex gap-1 shrink-0">
                    <button onClick={checkDriveStatus} className="p-1 rounded hover:bg-white/60 text-slate-400 transition-colors" title="새로고침">
                      <RefreshCw className="w-3 h-3" />
                    </button>
                    <button
                      onClick={() => setShowAiStudioGuide(v => !v)}
                      className="px-2 py-1 bg-violet-600 hover:bg-violet-700 text-white rounded text-[10px] font-bold transition-colors flex items-center gap-1"
                    >
                      <Settings className="w-3 h-3" /> 설정
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
                            { icon: '📤', title: '보고서 제출', desc: '각 교구에서 제출하면 Drive에 자동 업로드됩니다' },
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

              {/* Progress indicator */}
              {adminCompilationProgress && (
                <div className="mt-3 flex items-center gap-2 p-3 bg-slate-50 border border-dashed border-slate-200 rounded-xl">
                  <div className="w-5 h-5 rounded-full border-2 border-purple-200 border-t-purple-600 animate-spin shrink-0" />
                  <p className="text-xs font-bold text-purple-700 animate-pulse">{adminCompilationProgress}</p>
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

      {/* AI Smart Paste Modal */}
      {showAiPasteModal && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-3xl flex flex-col max-h-[85vh]">
            <div className="flex items-center justify-between p-5 border-b border-slate-200 bg-gradient-to-r from-purple-600 to-indigo-600 text-white rounded-t-2xl">
              <h2 className="text-xl font-bold flex items-center gap-2">
                <Bot className="w-6 h-6 text-purple-100" /> AI 스마트 복붙 (자동정리)
              </h2>
              <button disabled={isAiProcessing} onClick={() => setShowAiPasteModal(false)} className="text-white/80 hover:bg-white/10 hover:text-white rounded-full p-2 transition-colors">
                <X className="w-5 h-5" />
              </button>
            </div>
            
            <div className="p-6 overflow-y-auto space-y-4 bg-slate-50 flex-1 flex flex-col min-h-0">
              <div className="bg-blue-50 border border-blue-100 p-3 rounded-lg flex items-start gap-2 text-sm text-blue-800">
                <AlertCircle className="w-5 h-5 shrink-0 mt-0.5" />
                <p>외부 문서(엑셀, 워드, 웹페이지)에서 표, 이미지, 텍스트를 그대로 복사하여 아래 영역에 붙여넣으세요. AI가 분석하여 주간보고 양식으로 완벽히 분리해 줍니다.</p>
              </div>

              <div 
                ref={aiPasteRef}
                contentEditable={!isAiProcessing}
                className={`flex-1 min-h-[300px] max-h-[500px] overflow-y-auto bg-white border-2 border-dashed ${isAiProcessing ? 'border-slate-300 bg-slate-100 text-slate-400' : 'border-purple-300 focus:border-purple-500 focus:ring-4 focus:ring-purple-100'} rounded-xl p-4 text-sm focus:outline-none transition-all cursor-text text-slate-800`}
                data-placeholder="여기를 클릭하고 복사한 내용을 붙여넣으세요 (Ctrl+V)"
                style={{ emptyCells: 'show' }}
              />
              <style>{`
                [contenteditable][data-placeholder]:empty:before {
                  content: attr(data-placeholder);
                  color: #94a3b8;
                  pointer-events: none;
                  display: block;
                }
              `}</style>
            </div>

            <div className="p-5 border-t border-slate-200 flex justify-end gap-3 bg-white rounded-b-2xl">
              <button 
                disabled={isAiProcessing}
                onClick={() => setShowAiPasteModal(false)}
                className="px-5 py-2.5 text-slate-600 hover:bg-slate-100 rounded-lg font-medium transition-colors"
              >
                취소
              </button>
              <button 
                disabled={isAiProcessing}
                onClick={handleAiPasteSubmit}
                className="px-6 py-2.5 bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-700 hover:to-indigo-700 text-white rounded-lg font-bold flex items-center gap-2 shadow-md transition-all disabled:opacity-50"
              >
                {isAiProcessing ? <RefreshCw className="w-5 h-5 animate-spin" /> : <Sparkles className="w-5 h-5" />}
                {isAiProcessing ? 'AI가 분석하고 정리하는 중...' : 'AI 마법으로 양식에 맞춰 넣기'}
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
              
              {role === 'church' && (
                <div className="space-y-6">
                  <section>
                    <h3 className="font-bold text-lg text-slate-900 border-b border-slate-200 pb-2 mb-3 flex items-center gap-2">
                      <span className="bg-blue-100 text-blue-700 w-6 h-6 rounded-full flex items-center justify-center text-sm font-bold">1</span>
                      간단하게 작성하고 모바일로 저장하세요
                    </h3>
                    <p className="text-sm leading-relaxed mb-2">교회장님들은 복잡한 작업 없이 가장 기본적이고 핵심적인 내용만 기입해주시면 됩니다. 스마트폰이나 태블릿에서도 쉽게 입력할 수 있습니다.</p>
                  </section>
                  <section>
                    <h3 className="font-bold text-lg text-slate-900 border-b border-slate-200 pb-2 mb-3 flex items-center gap-2">
                      <span className="bg-emerald-100 text-emerald-700 w-6 h-6 rounded-full flex items-center justify-center text-sm font-bold">2</span>
                      단축키 및 스마트 복붙(AI) 활용
                    </h3>
                    <ul className="list-disc list-inside space-y-2 text-sm ml-2 mb-2">
                      <li><kbd className="px-1.5 py-0.5 bg-slate-100 border border-slate-300 rounded font-mono text-xs">Enter</kbd> : 같은 수준의 새 항목을 아래에 추가</li>
                      <li><kbd className="px-1.5 py-0.5 bg-slate-100 border border-slate-300 rounded font-mono text-xs">Tab</kbd> / <kbd className="px-1.5 py-0.5 bg-slate-100 border border-slate-300 rounded font-mono text-xs">Shift + Tab</kbd> : 항목의 수준(Level) 변경</li>
                    </ul>
                    <div className="bg-purple-50 p-3 rounded text-sm text-purple-900 border border-purple-100 mt-3">
                      <strong>🤖 AI 스마트 복붙 기능:</strong> 카카오톡이나 밴드에 올라온 텍스트를 그대로 복사해서 빈 칸에 붙여넣으면, AI가 자동으로 개조식으로 정리해줍니다!
                    </div>
                  </section>
                  <section>
                    <h3 className="font-bold text-lg text-slate-900 border-b border-slate-200 pb-2 mb-3 flex items-center gap-2">
                      <span className="bg-amber-100 text-amber-700 w-6 h-6 rounded-full flex items-center justify-center text-sm font-bold">3</span>
                      '제출 확정' 필수
                    </h3>
                    <div className="flex items-start gap-3 bg-amber-50 p-4 rounded-lg border border-amber-100">
                      <AlertCircle className="w-5 h-5 text-amber-500 shrink-0 mt-0.5" />
                      <div className="text-sm text-slate-800 leading-relaxed">
                        작성이 모두 끝났다면 화면 하단 또는 우측 하단의 <strong className="text-emerald-700">✅ 제출 확정</strong> 버튼을 반드시 눌러주세요. 그래야 교구 사무장님에게 보고서가 전달됩니다.
                      </div>
                    </div>
                  </section>
                </div>
              )}

              {role === 'manager' && (
                <div className="space-y-6">
                  <section>
                    <h3 className="font-bold text-lg text-slate-900 border-b border-slate-200 pb-2 mb-3 flex items-center gap-2">
                      <span className="bg-blue-100 text-blue-700 w-6 h-6 rounded-full flex items-center justify-center text-sm font-bold">1</span>
                      교회별 제출 현황 확인
                    </h3>
                    <p className="text-sm leading-relaxed mb-2">상단의 '교회' 선택 박스 밑에 교구 내 모든 교회의 제출 현황이 표시됩니다. <strong className="text-emerald-600">초록색 점</strong>은 '제출 확정'된 교회이며, <strong className="text-slate-400">회색 점</strong>은 아직 작성하지 않은 교회입니다.</p>
                  </section>
                  <section>
                    <h3 className="font-bold text-lg text-slate-900 border-b border-slate-200 pb-2 mb-3 flex items-center gap-2">
                      <span className="bg-purple-100 text-purple-700 w-6 h-6 rounded-full flex items-center justify-center text-sm font-bold">2</span>
                      교구 전체 AI 검토 및 수정
                    </h3>
                    <p className="text-sm leading-relaxed mb-2">개별 교회가 올린 보고서를 확인하시면서 필요 시 직접 수정할 수 있습니다. 각 교회 보고서에서 필요 시 <strong className="text-purple-700">🤖 AI 문맥 검토</strong> 기능을 활용하여 오탈자나 어색한 문맥을 한 번에 정리할 수 있습니다.</p>
                  </section>
                  <section>
                    <h3 className="font-bold text-lg text-slate-900 border-b border-slate-200 pb-2 mb-3 flex items-center gap-2">
                      <span className="bg-emerald-100 text-emerald-700 w-6 h-6 rounded-full flex items-center justify-center text-sm font-bold">3</span>
                      본부로 일괄 다운로드 및 보고
                    </h3>
                    <div className="flex items-start gap-3 bg-emerald-50 p-4 rounded-lg border border-emerald-100">
                      <Check className="w-5 h-5 text-emerald-600 shrink-0 mt-0.5" />
                      <div className="text-sm text-slate-800 leading-relaxed">
                        교회들의 취합이 끝났다면, 우측 하단의 <strong className="text-blue-700">📥 Word로 다운로드</strong> 버튼을 눌러주세요.<br/>
                        현재 교구 내의 전체 교회 데이터가 <strong>단 1개의 파일(.docx)</strong>로 깔끔하게 병합되어 다운로드됩니다. 이 파일을 본부에 보고하시면 됩니다.
                      </div>
                    </div>
                  </section>
                </div>
              )}

              {role === 'admin' && (
                <div className="space-y-6">
                  <section>
                    <h3 className="font-bold text-lg text-slate-900 border-b border-slate-200 pb-2 mb-3 flex items-center gap-2">
                      <span className="bg-purple-100 text-purple-700 w-6 h-6 rounded-full flex items-center justify-center text-sm font-bold">1</span>
                      전국 교구 모니터링 (관리자 콘솔)
                    </h3>
                    <p className="text-sm leading-relaxed mb-2">상단의 <strong>[⚙️ 관리자 콘솔]</strong> 탭을 클릭하면 전국 모든 교구 부서의 실시간 작성 및 제출율을 한 눈에 확인할 수 있습니다.</p>
                  </section>
                  <section>
                    <h3 className="font-bold text-lg text-slate-900 border-b border-slate-200 pb-2 mb-3 flex items-center gap-2">
                      <span className="bg-red-100 text-red-700 w-6 h-6 rounded-full flex items-center justify-center text-sm font-bold">2</span>
                      매주 주간보고 사이클 갱신 (전체 초기화)
                    </h3>
                    <div className="flex items-start gap-3 bg-red-50 p-4 rounded-lg border border-red-100">
                      <AlertCircle className="w-5 h-5 text-red-600 shrink-0 mt-0.5" />
                      <div className="text-sm text-slate-800 leading-relaxed">
                        매주 새로운 보고를 취합받기 전, 반드시 <strong>[전체 데이터 초기화]</strong> 버튼을 눌러야 합니다.<br/><br/>
                        초기화 시 <strong>양력 날짜와 천일국 천력 날짜</strong>를 기입하게 되며, 이 설정값은 시스템 전역에 등록되어 모든 사용자의 화면 상단에 표시됩니다.
                      </div>
                    </div>
                  </section>
                  <section>
                    <h3 className="font-bold text-lg text-slate-900 border-b border-slate-200 pb-2 mb-3 flex items-center gap-2">
                      <span className="bg-blue-100 text-blue-700 w-6 h-6 rounded-full flex items-center justify-center text-sm font-bold">3</span>
                      공지사항 관리 및 데이터 백업
                    </h3>
                    <p className="text-sm leading-relaxed mb-2"><strong>[공지 작성]</strong> 탭에서 전국 교회장/사무장에게 전달할 공지사항(PDF 첨부 가능)을 등록할 수 있으며, <strong>[추출]</strong> 버튼을 통해 언제든지 전체 데이터를 원본 JSON 형태로 백업하실 수 있습니다.</p>
                  </section>
                </div>
              )}

              <section className="bg-indigo-50 border border-indigo-100 p-4 rounded-lg mt-6">
                <h3 className="font-bold text-base text-indigo-900 pb-2 mb-3 border-b border-indigo-200 flex items-center gap-2">
                  <Check className="w-4 h-4 text-indigo-700" /> 주간보고 세부항목 표준 배열 순서
                </h3>
                <p className="text-xs text-indigo-800 mb-3 font-medium leading-relaxed">
                  작성 시 아래 순서를 권장하며, 순서가 섞이더라도 <strong>'AI 스마트 복붙'</strong>이나 <strong>'AI 문맥 검토'</strong>를 사용하시면 알아서 표준 순서에 맞춰 깔끔하게 재배치됩니다. (※ 사진은 무조건 맨 뒤로 자동 이동)
                </p>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-2 text-xs text-indigo-950 leading-relaxed font-medium">
                  <p><span className="text-[10px] bg-indigo-200 text-indigo-800 px-1.5 py-0.5 rounded font-bold mr-1.5">1순위</span> 일시</p>
                  <p><span className="text-[10px] bg-indigo-200 text-indigo-800 px-1.5 py-0.5 rounded font-bold mr-1.5">2순위</span> 장소</p>
                  <p><span className="text-[10px] bg-indigo-200 text-indigo-800 px-1.5 py-0.5 rounded font-bold mr-1.5">3순위</span> 대상 및 참석인원</p>
                  <p><span className="text-[10px] bg-indigo-200 text-indigo-800 px-1.5 py-0.5 rounded font-bold mr-1.5">4순위</span> 주제 및 목적</p>
                  <p><span className="text-[10px] bg-indigo-200 text-indigo-800 px-1.5 py-0.5 rounded font-bold mr-1.5">5순위</span> 주요 내용</p>
                  <p><span className="text-[10px] bg-indigo-200 text-indigo-800 px-1.5 py-0.5 rounded font-bold mr-1.5">6순위</span> 결과 및 향후 계획</p>
                  <p className="sm:col-span-2 text-indigo-700 mt-1 flex items-center gap-1">
                    <span className="text-[10px] bg-indigo-600 text-white px-1.5 py-0.5 rounded font-bold mr-1">마지막</span>
                    사진 및 첨부자료 <ImageIcon className="w-3.5 h-3.5 inline"/> (항상 최하단 배치)
                  </p>
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

      {/* 교회 보고서 미리보기 모달 */}
      {previewChurchModal && (
        <div className="fixed inset-0 z-[70] flex items-center justify-center p-4 bg-slate-900/70 backdrop-blur-sm" onClick={() => setPreviewChurchModal(null)}>
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg flex flex-col max-h-[85vh]" onClick={e => e.stopPropagation()}>
            {/* 헤더 */}
            <div className={`px-5 py-4 flex items-center justify-between rounded-t-2xl ${
              previewChurchModal.status === 'submitted' ? 'bg-emerald-600' :
              previewChurchModal.status === 'draft' ? 'bg-blue-600' : 'bg-slate-600'
            }`}>
              <div>
                <h3 className="text-base font-black text-white">{getDisplayParish(previewChurchModal.parish)} · {getDisplayChurch(previewChurchModal.church)}</h3>
                <p className="text-white/70 text-xs mt-0.5 flex items-center gap-1.5">
                  {previewChurchModal.status === 'submitted' ? <><CheckCircle className="w-3 h-3" />제출완료</> :
                   previewChurchModal.status === 'draft' ? <><Clock className="w-3 h-3" />작성중</> : '미작성'}
                  {previewChurchModal.timestamp && <span> · {getRelativeTime(previewChurchModal.timestamp) || previewChurchModal.timestamp}</span>}
                </p>
              </div>
              <button onClick={() => setPreviewChurchModal(null)} className="text-white/70 hover:text-white rounded-full p-1.5 hover:bg-white/20 transition-colors"><X className="w-5 h-5" /></button>
            </div>
            {/* 내용 */}
            <div className="overflow-y-auto flex-1 p-5 space-y-1">
              {previewChurchModal.data.length === 0 ? (
                <p className="text-slate-400 text-sm text-center py-8">작성된 내용이 없습니다.</p>
              ) : previewChurchModal.data.filter((item: any) => item.text?.trim()).map((item: any) => (
                <div key={item.id} className={`flex gap-2 text-sm ${
                  item.level === 0 ? 'mt-3' : ''
                }`}>
                  <span className="shrink-0 text-slate-400" style={{ marginLeft: `${item.level * 14}px` }}>
                    {item.level === 0 ? '▌' : item.level === 1 ? '•' : '–'}
                  </span>
                  <span className={`${item.level === 0 ? 'font-black text-slate-800' : item.level === 1 ? 'font-semibold text-slate-700' : 'text-slate-600'} leading-snug`}>
                    {item.text}
                  </span>
                </div>
              ))}
            </div>
            {/* 푸터 */}
            <div className="px-5 py-3 border-t border-slate-100 flex justify-end">
              <button
                onClick={() => { setActiveTab('report'); setParish(previewChurchModal.parish); setChurch(previewChurchModal.church); setPreviewChurchModal(null); }}
                className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-sm font-bold transition-colors flex items-center gap-1.5"
              >
                <ArrowRight className="w-3.5 h-3.5" /> 에디터에서 열기
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Reset Modal */}
      {/* 내 교회 변경 모달 */}
      {showChurchChangeModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm" onClick={() => setShowChurchChangeModal(false)}>
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-sm overflow-hidden border border-slate-100" onClick={e => e.stopPropagation()}>
            <div className="bg-amber-600 px-5 py-4 flex items-center justify-between">
              <div>
                <h3 className="text-base font-black text-white flex items-center gap-2"><ArrowRight className="w-4 h-4"/> 내 교회 변경</h3>
                <p className="text-amber-200 text-xs mt-0.5">교구와 교회를 선택하세요</p>
              </div>
              <button onClick={() => setShowChurchChangeModal(false)} className="text-white/70 hover:text-white rounded-full p-1.5 hover:bg-white/20 transition-colors"><X className="w-5 h-5" /></button>
            </div>
            <div className="p-5 space-y-4">
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1.5">교구 선택</label>
                <select
                  className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-amber-500 outline-none bg-white"
                  value={churchChangeParish}
                  onChange={e => {
                    const p = e.target.value;
                    setChurchChangeParish(p);
                    setChurchChangeChurch(PARISH_CHURCH_MAP[p][0]);
                  }}
                >
                  {Object.keys(PARISH_CHURCH_MAP).map(p => (
                    <option key={p} value={p}>{getDisplayParish(p)}</option>
                  ))}
                </select>
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1.5">교회 선택</label>
                <select
                  className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-amber-500 outline-none bg-white"
                  value={churchChangeChurch}
                  onChange={e => setChurchChangeChurch(e.target.value)}
                >
                  {(PARISH_CHURCH_MAP[churchChangeParish] || []).map(c => (
                    <option key={c} value={c}>{getDisplayChurch(c)}</option>
                  ))}
                </select>
              </div>
              <div className="flex gap-2 pt-2">
                <button
                  onClick={() => setShowChurchChangeModal(false)}
                  className="flex-1 py-2.5 rounded-lg border border-slate-200 text-slate-600 text-sm font-medium hover:bg-slate-50 transition-colors"
                >
                  취소
                </button>
                <button
                  onClick={() => {
                    localStorage.setItem('APP_PARISH', churchChangeParish);
                    localStorage.setItem('APP_CHURCH', churchChangeChurch);
                    setParish(churchChangeParish);
                    setChurch(churchChangeChurch);
                    setShowChurchChangeModal(false);
                  }}
                  className="flex-1 py-2.5 rounded-lg bg-amber-600 hover:bg-amber-700 text-white text-sm font-bold transition-colors flex items-center justify-center gap-2"
                >
                  <ArrowRight className="w-4 h-4"/> 변경하기
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {showResetModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm" onClick={() => setShowResetModal(false)}>
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl flex flex-col max-h-[90vh]" onClick={e => e.stopPropagation()}>
            {/* Header */}
            <div className="flex items-center justify-between px-6 py-4 border-b border-slate-200 bg-red-50 rounded-t-2xl">
              <h2 className="text-lg font-bold text-red-700 flex items-center gap-2">
                <Trash2 className="w-5 h-5" /> 데이터 초기화
              </h2>
              <button onClick={() => setShowResetModal(false)} className="text-slate-400 hover:text-slate-600 rounded-full p-1.5 hover:bg-slate-200 transition-colors"><X className="w-5 h-5" /></button>
            </div>

            {/* Mode Tabs */}
            <div className="flex border-b border-slate-200 bg-slate-50">
              <button
                onClick={() => setResetMode('quick')}
                className={`flex-1 py-2.5 text-sm font-semibold transition-colors ${resetMode === 'quick' ? 'text-red-600 border-b-2 border-red-500 bg-white' : 'text-slate-500 hover:text-slate-700'}`}
              >빠른 선택</button>
              <button
                onClick={() => setResetMode('custom')}
                className={`flex-1 py-2.5 text-sm font-semibold transition-colors ${resetMode === 'custom' ? 'text-red-600 border-b-2 border-red-500 bg-white' : 'text-slate-500 hover:text-slate-700'}`}
              >교회별 선택</button>
            </div>

            <div className="overflow-y-auto flex-1 p-5">
              {resetMode === 'quick' ? (
                <div className="space-y-3">
                  <p className="text-xs text-slate-500 mb-4">초기화할 범위를 선택하세요. 이 작업은 되돌릴 수 없습니다.</p>

                  {/* 현재 교회 */}
                  <button
                    onClick={() => { executeReset([{parish, church}], false); }}
                    className="w-full flex items-center gap-3 p-4 rounded-xl border-2 border-slate-200 hover:border-red-300 hover:bg-red-50 transition-all text-left group"
                  >
                    <div className="w-10 h-10 rounded-full bg-slate-100 group-hover:bg-red-100 flex items-center justify-center shrink-0">
                      <Trash2 className="w-5 h-5 text-slate-400 group-hover:text-red-500" />
                    </div>
                    <div>
                      <div className="font-bold text-slate-800 text-sm">현재 교회만 초기화</div>
                      <div className="text-xs text-slate-500 mt-0.5">{parish} · <span className="font-medium text-red-600">{church}</span> 1개 교회</div>
                    </div>
                  </button>

                  {/* 현재 교구 전체 */}
                  <button
                    onClick={() => { const targets = PARISH_CHURCH_MAP[parish].map(c => ({parish, church: c})); executeReset(targets, true); }}
                    className="w-full flex items-center gap-3 p-4 rounded-xl border-2 border-slate-200 hover:border-orange-300 hover:bg-orange-50 transition-all text-left group"
                  >
                    <div className="w-10 h-10 rounded-full bg-slate-100 group-hover:bg-orange-100 flex items-center justify-center shrink-0">
                      <Trash2 className="w-5 h-5 text-slate-400 group-hover:text-orange-500" />
                    </div>
                    <div>
                      <div className="font-bold text-slate-800 text-sm">현재 교구 전체 초기화</div>
                      <div className="text-xs text-slate-500 mt-0.5"><span className="font-medium text-orange-600">{parish}</span> 교구 전체 {PARISH_CHURCH_MAP[parish]?.length}개 교회</div>
                    </div>
                  </button>

                  {/* 전체 교구 초기화 */}
                  <button
                    onClick={() => { const targets = Object.keys(PARISH_CHURCH_MAP).flatMap(p => PARISH_CHURCH_MAP[p].map(c => ({parish: p, church: c}))); executeReset(targets, true); }}
                    className="w-full flex items-center gap-3 p-4 rounded-xl border-2 border-slate-200 hover:border-red-400 hover:bg-red-50 transition-all text-left group"
                  >
                    <div className="w-10 h-10 rounded-full bg-slate-100 group-hover:bg-red-100 flex items-center justify-center shrink-0">
                      <Trash2 className="w-5 h-5 text-slate-400 group-hover:text-red-600" />
                    </div>
                    <div>
                      <div className="font-bold text-red-700 text-sm">⚠️ 전체 초기화</div>
                      <div className="text-xs text-slate-500 mt-0.5">모든 교구 <span className="font-medium text-red-600">{Object.values(PARISH_CHURCH_MAP).flat().length}개</span> 전체 — 비밀번호 필요</div>
                    </div>
                  </button>
                </div>
              ) : (
                /* Custom church selection */
                <div>
                  <p className="text-xs text-slate-500 mb-3">초기화할 교구와 교회를 직접 선택하세요.</p>
                  <div className="space-y-3">
                    {Object.keys(PARISH_CHURCH_MAP).map(p => {
                      const parishChurches = PARISH_CHURCH_MAP[p];
                      const selectedChurchesInParish = (resetSelectedChurches[p] || []);
                      const allSelected = selectedChurchesInParish.length === parishChurches.length;
                      const someSelected = selectedChurchesInParish.length > 0 && !allSelected;
                      return (
                        <div key={p} className="border border-slate-200 rounded-xl overflow-hidden">
                          {/* Parish header */}
                          <button
                            onClick={() => {
                              if (allSelected) {
                                setResetSelectedChurches(prev => { const n = {...prev}; delete n[p]; return n; });
                              } else {
                                setResetSelectedChurches(prev => ({...prev, [p]: [...parishChurches]}));
                              }
                            }}
                            className={`w-full flex items-center gap-2 px-4 py-2.5 text-sm font-bold transition-colors ${allSelected ? 'bg-red-50 text-red-700' : someSelected ? 'bg-orange-50 text-orange-700' : 'bg-slate-50 text-slate-700 hover:bg-slate-100'}`}
                          >
                            <div className={`w-4 h-4 rounded border-2 flex items-center justify-center ${allSelected ? 'bg-red-500 border-red-500' : someSelected ? 'bg-orange-400 border-orange-400' : 'border-slate-300'}`}>
                              {(allSelected || someSelected) && <Check className="w-3 h-3 text-white" />}
                            </div>
                            {p}
                            <span className="ml-auto text-xs font-normal text-slate-400">{selectedChurchesInParish.length}/{parishChurches.length}</span>
                          </button>
                          {/* Church list */}
                          <div className="flex flex-wrap gap-1.5 p-3 bg-white">
                            {parishChurches.map(c => {
                              const isSelected = selectedChurchesInParish.includes(c);
                              return (
                                <button
                                  key={c}
                                  onClick={() => {
                                    setResetSelectedChurches(prev => {
                                      const cur = prev[p] || [];
                                      return {...prev, [p]: isSelected ? cur.filter(x => x !== c) : [...cur, c]};
                                    });
                                  }}
                                  className={`px-2.5 py-1 rounded-full text-xs font-medium border transition-colors ${isSelected ? 'bg-red-500 border-red-500 text-white' : 'bg-white border-slate-200 text-slate-600 hover:border-red-300 hover:text-red-600'}`}
                                >{c}</button>
                              );
                            })}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>

            {/* Footer */}
            {resetMode === 'custom' && (
              <div className="px-5 py-4 border-t border-slate-200 bg-slate-50 rounded-b-2xl flex items-center justify-between gap-3">
                <span className="text-sm text-slate-500">
                  {Object.values(resetSelectedChurches).flat().length}개 선택됨
                </span>
                <div className="flex gap-2">
                  <button onClick={() => setShowResetModal(false)} className="px-4 py-2 text-sm rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-100 transition-colors">취소</button>
                  <button
                    disabled={Object.values(resetSelectedChurches).flat().length === 0}
                    onClick={() => {
                      const targets = Object.entries(resetSelectedChurches).flatMap(([p, cs]) => (cs as string[]).map(c => ({parish: p, church: c})));
                      const requirePwd = targets.length > 1;
                      executeReset(targets, requirePwd);
                    }}
                    className="px-4 py-2 text-sm rounded-lg bg-red-600 hover:bg-red-700 disabled:opacity-40 text-white font-bold transition-colors flex items-center gap-1.5"
                  >
                    <Trash2 className="w-4 h-4" /> 선택 초기화
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      )}

    </div>
    </div>

    {/* 제출 완료 축하 배너 */}
    {showSubmitCelebration && (
      <div className="fixed inset-0 pointer-events-none z-[400] flex items-center justify-center">
        <div className="bg-gradient-to-r from-emerald-500 to-teal-500 text-white px-8 py-5 rounded-2xl shadow-2xl flex flex-col items-center gap-2 animate-in zoom-in-95 duration-300">
          <div className="text-4xl">🎉</div>
          <div className="text-xl font-black">{church} 보고서 제출 완료!</div>
          <div className="text-sm text-emerald-100">{parish} · {appConfig?.solarDate || ''}</div>
        </div>
      </div>
    )}

    {/* 범용 비밀번호 모달 */}
    {showPasswordModal && (
      <div className="fixed inset-0 bg-black/60 z-[350] flex items-center justify-center p-4" onClick={() => setShowPasswordModal(false)}>
        <div className="bg-white rounded-2xl shadow-2xl max-w-sm w-full overflow-hidden" onClick={e => e.stopPropagation()}>
          <div className="bg-gradient-to-br from-slate-700 to-slate-800 p-4 text-white flex items-center gap-3">
            <div className="bg-white/20 p-2 rounded-lg"><Key className="w-4 h-4" /></div>
            <h2 className="font-black text-base">{passwordModalTitle}</h2>
          </div>
          <div className="p-4 space-y-3">
            <input
              type="password"
              placeholder="비밀번호 입력"
              autoFocus
              className={`w-full border rounded-xl px-3.5 py-2.5 text-sm focus:ring-2 outline-none transition-all ${passwordModalError ? 'border-red-400 focus:ring-red-200' : 'border-slate-300 focus:ring-slate-300'}`}
              value={passwordModalInput}
              onChange={e => { setPasswordModalInput(e.target.value); setPasswordModalError(''); }}
              onKeyDown={e => e.key === 'Enter' && submitPasswordModal()}
            />
            {passwordModalError && (
              <p className="text-xs text-red-600 font-semibold flex items-center gap-1.5">
                <AlertCircle className="w-3.5 h-3.5 shrink-0" />{passwordModalError}
              </p>
            )}
            <div className="flex gap-2">
              <button onClick={() => setShowPasswordModal(false)} className="flex-1 py-2.5 border border-slate-300 text-slate-600 rounded-xl text-sm font-bold hover:bg-slate-50 transition-colors">취소</button>
              <button onClick={submitPasswordModal} className="flex-1 py-2.5 bg-slate-700 hover:bg-slate-800 text-white rounded-xl text-sm font-black transition-colors">확인</button>
            </div>
          </div>
        </div>
      </div>
    )}

    {/* 새 주간보고 시작 모달 */}
    {showNewWeekModal && (
      <div className="fixed inset-0 bg-black/60 z-[300] flex items-center justify-center p-4" onClick={() => !newWeekIsResetting && setShowNewWeekModal(false)}>
        <div className="bg-white rounded-2xl shadow-2xl max-w-md w-full overflow-hidden" onClick={e => e.stopPropagation()}>
          <div className="bg-gradient-to-br from-red-600 to-red-700 p-5 text-white">
            <div className="flex items-center gap-3">
              <div className="bg-white/20 p-2.5 rounded-xl">
                <RefreshCw className="w-5 h-5" />
              </div>
              <div>
                <h2 className="text-lg font-black">새 주간보고 시작</h2>
                <p className="text-red-100 text-xs mt-0.5">모든 교구/교회 데이터가 초기화됩니다</p>
              </div>
            </div>
          </div>
          <div className="p-5 space-y-4">
            <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 flex items-start gap-2">
              <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
              <p className="text-xs text-amber-800 leading-relaxed">이 작업을 실행하면 <strong>전 교구·전 교회의 주간보고 데이터</strong>가 모두 초기화되며 복구할 수 없습니다.</p>
            </div>

            <div>
              <label className="block text-xs font-bold text-slate-600 mb-1.5">관리자 비밀번호 <span className="text-red-500">*</span></label>
              <input
                type="password"
                placeholder="비밀번호 입력"
                className={`w-full border rounded-xl px-3.5 py-2.5 text-sm focus:ring-2 outline-none transition-all ${newWeekPasswordError && newWeekPassword !== 'chongmu2027' ? 'border-red-400 focus:ring-red-200' : 'border-slate-300 focus:ring-red-200 focus:border-red-400'}`}
                value={newWeekPassword}
                onChange={e => { setNewWeekPassword(e.target.value); setNewWeekPasswordError(''); }}
                autoFocus
              />
            </div>

            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1.5">양력 날짜 <span className="text-red-500">*</span></label>
                <input
                  type="text"
                  placeholder="예: 5월 4주차"
                  className="w-full border border-slate-300 rounded-xl px-3.5 py-2.5 text-sm focus:ring-2 focus:ring-red-200 focus:border-red-400 outline-none"
                  value={newWeekSolarDate}
                  onChange={e => { setNewWeekSolarDate(e.target.value); setNewWeekPasswordError(''); }}
                />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1.5">천력 날짜 <span className="text-red-500">*</span></label>
                <input
                  type="text"
                  placeholder="예: 4월 11일"
                  className="w-full border border-slate-300 rounded-xl px-3.5 py-2.5 text-sm focus:ring-2 focus:ring-red-200 focus:border-red-400 outline-none"
                  value={newWeekHeavenlyDate}
                  onChange={e => { setNewWeekHeavenlyDate(e.target.value); setNewWeekPasswordError(''); }}
                />
              </div>
            </div>
            <div>
              <label className="block text-xs font-bold text-slate-600 mb-1.5">제출 마감일 <span className="text-slate-400 font-normal">(선택)</span></label>
              <input
                type="text"
                placeholder="예: 목요일 오후 6시"
                className="w-full border border-slate-300 rounded-xl px-3.5 py-2.5 text-sm focus:ring-2 focus:ring-red-200 focus:border-red-400 outline-none"
                value={newWeekDeadline}
                onChange={e => setNewWeekDeadline(e.target.value)}
              />
            </div>

            {newWeekPasswordError && (
              <p className="text-xs text-red-600 font-semibold flex items-center gap-1.5">
                <AlertCircle className="w-3.5 h-3.5 shrink-0" />{newWeekPasswordError}
              </p>
            )}

            {newWeekSolarDate && newWeekHeavenlyDate && newWeekPassword === 'chongmu2027' && (
              <div className="bg-slate-50 border border-slate-200 rounded-xl p-3 text-xs text-slate-600 leading-relaxed">
                <span className="font-bold text-slate-800">[{newWeekSolarDate.trim()}]</span> (천력 {newWeekHeavenlyDate.trim()}) 주간보고를 새로 시작합니다.<br/>
                전체 교구의 보고서가 초기화됩니다.
              </div>
            )}

            <div className="flex gap-2 pt-1">
              <button
                onClick={() => setShowNewWeekModal(false)}
                disabled={newWeekIsResetting}
                className="flex-1 py-2.5 border border-slate-300 text-slate-600 rounded-xl text-sm font-bold hover:bg-slate-50 transition-colors disabled:opacity-50"
              >
                취소
              </button>
              <button
                onClick={executeNewWeekReset}
                disabled={newWeekIsResetting || !newWeekSolarDate.trim() || !newWeekHeavenlyDate.trim()}
                className="flex-1 py-2.5 bg-red-600 hover:bg-red-700 disabled:bg-red-300 text-white rounded-xl text-sm font-black transition-colors flex items-center justify-center gap-2"
              >
                {newWeekIsResetting ? (
                  <><RefreshCw className="w-4 h-4 animate-spin" /> 초기화 중...</>
                ) : (
                  <><RefreshCw className="w-4 h-4" /> 새 주간보고 시작</>
                )}
              </button>
            </div>
          </div>
        </div>
      </div>
    )}

    {/* 토스트 알림 컨테이너 */}
    <div className="fixed bottom-24 left-1/2 -translate-x-1/2 z-[200] flex flex-col items-center gap-2 pointer-events-none md:bottom-6">
      {toasts.map(t => {
        const styles: Record<ToastType, string> = {
          success: 'bg-emerald-600 text-white',
          error: 'bg-red-600 text-white',
          info: 'bg-blue-600 text-white',
          warning: 'bg-amber-500 text-white',
        };
        const icons: Record<ToastType, React.ReactNode> = {
          success: <CheckCircle className="w-4 h-4 shrink-0" />,
          error: <AlertCircle className="w-4 h-4 shrink-0" />,
          info: <Info className="w-4 h-4 shrink-0" />,
          warning: <AlertTriangle className="w-4 h-4 shrink-0" />,
        };
        return (
          <div
            key={t.id}
            className={`${styles[t.type]} px-4 py-2.5 rounded-xl shadow-2xl flex items-center gap-2.5 text-sm font-medium max-w-xs text-center animate-in fade-in slide-in-from-bottom-3 duration-300 pointer-events-auto`}
          >
            {icons[t.type]}
            <span>{t.message}</span>
          </div>
        );
      })}
    </div>
    </>
  );
}
