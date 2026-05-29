import React, { useState } from 'react';
import { User, Users, Shield, ArrowRight, CheckCircle } from 'lucide-react';

export type Role = 'church' | 'manager' | 'admin' | null;

interface RoleSelectionProps {
  onSelectRole: (role: Role, data?: any) => void;
  parishChurchMap: Record<string, string[]>;
  appConfig?: { solarDate: string; heavenlyDate: string; deadline?: string } | null;
}

const MANAGER_CODE_MAP: Record<string, string> = {
  'seoulbukbu': '서울북부',
  'seoulnambu': '서울남부',
  'gyeonggibukbu': '경기북부',
  'incheongyeonggiseobu': '인천경기서부',
  'gyeongginambu': '경기남부',
  'gangwon': '강원',
  'daejeonchungnam': '대전충남',
  'chungbuk': '충북',
  'jeonbuk': '전북',
  'gwangjujeonnamjeju': '광주전남제주',
  'daegugyeongbuk': '대구경북',
  'gyeongnam': '경남',
  'busanulsan': '부산울산',
  'cheonwonteukbyeol': '천원특별',
  'hyeophoe': '협회'
};

export default function RoleSelection({ onSelectRole, parishChurchMap, appConfig }: RoleSelectionProps) {
  const [selectedRole, setSelectedRole] = useState<Role>(null);

  const savedParish = localStorage.getItem('APP_PARISH') || '천원특별';
  const savedChurch = localStorage.getItem('APP_CHURCH') || parishChurchMap['천원특별'][0];
  const savedManagerCode = localStorage.getItem('APP_MANAGER_CODE') || '';
  const savedManagerParish = MANAGER_CODE_MAP[savedManagerCode] || '';

  const [parish, setParish] = useState(() => {
    const p = savedParish;
    return parishChurchMap[p] ? p : '천원특별';
  });
  const [church, setChurch] = useState(() => {
    const p = savedParish;
    const validParish = parishChurchMap[p] ? p : '천원특별';
    return parishChurchMap[validParish].includes(savedChurch) ? savedChurch : parishChurchMap[validParish][0];
  });
  const [managerCode, setManagerCode] = useState(savedManagerCode);
  const [adminCode, setAdminCode] = useState('');

  const handleStartChurch = () => {
    onSelectRole('church', { parish, church });
  };

  const handleStartManager = (code?: string) => {
    const codeToUse = (code ?? managerCode).toLowerCase();
    const matchedParish = MANAGER_CODE_MAP[codeToUse];
    if (matchedParish) {
      localStorage.setItem('APP_MANAGER_CODE', codeToUse);
      onSelectRole('manager', { parish: matchedParish });
    } else {
      alert('올바르지 않은 사무장 교구 코드입니다.');
    }
  };

  const handleStartAdmin = () => {
    if (adminCode === 'skmt0909!') {
      onSelectRole('admin');
    } else {
      alert('비밀번호가 일치하지 않습니다.');
    }
  };

  return (
    <div className="min-h-screen bg-gray-100 flex flex-col items-center justify-center p-4">
      <div className="max-w-md w-full bg-white rounded-2xl shadow-xl overflow-hidden border border-slate-100">
        <div className="bg-blue-800 p-8 text-center">
          <h1 className="text-2xl font-bold text-white mb-2">주간업무보고 시스템</h1>
          <p className="text-blue-200 text-sm">접속하실 모드를 선택해 주세요</p>
          {appConfig && (
            <div className="mt-3 inline-flex items-center gap-1.5 bg-blue-700/60 text-blue-100 text-xs font-semibold px-3 py-1.5 rounded-full border border-blue-600/50">
              <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse"></span>
              {appConfig.solarDate} (천력 {appConfig.heavenlyDate}) 취합 중
            </div>
          )}
        </div>

        <div className="p-6 space-y-6">
          {/* 교회장 모드 */}
          <div
            className={`border rounded-xl p-4 cursor-pointer transition-all ${selectedRole === 'church' ? 'border-blue-500 bg-blue-50 ring-2 ring-blue-200' : 'border-slate-200 hover:border-blue-300'}`}
            onClick={() => setSelectedRole('church')}
          >
            <div className="flex items-center gap-3 mb-2">
              <div className={`p-2 rounded-lg ${selectedRole === 'church' ? 'bg-blue-600 text-white' : 'bg-slate-100 text-slate-600'}`}>
                <User className="w-5 h-5" />
              </div>
              <div className="flex-1">
                <h3 className="font-bold text-slate-800">교회장 모드</h3>
                <p className="text-xs text-slate-500">내 교회 보고서 작성 및 관리</p>
              </div>
              {savedChurch && savedParish && selectedRole !== 'church' && (
                <span className="text-[10px] bg-blue-50 border border-blue-200 text-blue-600 px-2 py-0.5 rounded-full font-semibold">{savedParish} · {savedChurch}</span>
              )}
            </div>
            {selectedRole === 'church' && (
              <div className="mt-4 space-y-3 animate-in fade-in slide-in-from-top-2">
                <div>
                  <label className="block text-xs font-semibold text-slate-600 mb-1">교구 선택</label>
                  <select
                    className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 outline-none"
                    value={parish}
                    onChange={(e) => {
                      setParish(e.target.value);
                      setChurch(parishChurchMap[e.target.value][0]);
                    }}
                  >
                    {Object.keys(parishChurchMap).map(p => <option key={p} value={p}>{p}</option>)}
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-600 mb-1">교회 선택</label>
                  <select
                    className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 outline-none"
                    value={church}
                    onChange={(e) => setChurch(e.target.value)}
                  >
                    {parishChurchMap[parish].map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
                <button
                  onClick={(e) => { e.stopPropagation(); handleStartChurch(); }}
                  className="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-2.5 rounded-lg flex items-center justify-center gap-2 mt-2"
                >
                  시작하기 <ArrowRight className="w-4 h-4" />
                </button>
              </div>
            )}
          </div>

          {/* 사무장 모드 */}
          <div
            className={`border rounded-xl p-4 cursor-pointer transition-all ${selectedRole === 'manager' ? 'border-indigo-500 bg-indigo-50 ring-2 ring-indigo-200' : 'border-slate-200 hover:border-indigo-300'}`}
            onClick={() => setSelectedRole('manager')}
          >
            <div className="flex items-center gap-3 mb-2">
              <div className={`p-2 rounded-lg ${selectedRole === 'manager' ? 'bg-indigo-600 text-white' : 'bg-slate-100 text-slate-600'}`}>
                <Users className="w-5 h-5" />
              </div>
              <div className="flex-1">
                <h3 className="font-bold text-slate-800">사무장 모드</h3>
                <p className="text-xs text-slate-500">교구 전체 관리 및 취합</p>
              </div>
              {savedManagerParish && selectedRole !== 'manager' && (
                <span className="text-[10px] bg-indigo-50 border border-indigo-200 text-indigo-600 px-2 py-0.5 rounded-full font-semibold">{savedManagerParish}</span>
              )}
            </div>
            {selectedRole === 'manager' && (
              <div className="mt-4 space-y-3 animate-in fade-in slide-in-from-top-2">
                {savedManagerParish ? (
                  <div className="flex items-center gap-2 bg-indigo-50 border border-indigo-200 rounded-lg px-3 py-2.5">
                    <CheckCircle className="w-4 h-4 text-indigo-500 shrink-0" />
                    <div className="flex-1 min-w-0">
                      <p className="text-xs font-bold text-indigo-700">저장된 교구: {savedManagerParish}</p>
                      <p className="text-[10px] text-indigo-500">아래 버튼으로 바로 입장하거나, 코드를 변경하세요</p>
                    </div>
                  </div>
                ) : null}
                <div>
                  <label className="block text-xs font-semibold text-slate-600 mb-1">교구 영문 코드</label>
                  <input
                    type="text"
                    placeholder="예: seoulbukbu"
                    className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none"
                    value={managerCode}
                    onChange={(e) => setManagerCode(e.target.value)}
                    onKeyDown={(e) => e.key === 'Enter' && handleStartManager()}
                  />
                </div>
                {savedManagerParish && (
                  <button
                    onClick={(e) => { e.stopPropagation(); handleStartManager(savedManagerCode); }}
                    className="w-full bg-indigo-100 hover:bg-indigo-200 text-indigo-800 border border-indigo-300 font-bold py-2 rounded-lg flex items-center justify-center gap-2"
                  >
                    <CheckCircle className="w-4 h-4" />
                    {savedManagerParish} 바로 입장
                  </button>
                )}
                <button
                  onClick={(e) => { e.stopPropagation(); handleStartManager(); }}
                  className="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2.5 rounded-lg flex items-center justify-center gap-2"
                >
                  인증하기 <ArrowRight className="w-4 h-4" />
                </button>
              </div>
            )}
          </div>

          {/* 최고 관리자 모드 */}
          <div
            className={`border rounded-xl p-4 cursor-pointer transition-all ${selectedRole === 'admin' ? 'border-purple-500 bg-purple-50 ring-2 ring-purple-200' : 'border-slate-200 hover:border-purple-300'}`}
            onClick={() => setSelectedRole('admin')}
          >
            <div className="flex items-center gap-3 mb-2">
              <div className={`p-2 rounded-lg ${selectedRole === 'admin' ? 'bg-purple-600 text-white' : 'bg-slate-100 text-slate-600'}`}>
                <Shield className="w-5 h-5" />
              </div>
              <div>
                <h3 className="font-bold text-slate-800">최고 관리자 모드</h3>
                <p className="text-xs text-slate-500">전체 취합 및 AI 검토</p>
              </div>
            </div>
            {selectedRole === 'admin' && (
              <div className="mt-4 space-y-3 animate-in fade-in slide-in-from-top-2">
                <div>
                  <label className="block text-xs font-semibold text-slate-600 mb-1">관리자 비밀번호</label>
                  <input
                    type="password"
                    placeholder="비밀번호 입력"
                    className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-purple-500 outline-none"
                    value={adminCode}
                    onChange={(e) => setAdminCode(e.target.value)}
                    onKeyDown={(e) => e.key === 'Enter' && handleStartAdmin()}
                  />
                </div>
                <button
                  onClick={(e) => { e.stopPropagation(); handleStartAdmin(); }}
                  className="w-full bg-purple-600 hover:bg-purple-700 text-white font-bold py-2.5 rounded-lg flex items-center justify-center gap-2 mt-2"
                >
                  로그인 <ArrowRight className="w-4 h-4" />
                </button>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
