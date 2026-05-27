import React, { useState } from 'react';
import { User, Users, Shield, ArrowRight } from 'lucide-react';

export type Role = 'church' | 'manager' | 'admin' | null;

interface RoleSelectionProps {
  onSelectRole: (role: Role, data?: any) => void;
  parishChurchMap: Record<string, string[]>;
  appConfig?: { solarDate: string; heavenlyDate: string } | null;
}

export default function RoleSelection({ onSelectRole, parishChurchMap, appConfig }: RoleSelectionProps) {
  const [selectedRole, setSelectedRole] = useState<Role>(null);
  const [parish, setParish] = useState('천원특별');
  const [church, setChurch] = useState(parishChurchMap['천원특별'][0]);
  const [managerCode, setManagerCode] = useState('');
  const [adminCode, setAdminCode] = useState('');

  const handleStartChurch = () => {
    onSelectRole('church', { parish, church });
  };

  const handleStartManager = () => {
    // 임시 하드코딩 교구 코드 체크 (예: SEOUL -> 서울북부)
    // 실제로는 DB나 환경변수 연동이 좋지만, 요청에 따라 영문 코드(짧게) 사용
    const codeMap: Record<string, string> = {
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
    
    const matchedParish = codeMap[managerCode.toLowerCase()];
    if (matchedParish) {
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
              <div>
                <h3 className="font-bold text-slate-800">교회장 모드</h3>
                <p className="text-xs text-slate-500">내 교회 보고서 작성 및 관리</p>
              </div>
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
              <div>
                <h3 className="font-bold text-slate-800">사무장 모드</h3>
                <p className="text-xs text-slate-500">교구 전체 관리 및 취합</p>
              </div>
            </div>
            {selectedRole === 'manager' && (
              <div className="mt-4 space-y-3 animate-in fade-in slide-in-from-top-2">
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
                <button 
                  onClick={(e) => { e.stopPropagation(); handleStartManager(); }}
                  className="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2.5 rounded-lg flex items-center justify-center gap-2 mt-2"
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
