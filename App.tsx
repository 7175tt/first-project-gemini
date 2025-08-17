
import React, { useState, useCallback, useRef } from 'react';
import { read, utils, writeFile } from 'xlsx';
import { KPIS } from './constants';
import { KpiId } from './types';
import { generateReport } from './services/geminiService';

const App: React.FC = () => {
  const [selectedKpis, setSelectedKpis] = useState<Record<KpiId, boolean>>(
    Object.fromEntries(KPIS.map(kpi => [kpi.id, false])) as Record<KpiId, boolean>
  );

  const [kpiData, setKpiData] = useState<Record<KpiId, string>>(
    Object.fromEntries(KPIS.map(kpi => [kpi.id, ''])) as Record<KpiId, string>
  );

  const [audience, setAudience] = useState('');
  const [generatedReport, setGeneratedReport] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [copySuccess, setCopySuccess] = useState(false);
  const [importMessage, setImportMessage] = useState<{type: 'success' | 'error', text: string} | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleKpiToggle = (kpiId: KpiId) => {
    setSelectedKpis(prev => ({ ...prev, [kpiId]: !prev[kpiId] }));
  };

  const handleDataChange = (kpiId: KpiId, value: string) => {
    setKpiData(prev => ({ ...prev, [kpiId]: value }));
  };
  
  const isGenerateButtonDisabled = 
    !Object.values(selectedKpis).some(Boolean) || 
    KPIS.some(kpi => selectedKpis[kpi.id] && kpiData[kpi.id].trim() === '') ||
    isLoading;

  const handleSubmit = useCallback(async () => {
    setIsLoading(true);
    setError(null);
    setGeneratedReport('');

    const selectedKpisWithData: Partial<Record<KpiId, { label: string; data: string }>> = {};
    
    KPIS.forEach(kpi => {
        if (selectedKpis[kpi.id] && kpiData[kpi.id]) {
            selectedKpisWithData[kpi.id] = {
                label: kpi.label,
                data: kpiData[kpi.id]
            };
        }
    });

    if (Object.keys(selectedKpisWithData).length === 0) {
      setError("請至少選擇一個KPI並填寫對應的資料。");
      setIsLoading(false);
      return;
    }

    try {
      const report = await generateReport(selectedKpisWithData, audience);
      setGeneratedReport(report);
    } catch (e: any) {
      setError(e.message || "發生未知錯誤");
    } finally {
      setIsLoading(false);
    }
  }, [selectedKpis, kpiData, audience]);

  const handleCopy = () => {
    navigator.clipboard.writeText(generatedReport).then(() => {
        setCopySuccess(true);
        setTimeout(() => setCopySuccess(false), 2000);
    }, (err) => {
        console.error('Could not copy text: ', err);
    });
  };

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    setImportMessage(null);
    setError(null);
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        if (!data) {
            throw new Error("無法讀取檔案內容。");
        }
        const workbook = read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = utils.sheet_to_json<string[]>(worksheet, { header: 1 });
        
        let updatedCount = 0;
        const newKpiData = { ...kpiData };
        const newSelectedKpis = { ...selectedKpis };
        const validKpiIds = new Set(KPIS.map(k => k.id));
        
        const startingRow = (jsonData.length > 0 && !validKpiIds.has(jsonData[0][0] as KpiId)) ? 1 : 0;
        
        for (let i = startingRow; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || row.length < 2) continue;

            const kpiId = row[0]?.trim() as KpiId;
            const value = row[1]?.trim();

            if (validKpiIds.has(kpiId) && value) {
                newKpiData[kpiId] = value;
                newSelectedKpis[kpiId] = true;
                updatedCount++;
            }
        }

        if (updatedCount > 0) {
            setKpiData(newKpiData);
            setSelectedKpis(newSelectedKpis);
            setImportMessage({ type: 'success', text: `成功匯入並更新了 ${updatedCount} 個 KPI 項目。` });
        } else {
            setImportMessage({ type: 'error', text: "檔案中未找到有效的 KPI 資料。請檢查檔案格式是否正確（第一欄為 KPI ID，第二欄為資料）。" });
        }

      } catch (error) {
        console.error("Error parsing Excel file:", error);
        setImportMessage({ type: 'error', text: "讀取或解析檔案時發生錯誤，請確認檔案格式是否正確。" });
      } finally {
        if(event.target) {
            event.target.value = '';
        }
      }
    };

    reader.onerror = () => {
        setImportMessage({ type: 'error', text: "讀取檔案失敗。" });
        if(event.target) {
            event.target.value = '';
        }
    };

    reader.readAsArrayBuffer(file);
  };

  const handleExportTemplate = () => {
    const templateData = KPIS.map(kpi => ({
      'KPI ID': kpi.id,
      '成果資料': '' // Leave this empty for the user
    }));
    
    const worksheet = utils.json_to_sheet(templateData);
    const workbook = utils.book_new();
    utils.book_append_sheet(workbook, worksheet, 'KPI 資料範本');
    
    // Set column widths for better readability
    worksheet['!cols'] = [{ wch: 30 }, { wch: 80 }];

    writeFile(workbook, 'KPI_成果報告範本.xlsx');
  };

  const SparkleIcon: React.FC<{className?: string}> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor" className={className}>
      <path fillRule="evenodd" d="M9 4.5a.75.75 0 01.75.75v3.546a.75.75 0 01-1.5 0V5.25A.75.75 0 019 4.5zM12.75 8.663a.75.75 0 00-1.5 0v3.546a.75.75 0 001.5 0V8.663zM15 4.5a.75.75 0 01.75.75v3.546a.75.75 0 01-1.5 0V5.25A.75.75 0 0115 4.5zM18.75 8.663a.75.75 0 00-1.5 0v3.546a.75.75 0 001.5 0V8.663z" clipRule="evenodd" />
      <path d="M8.25 12.75a.75.75 0 01.75-.75h6a.75.75 0 010 1.5h-6a.75.75 0 01-.75-.75zM12 15a.75.75 0 01.75-.75h.008a.75.75 0 01.75.75v.008a.75.75 0 01-.75.75h-.008a.75.75 0 01-.75-.75v-.008zM15.75 12.75a.75.75 0 01.75-.75h.008a.75.75 0 01.75.75v.008a.75.75 0 01-.75.75h-.008a.75.75 0 01-.75-.75v-.008zM8.25 15.75a.75.75 0 01.75-.75h.008a.75.75 0 01.75.75v.008a.75.75 0 01-.75.75H9a.75.75 0 01-.75-.75v-.008zM9.75 18a.75.75 0 01.75-.75h3a.75.75 0 010 1.5h-3a.75.75 0 01-.75-.75z" />
      <path fillRule="evenodd" d="M12 2.25c-5.385 0-9.75 4.365-9.75 9.75s4.365 9.75 9.75 9.75 9.75-4.365 9.75-9.75S17.385 2.25 12 2.25zM3.5 12a8.5 8.5 0 1117 0 8.5 8.5 0 01-17 0z" clipRule="evenodd" />
    </svg>
  );

  const UploadIcon: React.FC<{className?: string}> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M12 16.5V9.75m0 0l-3.75 3.75M12 9.75l3.75 3.75M3.75 18A5.25 5.25 0 009 20.25h6a5.25 5.25 0 005.25-5.25c0-2.01-1.125-3.75-2.625-4.583A7.5 7.5 0 0012 3.75a7.5 7.5 0 00-6.375 3.417c-1.5 1.083-2.625 2.875-2.625 4.833z" />
    </svg>
  );

  const DownloadIcon: React.FC<{className?: string}> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className={className}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5M16.5 12L12 16.5m0 0L7.5 12m4.5 4.5V3" />
    </svg>
  );

  return (
    <div className="min-h-screen font-sans text-slate-800 p-4 md:p-8">
      <div className="max-w-7xl mx-auto">
        <header className="text-center mb-8">
          <h1 className="text-4xl font-bold text-slate-900">政府計畫成果報告產生器</h1>
          <p className="text-slate-600 mt-2">勾選項目、填寫成果，快速生成專業報告草稿</p>
        </header>

        <main className="grid grid-cols-1 lg:grid-cols-2 gap-8">
          {/* Input Panel */}
          <div className="bg-white p-6 rounded-2xl shadow-lg border border-slate-200 flex flex-col gap-6">
            <div>
              <h2 className="text-xl font-semibold text-slate-700 mb-3 border-b pb-2">1. 選擇要包含的成果項目 (KPI)</h2>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                {KPIS.map(kpi => (
                  <label key={kpi.id} className="flex items-center space-x-3 p-3 rounded-lg hover:bg-slate-100 transition-colors cursor-pointer">
                    <input
                      type="checkbox"
                      checked={selectedKpis[kpi.id]}
                      onChange={() => handleKpiToggle(kpi.id)}
                      className="h-5 w-5 rounded border-gray-300 text-indigo-600 focus:ring-indigo-500"
                    />
                    <span className="text-slate-700 font-medium">{kpi.label}</span>
                  </label>
                ))}
              </div>
            </div>
            
            <div>
                <div className="flex justify-between items-center flex-wrap gap-2 border-b pb-2 mb-3">
                    <h2 className="text-xl font-semibold text-slate-700">2. 填寫各項目的原始資料</h2>
                    <div className="flex items-center gap-2">
                        <button
                            onClick={handleExportTemplate}
                            className="flex items-center justify-center gap-2 bg-slate-600 text-white font-semibold py-2 px-3 rounded-lg hover:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-slate-500 transition-colors text-sm"
                        >
                            <DownloadIcon className="h-5 w-5" />
                            下載 Excel 範本
                        </button>
                        <button
                            onClick={() => fileInputRef.current?.click()}
                            className="flex items-center justify-center gap-2 bg-teal-600 text-white font-semibold py-2 px-3 rounded-lg hover:bg-teal-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-teal-500 transition-colors text-sm"
                        >
                            <UploadIcon className="h-5 w-5" />
                            從 Excel 匯入
                        </button>
                    </div>
                    <input
                        type="file"
                        ref={fileInputRef}
                        onChange={handleFileChange}
                        className="hidden"
                        accept=".xlsx, .xls, .csv"
                    />
                </div>
                 <p className="text-xs text-slate-500 mb-4">
                    提示：可下載範本檔案填寫後匯入。上傳的檔案應包含兩欄，第一欄為 KPI ID (例如 kpiA)，第二欄為對應的成果資料。
                </p>
                {importMessage && (
                    <div className={`p-3 rounded-md mb-4 text-sm ${importMessage.type === 'success' ? 'bg-green-100 text-green-800' : 'bg-red-100 text-red-800'}`}>
                        {importMessage.text}
                    </div>
                )}
                <div className="space-y-4">
                {KPIS.map(kpi => (
                    <div key={kpi.id} className={`${selectedKpis[kpi.id] ? 'opacity-100' : 'opacity-40 pointer-events-none'}`}>
                    <label htmlFor={kpi.id} className="block text-sm font-medium text-slate-600 mb-1">{kpi.label}</label>
                    <textarea
                        id={kpi.id}
                        rows={4}
                        value={kpiData[kpi.id]}
                        onChange={e => handleDataChange(kpi.id, e.target.value)}
                        placeholder={`請在此輸入關於 "${kpi.label}" 的具體成果、量化數據、質化描述等...`}
                        className="w-full p-2 border border-slate-300 rounded-md shadow-sm focus:ring-indigo-500 focus:border-indigo-500 transition"
                        disabled={!selectedKpis[kpi.id]}
                    />
                    </div>
                ))}
                </div>
            </div>

            <div>
              <h2 className="text-xl font-semibold text-slate-700 mb-3 border-b pb-2">3. (選填) 指定報告對象</h2>
               <input
                  type="text"
                  value={audience}
                  onChange={e => setAudience(e.target.value)}
                  placeholder="例如：上級督導單位、跨部門合作會議、民眾說明會..."
                  className="w-full p-2 border border-slate-300 rounded-md shadow-sm focus:ring-indigo-500 focus:border-indigo-500 transition"
                />
            </div>
            
            <button
              onClick={handleSubmit}
              disabled={isGenerateButtonDisabled}
              className="w-full flex items-center justify-center gap-2 bg-indigo-600 text-white font-bold py-3 px-4 rounded-lg hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 disabled:bg-slate-400 disabled:cursor-not-allowed transition-all duration-300 transform hover:scale-105 disabled:scale-100"
            >
              {isLoading ? (
                <>
                  <svg className="animate-spin -ml-1 mr-3 h-5 w-5 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                  </svg>
                  報告生成中...
                </>
              ) : (
                <>
                  <SparkleIcon className="h-6 w-6" />
                  生成報告草稿
                </>
              )}
            </button>
          </div>

          {/* Output Panel */}
          <div className="bg-white p-6 rounded-2xl shadow-lg border border-slate-200 flex flex-col">
            <div className="flex justify-between items-center border-b pb-2 mb-3">
                <h2 className="text-xl font-semibold text-slate-700">產生的報告草稿</h2>
                {generatedReport && !isLoading && (
                    <button 
                        onClick={handleCopy}
                        className={`px-4 py-2 text-sm font-medium rounded-md transition-colors ${copySuccess ? 'bg-green-600 text-white' : 'bg-slate-200 text-slate-700 hover:bg-slate-300'}`}
                    >
                        {copySuccess ? '已複製！' : '複製內容'}
                    </button>
                )}
            </div>

            <div className="flex-grow bg-slate-50 rounded-md p-4 whitespace-pre-wrap overflow-y-auto min-h-[300px] text-slate-800 leading-relaxed font-serif">
              {isLoading && (
                 <div className="flex flex-col items-center justify-center h-full text-slate-500">
                    <svg className="animate-spin h-8 w-8 text-indigo-500 mb-4" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                        <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                        <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                    </svg>
                    <p className="font-semibold">AI 正在為您撰寫報告...</p>
                    <p className="text-sm">請稍候片刻</p>
                </div>
              )}
              {error && <div className="text-red-600 bg-red-100 p-4 rounded-md">{error}</div>}
              {!isLoading && !error && !generatedReport && (
                <div className="flex items-center justify-center h-full text-slate-400">
                    <p>報告結果將會顯示於此</p>
                </div>
              )}
              {generatedReport}
            </div>
          </div>
        </main>
      </div>
    </div>
  );
};

export default App;
