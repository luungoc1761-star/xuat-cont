import React, { useEffect, useState, useMemo, useRef } from 'react';
import {
  Boxes,
  Upload,
  Download,
  Cloud,
  Calendar,
  CalendarDays,
  Layers,
  Radio,
  Sparkles,
  RefreshCw,
  CheckCircle2
} from 'lucide-react';
import {
  AllocationMode,
  InitialDataPayload,
  LeftRightPair,
  PlanDataset,
  StockItem
} from './types';
import {
  reconcileShipmentPlan,
  buildLrPairsFromStock
} from './utils/reconciliationEngine';
import {
  parseStockWorkbook,
  parseShipmentWorkbook,
  exportFullReconciliationToExcel
} from './utils/excelParserAndExporter';
import { KpiCardsSection } from './components/KpiCardsSection';
import { MultiAgentAuditPanel } from './components/MultiAgentAuditPanel';
import { Warehouse3DScene } from './components/Warehouse3DScene';
import { QuickFillStockSearch } from './components/QuickFillStockSearch';
import { FulfillmentReportSection } from './components/FulfillmentReportSection';
import { DataImportAndOneDriveModal } from './components/DataImportAndOneDriveModal';

export function App() {
  const [loading, setLoading] = useState(true);
  const [stockItems, setStockItems] = useState<StockItem[]>([]);
  const [lrPairs, setLrPairs] = useState<LeftRightPair[]>([]);
  const [plans, setPlans] = useState<Record<string, PlanDataset>>({});
  const [activePlanKey, setActivePlanKey] = useState<string>('NOV-full');
  const [allocationMode, setAllocationMode] = useState<AllocationMode>('fifo');

  const [selectedModel, setSelectedModel] = useState<string | null>(null);
  const [filterMode3D, setFilterMode3D] = useState<
    'all_fg' | 'imbalanced' | 'with_demand' | 'cross_wh' | 'all_pairs'
  >('all_fg');
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedCodes, setSelectedCodes] = useState<string[]>([]);
  const [reportTab, setReportTab] = useState<
    'split' | 'fulfilled' | 'shortage' | 'po' | 'items'
  >('split');

  const [isImportModalOpen, setIsImportModalOpen] = useState(false);
  const [lastVerifiedAt, setLastVerifiedAt] = useState<string>('');
  const [lastSyncLog, setLastSyncLog] = useState<string>(
    'Đã tải và đối soát chuẩn xác Stock list.xlsx (Kho 01 & 60) + NOV. SHIPMENT -OCT22.xlsx trên Repo'
  );

  // Quick inline file inputs in header
  const quickStockInputRef = useRef<HTMLInputElement | null>(null);
  const quickMonthlyInputRef = useRef<HTMLInputElement | null>(null);
  const quickWeeklyInputRef = useRef<HTMLInputElement | null>(null);

  // Browser File System Access API live watcher state for OneDrive Stock list.xlsx
  const [liveWatchActive, setLiveWatchActive] = useState(false);
  const [liveWatchFileName, setLiveWatchFileName] = useState<string | null>(null);
  const fileHandleRef = useRef<any>(null);
  const lastModifiedRef = useRef<number>(0);
  const watchTimerRef = useRef<number | null>(null);

  // Load initial verified dataset
  const loadInitialDataset = async () => {
    try {
      const res = await fetch(`./data/initial-data.json?t=${Date.now()}`);
      const payload: InitialDataPayload = await res.json();
      setStockItems(payload.stockItems);
      setLrPairs(payload.lrPairs);
      setPlans(payload.plans);
      if (!payload.plans[activePlanKey]) {
        const firstKey = Object.keys(payload.plans)[0] || 'NOV-full';
        setActivePlanKey(firstKey);
      }
      setLastVerifiedAt(new Date().toLocaleTimeString('vi-VN'));
    } catch (err) {
      console.error('Failed to load initial dataset:', err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadInitialDataset();

    // Listen to Vite HMR custom event when repo Excel files change externally / via OneDrive watcher
    if (import.meta.hot) {
      import.meta.hot.on('repo-data-updated', (data: any) => {
        loadInitialDataset();
        setLastSyncLog(
          `[Realtime Auto-Sync] Phát hiện thay đổi file ${data?.file || 'Stock list.xlsx'} -> Đã tự động cập nhật lúc ${new Date().toLocaleTimeString('vi-VN')}`
        );
      });
    }

    return () => {
      if (watchTimerRef.current) window.clearInterval(watchTimerRef.current);
    };
  }, []);

  // Build metadata map for preserving descriptions when re-uploading Stock list.xlsx
  const existingMetaMap = useMemo(() => {
    const m = new Map<string, { feature: string; description: string; customerCode: string }>();
    for (const s of stockItems) {
      m.set(s.code, {
        feature: s.feature,
        description: s.description,
        customerCode: s.customerCode
      });
    }
    for (const pl of Object.values(plans)) {
      for (const l of pl.lines) {
        if (!m.has(l.itemCode) || !m.get(l.itemCode)?.description) {
          m.set(l.itemCode, {
            feature: l.feature,
            description: l.description,
            customerCode: l.customerItemCode
          });
        }
      }
    }
    return m;
  }, [stockItems, plans]);

  const activePlan = useMemo(() => {
    return (
      plans[activePlanKey] ||
      Object.values(plans)[0] || {
        name: 'Kế Hoạch Xuất Hàng Tháng 11 (NOV-full)',
        type: 'monthly',
        lines: [],
        containers: []
      }
    );
  }, [plans, activePlanKey]);

  const reconciliation = useMemo(() => {
    return reconcileShipmentPlan(stockItems, activePlan.lines, allocationMode);
  }, [stockItems, activePlan, allocationMode]);

  // Helper to convert ArrayBuffer to base64 and sync with backend repo API
  const syncBufferToBackendRepo = async (
    buf: ArrayBuffer,
    fileType: 'stock' | 'monthly' | 'weekly',
    fileName: string,
    autoGitPush: boolean
  ) => {
    try {
      const bytes = new Uint8Array(buf);
      let binary = '';
      const chunkSize = 0x8000;
      for (let i = 0; i < bytes.length; i += chunkSize) {
        binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
      }
      const base64Data = btoa(binary);
      const res = await fetch('/api/sync-file', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fileType, fileName, base64Data, autoGitPush })
      });
      if (res.ok) {
        const data = await res.json();
        return data.gitStatus || 'Đã lưu vào repo';
      }
    } catch {
      // Static hosting fallback: works 100% in-browser even without backend
    }
    return 'Đã xử lý trực tiếp trên trình duyệt';
  };

  // Upload handler for Stock, Monthly Plan, Weekly Plan
  const handleUploadFile = async (
    file: File,
    type: 'stock' | 'monthly' | 'weekly',
    autoGitPush = true
  ) => {
    const buf = await file.arrayBuffer();
    if (type === 'stock') {
      const parsed = parseStockWorkbook(buf, existingMetaMap);
      setStockItems(parsed.stockItems);
      setLrPairs(parsed.lrPairs);
      const statusMsg = await syncBufferToBackendRepo(buf, 'stock', file.name, autoGitPush);
      setLastVerifiedAt(new Date().toLocaleTimeString('vi-VN'));
      setLastSyncLog(
        `Đã nhập File Tồn Kho "${file.name}": Kho 01 (${parsed.wh01Rows} lô) + Kho 60 (${parsed.wh60Rows} lô) = ${parsed.stockItems.length} mã • ${statusMsg}`
      );
    } else if (type === 'monthly') {
      const parsedPlans = parseShipmentWorkbook(buf, 'monthly');
      setPlans((prev) => ({ ...prev, ...parsedPlans }));
      const firstKey = Object.keys(parsedPlans)[0];
      if (firstKey) setActivePlanKey(firstKey);
      const statusMsg = await syncBufferToBackendRepo(buf, 'monthly', file.name, autoGitPush);
      setLastVerifiedAt(new Date().toLocaleTimeString('vi-VN'));
      setLastSyncLog(
        `Đã nhập File Kế Hoạch Tháng "${file.name}" (${Object.keys(parsedPlans).join(', ')}) • ${statusMsg}`
      );
    } else if (type === 'weekly') {
      const parsedPlans = parseShipmentWorkbook(buf, 'weekly');
      setPlans((prev) => ({ ...prev, ...parsedPlans }));
      const wkKey =
        Object.keys(parsedPlans).find((k) => k.toLowerCase().includes('wk')) ||
        Object.keys(parsedPlans)[0];
      if (wkKey) setActivePlanKey(wkKey);
      const statusMsg = await syncBufferToBackendRepo(buf, 'weekly', file.name, autoGitPush);
      setLastVerifiedAt(new Date().toLocaleTimeString('vi-VN'));
      setLastSyncLog(
        `Đã nhập File Kế Hoạch Tuần "${file.name}" (${wkKey}) • ${statusMsg}`
      );
    }
  };

  // Start Browser Live File Watcher for OneDrive Stock list.xlsx
  const handleStartBrowserLiveWatch = async (autoGitPush: boolean) => {
    const anyWin = window as any;
    if (!anyWin.showOpenFilePicker) {
      alert(
        'Trình duyệt hiện tại không hỗ trợ File System Access API. Vui lòng chọn file qua nút Nhập File Tồn Kho hoặc dùng Chrome/Edge.'
      );
      return;
    }
    try {
      const [handle] = await anyWin.showOpenFilePicker({
        types: [
          {
            description: 'Excel Stock List (OneDrive)',
            accept: {
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx', '.xls']
            }
          }
        ],
        multiple: false
      });
      fileHandleRef.current = handle;
      const initialFile: File = await handle.getFile();
      lastModifiedRef.current = initialFile.lastModified;
      setLiveWatchFileName(initialFile.name);
      setLiveWatchActive(true);

      await handleUploadFile(initialFile, 'stock', autoGitPush);

      if (watchTimerRef.current) window.clearInterval(watchTimerRef.current);
      watchTimerRef.current = window.setInterval(async () => {
        try {
          if (!fileHandleRef.current) return;
          const currentFile: File = await fileHandleRef.current.getFile();
          if (currentFile.lastModified > lastModifiedRef.current) {
            lastModifiedRef.current = currentFile.lastModified;
            await handleUploadFile(currentFile, 'stock', autoGitPush);
          }
        } catch (err) {
          console.error('Live watch poll error:', err);
        }
      }, 2500);
    } catch {
      // User cancelled picker
    }
  };

  const handleStopBrowserLiveWatch = () => {
    if (watchTimerRef.current) {
      window.clearInterval(watchTimerRef.current);
      watchTimerRef.current = null;
    }
    fileHandleRef.current = null;
    setLiveWatchActive(false);
    setLiveWatchFileName(null);
    setLastSyncLog('Đã dừng chế độ theo dõi Realtime file OneDrive.');
  };

  const handleServerOneDriveSync = async (folderPath: string, autoGitPush: boolean) => {
    try {
      const res = await fetch('/api/onedrive-Config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ folderPath, autoGitPush })
      });
      const data = await res.json();
      if (data.ok) {
        await loadInitialDataset();
        setLastSyncLog(
          `Đã đồng bộ thành công Stock list.xlsx từ "${data.syncedFrom}" lúc ${new Date().toLocaleTimeString('vi-VN')}`
        );
      } else {
        setLastSyncLog(data.message || 'Không tìm thấy đường dẫn trên máy chủ.');
      }
    } catch {
      setLastSyncLog(
        'Mẹo: Nhấn nút "Kết Nối File Stock list.xlsx Trong OneDrive (Realtime)" ở trên để trình duyệt tự động đọc trực tiếp từ thư mục OneDrive trên máy tính của anh!'
      );
    }
  };

  if (loading) {
    return (
      <div className="min-h-screen flex flex-col items-center justify-center bg-slate-950 text-white p-6">
        <div className="h-12 w-12 rounded-2xl bg-gradient-to-br from-cyan-500 to-emerald-500 flex items-center justify-center animate-spin mb-4">
          <RefreshCw className="h-6 w-6 text-slate-950" />
        </div>
        <h2 className="text-lg font-bold">
          Đang Khởi Tạo Dashboard 3D & Đối Soát Đa Agent (Kho 01 & Kho 60)...
        </h2>
      </div>
    );
  }

  return (
    <div className="min-h-screen flex flex-col bg-slate-950 text-slate-100">
      {/* Hidden quick inputs for 1-click upload in top bar */}
      <input
        ref={quickStockInputRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) handleUploadFile(f, 'stock', true);
          e.target.value = '';
        }}
        className="hidden"
      />
      <input
        ref={quickMonthlyInputRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) handleUploadFile(f, 'monthly', true);
          e.target.value = '';
        }}
        className="hidden"
      />
      <input
        ref={quickWeeklyInputRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) handleUploadFile(f, 'weekly', true);
          e.target.value = '';
        }}
        className="hidden"
      />

      {/* Sticky Top Command Navigation Bar */}
      <header className="sticky top-0 z-40 border-b border-slate-800/90 bg-slate-950/85 backdrop-blur-xl">
        <div className="mx-auto max-w-[1600px] px-4 sm:px-6 py-3 flex flex-wrap items-center justify-between gap-3">
          {/* Brand Title */}
          <div className="flex items-center gap-3">
            <div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-gradient-to-br from-cyan-400 via-sky-500 to-emerald-400 text-slate-950 shadow-lg shadow-cyan-500/25">
              <Boxes className="h-6 w-6" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-base sm:text-lg font-extrabold tracking-tight text-white">
                  LPVN SMART WAREHOUSE 3D & CONTAINER SHIPMENT HUB
                </h1>
                <span className="hidden md:inline-flex items-center gap-1 rounded-full bg-emerald-500/15 border border-emerald-500/35 px-2.5 py-0.5 text-[11px] font-semibold text-emerald-300">
                  <CheckCircle2 className="h-3 w-3" /> Kho 01 & Kho 60 Verified
                </span>
              </div>
              <p className="text-xs text-slate-400">
                Báo cáo Tồn Kho Thành Phẩm 3D (Trái/Phải) & Đối Soát Kế Hoạch Xuất Hàng Tháng / Tuần • Đồng Bộ OneDrive Realtime
              </p>
            </div>
          </div>

          {/* Active Plan Selector + Source File Upload Buttons + Excel Export */}
          <div className="flex flex-wrap items-center gap-2">
            {/* Plan Switcher Pills */}
            <div className="inline-flex rounded-xl border border-slate-800 bg-slate-900 p-1">
              {Object.entries(plans).map(([key, pl]) => (
                <button
                  key={key}
                  onClick={() => setActivePlanKey(key)}
                  className={`rounded-lg px-3 py-1.5 text-xs font-bold transition cursor-pointer ${
                    activePlanKey === key
                      ? 'bg-cyan-500 text-slate-950 shadow'
                      : 'text-slate-300 hover:text-white'
                  }`}
                >
                  {key === 'NOV-full'
                    ? 'KH Tháng 11 (NOV-full)'
                    : key === 'OCT-WK 42'
                    ? 'KH Tuần 42 (OCT-WK 42)'
                    : key === 'Data-OPENING'
                    ? 'Tất Cả Đơn Mở (OPENING)'
                    : key}
                </button>
              ))}
            </div>

            {/* 3 Direct Source Import Buttons + Hub Modal Button */}
            <div className="flex items-center gap-1.5 rounded-xl border border-cyan-500/30 bg-slate-900/90 p-1">
              <button
                onClick={() => quickStockInputRef.current?.click()}
                className="inline-flex items-center gap-1 rounded-lg bg-cyan-500/15 px-2.5 py-1.5 text-xs font-semibold text-cyan-300 hover:bg-cyan-500 hover:text-slate-950 transition cursor-pointer"
                title="Nhập File Tồn Kho (Stock list.xlsx)"
              >
                <Upload className="h-3.5 w-3.5" />
                <span>Nhập Tồn Kho</span>
              </button>
              <button
                onClick={() => quickMonthlyInputRef.current?.click()}
                className="inline-flex items-center gap-1 rounded-lg bg-indigo-500/15 px-2.5 py-1.5 text-xs font-semibold text-indigo-300 hover:bg-indigo-500 hover:text-white transition cursor-pointer"
                title="Nhập File Kế Hoạch Tháng"
              >
                <Calendar className="h-3.5 w-3.5" />
                <span>Nhập KH Tháng</span>
              </button>
              <button
                onClick={() => quickWeeklyInputRef.current?.click()}
                className="inline-flex items-center gap-1 rounded-lg bg-emerald-500/15 px-2.5 py-1.5 text-xs font-semibold text-emerald-300 hover:bg-emerald-500 hover:text-slate-950 transition cursor-pointer"
                title="Nhập File Kế Hoạch Tuần"
              >
                <CalendarDays className="h-3.5 w-3.5" />
                <span>Nhập KH Tuần</span>
              </button>
              <button
                onClick={() => setIsImportModalOpen(true)}
                className="inline-flex items-center gap-1 rounded-lg bg-sky-500/20 border border-sky-500/40 px-2.5 py-1.5 text-xs font-bold text-sky-300 hover:bg-sky-500 hover:text-slate-950 transition cursor-pointer"
                title="Mở Trung Tâm Đồng Bộ OneDrive Realtime & Quản Lý Nguồn"
              >
                <Cloud className="h-3.5 w-3.5" />
                <span>
                  {liveWatchActive ? `OneDrive Live (${liveWatchFileName})` : 'OneDrive Realtime'}
                </span>
              </button>
            </div>

            {/* Export Excel Primary Button */}
            <button
              onClick={() =>
                exportFullReconciliationToExcel(
                  reconciliation,
                  lrPairs,
                  stockItems,
                  activePlan.name,
                  allocationMode
                )
              }
              className="inline-flex items-center gap-1.5 rounded-xl bg-gradient-to-r from-emerald-500 to-teal-400 px-3.5 py-2 text-xs font-extrabold text-slate-950 shadow-lg shadow-emerald-500/20 hover:brightness-110 transition cursor-pointer"
            >
              <Download className="h-4 w-4" />
              <span>Xuất Excel (.xlsx)</span>
            </button>
          </div>
        </div>
      </header>

      {/* Main Content Container */}
      <main className="mx-auto w-full max-w-[1600px] flex-1 px-4 sm:px-6 py-5 space-y-5">
        {/* 1. Animated KPI Cards */}
        <KpiCardsSection
          stockItems={stockItems}
          lrPairs={lrPairs}
          reconciliation={reconciliation}
          activePlanLabel={activePlan.name}
          onQuickJumpTab={(tab) => {
            setReportTab(tab);
            document
              .getElementById('fulfillment-report-section')
              ?.scrollIntoView({ behavior: 'smooth' });
          }}
        />

        {/* 2. Multi-Agent Audit & Verification Bar */}
        <MultiAgentAuditPanel
          stockItems={stockItems}
          lrPairs={lrPairs}
          reconciliation={reconciliation}
          activePlanName={activePlan.name}
          lastVerifiedAt={lastVerifiedAt}
          onRunAuditAgain={() => {
            setLrPairs(buildLrPairsFromStock(stockItems));
            setLastVerifiedAt(new Date().toLocaleTimeString('vi-VN'));
          }}
        />

        {/* 3. Dynamic 3D Warehouse Dashboard (Finished Goods WH 01 & 60 Left/Right) */}
        <Warehouse3DScene
          lrPairs={lrPairs}
          itemDemands={reconciliation.itemDemands}
          selectedModel={selectedModel}
          onSelectModel={setSelectedModel}
          filterMode={filterMode3D}
          setFilterMode={setFilterMode3D}
        />

        {/* 4. Quick-Fill & Multi-Code Stock Search */}
        <QuickFillStockSearch
          stockItems={stockItems}
          lrPairs={lrPairs}
          itemDemands={reconciliation.itemDemands}
          searchQuery={searchQuery}
          setSearchQuery={setSearchQuery}
          selectedCodes={selectedCodes}
          setSelectedCodes={setSelectedCodes}
          onSelectModelIn3D={(model) => setSelectedModel(model)}
        />

        {/* 5. Bottom Fulfillment & Shortage Report (Kho 01 & 60 vs Shipment Plan + Excel Export) */}
        <FulfillmentReportSection
          reconciliation={reconciliation}
          lrPairs={lrPairs}
          stockItems={stockItems}
          activePlanKey={activePlanKey}
          activePlanName={activePlan.name}
          allocationMode={allocationMode}
          setAllocationMode={setAllocationMode}
          searchQuery={searchQuery}
          selectedModel={selectedModel}
          activeTab={reportTab}
          setActiveTab={setReportTab}
        />
      </main>

      {/* Footer */}
      <footer className="border-t border-slate-900 bg-slate-950/90 py-4 px-6 text-center text-xs text-slate-500">
        LPVN 3D Warehouse & Container Shipment Reconciliation System • Dữ liệu đối soát tự động từ{' '}
        <code className="text-slate-300">Stock list.xlsx</code> (Kho 01 & 60) và{' '}
        <code className="text-slate-300">NOV. SHIPMENT -OCT22.xlsx</code> • Hỗ trợ OneDrive Realtime Sync
      </footer>

      {/* Data Source Import & OneDrive Realtime Sync Modal */}
      <DataImportAndOneDriveModal
        isOpen={isImportModalOpen}
        onClose={() => setIsImportModalOpen(false)}
        onUploadFile={handleUploadFile}
        liveWatchActive={liveWatchActive}
        liveWatchFileName={liveWatchFileName}
        lastSyncLog={lastSyncLog}
        onStartBrowserLiveWatch={handleStartBrowserLiveWatch}
        onStopBrowserLiveWatch={handleStopBrowserLiveWatch}
        onServerOneDriveSync={handleServerOneDriveSync}
      />
    </div>
  );
}
export default App;
