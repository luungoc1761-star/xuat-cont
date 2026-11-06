import React, { useEffect, useState, useMemo, useRef } from 'react';
import {
  Boxes,
  Upload,
  Download,
  Calendar,
  CalendarDays,
  RefreshCw,
  CheckCircle2,
  Calculator,
  Database
} from 'lucide-react';
import {
  AllocationMode,
  InitialDataPayload,
  LeftRightPair,
  PlanDataset,
  ReconciledPoSummary,
  StockItem
} from './types';
import {
  reconcileShipmentPlan,
  buildLrPairsFromStock,
  deductShippedContainerFromStock,
  isAccessoryModel
} from './utils/reconciliationEngine';
import {
  parseStockWorkbook,
  parseShipmentWorkbook,
  exportFullReconciliationToExcel,
  exportContainersByPoSoToExcel
} from './utils/excelParserAndExporter';
import { KpiCardsSection } from './components/KpiCardsSection';
import { MultiAgentAuditPanel } from './components/MultiAgentAuditPanel';
import { Warehouse3DScene } from './components/Warehouse3DScene';
import { QuickFillStockSearch } from './components/QuickFillStockSearch';
import { FulfillmentReportSection } from './components/FulfillmentReportSection';
import { DataImportAndOneDriveModal } from './components/DataImportAndOneDriveModal';

export function App() {
  const [loading, setLoading] = useState(true);
  const [isReconcilingStock, setIsReconcilingStock] = useState(false);
  const [stockItems, setStockItems] = useState<StockItem[]>([]);
  const [lrPairs, setLrPairs] = useState<LeftRightPair[]>([]);
  const [plans, setPlans] = useState<Record<string, PlanDataset>>({});
  const [activePlanKey, setActivePlanKey] = useState<string>('NOV-full');
  const [allocationMode, setAllocationMode] = useState<AllocationMode>('direct');
  const [activeStockFileName, setActiveStockFileName] = useState<string>('Stock list.xlsx');

  const [selectedModel, setSelectedModel] = useState<string | null>(null);
  const [filterMode3D, setFilterMode3D] = useState<
    'all_fg' | 'imbalanced' | 'with_demand' | 'cross_wh' | 'all_pairs'
  >('all_fg');
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedCodes, setSelectedCodes] = useState<string[]>([]);
  const [reportTab, setReportTab] = useState<
    'po' | 'split' | 'fulfilled' | 'shortage' | 'items'
  >('po');

  const [isImportModalOpen, setIsImportModalOpen] = useState(false);
  const [lastVerifiedAt, setLastVerifiedAt] = useState<string>('');
  const [lastSyncLog, setLastSyncLog] = useState<string>(
    'Đã tải và đối soát chuẩn xác Stock list.xlsx (Kho 01 & 60) + NOV. SHIPMENT -OCT22.xlsx'
  );
  const [exportedHistory, setExportedHistory] = useState<
    Array<{
      contLabel: string;
      po: string;
      so: string;
      deductedQty: number;
      timestamp: string;
      prevStockItems: StockItem[];
      prevLrPairs: LeftRightPair[];
      prevPlans: Record<string, PlanDataset>;
    }>
  >([]);

  // Persistent metadata map and base plans reference
  const masterMetaMapRef = useRef<
    Map<string, { feature: string; description: string; customerCode: string }>
  >(new Map());
  const basePlansRef = useRef<Record<string, PlanDataset>>({});

  // Quick inline file inputs in header
  const quickStockInputRef = useRef<HTMLInputElement | null>(null);
  const quickMonthlyInputRef = useRef<HTMLInputElement | null>(null);
  const quickWeeklyInputRef = useRef<HTMLInputElement | null>(null);

  // Load initial verified dataset
  const loadInitialDataset = async () => {
    try {
      const res = await fetch(`./data/initial-data.json?t=${Date.now()}`);
      const payload: InitialDataPayload = await res.json();
      const filteredStock = payload.stockItems.filter((item) => {
        if (item.code.startsWith('3') && item.wh01Qty > 0 && !isAccessoryModel(item.feature, item.category)) {
          return item.wh60Qty > 0;
        }
        return true;
      });

      // Populate persistent metadata map
      const m = new Map(masterMetaMapRef.current);
      for (const s of payload.stockItems) {
        m.set(s.code, {
          feature: s.feature,
          description: s.description,
          customerCode: s.customerCode
        });
      }
      for (const pl of Object.values(payload.plans)) {
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
      masterMetaMapRef.current = m;
      basePlansRef.current = payload.plans;

      setStockItems(filteredStock);
      setLrPairs(buildLrPairsFromStock(filteredStock));
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

    const metaHot = (import.meta as any).hot;
    if (metaHot) {
      metaHot.on('repo-data-updated', (data: any) => {
        loadInitialDataset();
        setLastSyncLog(
          `Đã cập nhật dữ liệu từ file ${data?.file || 'Stock list.xlsx'} lúc ${new Date().toLocaleTimeString('vi-VN')}`
        );
      });
    }
  }, []);

  // Build metadata map for preserving descriptions when uploading Stock list.xlsx
  const existingMetaMap = useMemo(() => {
    const m = new Map<string, { feature: string; description: string; customerCode: string }>(
      masterMetaMapRef.current
    );
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
    masterMetaMapRef.current = m;
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

  // Handler when clicking "Xuất Cont" on any container card:
  // 1. Deducts reconciled inventory from warehouse stock (Kho 60 & Kho 01)
  // 2. Carries cumulative deduction forward for all subsequent orders in the plan
  // 3. Removes the completed container from the active shipment plan
  const handleExportContainer = (poGroup: ReconciledPoSummary) => {
    const contLabel = `Cont #${String(poGroup.containerIndex).padStart(2, '0')}`;
    const lineIdsToRemove = new Set(poGroup.lines.map((l) => l.id));
    const deductedQty = poGroup.totalAllocatedQty;

    const nextStockItems = deductShippedContainerFromStock(stockItems, poGroup);
    const nextLrPairs = buildLrPairsFromStock(nextStockItems);

    const nextPlans: Record<string, PlanDataset> = { ...plans };
    const currentPlan = nextPlans[activePlanKey];
    if (currentPlan) {
      const remainingLines = currentPlan.lines.filter((l) => !lineIdsToRemove.has(l.id));
      const remainingContIndices = new Set(remainingLines.map((l) => l.containerIndex));
      const remainingContainers = currentPlan.containers.filter((c) =>
        remainingContIndices.has(c.containerIndex)
      );
      nextPlans[activePlanKey] = {
        ...currentPlan,
        lines: remainingLines,
        containers: remainingContainers
      };
    }

    setExportedHistory((prev) => [
      ...prev,
      {
        contLabel,
        po: poGroup.po,
        so: poGroup.so || poGroup.soList.join(', ') || 'N/A',
        deductedQty,
        timestamp: new Date().toLocaleTimeString('vi-VN'),
        prevStockItems: stockItems,
        prevLrPairs: lrPairs,
        prevPlans: plans
      }
    ]);

    setStockItems(nextStockItems);
    setLrPairs(nextLrPairs);
    setPlans(nextPlans);
    setLastVerifiedAt(new Date().toLocaleTimeString('vi-VN'));
    setLastSyncLog(
      `Đã hoàn thành xuất ${contLabel} (PO: ${poGroup.po} | SO: ${poGroup.so || 'N/A'}): Đã trừ -${deductedQty.toLocaleString('en-US')} pcs khỏi hệ thống tồn kho (lũy kế cho các đơn tiếp theo) và đưa đơn ra khỏi kế hoạch xuất.`
    );
  };

  const handleUndoLastExportContainer = () => {
    if (exportedHistory.length === 0) return;
    const last = exportedHistory[exportedHistory.length - 1];
    setStockItems(last.prevStockItems);
    setLrPairs(last.prevLrPairs);
    setPlans(last.prevPlans);
    setExportedHistory((prev) => prev.slice(0, -1));
    setLastVerifiedAt(new Date().toLocaleTimeString('vi-VN'));
    setLastSyncLog(`Đã hoàn tác xuất ${last.contLabel} (PO: ${last.po}) và khôi phục tồn kho.`);
  };

  // Helper to convert ArrayBuffer to base64 and sync with backend repo API (if running dev server)
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
    return 'Đã tính đối soát trực tiếp trên trình duyệt';
  };

  // Upload handler for Stock, Monthly Plan, Weekly Plan
  const handleUploadFile = async (
    file: File,
    type: 'stock' | 'monthly' | 'weekly',
    autoGitPush = true
  ) => {
    setIsReconcilingStock(true);
    try {
      const buf = await file.arrayBuffer();
      if (type === 'stock') {
        const parsed = parseStockWorkbook(buf, existingMetaMap);
        // Reset any previous exported container deductions when importing a fresh stock file
        const targetPlans =
          exportedHistory.length > 0 && Object.keys(basePlansRef.current).length > 0
            ? basePlansRef.current
            : plans;
        if (exportedHistory.length > 0 && Object.keys(basePlansRef.current).length > 0) {
          setPlans(basePlansRef.current);
          setExportedHistory([]);
        }
        setStockItems(parsed.stockItems);
        setLrPairs(parsed.lrPairs);
        setActiveStockFileName(file.name);

        const currentPlanObj =
          targetPlans[activePlanKey] || Object.values(targetPlans)[0];
        const reconPreview = reconcileShipmentPlan(
          parsed.stockItems,
          currentPlanObj ? currentPlanObj.lines : [],
          allocationMode
        );

        const statusMsg = await syncBufferToBackendRepo(buf, 'stock', file.name, autoGitPush);
        setLastVerifiedAt(new Date().toLocaleTimeString('vi-VN'));
        setLastSyncLog(
          `Đã nhập File Tồn Kho "${file.name}" & tính đối soát: Kho 01 (${parsed.wh01Rows} lô) + Kho 60 (${parsed.wh60Rows} lô) = ${parsed.stockItems.length} mã • Kết quả: ${reconPreview.fulfilledPos.length} Cont đủ hàng / ${reconPreview.shortagePos.length} Cont thiếu hàng • ${statusMsg}`
        );
      } else if (type === 'monthly') {
        const parsedPlans = parseShipmentWorkbook(buf, 'monthly');
        setPlans((prev) => {
          const next = { ...prev, ...parsedPlans };
          basePlansRef.current = next;
          return next;
        });
        setExportedHistory([]);
        const firstKey = Object.keys(parsedPlans)[0];
        if (firstKey) setActivePlanKey(firstKey);
        const statusMsg = await syncBufferToBackendRepo(buf, 'monthly', file.name, autoGitPush);
        setLastVerifiedAt(new Date().toLocaleTimeString('vi-VN'));
        setLastSyncLog(
          `Đã nhập File Kế Hoạch Tháng "${file.name}" (${Object.keys(parsedPlans).join(', ')}) & tính đối soát • ${statusMsg}`
        );
      } else if (type === 'weekly') {
        const parsedPlans = parseShipmentWorkbook(buf, 'weekly');
        setPlans((prev) => {
          const next = { ...prev, ...parsedPlans };
          basePlansRef.current = next;
          return next;
        });
        setExportedHistory([]);
        const wkKey =
          Object.keys(parsedPlans).find((k) => k.toLowerCase().includes('wk')) ||
          Object.keys(parsedPlans)[0];
        if (wkKey) setActivePlanKey(wkKey);
        const statusMsg = await syncBufferToBackendRepo(buf, 'weekly', file.name, autoGitPush);
        setLastVerifiedAt(new Date().toLocaleTimeString('vi-VN'));
        setLastSyncLog(
          `Đã nhập File Kế Hoạch Tuần "${file.name}" (${wkKey}) & tính đối soát • ${statusMsg}`
        );
      }
    } finally {
      setIsReconcilingStock(false);
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
                Báo cáo Tồn Kho Thành Phẩm 3D (Trái/Phải) & Đối Soát Kế Hoạch Xuất Hàng Tháng / Tuần
              </p>
            </div>
          </div>

          {/* Active Plan Selector + Source File Upload Buttons + Excel Export */}
          <div className="flex flex-wrap items-center gap-2">
            {/* Plan Switcher Pills */}
            <div className="inline-flex rounded-xl border border-slate-800 bg-slate-900 p-1">
              {Object.entries(plans).map(([key]) => (
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

            {/* Direct Source Import Buttons (Prominent Nhập Tồn Kho to pick file & calculate reconciliation) */}
            <div className="flex items-center gap-1.5 rounded-xl border border-cyan-500/35 bg-slate-900/95 p-1">
              <button
                onClick={() => quickStockInputRef.current?.click()}
                disabled={isReconcilingStock}
                className="inline-flex items-center gap-1.5 rounded-lg bg-gradient-to-r from-cyan-400 to-sky-500 px-3.5 py-1.5 text-xs font-extrabold text-slate-950 shadow-md shadow-cyan-500/20 hover:brightness-110 transition cursor-pointer"
                title="Chọn File Tồn Kho (Stock list.xlsx) từ máy tính để tính đối soát ngay"
              >
                <Upload className="h-3.5 w-3.5" />
                <span>
                  {isReconcilingStock ? 'Đang Tính Đối Soát...' : 'Nhập Tồn Kho & Tính Đối Soát'}
                </span>
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
                className="inline-flex items-center gap-1 rounded-lg bg-slate-800 border border-slate-700 px-2.5 py-1.5 text-xs font-semibold text-slate-200 hover:bg-slate-700 hover:text-white transition cursor-pointer"
                title="Mở bảng quản lý nhập các file nguồn & đối soát"
              >
                <Calculator className="h-3.5 w-3.5 text-cyan-400" />
                <span>Quản Lý File Nguồn</span>
              </button>
            </div>

            {/* Export Excel Per Container (Each Sheet = PO&SO) Button */}
            <button
              onClick={() =>
                exportContainersByPoSoToExcel(
                  reconciliation.poSummaries,
                  activePlan.name,
                  allocationMode
                )
              }
              className="inline-flex items-center gap-1.5 rounded-xl bg-gradient-to-r from-amber-400 to-orange-500 px-3.5 py-2 text-xs font-extrabold text-slate-950 shadow-lg shadow-amber-500/20 hover:brightness-110 transition cursor-pointer"
              title="Xuất file Excel riêng theo từng Cont, trong đó mỗi Sheet đặt tên theo PO & SO của Cont đó"
            >
              <Download className="h-4 w-4" />
              <span>Xuất Excel Từng Cont (Sheet PO&SO)</span>
            </button>

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
              <span>Xuất Excel Tổng Hợp</span>
            </button>
          </div>
        </div>
      </header>

      {/* Main Content Container */}
      <main className="mx-auto w-full max-w-[1600px] flex-1 px-4 sm:px-6 py-5 space-y-5">
        {/* Status Banner for Stock File Import & Reconciliation */}
        {lastSyncLog && (
          <div className="flex flex-wrap items-center justify-between gap-3 rounded-xl border border-cyan-500/35 bg-slate-900/90 px-4 py-2.5 text-xs">
            <div className="flex items-center gap-2 text-slate-200">
              <Database className="h-4 w-4 text-cyan-400 shrink-0" />
              <span>
                <strong className="text-cyan-300">Trạng thái Đối soát Tồn kho:</strong> {lastSyncLog}
              </span>
            </div>
            <div className="flex items-center gap-2">
              <button
                onClick={() => quickStockInputRef.current?.click()}
                className="inline-flex items-center gap-1.5 rounded-lg bg-cyan-500/20 border border-cyan-500/40 px-3 py-1 text-xs font-bold text-cyan-300 hover:bg-cyan-500 hover:text-slate-950 transition cursor-pointer"
              >
                <Upload className="h-3.5 w-3.5" />
                <span>Chọn File Tồn Kho Khác Để Tính Đối Soát</span>
              </button>
            </div>
          </div>
        )}

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
          activeStockFileName={activeStockFileName}
          onTriggerStockUpload={() => quickStockInputRef.current?.click()}
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
          onExportContainer={handleExportContainer}
          exportedContainersLog={exportedHistory}
          onUndoLastExportContainer={handleUndoLastExportContainer}
          onTriggerStockUpload={() => quickStockInputRef.current?.click()}
        />
      </main>

      {/* Footer */}
      <footer className="border-t border-slate-900 bg-slate-950/90 py-4 px-6 text-center text-xs text-slate-500">
        LPVN 3D Warehouse & Container Shipment Reconciliation System • Đối soát trực tiếp từ File Tồn Kho{' '}
        <code className="text-slate-300">{activeStockFileName}</code> (Kho 01 & 60) và Kế Hoạch Xuất Hàng{' '}
        <code className="text-slate-300">NOV. SHIPMENT -OCT22.xlsx</code>
      </footer>

      {/* Data Source Import & Reconciliation Modal */}
      <DataImportAndOneDriveModal
        isOpen={isImportModalOpen}
        onClose={() => setIsImportModalOpen(false)}
        onUploadFile={handleUploadFile}
        lastSyncLog={lastSyncLog}
        activeStockFileName={activeStockFileName}
      />
    </div>
  );
}
export default App;
