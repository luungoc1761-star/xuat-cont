import React, { useState, useMemo } from 'react';
import {
  CheckCircle2,
  AlertOctagon,
  Download,
  FileSpreadsheet,
  Filter,
  Layers,
  PackageCheck,
  PackageX,
  Ship,
  Columns
} from 'lucide-react';
import {
  AllocationMode,
  LeftRightPair,
  ReconciledOrderLine,
  ReconciledPoSummary,
  StockItem
} from '../types';
import { ReconciliationResult } from '../utils/reconciliationEngine';
import {
  exportFullReconciliationToExcel,
  exportContainersByPoSoToExcel,
  exportSingleContainerPoSoToExcel
} from '../utils/excelParserAndExporter';

interface FulfillmentReportSectionProps {
  reconciliation: ReconciliationResult;
  lrPairs: LeftRightPair[];
  stockItems: StockItem[];
  activePlanKey: string;
  activePlanName: string;
  allocationMode: AllocationMode;
  setAllocationMode: (mode: AllocationMode) => void;
  searchQuery: string;
  selectedModel: string | null;
  activeTab: 'po' | 'split' | 'fulfilled' | 'shortage' | 'items';
  setActiveTab: (t: 'po' | 'split' | 'fulfilled' | 'shortage' | 'items') => void;
  onExportContainer?: (poGroup: ReconciledPoSummary) => void;
  exportedContainersLog?: Array<{
    contLabel: string;
    po: string;
    so: string;
    deductedQty: number;
    timestamp: string;
  }>;
  onUndoLastExportContainer?: () => void;
}

export const FulfillmentReportSection: React.FC<FulfillmentReportSectionProps> = ({
  reconciliation,
  lrPairs,
  stockItems,
  activePlanName,
  allocationMode,
  setAllocationMode,
  searchQuery,
  selectedModel,
  activeTab,
  setActiveTab,
  onExportContainer,
  exportedContainersLog = [],
  onUndoLastExportContainer
}) => {
  const [customerFilter, setCustomerFilter] = useState<string>('ALL');
  const [poFilter, setPoFilter] = useState<string>('ALL');
  const [contStatusFilter, setContStatusFilter] = useState<'ALL' | 'SHORTAGE' | 'FULFILLED'>('ALL');

  const customers = useMemo(() => {
    return Array.from(
      new Set(reconciliation.reconciledLines.map((x) => x.customer).filter(Boolean))
    ).sort();
  }, [reconciliation.reconciledLines]);

  const pos = useMemo(() => {
    const base =
      customerFilter === 'ALL'
        ? reconciliation.reconciledLines
        : reconciliation.reconciledLines.filter((x) => x.customer === customerFilter);
    return Array.from(new Set(base.map((x) => x.po).filter(Boolean))).sort();
  }, [reconciliation.reconciledLines, customerFilter]);

  const filterLine = (l: ReconciledOrderLine) => {
    if (customerFilter !== 'ALL' && l.customer !== customerFilter) return false;
    if (poFilter !== 'ALL' && l.po !== poFilter) return false;
    if (selectedModel && l.feature !== selectedModel && !l.itemCode.includes(selectedModel)) {
      return false;
    }
    if (searchQuery.trim()) {
      const tokens = searchQuery
        .split(/[\s,;]+/)
        .map((t) => t.trim().toLowerCase())
        .filter(Boolean);
      const hay = `${l.itemCode} ${l.feature} ${l.description} ${l.po} ${l.so} ${l.customer}`.toLowerCase();
      if (!tokens.some((tk) => hay.includes(tk))) return false;
    }
    return true;
  };

  const filteredFulfilledLines = useMemo(
    () => reconciliation.fulfilledLines.filter(filterLine),
    [reconciliation.fulfilledLines, customerFilter, poFilter, selectedModel, searchQuery]
  );

  const filteredShortageLines = useMemo(
    () => reconciliation.shortageLines.filter(filterLine),
    [reconciliation.shortageLines, customerFilter, poFilter, selectedModel, searchQuery]
  );

  const filteredPoSummaries = useMemo(() => {
    return reconciliation.poSummaries.filter((p) => {
      if (contStatusFilter === 'SHORTAGE' && p.status !== 'SHORTAGE') return false;
      if (contStatusFilter === 'FULFILLED' && p.status !== 'FULFILLED') return false;
      if (customerFilter !== 'ALL' && p.customer !== customerFilter) return false;
      if (poFilter !== 'ALL' && p.po !== poFilter) return false;
      if (selectedModel) {
        return p.lines.some(
          (l) => l.feature === selectedModel || l.itemCode.includes(selectedModel)
        );
      }
      if (searchQuery.trim()) {
        const tokens = searchQuery
          .split(/[\s,;]+/)
          .map((t) => t.trim().toLowerCase())
          .filter(Boolean);
        return p.lines.some((l) => {
          const hay = `${l.itemCode} ${l.feature} ${l.description} ${l.po} ${l.so} ${l.customer}`.toLowerCase();
          return tokens.some((tk) => hay.includes(tk));
        });
      }
      return true;
    });
  }, [
    reconciliation.poSummaries,
    contStatusFilter,
    customerFilter,
    poFilter,
    selectedModel,
    searchQuery
  ]);

  return (
    <section
      id="fulfillment-report-section"
      className="rounded-2xl border border-slate-800/90 bg-slate-900/85 backdrop-blur-xl shadow-2xl overflow-hidden"
    >
      {/* Top Report Header & Excel Export Actions */}
      <div className="flex flex-wrap items-center justify-between gap-4 border-b border-slate-800 bg-slate-950/90 px-6 py-4">
        <div>
          <div className="flex flex-wrap items-center gap-2.5">
            <h2 className="text-lg font-extrabold text-white">
              Báo Cáo Tính Riêng Theo Từng Cont Xuất (Dựa Trên SO & PO) — Đơn Đủ Hàng & Mã Thiếu
            </h2>
            <span className="rounded-full bg-cyan-500/15 border border-cyan-500/30 px-2.5 py-0.5 text-xs font-semibold text-cyan-300">
              {activePlanName}
            </span>
          </div>
          <p className="mt-0.5 text-xs text-slate-400">
            Tính chi tiết từng Cont hàng theo <strong className="text-cyan-300">SO & PO</strong> đối chiếu với tồn kho <strong className="text-blue-300">Kho 01</strong> + <strong className="text-emerald-300">Kho 60</strong> để biết chính xác từng Cont thiếu những mã nào, số lượng bao nhiêu và xuất ra từng Sheet Excel theo <code className="text-amber-300">PO & SO</code>
          </p>
        </div>

        {/* Allocation Mode Toggle + Export Excel Buttons */}
        <div className="flex flex-wrap items-center gap-2.5">
          <div className="inline-flex rounded-xl border border-slate-700/80 bg-slate-900 p-1">
            <button
              onClick={() => setAllocationMode('direct')}
              className={`rounded-lg px-3 py-1.5 text-xs font-semibold transition cursor-pointer ${
                allocationMode === 'direct'
                  ? 'bg-cyan-500 text-slate-950 shadow'
                  : 'text-slate-300 hover:text-white'
              }`}
              title="Tính riêng độc lập từng Cont (SO & PO) so với Tổng Tồn Kho 01 + Kho 60"
            >
              Tính Riêng Độc Lập Từng Cont (SO & PO)
            </button>
            <button
              onClick={() => setAllocationMode('fifo')}
              className={`rounded-lg px-3 py-1.5 text-xs font-semibold transition cursor-pointer ${
                allocationMode === 'fifo'
                  ? 'bg-cyan-500 text-slate-950 shadow'
                  : 'text-slate-300 hover:text-white'
              }`}
              title="Trừ tồn kho lũy kế tuần tự từ Cont #1 đến Cont cuối"
            >
              Tính Lũy Kế FIFO Giữa Các Cont
            </button>
          </div>

          <button
            onClick={() =>
              exportContainersByPoSoToExcel(
                filteredPoSummaries,
                activePlanName,
                allocationMode
              )
            }
            className="inline-flex items-center gap-2 rounded-xl bg-gradient-to-r from-amber-400 to-orange-500 px-4 py-2.5 text-xs font-extrabold text-slate-950 shadow-lg shadow-amber-500/25 hover:brightness-110 transition cursor-pointer"
            title="Xuất file Excel trong đó mỗi Sheet là 1 PO & SO của từng Cont"
          >
            <FileSpreadsheet className="h-4 w-4" />
            <span>Xuất Excel Theo Từng Cont (Mỗi Sheet = PO & SO)</span>
          </button>

          <button
            onClick={() =>
              exportFullReconciliationToExcel(
                reconciliation,
                lrPairs,
                stockItems,
                activePlanName,
                allocationMode
              )
            }
            className="inline-flex items-center gap-2 rounded-xl bg-gradient-to-r from-emerald-500 to-teal-400 px-4 py-2.5 text-xs font-extrabold text-slate-950 shadow-lg shadow-emerald-500/25 hover:brightness-110 transition cursor-pointer"
          >
            <Download className="h-4 w-4" />
            <span>Xuất Excel Toàn Bộ Báo Cáo</span>
          </button>
        </div>
      </div>

      {/* Sub-navigation Tabs & Customer/PO Filters */}
      <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-800/80 bg-slate-900/95 px-6 py-3">
        <div className="flex flex-wrap items-center gap-1.5">
          <button
            onClick={() => setActiveTab('po')}
            className={`inline-flex items-center gap-1.5 rounded-xl px-3.5 py-2 text-xs font-bold transition cursor-pointer ${
              activeTab === 'po'
                ? 'bg-amber-400 text-slate-950 shadow-md shadow-amber-400/25'
                : 'bg-slate-800/90 text-amber-300 hover:bg-slate-800'
            }`}
          >
            <Ship className="h-3.5 w-3.5" />
            <span>
              Theo Từng Cont Xuất PO & SO ({reconciliation.poSummaries.length} Cont: {reconciliation.fulfilledPos.length} Đủ / {reconciliation.shortagePos.length} Thiếu)
            </span>
          </button>

          <button
            onClick={() => setActiveTab('split')}
            className={`inline-flex items-center gap-1.5 rounded-xl px-3.5 py-2 text-xs font-bold transition cursor-pointer ${
              activeTab === 'split'
                ? 'bg-cyan-500 text-slate-950 shadow-md shadow-cyan-500/20'
                : 'bg-slate-800/90 text-slate-300 hover:text-white'
            }`}
          >
            <Columns className="h-3.5 w-3.5" />
            <span>Song Song: Đơn Đủ & Hàng Thiếu</span>
          </button>

          <button
            onClick={() => setActiveTab('fulfilled')}
            className={`inline-flex items-center gap-1.5 rounded-xl px-3.5 py-2 text-xs font-bold transition cursor-pointer ${
              activeTab === 'fulfilled'
                ? 'bg-emerald-500 text-slate-950 shadow-md shadow-emerald-500/20'
                : 'bg-slate-800/90 text-emerald-300 hover:bg-slate-800'
            }`}
          >
            <PackageCheck className="h-3.5 w-3.5" />
            <span>Dòng Đủ Hàng Xuất ({filteredFulfilledLines.length})</span>
          </button>

          <button
            onClick={() => setActiveTab('shortage')}
            className={`inline-flex items-center gap-1.5 rounded-xl px-3.5 py-2 text-xs font-bold transition cursor-pointer ${
              activeTab === 'shortage'
                ? 'bg-rose-500 text-white shadow-md shadow-rose-500/20'
                : 'bg-slate-800/90 text-rose-300 hover:bg-slate-800'
            }`}
          >
            <PackageX className="h-3.5 w-3.5" />
            <span>Dòng Hàng Thiếu ({filteredShortageLines.length})</span>
          </button>

          <button
            onClick={() => setActiveTab('items')}
            className={`inline-flex items-center gap-1.5 rounded-xl px-3.5 py-2 text-xs font-bold transition cursor-pointer ${
              activeTab === 'items'
                ? 'bg-indigo-500 text-white shadow-md shadow-indigo-500/20'
                : 'bg-slate-800/90 text-indigo-300 hover:bg-slate-800'
            }`}
          >
            <Layers className="h-3.5 w-3.5" />
            <span>Tổng Cầu Theo Mã Hàng ({reconciliation.itemDemands.length} mã)</span>
          </button>
        </div>

        {/* Customer & PO Select Filters */}
        <div className="flex flex-wrap items-center gap-2">
          <div className="flex items-center gap-1.5 text-xs text-slate-400">
            <Filter className="h-3.5 w-3.5 text-cyan-400" />
            <span>Lọc:</span>
          </div>
          <select
            value={customerFilter}
            onChange={(e) => {
              setCustomerFilter(e.target.value);
              setPoFilter('ALL');
            }}
            className="rounded-xl border border-slate-700 bg-slate-950 px-3 py-1.5 text-xs text-white focus:border-cyan-400 focus:outline-none"
          >
            <option value="ALL">Tất cả khách hàng ({customers.length})</option>
            {customers.map((c) => (
              <option key={c} value={c}>
                {c}
              </option>
            ))}
          </select>

          <select
            value={poFilter}
            onChange={(e) => setPoFilter(e.target.value)}
            className="rounded-xl border border-slate-700 bg-slate-950 px-3 py-1.5 text-xs font-mono text-white focus:border-cyan-400 focus:outline-none"
          >
            <option value="ALL">Tất cả PO ({pos.length})</option>
            {pos.map((p) => (
              <option key={p} value={p}>
                PO: {p}
              </option>
            ))}
          </select>
        </div>
      </div>

      {/* TAB 1 (DEFAULT): DETAILED PER-CONTAINER REPORT BY PO & SO */}
      {activeTab === 'po' && (
        <div className="p-5 space-y-4">
          {/* Quick Status Filter Bar for Containers */}
          <div className="flex flex-wrap items-center justify-between gap-3 rounded-xl border border-slate-800 bg-slate-950/80 px-4 py-3">
            <div className="flex flex-wrap items-center gap-2">
              <span className="text-xs font-bold text-slate-300">Lọc nhanh trạng thái Cont (PO & SO):</span>
              <button
                onClick={() => setContStatusFilter('ALL')}
                className={`rounded-lg px-3 py-1.5 text-xs font-bold transition cursor-pointer ${
                  contStatusFilter === 'ALL'
                    ? 'bg-cyan-500 text-slate-950'
                    : 'bg-slate-800 text-slate-300 hover:text-white'
                }`}
              >
                Tất Cả Cont ({reconciliation.poSummaries.length})
              </button>
              <button
                onClick={() => setContStatusFilter('SHORTAGE')}
                className={`rounded-lg px-3 py-1.5 text-xs font-bold transition cursor-pointer ${
                  contStatusFilter === 'SHORTAGE'
                    ? 'bg-rose-500 text-white'
                    : 'bg-rose-950/40 border border-rose-500/30 text-rose-300 hover:bg-rose-900/50'
                }`}
              >
                🚨 Cont Đang Thiếu Hàng ({reconciliation.shortagePos.length} Cont)
              </button>
              <button
                onClick={() => setContStatusFilter('FULFILLED')}
                className={`rounded-lg px-3 py-1.5 text-xs font-bold transition cursor-pointer ${
                  contStatusFilter === 'FULFILLED'
                    ? 'bg-emerald-500 text-slate-950'
                    : 'bg-emerald-950/40 border border-emerald-500/30 text-emerald-300 hover:bg-emerald-900/50'
                }`}
              >
                ✅ Cont Đủ Hàng Xuất 100% ({reconciliation.fulfilledPos.length} Cont)
              </button>
            </div>

            <div className="text-xs text-slate-400 font-mono">
              Mỗi Cont bên dưới tương ứng 1 Sheet Excel tên <span className="text-amber-300 font-bold">PO..._SO...</span>
            </div>
          </div>

          {exportedContainersLog.length > 0 && (
            <div className="flex flex-wrap items-center justify-between gap-3 rounded-xl border border-emerald-500/40 bg-emerald-950/30 px-4 py-2.5 text-xs">
              <div className="flex flex-wrap items-center gap-2 text-emerald-200">
                <CheckCircle2 className="h-4 w-4 text-emerald-400 shrink-0" />
                <span>
                  Đã hoàn thành xuất <strong className="text-white">{exportedContainersLog.length} Cont</strong> (Đã trừ tồn kho đối soát lũy kế cho các đơn tiếp theo & đưa khỏi kế hoạch xuất):{' '}
                  {exportedContainersLog.map((c, i) => (
                    <span
                      key={`${c.contLabel}-${c.po}-${i}`}
                      className="inline-flex items-center gap-1 rounded bg-emerald-500/20 border border-emerald-500/30 px-2 py-0.5 font-mono text-[11px] text-emerald-300 mr-1.5"
                    >
                      {c.contLabel} (PO: {c.po} | -{c.deductedQty.toLocaleString('en-US')} pcs)
                    </span>
                  ))}
                </span>
              </div>
              {onUndoLastExportContainer && (
                <button
                  onClick={onUndoLastExportContainer}
                  className="rounded-lg border border-slate-700 bg-slate-900 px-2.5 py-1 text-[11px] font-semibold text-slate-300 hover:border-cyan-400 hover:text-white transition cursor-pointer"
                >
                  Hoàn tác Cont vừa xuất
                </button>
              )}
            </div>
          )}

          {/* List of Container (PO & SO) Detailed Cards */}
          <div className="space-y-5">
            {filteredPoSummaries.map((poGroup) => {
              const isFull = poGroup.status === 'FULFILLED';
              return (
                <div
                  key={poGroup.id}
                  className={`rounded-2xl border overflow-hidden shadow-lg transition ${
                    isFull
                      ? 'border-emerald-500/35 bg-slate-950/80'
                      : 'border-rose-500/45 bg-slate-950/85'
                  }`}
                >
                  {/* Container Card Header */}
                  <div
                    className={`flex flex-wrap items-center justify-between gap-3 px-5 py-3.5 border-b ${
                      isFull
                        ? 'border-emerald-500/25 bg-emerald-950/25'
                        : 'border-rose-500/30 bg-rose-950/30'
                    }`}
                  >
                    <div className="flex flex-wrap items-center gap-3">
                      <span className="rounded-lg bg-slate-900 border border-slate-700 px-2.5 py-1 text-xs font-mono font-extrabold text-cyan-300">
                        CONT #{String(poGroup.containerIndex).padStart(2, '0')}
                      </span>
                      <div>
                        <div className="flex flex-wrap items-center gap-2">
                          <span className="text-sm sm:text-base font-extrabold font-mono text-white">
                            PO: <span className="text-amber-300">{poGroup.po}</span>
                          </span>
                          <span className="text-slate-500">|</span>
                          <span className="text-sm sm:text-base font-extrabold font-mono text-white">
                            SO: <span className="text-cyan-300">{poGroup.so || poGroup.soList.join(', ') || 'N/A'}</span>
                          </span>
                          <span className="rounded bg-slate-800/90 border border-slate-700 px-2 py-0.5 text-[11px] font-mono text-slate-300">
                            Sheet Excel: <strong className="text-amber-300">{poGroup.sheetNameExcel}</strong>
                          </span>
                          {isFull ? (
                            <span className="inline-flex items-center gap-1 rounded-full bg-emerald-500/20 border border-emerald-500/40 px-2.5 py-0.5 text-xs font-bold text-emerald-300">
                              <CheckCircle2 className="h-3.5 w-3.5" /> ĐỦ HÀNG XUẤT CONT 100% ({poGroup.totalLines}/{poGroup.totalLines} mã)
                            </span>
                          ) : (
                            <span className="inline-flex items-center gap-1 rounded-full bg-rose-500/25 border border-rose-500/50 px-2.5 py-0.5 text-xs font-bold text-rose-300">
                              <AlertOctagon className="h-3.5 w-3.5" /> THIẾU {poGroup.shortageLinesCount}/{poGroup.totalLines} MÃ (Thiếu -{poGroup.totalShortageQty.toLocaleString('en-US')} pcs)
                            </span>
                          )}
                        </div>
                        <div className="mt-0.5 text-xs text-slate-300">
                          Khách hàng: <strong className="text-white">{poGroup.customer}</strong> • Ngày đóng Cont:{' '}
                          <strong className="font-mono text-cyan-300">{poGroup.loadingDate || 'Chưa cập nhật'}</strong> • Tổng kiện:{' '}
                          <strong className="font-mono text-white">{poGroup.totalPkgs.toLocaleString('en-US')} pkg</strong> • Thể tích:{' '}
                          <strong className="font-mono text-white">{poGroup.totalCbm.toFixed(2)} CBM</strong>
                        </div>
                      </div>
                    </div>

                    <div className="flex items-center gap-3">
                      <div className="text-right font-mono text-xs">
                        <div className="text-slate-400">Tổng SL Cont</div>
                        <div className="text-base font-extrabold text-white">
                          {poGroup.totalRequiredQty.toLocaleString('en-US')} pcs
                        </div>
                      </div>
                      <button
                        onClick={() =>
                          exportSingleContainerPoSoToExcel(poGroup, allocationMode)
                        }
                        className="inline-flex items-center gap-1.5 rounded-xl bg-amber-400/20 border border-amber-400/40 px-3 py-2 text-xs font-bold text-amber-200 hover:bg-amber-400 hover:text-slate-950 transition cursor-pointer"
                        title={`Xuất riêng Sheet Excel ${poGroup.sheetNameExcel} của Cont này`}
                      >
                        <Download className="h-3.5 w-3.5" />
                        <span>Xuất Sheet Cont Này</span>
                      </button>
                      {onExportContainer && (
                        <button
                          onClick={() => onExportContainer(poGroup)}
                          className="inline-flex items-center gap-1.5 rounded-xl bg-gradient-to-r from-emerald-500 to-cyan-400 px-3.5 py-2 text-xs font-extrabold text-slate-950 shadow-md shadow-emerald-500/20 hover:brightness-110 transition cursor-pointer"
                          title="Xuất Cont: Trừ tồn đối soát khỏi hệ thống tồn kho, trừ lũy kế cho các đơn tiếp theo và đưa đơn này ra khỏi kế hoạch xuất"
                        >
                          <Ship className="h-3.5 w-3.5" />
                          <span>Xuất Cont</span>
                        </button>
                      )}
                    </div>
                  </div>

                  {/* Highlighted Box: Exact Missing Codes & Quantities for this Cont (if any) */}
                  {poGroup.shortageLines.length > 0 && (
                    <div className="m-4 rounded-xl border border-rose-500/40 bg-rose-950/25 p-4">
                      <div className="flex flex-wrap items-center justify-between gap-2 mb-2.5">
                        <span className="text-xs font-extrabold uppercase tracking-wider text-rose-300 flex items-center gap-1.5">
                          <AlertOctagon className="h-4 w-4 text-rose-400" />
                          Danh Sách Mã Hàng Đang Thiếu Của Cont (PO: {poGroup.po} & SO: {poGroup.so || 'N/A'}) — Thiếu {poGroup.shortageLines.length} Mã:
                        </span>
                        <span className="text-xs font-mono font-bold text-rose-300">
                          Tổng số lượng thiếu của Cont: -{poGroup.totalShortageQty.toLocaleString('en-US')} pcs
                        </span>
                      </div>
                      <div className="overflow-x-auto rounded-lg border border-rose-500/30 bg-slate-950/90">
                        <table className="w-full text-left border-collapse text-xs">
                          <thead className="bg-rose-950/50 text-rose-200 border-b border-rose-500/30">
                            <tr>
                              <th className="py-2 px-3">Mã Hàng Thiếu</th>
                              <th className="py-2 px-3">Model / Bên</th>
                              <th className="py-2 px-3">Mô Tả</th>
                              <th className="py-2 px-3 text-right">SL Cont Cần</th>
                              <th className="py-2 px-3 text-right text-blue-300">Tồn Kho 01</th>
                              <th className="py-2 px-3 text-right text-cyan-300">Tồn Kho 60</th>
                              <th className="py-2 px-3 text-right">Tồn Khả Dụng (01+60)</th>
                              <th className="py-2 px-3 text-right text-emerald-300">Đáp Ứng Được</th>
                              <th className="py-2 px-3 text-right text-rose-400 font-extrabold">SỐ LƯỢNG THIẾU</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-rose-500/20 font-mono">
                            {poGroup.shortageLines.map((l) => (
                              <tr key={l.id} className="hover:bg-rose-950/30">
                                <td className="py-2 px-3 font-bold text-amber-300">{l.itemCode}</td>
                                <td className="py-2 px-3">
                                  <span className="text-slate-300 mr-1">#{l.feature}</span>
                                  <span
                                    className={`rounded px-1.5 py-0.5 text-[10px] font-bold ${
                                      l.side === 'L'
                                        ? 'bg-cyan-500/20 text-cyan-300'
                                        : l.side === 'R'
                                        ? 'bg-emerald-500/20 text-emerald-300'
                                        : 'bg-slate-800 text-slate-400'
                                    }`}
                                  >
                                    {l.side === 'L' ? 'Trái (LHF)' : l.side === 'R' ? 'Phải (RHF)' : 'Chung'}
                                  </span>
                                </td>
                                <td className="py-2 px-3 font-sans text-slate-200">{l.description}</td>
                                <td className="py-2 px-3 text-right font-bold text-white">
                                  {l.qty.toLocaleString('en-US')}
                                </td>
                                <td className="py-2 px-3 text-right text-blue-300">
                                  {l.wh01Initial.toLocaleString('en-US')}
                                </td>
                                <td className="py-2 px-3 text-right text-cyan-300">
                                  {l.wh60Initial.toLocaleString('en-US')}
                                </td>
                                <td className="py-2 px-3 text-right text-slate-200">
                                  {l.stockBeforeLine.toLocaleString('en-US')}
                                </td>
                                <td className="py-2 px-3 text-right text-emerald-300">
                                  {l.allocatedTotal.toLocaleString('en-US')}
                                </td>
                                <td className="py-2 px-3 text-right font-extrabold text-rose-400 text-sm">
                                  -{l.shortageQty.toLocaleString('en-US')}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}

                  {/* Full Table of All Lines in this Container (PO & SO) */}
                  <div className="px-4 pb-4 pt-2">
                    <div className="overflow-x-auto rounded-xl border border-slate-800 bg-slate-900/60">
                      <table className="w-full text-left border-collapse text-xs">
                        <thead className="bg-slate-900 text-slate-300 border-b border-slate-800">
                          <tr>
                            <th className="py-2 px-3">STT</th>
                            <th className="py-2 px-3">PO / SO</th>
                            <th className="py-2 px-3">Mã Hàng (LPVN Code)</th>
                            <th className="py-2 px-3">Model / Bên</th>
                            <th className="py-2 px-3">Mô Tả Hàng Hóa</th>
                            <th className="py-2 px-3 text-right">SL Cont Cần</th>
                            <th className="py-2 px-3 text-right text-blue-300">Kho 01</th>
                            <th className="py-2 px-3 text-right text-cyan-300">Kho 60</th>
                            <th className="py-2 px-3 text-right">Tồn Khả Dụng</th>
                            <th className="py-2 px-3 text-right text-emerald-300">SL Đáp Ứng</th>
                            <th className="py-2 px-3 text-right text-rose-400">SL Thiếu</th>
                            <th className="py-2 px-3 text-center">Trạng Thái</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-800/60 font-mono">
                          {poGroup.lines.map((l, idx) => (
                            <tr
                              key={l.id}
                              className={
                                l.shortageQty > 0
                                  ? 'bg-rose-950/15 hover:bg-rose-950/25'
                                  : 'hover:bg-slate-800/40'
                              }
                            >
                              <td className="py-2 px-3 text-slate-400">{idx + 1}</td>
                              <td className="py-2 px-3">
                                <span className="text-amber-300 font-bold">{l.po}</span>
                                <span className="text-slate-500 mx-1">/</span>
                                <span className="text-cyan-300">{l.so}</span>
                              </td>
                              <td className="py-2 px-3 font-bold text-white">{l.itemCode}</td>
                              <td className="py-2 px-3">
                                <span className="text-slate-300 mr-1">#{l.feature}</span>
                                <span
                                  className={`rounded px-1.5 py-0.5 text-[10px] font-bold ${
                                    l.side === 'L'
                                      ? 'bg-cyan-500/20 text-cyan-300'
                                      : l.side === 'R'
                                      ? 'bg-emerald-500/20 text-emerald-300'
                                      : 'bg-slate-800 text-slate-400'
                                  }`}
                                >
                                  {l.side === 'L' ? 'Trái' : l.side === 'R' ? 'Phải' : 'Chung'}
                                </span>
                              </td>
                              <td className="py-2 px-3 font-sans text-slate-200 max-w-[220px] truncate">
                                {l.description}
                              </td>
                              <td className="py-2 px-3 text-right font-bold text-white">
                                {l.qty.toLocaleString('en-US')}
                              </td>
                              <td className="py-2 px-3 text-right text-blue-300">
                                {l.wh01Initial.toLocaleString('en-US')}
                              </td>
                              <td className="py-2 px-3 text-right text-cyan-300">
                                {l.wh60Initial.toLocaleString('en-US')}
                              </td>
                              <td className="py-2 px-3 text-right text-slate-300">
                                {l.stockBeforeLine.toLocaleString('en-US')}
                              </td>
                              <td className="py-2 px-3 text-right text-emerald-300 font-semibold">
                                {l.allocatedTotal.toLocaleString('en-US')}
                              </td>
                              <td className="py-2 px-3 text-right font-bold">
                                {l.shortageQty > 0 ? (
                                  <span className="text-rose-400">
                                    -{l.shortageQty.toLocaleString('en-US')}
                                  </span>
                                ) : (
                                  <span className="text-slate-600">0</span>
                                )}
                              </td>
                              <td className="py-2 px-3 text-center">
                                {l.shortageQty > 0 ? (
                                  <span className="rounded bg-rose-500/20 border border-rose-500/40 px-2 py-0.5 text-[10px] font-bold text-rose-300">
                                    THIẾU -{l.shortageQty.toLocaleString('en-US')}
                                  </span>
                                ) : (
                                  <span className="rounded bg-emerald-500/15 border border-emerald-500/30 px-2 py-0.5 text-[10px] font-bold text-emerald-300">
                                    ĐỦ HÀNG
                                  </span>
                                )}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* TAB 2: SPLIT VIEW (Both Fulfilled Orders & Shortage Orders visible together) */}
      {activeTab === 'split' && (
        <div className="p-5 grid grid-cols-1 xl:grid-cols-2 gap-5">
          {/* Left Panel: Fulfilled Order Lines */}
          <div className="rounded-xl border border-emerald-500/35 bg-slate-950/70 overflow-hidden flex flex-col">
            <div className="flex items-center justify-between border-b border-emerald-500/25 bg-emerald-950/30 px-4 py-3">
              <div className="flex items-center gap-2">
                <CheckCircle2 className="h-4 w-4 text-emerald-400" />
                <h3 className="text-sm font-bold text-emerald-300">
                  ĐƠN ĐỦ HÀNG XUẤT (KHO 01 + KHO 60 ĐÁP ỨNG 100%)
                </h3>
              </div>
              <span className="rounded-full bg-emerald-500/20 border border-emerald-500/40 px-2.5 py-0.5 text-xs font-mono font-bold text-emerald-300">
                {filteredFulfilledLines.length} dòng ({reconciliation.fulfilledPos.length} Cont đủ 100%)
              </span>
            </div>
            <div className="max-h-[500px] overflow-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead className="sticky top-0 bg-slate-900/95 text-slate-300 border-b border-slate-800 z-10">
                  <tr>
                    <th className="py-2.5 px-3">Cont / PO / SO</th>
                    <th className="py-2.5 px-3">Mã Hàng & Bên</th>
                    <th className="py-2.5 px-3 text-right">Cần Xuất</th>
                    <th className="py-2.5 px-3 text-right text-blue-300">Kho 01</th>
                    <th className="py-2.5 px-3 text-right text-cyan-300">Kho 60</th>
                    <th className="py-2.5 px-3 text-right text-emerald-300">Còn Dư</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/60 font-mono">
                  {filteredFulfilledLines.map((l) => (
                    <tr key={l.id} className="hover:bg-emerald-950/15 transition">
                      <td className="py-2 px-3">
                        <div className="font-sans font-semibold text-white">
                          Cont #{String(l.containerIndex).padStart(2, '0')} • {l.customer}
                        </div>
                        <div className="text-[11px] text-slate-400">
                          PO: <span className="text-emerald-300">{l.po}</span> | SO:{' '}
                          <span className="text-cyan-300">{l.so}</span>
                        </div>
                      </td>
                      <td className="py-2 px-3">
                        <div className="flex items-center gap-1.5">
                          <span className="font-bold text-cyan-300">{l.itemCode}</span>
                          <span
                            className={`rounded px-1 py-0.2 text-[10px] ${
                              l.side === 'L'
                                ? 'bg-cyan-500/20 text-cyan-300'
                                : l.side === 'R'
                                ? 'bg-emerald-500/20 text-emerald-300'
                                : 'bg-slate-800 text-slate-400'
                            }`}
                          >
                            {l.side === 'L' ? 'Trái' : l.side === 'R' ? 'Phải' : 'Chung'}
                          </span>
                        </div>
                        <div className="font-sans text-[11px] text-slate-400 truncate max-w-[190px]">
                          {l.description}
                        </div>
                      </td>
                      <td className="py-2 px-3 text-right font-bold text-white">
                        {l.qty.toLocaleString('en-US')}
                      </td>
                      <td className="py-2 px-3 text-right text-blue-300">
                        {l.wh01Initial.toLocaleString('en-US')}
                      </td>
                      <td className="py-2 px-3 text-right text-cyan-300">
                        {l.wh60Initial.toLocaleString('en-US')}
                      </td>
                      <td className="py-2 px-3 text-right font-bold text-emerald-400">
                        +{l.stockAfterLine.toLocaleString('en-US')}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* Right Panel: Shortage Order Lines */}
          <div className="rounded-xl border border-rose-500/35 bg-slate-950/70 overflow-hidden flex flex-col">
            <div className="flex items-center justify-between border-b border-rose-500/25 bg-rose-950/30 px-4 py-3">
              <div className="flex items-center gap-2">
                <AlertOctagon className="h-4 w-4 text-rose-400" />
                <h3 className="text-sm font-bold text-rose-300">
                  ĐƠN & MÃ HÀNG THIẾU (KHÔNG ĐỦ TỒN KHO 01 + KHO 60)
                </h3>
              </div>
              <span className="rounded-full bg-rose-500/20 border border-rose-500/40 px-2.5 py-0.5 text-xs font-mono font-bold text-rose-300">
                {filteredShortageLines.length} dòng thiếu ({reconciliation.shortagePos.length} Cont thiếu)
              </span>
            </div>
            <div className="max-h-[500px] overflow-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead className="sticky top-0 bg-slate-900/95 text-slate-300 border-b border-slate-800 z-10">
                  <tr>
                    <th className="py-2.5 px-3">Cont / PO / SO</th>
                    <th className="py-2.5 px-3">Mã Hàng & Bên</th>
                    <th className="py-2.5 px-3 text-right">Cần Xuất</th>
                    <th className="py-2.5 px-3 text-right">Tồn Khả Dụng</th>
                    <th className="py-2.5 px-3 text-right text-rose-400">SL Thiếu</th>
                    <th className="py-2.5 px-3 text-center">Mức Độ</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/60 font-mono">
                  {filteredShortageLines.map((l) => (
                    <tr key={l.id} className="hover:bg-rose-950/20 transition">
                      <td className="py-2 px-3">
                        <div className="font-sans font-semibold text-white">
                          Cont #{String(l.containerIndex).padStart(2, '0')} • {l.customer}
                        </div>
                        <div className="text-[11px] text-slate-400">
                          PO: <span className="text-rose-300">{l.po}</span> | SO:{' '}
                          <span className="text-cyan-300">{l.so}</span>
                        </div>
                      </td>
                      <td className="py-2 px-3">
                        <div className="flex items-center gap-1.5">
                          <span className="font-bold text-amber-300">{l.itemCode}</span>
                          <span
                            className={`rounded px-1 py-0.2 text-[10px] ${
                              l.side === 'L'
                                ? 'bg-cyan-500/20 text-cyan-300'
                                : l.side === 'R'
                                ? 'bg-emerald-500/20 text-emerald-300'
                                : 'bg-slate-800 text-slate-400'
                            }`}
                          >
                            {l.side === 'L' ? 'Trái' : l.side === 'R' ? 'Phải' : 'Chung'}
                          </span>
                        </div>
                        <div className="font-sans text-[11px] text-slate-400 truncate max-w-[190px]">
                          {l.description}
                        </div>
                      </td>
                      <td className="py-2 px-3 text-right font-bold text-white">
                        {l.qty.toLocaleString('en-US')}
                      </td>
                      <td className="py-2 px-3 text-right text-slate-300">
                        <div>{l.stockBeforeLine.toLocaleString('en-US')}</div>
                        <div className="text-[10px] text-slate-500">
                          (01:{l.wh01Initial} | 60:{l.wh60Initial})
                        </div>
                      </td>
                      <td className="py-2 px-3 text-right font-extrabold text-rose-400">
                        -{l.shortageQty.toLocaleString('en-US')}
                      </td>
                      <td className="py-2 px-3 text-center">
                        {l.fulfillmentStatus === 'ZERO_STOCK' ? (
                          <span className="rounded bg-rose-500/25 border border-rose-500/40 px-1.5 py-0.5 text-[10px] font-bold text-rose-300">
                            Tồn=0
                          </span>
                        ) : (
                          <span className="rounded bg-amber-500/25 border border-amber-500/40 px-1.5 py-0.5 text-[10px] font-bold text-amber-300">
                            Đủ {l.fulfillmentRate.toFixed(0)}%
                          </span>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* TAB 3: FULL TABLE OF FULFILLED ORDERS */}
      {activeTab === 'fulfilled' && (
        <div className="p-5">
          <div className="max-h-[560px] overflow-auto rounded-xl border border-emerald-500/30 bg-slate-950/75">
            <table className="w-full text-left border-collapse text-xs">
              <thead className="sticky top-0 bg-slate-900 text-slate-300 border-b border-slate-800 z-10">
                <tr>
                  <th className="py-3 px-3">Cont #</th>
                  <th className="py-3 px-3">Khách Hàng</th>
                  <th className="py-3 px-3">PO & SO</th>
                  <th className="py-3 px-3">Mã Hàng (LPVN Code)</th>
                  <th className="py-3 px-3">Model / Bên</th>
                  <th className="py-3 px-3">Mô Tả Hàng Hóa</th>
                  <th className="py-3 px-3">Ngày Xuất</th>
                  <th className="py-3 px-3 text-right">SL Đặt Xuất</th>
                  <th className="py-3 px-3 text-right text-blue-300">Tồn Đầu Kho 01</th>
                  <th className="py-3 px-3 text-right text-cyan-300">Tồn Đầu Kho 60</th>
                  <th className="py-3 px-3 text-right text-emerald-300">Tồn Khả Dụng</th>
                  <th className="py-3 px-3 text-right">Tồn Sau Đơn</th>
                  <th className="py-3 px-3 text-center">Trạng Thái</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/60 font-mono">
                {filteredFulfilledLines.map((l) => (
                  <tr key={l.id} className="hover:bg-emerald-950/20 transition">
                    <td className="py-2.5 px-3 text-cyan-300 font-bold">
                      Cont #{String(l.containerIndex).padStart(2, '0')}
                    </td>
                    <td className="py-2.5 px-3 font-sans font-semibold text-white">{l.customer}</td>
                    <td className="py-2.5 px-3">
                      <div className="text-emerald-300 font-bold">PO: {l.po}</div>
                      {l.so && <div className="text-[10px] text-cyan-300">SO: {l.so}</div>}
                    </td>
                    <td className="py-2.5 px-3 font-bold text-cyan-300">{l.itemCode}</td>
                    <td className="py-2.5 px-3">
                      <span className="text-slate-300 mr-1.5">#{l.feature}</span>
                      <span
                        className={`rounded px-1.5 py-0.5 text-[10px] font-bold ${
                          l.side === 'L'
                            ? 'bg-cyan-500/20 text-cyan-300'
                            : l.side === 'R'
                            ? 'bg-emerald-500/20 text-emerald-300'
                            : 'bg-slate-800 text-slate-400'
                        }`}
                      >
                        {l.side === 'L' ? 'Trái (LHF)' : l.side === 'R' ? 'Phải (RHF)' : 'Chung'}
                      </span>
                    </td>
                    <td className="py-2.5 px-3 font-sans text-slate-200 max-w-[220px] truncate">
                      {l.description}
                    </td>
                    <td className="py-2.5 px-3 text-slate-300">
                      {l.updatedLoadingDate || l.loadingDate || l.crd}
                    </td>
                    <td className="py-2.5 px-3 text-right font-bold text-white">
                      {l.qty.toLocaleString('en-US')}
                    </td>
                    <td className="py-2.5 px-3 text-right text-blue-300">
                      {l.wh01Initial.toLocaleString('en-US')}
                    </td>
                    <td className="py-2.5 px-3 text-right text-cyan-300">
                      {l.wh60Initial.toLocaleString('en-US')}
                    </td>
                    <td className="py-2.5 px-3 text-right font-semibold text-emerald-300">
                      {l.stockBeforeLine.toLocaleString('en-US')}
                    </td>
                    <td className="py-2.5 px-3 text-right text-slate-300">
                      {l.stockAfterLine.toLocaleString('en-US')}
                    </td>
                    <td className="py-2.5 px-3 text-center">
                      <span className="inline-flex items-center gap-1 rounded-full bg-emerald-500/20 border border-emerald-500/40 px-2.5 py-0.5 text-[10px] font-bold text-emerald-300">
                        <CheckCircle2 className="h-3 w-3" /> ĐỦ HÀNG 100%
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* TAB 4: FULL TABLE OF SHORTAGE ORDERS */}
      {activeTab === 'shortage' && (
        <div className="p-5">
          <div className="max-h-[560px] overflow-auto rounded-xl border border-rose-500/30 bg-slate-950/75">
            <table className="w-full text-left border-collapse text-xs">
              <thead className="sticky top-0 bg-slate-900 text-slate-300 border-b border-slate-800 z-10">
                <tr>
                  <th className="py-3 px-3">Cont #</th>
                  <th className="py-3 px-3">Khách Hàng</th>
                  <th className="py-3 px-3">PO & SO</th>
                  <th className="py-3 px-3">Mã Hàng (LPVN Code)</th>
                  <th className="py-3 px-3">Model / Bên</th>
                  <th className="py-3 px-3">Mô Tả Hàng Hóa</th>
                  <th className="py-3 px-3">Ngày Xuất</th>
                  <th className="py-3 px-3 text-right">SL Cần Xuất</th>
                  <th className="py-3 px-3 text-right text-blue-300">Tồn Đầu Kho 01</th>
                  <th className="py-3 px-3 text-right text-cyan-300">Tồn Đầu Kho 60</th>
                  <th className="py-3 px-3 text-right">Tồn Trước Đơn</th>
                  <th className="py-3 px-3 text-right text-emerald-300">Đáp Ứng Được</th>
                  <th className="py-3 px-3 text-right text-rose-400">SL THIẾU HỤT</th>
                  <th className="py-3 px-3 text-center">Trạng Thái Thiếu</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/60 font-mono">
                {filteredShortageLines.map((l) => (
                  <tr key={l.id} className="hover:bg-rose-950/20 transition">
                    <td className="py-2.5 px-3 text-rose-300 font-bold">
                      Cont #{String(l.containerIndex).padStart(2, '0')}
                    </td>
                    <td className="py-2.5 px-3 font-sans font-semibold text-white">{l.customer}</td>
                    <td className="py-2.5 px-3">
                      <div className="text-rose-300 font-bold">PO: {l.po}</div>
                      {l.so && <div className="text-[10px] text-cyan-300">SO: {l.so}</div>}
                    </td>
                    <td className="py-2.5 px-3 font-bold text-amber-300">{l.itemCode}</td>
                    <td className="py-2.5 px-3">
                      <span className="text-slate-300 mr-1.5">#{l.feature}</span>
                      <span
                        className={`rounded px-1.5 py-0.5 text-[10px] font-bold ${
                          l.side === 'L'
                            ? 'bg-cyan-500/20 text-cyan-300'
                            : l.side === 'R'
                            ? 'bg-emerald-500/20 text-emerald-300'
                            : 'bg-slate-800 text-slate-400'
                        }`}
                      >
                        {l.side === 'L' ? 'Trái (LHF)' : l.side === 'R' ? 'Phải (RHF)' : 'Chung'}
                      </span>
                    </td>
                    <td className="py-2.5 px-3 font-sans text-slate-200 max-w-[220px] truncate">
                      {l.description}
                    </td>
                    <td className="py-2.5 px-3 text-slate-300">
                      {l.updatedLoadingDate || l.loadingDate || l.crd}
                    </td>
                    <td className="py-2.5 px-3 text-right font-bold text-white">
                      {l.qty.toLocaleString('en-US')}
                    </td>
                    <td className="py-2.5 px-3 text-right text-blue-300">
                      {l.wh01Initial.toLocaleString('en-US')}
                    </td>
                    <td className="py-2.5 px-3 text-right text-cyan-300">
                      {l.wh60Initial.toLocaleString('en-US')}
                    </td>
                    <td className="py-2.5 px-3 text-right text-slate-300">
                      {l.stockBeforeLine.toLocaleString('en-US')}
                    </td>
                    <td className="py-2.5 px-3 text-right text-emerald-300">
                      {l.allocatedTotal.toLocaleString('en-US')}
                    </td>
                    <td className="py-2.5 px-3 text-right font-extrabold text-rose-400">
                      -{l.shortageQty.toLocaleString('en-US')}
                    </td>
                    <td className="py-2.5 px-3 text-center">
                      {l.fulfillmentStatus === 'ZERO_STOCK' ? (
                        <span className="inline-flex items-center gap-1 rounded-full bg-rose-500/20 border border-rose-500/40 px-2.5 py-0.5 text-[10px] font-bold text-rose-300">
                          HẾT TỒN KHO (0)
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1 rounded-full bg-amber-500/20 border border-amber-500/40 px-2.5 py-0.5 text-[10px] font-bold text-amber-300">
                          THIẾU 1 PHẦN ({l.fulfillmentRate.toFixed(0)}%)
                        </span>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* TAB 5: AGGREGATED BY ITEM CODE */}
      {activeTab === 'items' && (
        <div className="p-5">
          <div className="max-h-[560px] overflow-auto rounded-xl border border-slate-800 bg-slate-950/75">
            <table className="w-full text-left border-collapse text-xs">
              <thead className="sticky top-0 bg-slate-900 text-slate-300 border-b border-slate-800 z-10">
                <tr>
                  <th className="py-3 px-3">Mã Hàng (LPVN Code)</th>
                  <th className="py-3 px-3">Model / Bên</th>
                  <th className="py-3 px-3">Mô Tả</th>
                  <th className="py-3 px-3 text-center">Số PO Đặt</th>
                  <th className="py-3 px-3 text-right text-blue-300">Tồn Kho 01</th>
                  <th className="py-3 px-3 text-right text-cyan-300">Tồn Kho 60</th>
                  <th className="py-3 px-3 text-right text-white">Tổng Tồn (01+60)</th>
                  <th className="py-3 px-3 text-right text-amber-300">Tổng Cầu Xuất</th>
                  <th className="py-3 px-3 text-right">Chênh Lệch (Tồn - Cầu)</th>
                  <th className="py-3 px-3 text-center">Kết Luận</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/60 font-mono">
                {reconciliation.itemDemands.map((d) => (
                  <tr key={d.itemCode} className="hover:bg-slate-800/50 transition">
                    <td className="py-2.5 px-3 font-bold text-cyan-300">{d.itemCode}</td>
                    <td className="py-2.5 px-3">
                      <span className="text-slate-300 mr-1.5">#{d.feature}</span>
                      <span
                        className={`rounded px-1.5 py-0.5 text-[10px] font-bold ${
                          d.side === 'L'
                            ? 'bg-cyan-500/20 text-cyan-300'
                            : d.side === 'R'
                            ? 'bg-emerald-500/20 text-emerald-300'
                            : 'bg-slate-800 text-slate-400'
                        }`}
                      >
                        {d.side === 'L' ? 'Trái' : d.side === 'R' ? 'Phải' : 'Chung'}
                      </span>
                    </td>
                    <td className="py-2.5 px-3 font-sans text-slate-200 max-w-[240px] truncate">
                      {d.description}
                    </td>
                    <td className="py-2.5 px-3 text-center text-slate-300">
                      {d.posList.length} PO ({d.orderLinesCount} dòng)
                    </td>
                    <td className="py-2.5 px-3 text-right text-blue-300">
                      {d.wh01Qty.toLocaleString('en-US')}
                    </td>
                    <td className="py-2.5 px-3 text-right text-cyan-300">
                      {d.wh60Qty.toLocaleString('en-US')}
                    </td>
                    <td className="py-2.5 px-3 text-right font-bold text-white">
                      {d.totalStock.toLocaleString('en-US')}
                    </td>
                    <td className="py-2.5 px-3 text-right font-bold text-amber-300">
                      {d.totalRequiredQty.toLocaleString('en-US')}
                    </td>
                    <td className="py-2.5 px-3 text-right font-extrabold">
                      {d.shortageQty > 0 ? (
                        <span className="text-rose-400">
                          -{d.shortageQty.toLocaleString('en-US')}
                        </span>
                      ) : (
                        <span className="text-emerald-400">
                          +{d.surplusQty.toLocaleString('en-US')}
                        </span>
                      )}
                    </td>
                    <td className="py-2.5 px-3 text-center">
                      {d.status === 'ENOUGH' ? (
                        <span className="rounded-full bg-emerald-500/20 border border-emerald-500/40 px-2.5 py-0.5 text-[10px] font-bold text-emerald-300">
                          ĐỦ TỔNG CẦU
                        </span>
                      ) : (
                        <span className="rounded-full bg-rose-500/20 border border-rose-500/40 px-2.5 py-0.5 text-[10px] font-bold text-rose-300">
                          THIẾU HÀNG
                        </span>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </section>
  );
};
