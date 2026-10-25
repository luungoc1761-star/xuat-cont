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
  ArrowUpDown,
  Columns
} from 'lucide-react';
import { AllocationMode, LeftRightPair, ReconciledOrderLine, StockItem } from '../types';
import { ReconciliationResult } from '../utils/reconciliationEngine';
import { exportFullReconciliationToExcel } from '../utils/excelParserAndExporter';

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
  activeTab: 'split' | 'fulfilled' | 'shortage' | 'po' | 'items';
  setActiveTab: (t: 'split' | 'fulfilled' | 'shortage' | 'po' | 'items') => void;
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
  setActiveTab
}) => {
  const [customerFilter, setCustomerFilter] = useState<string>('ALL');
  const [poFilter, setPoFilter] = useState<string>('ALL');

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

  // Filter predicate for order lines
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
  }, [reconciliation.poSummaries, customerFilter, poFilter, selectedModel, searchQuery]);

  const handleExportExcel = () => {
    exportFullReconciliationToExcel(
      reconciliation,
      lrPairs,
      stockItems,
      activePlanName,
      allocationMode
    );
  };

  return (
    <section
      id="fulfillment-report-section"
      className="rounded-2xl border border-slate-800/90 bg-slate-900/85 backdrop-blur-xl shadow-2xl overflow-hidden"
    >
      {/* Top Report Header & Excel Export Action */}
      <div className="flex flex-wrap items-center justify-between gap-4 border-b border-slate-800 bg-slate-950/85 px-6 py-4">
        <div>
          <div className="flex flex-wrap items-center gap-2.5">
            <h2 className="text-lg font-extrabold text-white">
              Báo Cáo Đối Soát Xuất Hàng Với Kho 01 & Kho 60 — Đơn Đủ Hàng Xuất & Đơn Hàng Thiếu
            </h2>
            <span className="rounded-full bg-cyan-500/15 border border-cyan-500/30 px-2.5 py-0.5 text-xs font-semibold text-cyan-300">
              {activePlanName}
            </span>
          </div>
          <p className="mt-0.5 text-xs text-slate-400">
            Chỉ lấy tồn kho tại <strong className="text-cyan-300">Kho 01</strong> và{' '}
            <strong className="text-emerald-300">Kho 60</strong> để đối chiếu chính xác từng đơn hàng, PO, Container và từng mã hàng Trái/Phải
          </p>
        </div>

        {/* Allocation Mode Toggle + Export Excel Button */}
        <div className="flex flex-wrap items-center gap-2.5">
          <div className="inline-flex rounded-xl border border-slate-700/80 bg-slate-900 p-1">
            <button
              onClick={() => setAllocationMode('fifo')}
              className={`rounded-lg px-3 py-1.5 text-xs font-semibold transition cursor-pointer ${
                allocationMode === 'fifo'
                  ? 'bg-cyan-500 text-slate-950 shadow'
                  : 'text-slate-300 hover:text-white'
              }`}
              title="Trừ tồn kho lũy kế theo thứ tự đơn hàng/lịch đóng cont"
            >
              Phân Bổ Lũy Kế FIFO
            </button>
            <button
              onClick={() => setAllocationMode('direct')}
              className={`rounded-lg px-3 py-1.5 text-xs font-semibold transition cursor-pointer ${
                allocationMode === 'direct'
                  ? 'bg-cyan-500 text-slate-950 shadow'
                  : 'text-slate-300 hover:text-white'
              }`}
              title="So sánh trực tiếp số lượng từng dòng với tổng tồn Kho 01 + 60"
            >
              Đối Chiếu Trực Tiếp Từng Dòng
            </button>
          </div>

          <button
            onClick={handleExportExcel}
            className="inline-flex items-center gap-2 rounded-xl bg-gradient-to-r from-emerald-500 to-teal-400 px-4 py-2.5 text-xs font-extrabold text-slate-950 shadow-lg shadow-emerald-500/25 hover:brightness-110 transition cursor-pointer"
          >
            <Download className="h-4 w-4" />
            <span>Xuất File Excel (.xlsx)</span>
          </button>
        </div>
      </div>

      {/* Sub-navigation Tabs & Customer/PO Filters */}
      <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-800/80 bg-slate-900/95 px-6 py-3">
        <div className="flex flex-wrap items-center gap-1.5">
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
            <span>Đơn Đủ Hàng Xuất ({filteredFulfilledLines.length})</span>
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
            <span>Đơn & Mã Hàng Thiếu ({filteredShortageLines.length})</span>
          </button>

          <button
            onClick={() => setActiveTab('po')}
            className={`inline-flex items-center gap-1.5 rounded-xl px-3.5 py-2 text-xs font-bold transition cursor-pointer ${
              activeTab === 'po'
                ? 'bg-indigo-500 text-white shadow-md shadow-indigo-500/20'
                : 'bg-slate-800/90 text-indigo-300 hover:bg-slate-800'
            }`}
          >
            <Ship className="h-3.5 w-3.5" />
            <span>Theo PO / Container ({filteredPoSummaries.length} PO)</span>
          </button>

          <button
            onClick={() => setActiveTab('items')}
            className={`inline-flex items-center gap-1.5 rounded-xl px-3.5 py-2 text-xs font-bold transition cursor-pointer ${
              activeTab === 'items'
                ? 'bg-amber-500 text-slate-950 shadow-md shadow-amber-500/20'
                : 'bg-slate-800/90 text-amber-300 hover:bg-slate-800'
            }`}
          >
            <Layers className="h-3.5 w-3.5" />
            <span>Tổng Hợp Theo Mã Hàng ({reconciliation.itemDemands.length} mã)</span>
          </button>
        </div>

        {/* Customer & PO Select Filters */}
        <div className="flex flex-wrap items-center gap-2">
          <div className="flex items-center gap-1.5 text-xs text-slate-400">
            <Filter className="h-3.5 w-3.5 text-cyan-400" />
            <span>Khách hàng:</span>
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

      {/* TAB 1: SPLIT VIEW (Both Fulfilled Orders & Shortage Orders visible together) */}
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
                {filteredFulfilledLines.length} dòng đơn ({reconciliation.fulfilledPos.length} PO đủ 100%)
              </span>
            </div>
            <div className="max-h-[500px] overflow-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead className="sticky top-0 bg-slate-900/95 text-slate-300 border-b border-slate-800 z-10">
                  <tr>
                    <th className="py-2.5 px-3">Khách / PO</th>
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
                        <div className="font-sans font-semibold text-white">{l.customer}</div>
                        <div className="text-[11px] text-slate-400">
                          PO: <span className="text-emerald-300">{l.po}</span> • {l.updatedLoadingDate || l.loadingDate}
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
                {filteredShortageLines.length} dòng thiếu ({reconciliation.shortagePos.length} PO thiếu)
              </span>
            </div>
            <div className="max-h-[500px] overflow-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead className="sticky top-0 bg-slate-900/95 text-slate-300 border-b border-slate-800 z-10">
                  <tr>
                    <th className="py-2.5 px-3">Khách / PO</th>
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
                        <div className="font-sans font-semibold text-white">{l.customer}</div>
                        <div className="text-[11px] text-slate-400">
                          PO: <span className="text-rose-300">{l.po}</span> • {l.updatedLoadingDate || l.loadingDate}
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

      {/* TAB 2: FULL TABLE OF FULFILLED ORDERS */}
      {activeTab === 'fulfilled' && (
        <div className="p-5">
          <div className="max-h-[560px] overflow-auto rounded-xl border border-emerald-500/30 bg-slate-950/75">
            <table className="w-full text-left border-collapse text-xs">
              <thead className="sticky top-0 bg-slate-900 text-slate-300 border-b border-slate-800 z-10">
                <tr>
                  <th className="py-3 px-3">Dòng</th>
                  <th className="py-3 px-3">Khách Hàng</th>
                  <th className="py-3 px-3">SO / PO</th>
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
                    <td className="py-2.5 px-3 text-slate-400">#{l.rowNumber}</td>
                    <td className="py-2.5 px-3 font-sans font-semibold text-white">{l.customer}</td>
                    <td className="py-2.5 px-3">
                      <div className="text-emerald-300 font-bold">PO: {l.po}</div>
                      {l.so && <div className="text-[10px] text-slate-400">SO: {l.so}</div>}
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

      {/* TAB 3: FULL TABLE OF SHORTAGE ORDERS */}
      {activeTab === 'shortage' && (
        <div className="p-5">
          <div className="max-h-[560px] overflow-auto rounded-xl border border-rose-500/30 bg-slate-950/75">
            <table className="w-full text-left border-collapse text-xs">
              <thead className="sticky top-0 bg-slate-900 text-slate-300 border-b border-slate-800 z-10">
                <tr>
                  <th className="py-3 px-3">Dòng</th>
                  <th className="py-3 px-3">Khách Hàng</th>
                  <th className="py-3 px-3">SO / PO</th>
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
                    <td className="py-2.5 px-3 text-slate-400">#{l.rowNumber}</td>
                    <td className="py-2.5 px-3 font-sans font-semibold text-white">{l.customer}</td>
                    <td className="py-2.5 px-3">
                      <div className="text-rose-300 font-bold">PO: {l.po}</div>
                      {l.so && <div className="text-[10px] text-slate-400">SO: {l.so}</div>}
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

      {/* TAB 4: GROUPED BY PO / CONTAINER */}
      {activeTab === 'po' && (
        <div className="p-5 grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4 max-h-[600px] overflow-y-auto">
          {filteredPoSummaries.map((po) => {
            const isFull = po.status === 'FULFILLED';
            return (
              <div
                key={po.po}
                className={`rounded-xl border p-4 transition ${
                  isFull
                    ? 'border-emerald-500/35 bg-emerald-950/15 hover:border-emerald-400'
                    : 'border-rose-500/35 bg-rose-950/15 hover:border-rose-400'
                }`}
              >
                <div className="flex items-start justify-between gap-2">
                  <div>
                    <div className="flex items-center gap-2">
                      <span className="font-mono text-sm font-extrabold text-white">
                        PO: {po.po}
                      </span>
                      {isFull ? (
                        <span className="rounded-full bg-emerald-500/20 border border-emerald-500/40 px-2 py-0.5 text-[10px] font-bold text-emerald-300">
                          ĐỦ HÀNG XUẤT CONT 100%
                        </span>
                      ) : (
                        <span className="rounded-full bg-rose-500/20 border border-rose-500/40 px-2 py-0.5 text-[10px] font-bold text-rose-300">
                          THIẾU {po.shortageLinesCount}/{po.totalLines} MÃ
                        </span>
                      )}
                    </div>
                    <div className="mt-1 text-xs font-semibold text-slate-200">
                      {po.customer}{' '}
                      {po.soList.length > 0 && (
                        <span className="font-mono text-slate-400">• SO: {po.soList.join(', ')}</span>
                      )}
                    </div>
                  </div>
                  <div className="text-right font-mono text-xs text-slate-300">
                    <div>{po.loadingDate || 'Chưa có lịch'}</div>
                    <div className="text-[11px] text-cyan-300">{po.totalCbm.toFixed(1)} CBM</div>
                  </div>
                </div>

                {/* Progress Bar */}
                <div className="mt-3 h-2 w-full rounded-full bg-slate-800 overflow-hidden">
                  <div
                    className={`h-full ${
                      isFull
                        ? 'bg-gradient-to-r from-emerald-500 to-teal-400'
                        : 'bg-gradient-to-r from-amber-500 to-rose-500'
                    }`}
                    style={{ width: `${Math.min(100, po.fulfillmentRate)}%` }}
                  />
                </div>

                <div className="mt-2.5 grid grid-cols-3 gap-2 text-center font-mono text-xs bg-slate-950/60 rounded-lg p-2">
                  <div>
                    <div className="text-[10px] text-slate-400">Tổng Cần</div>
                    <div className="font-bold text-white">
                      {po.totalRequiredQty.toLocaleString('en-US')}
                    </div>
                  </div>
                  <div>
                    <div className="text-[10px] text-slate-400">Đáp Ứng</div>
                    <div className="font-bold text-emerald-300">
                      {po.totalAllocatedQty.toLocaleString('en-US')}
                    </div>
                  </div>
                  <div>
                    <div className="text-[10px] text-slate-400">Thiếu Hụt</div>
                    <div
                      className={`font-bold ${
                        po.totalShortageQty > 0 ? 'text-rose-400' : 'text-slate-500'
                      }`}
                    >
                      {po.totalShortageQty > 0
                        ? `-${po.totalShortageQty.toLocaleString('en-US')}`
                        : '0'}
                    </div>
                  </div>
                </div>

                {/* Mini list of lines inside PO */}
                <div className="mt-2.5 max-h-32 overflow-y-auto space-y-1 text-[11px] font-mono border-t border-slate-800/80 pt-2">
                  {po.lines.map((line) => (
                    <div key={line.id} className="flex items-center justify-between">
                      <span className="text-slate-300 truncate max-w-[170px]">
                        {line.itemCode} ({line.side === 'L' ? 'Trái' : line.side === 'R' ? 'Phải' : 'C'})
                      </span>
                      {line.shortageQty > 0 ? (
                        <span className="text-rose-400 font-bold">
                          Cần {line.qty} / Thiếu -{line.shortageQty}
                        </span>
                      ) : (
                        <span className="text-emerald-400">Đủ ({line.qty})</span>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            );
          })}
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
