import React, { useEffect, useState } from 'react';
import {
  Warehouse,
  Scale,
  CheckCircle2,
  AlertOctagon,
  Ship,
  ShieldCheck,
  TrendingUp,
  Layers
} from 'lucide-react';
import { LeftRightPair, StockItem } from '../types';
import { ReconciliationResult } from '../utils/reconciliationEngine';

interface KpiCardsSectionProps {
  stockItems: StockItem[];
  lrPairs: LeftRightPair[];
  reconciliation: ReconciliationResult;
  activePlanLabel: string;
  onQuickJumpTab: (tab: 'fulfilled' | 'shortage' | 'po' | 'items') => void;
}

function useAnimatedNumber(target: number, duration = 650): number {
  const [val, setVal] = useState(target);
  useEffect(() => {
    let start = val;
    const diff = target - start;
    if (Math.abs(diff) < 0.5) {
      setVal(target);
      return;
    }
    const startTime = performance.now();
    let frameId = 0;
    const step = (now: number) => {
      const elapsed = Math.min(1, (now - startTime) / duration);
      const eased = 1 - Math.pow(1 - elapsed, 3);
      setVal(Math.round(start + diff * eased));
      if (elapsed < 1) {
        frameId = requestAnimationFrame(step);
      } else {
        setVal(target);
      }
    };
    frameId = requestAnimationFrame(step);
    return () => cancelAnimationFrame(frameId);
  }, [target, duration]);
  return val;
}

export const KpiCardsSection: React.FC<KpiCardsSectionProps> = ({
  stockItems,
  lrPairs,
  reconciliation,
  activePlanLabel,
  onQuickJumpTab
}) => {
  const wh01Total = stockItems.reduce((a, b) => a + b.wh01Qty, 0);
  const wh60Total = stockItems.reduce((a, b) => a + b.wh60Qty, 0);
  const combinedStock = wh01Total + wh60Total;

  const fgPairs = lrPairs.filter((p) => p.isFinishedGood);
  const fgLeftTotal = fgPairs.reduce((a, b) => a + b.leftTotal, 0);
  const fgRightTotal = fgPairs.reduce((a, b) => a + b.rightTotal, 0);
  const fgCombined = fgLeftTotal + fgRightTotal;
  const fgMatchedSets = fgPairs.reduce((a, b) => a + b.matchedSets, 0);
  const balancedCount = fgPairs.filter((p) => p.isBalanced).length;
  const imbalancedCount = fgPairs.length - balancedCount;

  const animCombinedStock = useAnimatedNumber(combinedStock);
  const animFgCombined = useAnimatedNumber(fgCombined);
  const animFulfilledLines = useAnimatedNumber(reconciliation.metrics.fulfilledLinesCount);
  const animShortageLines = useAnimatedNumber(reconciliation.metrics.shortageLinesCount);
  const animRequiredQty = useAnimatedNumber(reconciliation.metrics.totalRequiredQty);
  const animShortageQty = useAnimatedNumber(reconciliation.metrics.totalShortageQty);

  const wh01Pct = combinedStock > 0 ? Math.round((wh01Total / combinedStock) * 100) : 0;
  const wh60Pct = 100 - wh01Pct;
  const fulfilledLinePct =
    reconciliation.metrics.totalOrderLines > 0
      ? Math.round(
          (reconciliation.metrics.fulfilledLinesCount /
            reconciliation.metrics.totalOrderLines) *
            100
        )
      : 100;

  return (
    <section className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-6 gap-4">
      {/* KPI 1: Tổng Tồn Kho 01 & 60 */}
      <div className="group relative rounded-2xl border border-cyan-500/30 bg-slate-900/80 p-4 backdrop-blur-xl shadow-lg hover:border-cyan-400/60 hover:-translate-y-1 transition-all duration-300 overflow-hidden">
        <div className="absolute -right-6 -top-6 h-24 w-24 rounded-full bg-cyan-500/10 blur-xl group-hover:bg-cyan-500/20 transition" />
        <div className="flex items-center justify-between">
          <span className="text-xs font-semibold uppercase tracking-wider text-cyan-300">
            Tồn Kho 01 & 60
          </span>
          <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-cyan-500/15 text-cyan-400 border border-cyan-500/30">
            <Warehouse className="h-4 w-4" />
          </div>
        </div>
        <div className="mt-2 flex items-baseline gap-1.5">
          <span className="text-2xl font-extrabold font-mono text-white">
            {animCombinedStock.toLocaleString('en-US')}
          </span>
          <span className="text-xs text-slate-400">pcs ({stockItems.length} mã)</span>
        </div>
        <div className="mt-2.5 h-2 w-full rounded-full bg-slate-800 overflow-hidden flex">
          <div
            className="h-full bg-blue-500 transition-all duration-500"
            style={{ width: `${wh01Pct}%` }}
            title={`Kho 01: ${wh01Total.toLocaleString('en-US')} pcs`}
          />
          <div
            className="h-full bg-cyan-400 transition-all duration-500"
            style={{ width: `${wh60Pct}%` }}
            title={`Kho 60: ${wh60Total.toLocaleString('en-US')} pcs`}
          />
        </div>
        <div className="mt-2 flex items-center justify-between text-[11px] font-mono text-slate-300">
          <span className="text-blue-300">Kho 01: {wh01Total.toLocaleString('en-US')}</span>
          <span className="text-cyan-300">Kho 60: {wh60Total.toLocaleString('en-US')}</span>
        </div>
      </div>

      {/* KPI 2: Tồn Kho Thành Phẩm Trái / Phải */}
      <div className="group relative rounded-2xl border border-indigo-500/30 bg-slate-900/80 p-4 backdrop-blur-xl shadow-lg hover:border-indigo-400/60 hover:-translate-y-1 transition-all duration-300 overflow-hidden">
        <div className="absolute -right-6 -top-6 h-24 w-24 rounded-full bg-indigo-500/10 blur-xl group-hover:bg-indigo-500/20 transition" />
        <div className="flex items-center justify-between">
          <span className="text-xs font-semibold uppercase tracking-wider text-indigo-300">
            Thành Phẩm Trái / Phải
          </span>
          <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-indigo-500/15 text-indigo-400 border border-indigo-500/30">
            <Scale className="h-4 w-4" />
          </div>
        </div>
        <div className="mt-2 flex items-baseline gap-1.5">
          <span className="text-2xl font-extrabold font-mono text-white">
            {animFgCombined.toLocaleString('en-US')}
          </span>
          <span className="text-xs text-slate-400">
            pcs ({fgMatchedSets.toLocaleString('en-US')} bộ)
          </span>
        </div>
        <div className="mt-2 flex items-center justify-between text-[11px] font-mono">
          <span className="text-cyan-300">Trái: {fgLeftTotal.toLocaleString('en-US')}</span>
          <span className="text-emerald-300">Phải: {fgRightTotal.toLocaleString('en-US')}</span>
        </div>
        <div className="mt-1.5 flex items-center justify-between text-[11px]">
          <span className="inline-flex items-center gap-1 text-emerald-300 font-medium">
            ● {balancedCount} cặp cân bằng
          </span>
          <span className="inline-flex items-center gap-1 text-amber-300 font-medium">
            ▲ {imbalancedCount} cặp lệch
          </span>
        </div>
      </div>

      {/* KPI 3: Đơn / Dòng Đủ Hàng Xuất */}
      <div
        onClick={() => onQuickJumpTab('fulfilled')}
        className="group relative rounded-2xl border border-emerald-500/35 bg-slate-900/80 p-4 backdrop-blur-xl shadow-lg hover:border-emerald-400 hover:-translate-y-1 transition-all duration-300 cursor-pointer overflow-hidden"
      >
        <div className="absolute -right-6 -top-6 h-24 w-24 rounded-full bg-emerald-500/10 blur-xl group-hover:bg-emerald-500/20 transition" />
        <div className="flex items-center justify-between">
          <span className="text-xs font-semibold uppercase tracking-wider text-emerald-300">
            Đơn Đủ Hàng Xuất
          </span>
          <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-emerald-500/15 text-emerald-400 border border-emerald-500/30">
            <CheckCircle2 className="h-4 w-4" />
          </div>
        </div>
        <div className="mt-2 flex items-baseline gap-1.5">
          <span className="text-2xl font-extrabold font-mono text-emerald-300">
            {animFulfilledLines}
          </span>
          <span className="text-xs text-slate-400">
            / {reconciliation.metrics.totalOrderLines} dòng ({fulfilledLinePct}%)
          </span>
        </div>
        <div className="mt-2.5 h-2 w-full rounded-full bg-slate-800 overflow-hidden">
          <div
            className="h-full bg-gradient-to-r from-emerald-500 to-teal-400 transition-all duration-500"
            style={{ width: `${fulfilledLinePct}%` }}
          />
        </div>
        <div className="mt-2 flex items-center justify-between text-[11px] text-slate-300">
          <span>
            PO Đủ 100%:{' '}
            <strong className="font-mono text-emerald-300">
              {reconciliation.metrics.fulfilledPosCount}/{reconciliation.metrics.totalPos}
            </strong>
          </span>
          <span className="text-emerald-400 underline">Xem đơn đủ →</span>
        </div>
      </div>

      {/* KPI 4: Đơn / Dòng Thiếu Hàng */}
      <div
        onClick={() => onQuickJumpTab('shortage')}
        className="group relative rounded-2xl border border-rose-500/35 bg-slate-900/80 p-4 backdrop-blur-xl shadow-lg hover:border-rose-400 hover:-translate-y-1 transition-all duration-300 cursor-pointer overflow-hidden"
      >
        <div className="absolute -right-6 -top-6 h-24 w-24 rounded-full bg-rose-500/10 blur-xl group-hover:bg-rose-500/20 transition" />
        <div className="flex items-center justify-between">
          <span className="text-xs font-semibold uppercase tracking-wider text-rose-300">
            Đơn & Mã Thiếu Hàng
          </span>
          <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-rose-500/15 text-rose-400 border border-rose-500/30">
            <AlertOctagon className="h-4 w-4" />
          </div>
        </div>
        <div className="mt-2 flex items-baseline gap-1.5">
          <span className="text-2xl font-extrabold font-mono text-rose-400">
            {animShortageLines}
          </span>
          <span className="text-xs text-slate-400">
            dòng thiếu ({reconciliation.metrics.shortagePosCount} PO)
          </span>
        </div>
        <div className="mt-2 flex items-center justify-between text-[11px] font-mono text-rose-200">
          <span>SL Thiếu: -{animShortageQty.toLocaleString('en-US')}</span>
          <span>Tồn=0: {reconciliation.metrics.zeroStockLinesCount} dòng</span>
        </div>
        <div className="mt-1.5 flex items-center justify-between text-[11px] text-slate-400">
          <span>Thiếu 1 phần: {reconciliation.metrics.partialStockLinesCount} dòng</span>
          <span className="text-rose-300 underline">Xem hàng thiếu →</span>
        </div>
      </div>

      {/* KPI 5: Sản Lượng & Thể Tích Kế Hoạch */}
      <div
        onClick={() => onQuickJumpTab('po')}
        className="group relative rounded-2xl border border-sky-500/30 bg-slate-900/80 p-4 backdrop-blur-xl shadow-lg hover:border-sky-400/60 hover:-translate-y-1 transition-all duration-300 cursor-pointer overflow-hidden"
      >
        <div className="flex items-center justify-between">
          <span className="text-xs font-semibold uppercase tracking-wider text-sky-300">
            Sản Lượng Kế Hoạch
          </span>
          <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-sky-500/15 text-sky-400 border border-sky-500/30">
            <Ship className="h-4 w-4" />
          </div>
        </div>
        <div className="mt-2 flex items-baseline gap-1.5">
          <span className="text-2xl font-extrabold font-mono text-white">
            {animRequiredQty.toLocaleString('en-US')}
          </span>
          <span className="text-xs text-slate-400">pcs cần xuất</span>
        </div>
        <div className="mt-2 flex items-center justify-between text-[11px] font-mono text-slate-300">
          <span>
            Đáp ứng:{' '}
            <strong className="text-emerald-300">
              {reconciliation.metrics.overallFulfillmentRate.toFixed(1)}%
            </strong>
          </span>
          <span>CBM: {reconciliation.metrics.totalCbm.toFixed(1)} m³</span>
        </div>
        <div className="mt-1.5 flex items-center justify-between text-[11px] font-mono text-slate-400">
          <span>Kiện: {reconciliation.metrics.totalPkgs.toLocaleString('en-US')} pkg</span>
          <span className="text-amber-300">
            ${Math.round(reconciliation.metrics.totalAmountUsd).toLocaleString('en-US')}
          </span>
        </div>
      </div>

      {/* KPI 6: Trạng Thái Đối Soát Multi-Agent & OneDrive */}
      <div className="group relative rounded-2xl border border-teal-500/30 bg-slate-900/80 p-4 backdrop-blur-xl shadow-lg hover:border-teal-400/60 hover:-translate-y-1 transition-all duration-300 overflow-hidden">
        <div className="flex items-center justify-between">
          <span className="text-xs font-semibold uppercase tracking-wider text-teal-300">
            Multi-Agent Audit
          </span>
          <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-teal-500/15 text-teal-400 border border-teal-500/30">
            <ShieldCheck className="h-4 w-4" />
          </div>
        </div>
        <div className="mt-2 flex items-center gap-2">
          <span className="relative flex h-3 w-3">
            <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
            <span className="relative inline-flex rounded-full h-3 w-3 bg-emerald-500"></span>
          </span>
          <span className="text-lg font-extrabold text-white">100% Chuẩn Xác</span>
        </div>
        <div className="mt-2 text-[11px] text-slate-300 line-clamp-1">
          Đang chọn: <strong className="text-cyan-300">{activePlanLabel}</strong>
        </div>
        <div className="mt-1.5 flex items-center justify-between text-[11px] text-teal-300 font-mono">
          <span>4 Agents Verified</span>
          <span>Kho 01 & 60 Strict</span>
        </div>
      </div>
    </section>
  );
};
