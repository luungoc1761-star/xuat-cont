import React, { useState, useMemo } from 'react';
import {
  Search,
  Zap,
  X,
  CheckCircle2,
  AlertTriangle,
  Sparkles,
  Copy,
  Check,
  Filter,
  ArrowUpRight
} from 'lucide-react';
import { StockItem, LeftRightPair, ReconciledItemDemand } from '../types';

interface QuickFillStockSearchProps {
  stockItems: StockItem[];
  lrPairs: LeftRightPair[];
  itemDemands: ReconciledItemDemand[];
  searchQuery: string;
  setSearchQuery: (q: string) => void;
  selectedCodes: string[];
  setSelectedCodes: (codes: string[]) => void;
  onSelectModelIn3D: (model: string | null) => void;
}

export const QuickFillStockSearch: React.FC<QuickFillStockSearchProps> = ({
  stockItems,
  lrPairs,
  itemDemands,
  searchQuery,
  setSearchQuery,
  selectedCodes,
  setSelectedCodes,
  onSelectModelIn3D
}) => {
  const [showBatchModal, setShowBatchModal] = useState(false);
  const [batchText, setBatchText] = useState('');
  const [copiedCode, setCopiedCode] = useState<string | null>(null);
  const [whFilter, setWhFilter] = useState<'ALL' | '01' | '60' | 'BOTH' | 'FG'>('ALL');

  const shortageCodeSet = useMemo(() => {
    return new Set(itemDemands.filter((d) => d.shortageQty > 0).map((d) => d.itemCode));
  }, [itemDemands]);

  const demandMap = useMemo(() => {
    const m = new Map<string, ReconciledItemDemand>();
    for (const d of itemDemands) m.set(d.itemCode, d);
    return m;
  }, [itemDemands]);

  // Quick-Fill Preset Actions
  const applyPreset = (presetId: string) => {
    if (presetId === 'imbalanced_5') {
      const codes = lrPairs
        .filter((p) => p.isFinishedGood && !p.isBalanced)
        .flatMap((p) => [p.leftCode, p.rightCode]);
      setSelectedCodes(codes);
      setSearchQuery('3565, 3611, 5025, 5171, 8695');
    } else if (presetId === 'cross_8333') {
      setSelectedCodes(['8833310004', '8833320004']);
      setSearchQuery('8333');
      onSelectModelIn3D('8333');
    } else if (presetId === 'top_stock') {
      setSelectedCodes([
        '8863410204',
        '8863420204',
        '8835310004',
        '8835320004',
        '8832310004',
        '8832320004',
        '8869510104',
        '8869520104',
        '8833310004',
        '8833320004',
        '8356910204',
        '8356920204'
      ]);
      setSearchQuery('8634, 8353, 8323, 8695, 8333, 3569');
    } else if (presetId === 'shortage_in_stock') {
      const codes = stockItems
        .filter((s) => shortageCodeSet.has(s.code))
        .map((s) => s.code);
      setSelectedCodes(codes);
      setSearchQuery(codes.join(', '));
    } else if (presetId === 'both_wh') {
      const codes = stockItems
        .filter((s) => s.wh01Qty > 0 && s.wh60Qty > 0)
        .map((s) => s.code);
      setSelectedCodes(codes);
      setSearchQuery('4151, 4154, 5152, 8333, 8334, 8352, 8362, 8393, 8695, 8817');
    } else if (presetId === 'semi_lr') {
      setSelectedCodes([
        '1220190004',
        '1220200004',
        '1421710201',
        '1421720201',
        '1427790001',
        '1427800001',
        '1455350001',
        '1455360001'
      ]);
      setSearchQuery('1220, 1421, 1427, 1455');
    } else if (presetId === 'clear') {
      setSelectedCodes([]);
      setSearchQuery('');
      onSelectModelIn3D(null);
    }
  };

  // Filtered Stock Items
  const filteredStock = useMemo(() => {
    const tokens = searchQuery
      .split(/[\s,;]+/)
      .map((t) => t.trim().toLowerCase())
      .filter(Boolean);

    return stockItems.filter((item) => {
      if (whFilter === '01' && item.wh01Qty <= 0) return false;
      if (whFilter === '60' && item.wh60Qty <= 0) return false;
      if (whFilter === 'BOTH' && (item.wh01Qty <= 0 || item.wh60Qty <= 0)) return false;
      if (whFilter === 'FG' && !item.isFinishedGood) return false;

      if (selectedCodes.length > 0 && selectedCodes.includes(item.code)) {
        return true;
      }

      if (tokens.length === 0) return selectedCodes.length === 0;

      const hay = `${item.code} ${item.feature} ${item.description} ${item.category}`.toLowerCase();
      return tokens.some((tk) => hay.includes(tk));
    });
  }, [stockItems, searchQuery, selectedCodes, whFilter]);

  const handleBatchApply = () => {
    const tokens = batchText
      .split(/[\s,;\n]+/)
      .map((x) => x.trim())
      .filter(Boolean);
    setSearchQuery(tokens.join(', '));
    setSelectedCodes([]);
    setShowBatchModal(false);
  };

  const copyCode = (code: string) => {
    navigator.clipboard?.writeText(code);
    setCopiedCode(code);
    setTimeout(() => setCopiedCode(null), 1200);
  };

  return (
    <section className="rounded-2xl border border-slate-800/90 bg-slate-900/80 backdrop-blur-xl p-5 shadow-xl">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div className="flex items-center gap-2.5">
          <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-amber-500/15 border border-amber-500/30 text-amber-400">
            <Zap className="h-5 w-5" />
          </div>
          <div>
            <h2 className="text-base font-bold text-white flex items-center gap-2">
              Điền Nhanh & Tra Cứu Tức Thì Mã Hàng Trong Stock Tồn Kho (Kho 01 & 60)
              <span className="rounded-full bg-amber-500/15 border border-amber-500/30 px-2 py-0.5 text-[11px] font-mono text-amber-300">
                {filteredStock.length} / {stockItems.length} mã
              </span>
            </h2>
            <p className="text-xs text-slate-400">
              Click nút điền nhanh bên dưới hoặc nhập nhiều mã hàng / Model cách nhau bởi dấu phẩy để lọc tức thì cả bảng tồn kho và báo cáo xuất hàng
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2">
          <button
            onClick={() => setShowBatchModal(!showBatchModal)}
            className="inline-flex items-center gap-1.5 rounded-xl bg-indigo-500/20 border border-indigo-500/40 px-3.5 py-2 text-xs font-semibold text-indigo-200 hover:bg-indigo-500/30 transition cursor-pointer"
          >
            <Sparkles className="h-3.5 w-3.5 text-indigo-400" />
            Dán Nhanh Danh Sách Mã (Batch Fill)
          </button>
          {(searchQuery || selectedCodes.length > 0) && (
            <button
              onClick={() => applyPreset('clear')}
              className="inline-flex items-center gap-1 rounded-xl bg-rose-500/20 border border-rose-500/40 px-3 py-2 text-xs font-semibold text-rose-200 hover:bg-rose-500/30 transition cursor-pointer"
            >
              <X className="h-3.5 w-3.5" />
              Xóa Bộ Lọc
            </button>
          )}
        </div>
      </div>

      {/* Search Input & Warehouse Quick Filter */}
      <div className="mt-4 flex flex-col md:flex-row gap-3">
        <div className="relative flex-1">
          <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 h-4 w-4 text-cyan-400" />
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => {
              setSearchQuery(e.target.value);
              setSelectedCodes([]);
            }}
            placeholder="Nhập hoặc bấm nút Điền Nhanh bên dưới: VD 8833310004, 8333, 8634, 3611, 1220190004, LHF..."
            className="w-full rounded-xl border border-slate-700/80 bg-slate-950/90 pl-10 pr-10 py-2.5 text-sm text-white placeholder-slate-500 focus:border-cyan-400 focus:outline-none focus:ring-2 focus:ring-cyan-500/20 font-mono"
          />
          {searchQuery && (
            <button
              onClick={() => {
                setSearchQuery('');
                setSelectedCodes([]);
              }}
              className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-white cursor-pointer"
            >
              <X className="h-4 w-4" />
            </button>
          )}
        </div>

        {/* Warehouse Scope Filter Buttons */}
        <div className="flex flex-wrap items-center gap-1.5">
          {[
            { id: 'ALL', label: 'Tất Cả (220)' },
            { id: 'FG', label: 'Thành Phẩm Đầu 8 (101)' },
            { id: '01', label: 'Có Ở Kho 01 (140)' },
            { id: '60', label: 'Có Ở Kho 60 (100)' },
            { id: 'BOTH', label: 'Cả 2 Kho 01 & 60 (20)' }
          ].map((b) => (
            <button
              key={b.id}
              onClick={() => setWhFilter(b.id as any)}
              className={`rounded-xl px-3 py-2 text-xs font-semibold transition cursor-pointer ${
                whFilter === b.id
                  ? 'bg-cyan-500 text-slate-950 shadow-md shadow-cyan-500/20'
                  : 'bg-slate-800/80 text-slate-300 hover:bg-slate-800 border border-slate-700/60'
              }`}
            >
              {b.label}
            </button>
          ))}
        </div>
      </div>

      {/* Quick-Fill Smart Preset Buttons */}
      <div className="mt-3 flex flex-wrap items-center gap-2">
        <span className="text-xs font-semibold text-amber-300 flex items-center gap-1">
          <Zap className="h-3.5 w-3.5" /> Nút Điền Nhanh:
        </span>
        {[
          { id: 'imbalanced_5', label: '⚡ 5 Cặp Lệch Trái/Phải (3565, 3611, 5025, 5171, 8695)' },
          { id: 'cross_8333', label: '🔄 Cặp #8333 Cân Bằng Gộp Kho 01+60' },
          { id: 'top_stock', label: '🏆 Top 6 Model Tồn Lớn Nhất (8634, 8353, 8323...)' },
          { id: 'shortage_in_stock', label: '🚨 Các Mã Có Tồn Nhưng Đang Thiếu So Với Đơn' },
          { id: 'both_wh', label: '📦 10 Cặp Nằm Cả Ở Kho 01 & Kho 60' },
          { id: 'semi_lr', label: '⚙️ Bán Thành Phẩm Trái/Phải (1220, 1421, 1427, 1455)' }
        ].map((chip) => (
          <button
            key={chip.id}
            onClick={() => applyPreset(chip.id)}
            className="rounded-lg border border-slate-700/80 bg-slate-800/70 px-2.5 py-1 text-xs font-medium text-slate-200 hover:border-cyan-400/60 hover:bg-cyan-500/15 hover:text-cyan-200 transition cursor-pointer"
          >
            {chip.label}
          </button>
        ))}

        {/* Individual Quick Model Tag Pills */}
        {['8634', '8353', '8323', '8695', '8333', '3569', '5171', '5633', '8817', '4151'].map(
          (modelCode) => (
            <button
              key={modelCode}
              onClick={() => {
                setSearchQuery(modelCode);
                setSelectedCodes([]);
                onSelectModelIn3D(modelCode);
              }}
              className={`rounded-md px-2 py-0.5 text-xs font-mono font-semibold transition cursor-pointer ${
                searchQuery === modelCode
                  ? 'bg-cyan-400 text-slate-950'
                  : 'bg-slate-950 border border-slate-800 text-cyan-300 hover:border-cyan-500/50'
              }`}
            >
              #{modelCode}
            </button>
          )
        )}
      </div>

      {/* Multi-Code Batch Paste Drawer */}
      {showBatchModal && (
        <div className="mt-3 rounded-xl border border-indigo-500/40 bg-slate-950/95 p-4">
          <div className="flex items-center justify-between mb-2">
            <span className="text-xs font-bold text-indigo-300">
              Dán Danh Sách Mã Hàng (Stock Code) Hoặc Mã Model Từ Excel / OneDrive:
            </span>
            <button
              onClick={() => setShowBatchModal(false)}
              className="text-xs text-slate-400 hover:text-white cursor-pointer"
            >
              Đóng
            </button>
          </div>
          <textarea
            rows={3}
            value={batchText}
            onChange={(e) => setBatchText(e.target.value)}
            placeholder="Dán 1 hoặc nhiều mã hàng (mỗi dòng 1 mã hoặc cách nhau bằng dấu phẩy / khoảng trắng)... VD: 8833310004 8833320004 8580210004"
            className="w-full rounded-lg border border-slate-800 bg-slate-900 p-2.5 text-xs font-mono text-white focus:border-indigo-400 focus:outline-none"
          />
          <div className="mt-2 flex justify-end gap-2">
            <button
              onClick={() =>
                setBatchText('8833310004\n8833320004\n8580210004\n8580220004\n8835310004\n8835320004')
              }
              className="rounded-lg bg-slate-800 px-3 py-1.5 text-xs text-slate-300 hover:text-white cursor-pointer"
            >
              Điền Mẫu Thử
            </button>
            <button
              onClick={handleBatchApply}
              className="rounded-lg bg-indigo-500 px-4 py-1.5 text-xs font-bold text-white hover:bg-indigo-400 cursor-pointer"
            >
              Áp Dụng Tìm Kiếm Hàng Loạt
            </button>
          </div>
        </div>
      )}

      {/* Compact Stock Lookup Table when searching or filtering */}
      <div className="mt-4 max-h-64 overflow-y-auto rounded-xl border border-slate-800/90 bg-slate-950/70">
        <table className="w-full text-left border-collapse text-xs">
          <thead className="sticky top-0 bg-slate-900/95 text-slate-300 border-b border-slate-800 z-10">
            <tr>
              <th className="py-2.5 px-3 font-semibold">Mã Hàng (Stock Code)</th>
              <th className="py-2.5 px-3 font-semibold">Model</th>
              <th className="py-2.5 px-3 font-semibold">Bên</th>
              <th className="py-2.5 px-3 font-semibold">Mô Tả Hàng Hóa</th>
              <th className="py-2.5 px-3 font-semibold text-right text-blue-300">Tồn Kho 01</th>
              <th className="py-2.5 px-3 font-semibold text-right text-cyan-300">Tồn Kho 60</th>
              <th className="py-2.5 px-3 font-semibold text-right text-white">Tổng Tồn (01+60)</th>
              <th className="py-2.5 px-3 font-semibold text-right">Cầu Xuất / Trạng Thái</th>
              <th className="py-2.5 px-3 font-semibold text-center">3D / Copy</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-800/60 font-mono">
            {filteredStock.slice(0, 60).map((item) => {
              const dem = demandMap.get(item.code);
              return (
                <tr key={item.code} className="hover:bg-slate-800/50 transition">
                  <td className="py-2 px-3 font-bold text-cyan-300">{item.code}</td>
                  <td className="py-2 px-3 text-slate-300">#{item.feature}</td>
                  <td className="py-2 px-3">
                    {item.side === 'L' ? (
                      <span className="rounded bg-cyan-500/20 border border-cyan-500/30 px-1.5 py-0.5 text-[10px] font-bold text-cyan-300">
                        TRÁI (LHF)
                      </span>
                    ) : item.side === 'R' ? (
                      <span className="rounded bg-emerald-500/20 border border-emerald-500/30 px-1.5 py-0.5 text-[10px] font-bold text-emerald-300">
                        PHẢI (RHF)
                      </span>
                    ) : (
                      <span className="rounded bg-slate-800 px-1.5 py-0.5 text-[10px] text-slate-400">
                        CHUNG
                      </span>
                    )}
                  </td>
                  <td className="py-2 px-3 font-sans text-slate-200 max-w-[230px] truncate">
                    {item.description}
                  </td>
                  <td className="py-2 px-3 text-right text-blue-300">
                    {item.wh01Qty.toLocaleString('en-US')}
                  </td>
                  <td className="py-2 px-3 text-right text-cyan-300">
                    {item.wh60Qty.toLocaleString('en-US')}
                  </td>
                  <td className="py-2 px-3 text-right font-bold text-white">
                    {item.totalQty.toLocaleString('en-US')}
                  </td>
                  <td className="py-2 px-3 text-right">
                    {dem ? (
                      dem.shortageQty > 0 ? (
                        <span className="inline-flex items-center gap-1 rounded bg-rose-500/20 px-2 py-0.5 text-[11px] font-bold text-rose-300">
                          Cầu {dem.totalRequiredQty.toLocaleString('en-US')} (Thiếu -{dem.shortageQty.toLocaleString('en-US')})
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1 rounded bg-emerald-500/15 px-2 py-0.5 text-[11px] font-bold text-emerald-300">
                          Cầu {dem.totalRequiredQty.toLocaleString('en-US')} (Đủ)
                        </span>
                      )
                    ) : (
                      <span className="text-slate-500 text-[11px]">Không có đơn</span>
                    )}
                  </td>
                  <td className="py-2 px-3 text-center">
                    <div className="inline-flex items-center gap-1">
                      {item.isFinishedGood && (
                        <button
                          onClick={() => onSelectModelIn3D(item.feature)}
                          className="rounded bg-cyan-500/20 p-1 text-cyan-300 hover:bg-cyan-500 hover:text-slate-950 transition cursor-pointer"
                          title="Xem trên Dashboard 3D"
                        >
                          <ArrowUpRight className="h-3.5 w-3.5" />
                        </button>
                      )}
                      <button
                        onClick={() => copyCode(item.code)}
                        className="rounded bg-slate-800 p-1 text-slate-300 hover:bg-slate-700 hover:text-white transition cursor-pointer"
                        title="Copy mã hàng"
                      >
                        {copiedCode === item.code ? (
                          <Check className="h-3.5 w-3.5 text-emerald-400" />
                        ) : (
                          <Copy className="h-3.5 w-3.5" />
                        )}
                      </button>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </section>
  );
};
