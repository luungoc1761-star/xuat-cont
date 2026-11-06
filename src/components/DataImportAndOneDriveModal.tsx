import React, { useRef, useState } from 'react';
import {
  Upload,
  CheckCircle2,
  Calendar,
  CalendarDays,
  Database,
  X,
  GitBranch,
  Calculator
} from 'lucide-react';

interface DataImportAndOneDriveModalProps {
  isOpen: boolean;
  onClose: () => void;
  onUploadFile: (file: File, type: 'stock' | 'monthly' | 'weekly', autoGitPush: boolean) => Promise<void>;
  lastSyncLog: string;
  activeStockFileName?: string;
}

export const DataImportAndOneDriveModal: React.FC<DataImportAndOneDriveModalProps> = ({
  isOpen,
  onClose,
  onUploadFile,
  lastSyncLog,
  activeStockFileName = 'Stock list.xlsx'
}) => {
  const stockInputRef = useRef<HTMLInputElement | null>(null);
  const monthlyInputRef = useRef<HTMLInputElement | null>(null);
  const weeklyInputRef = useRef<HTMLInputElement | null>(null);

  const [autoGitPush, setAutoGitPush] = useState(true);
  const [uploadingType, setUploadingType] = useState<string | null>(null);

  if (!isOpen) return null;

  const handleFileChange = async (
    e: React.ChangeEvent<HTMLInputElement>,
    type: 'stock' | 'monthly' | 'weekly'
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploadingType(type);
    try {
      await onUploadFile(file, type, autoGitPush);
    } finally {
      setUploadingType(null);
      e.target.value = '';
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/80 backdrop-blur-md p-4 overflow-y-auto">
      <div className="relative w-full max-w-4xl rounded-2xl border border-cyan-500/40 bg-slate-900 shadow-2xl overflow-hidden my-8">
        {/* Header */}
        <div className="flex items-center justify-between border-b border-slate-800 bg-slate-950/80 px-6 py-4">
          <div className="flex items-center gap-3">
            <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-cyan-500/20 border border-cyan-500/40 text-cyan-300">
              <Calculator className="h-5 w-5" />
            </div>
            <div>
              <h2 className="text-lg font-bold text-white">
                Nhập File Tồn Kho & Tính Đối Soát Kế Hoạch Xuất Hàng
              </h2>
              <p className="text-xs text-slate-400">
                Chọn file tồn kho (Stock list.xlsx) hoặc file kế hoạch xuất hàng để đọc dữ liệu Kho 01 & Kho 60 và tính đối soát tự động
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="rounded-lg p-2 text-slate-400 hover:bg-slate-800 hover:text-white cursor-pointer"
          >
            <X className="h-5 w-5" />
          </button>
        </div>

        <div className="p-6 space-y-6">
          {/* Auto Git Commit & Push Toggle */}
          <div className="flex flex-wrap items-center justify-between gap-3 rounded-xl border border-emerald-500/30 bg-emerald-950/20 px-4 py-3">
            <div className="flex items-center gap-2.5">
              <GitBranch className="h-5 w-5 text-emerald-400" />
              <div>
                <div className="text-xs font-bold text-white">
                  Tự Động Lưu Vào Repo Khi Nhập File Mới
                </div>
                <div className="text-[11px] text-slate-300">
                  Mỗi khi chọn file mới, hệ thống tự động tính đối soát trên trình duyệt và lưu cập nhật vào repo (nếu có server local)
                </div>
              </div>
            </div>
            <label className="inline-flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={autoGitPush}
                onChange={(e) => setAutoGitPush(e.target.checked)}
                className="h-4 w-4 accent-emerald-500 rounded"
              />
              <span className="text-xs font-semibold text-emerald-300">
                Lưu & Cập Nhật Repo
              </span>
            </label>
          </div>

          {/* 3 Source File Upload Cards */}
          <div>
            <h3 className="text-xs font-bold uppercase tracking-wider text-cyan-300 mb-3">
              Chọn File Nguồn Để Tính Đối Soát (.xlsx / .xls / .csv)
            </h3>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              {/* Card 1: File Tồn Kho */}
              <div className="rounded-xl border-2 border-cyan-500/50 bg-slate-950/90 p-4 flex flex-col justify-between hover:border-cyan-400 transition shadow-lg shadow-cyan-500/10">
                <div>
                  <div className="flex items-center justify-between">
                    <span className="inline-flex items-center gap-1.5 text-xs font-bold text-cyan-300">
                      <Database className="h-4 w-4" /> 1. FILE TỒN KHO
                    </span>
                    <span className="rounded bg-cyan-500/15 px-2 py-0.5 text-[10px] font-mono text-cyan-300">
                      Kho 01 & 60
                    </span>
                  </div>
                  <p className="mt-2 text-xs text-slate-200 font-semibold">
                    Đang dùng: <span className="text-cyan-300">{activeStockFileName}</span>
                  </p>
                  <p className="mt-1 text-[11px] text-slate-400">
                    Chọn file tồn kho để hệ thống tự động lọc Kho 01 & Kho 60, ghép cặp Trái (LHF) / Phải (RHF) và tính đối soát ngay với kế hoạch xuất Cont.
                  </p>
                </div>
                <div className="mt-4">
                  <input
                    ref={stockInputRef}
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    onChange={(e) => handleFileChange(e, 'stock')}
                    className="hidden"
                  />
                  <button
                    onClick={() => stockInputRef.current?.click()}
                    disabled={uploadingType === 'stock'}
                    className="w-full inline-flex items-center justify-center gap-2 rounded-xl bg-gradient-to-r from-cyan-400 to-sky-500 px-4 py-2.5 text-xs font-extrabold text-slate-950 hover:brightness-110 transition cursor-pointer shadow-md shadow-cyan-500/20"
                  >
                    <Upload className="h-4 w-4" />
                    <span>
                      {uploadingType === 'stock'
                        ? 'Đang Đọc & Tính Đối Soát...'
                        : 'Chọn File Tồn Kho & Tính Đối Soát'}
                    </span>
                  </button>
                </div>
              </div>

              {/* Card 2: File Kế Hoạch Tháng */}
              <div className="rounded-xl border border-indigo-500/30 bg-slate-950/70 p-4 flex flex-col justify-between hover:border-indigo-400 transition">
                <div>
                  <div className="flex items-center justify-between">
                    <span className="inline-flex items-center gap-1.5 text-xs font-bold text-indigo-300">
                      <Calendar className="h-4 w-4" /> 2. FILE KẾ HOẠCH THÁNG
                    </span>
                    <span className="rounded bg-indigo-500/15 px-2 py-0.5 text-[10px] font-mono text-indigo-300">
                      NOV-full / Data
                    </span>
                  </div>
                  <p className="mt-2 text-xs text-slate-300 font-semibold">
                    NOV. SHIPMENT -OCT22.xlsx
                  </p>
                  <p className="mt-1 text-[11px] text-slate-400">
                    Đọc kế hoạch xuất hàng theo tháng (Sheet NOV-full, Data OPENING) và đối soát đơn đủ / đơn thiếu.
                  </p>
                </div>
                <div className="mt-4">
                  <input
                    ref={monthlyInputRef}
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    onChange={(e) => handleFileChange(e, 'monthly')}
                    className="hidden"
                  />
                  <button
                    onClick={() => monthlyInputRef.current?.click()}
                    disabled={uploadingType === 'monthly'}
                    className="w-full inline-flex items-center justify-center gap-2 rounded-xl bg-indigo-500 px-4 py-2.5 text-xs font-bold text-white hover:bg-indigo-400 transition cursor-pointer"
                  >
                    <Upload className="h-4 w-4" />
                    <span>
                      {uploadingType === 'monthly'
                        ? 'Đang Đọc Kế Hoạch Tháng...'
                        : 'Chọn File Kế Hoạch Tháng'}
                    </span>
                  </button>
                </div>
              </div>

              {/* Card 3: File Kế Hoạch Tuần */}
              <div className="rounded-xl border border-emerald-500/30 bg-slate-950/70 p-4 flex flex-col justify-between hover:border-emerald-400 transition">
                <div>
                  <div className="flex items-center justify-between">
                    <span className="inline-flex items-center gap-1.5 text-xs font-bold text-emerald-300">
                      <CalendarDays className="h-4 w-4" /> 3. FILE KẾ HOẠCH TUẦN
                    </span>
                    <span className="rounded bg-emerald-500/15 px-2 py-0.5 text-[10px] font-mono text-emerald-300">
                      OCT-WK 42 / Weekly
                    </span>
                  </div>
                  <p className="mt-2 text-xs text-slate-300 font-semibold">
                    Sheet OCT-WK 42 hoặc File Tuần riêng
                  </p>
                  <p className="mt-1 text-[11px] text-slate-400">
                    Nhập file kế hoạch xuất hàng theo tuần để kiểm tra nhanh tiến độ đóng cont trong tuần đối chiếu với Kho 01 & 60.
                  </p>
                </div>
                <div className="mt-4">
                  <input
                    ref={weeklyInputRef}
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    onChange={(e) => handleFileChange(e, 'weekly')}
                    className="hidden"
                  />
                  <button
                    onClick={() => weeklyInputRef.current?.click()}
                    disabled={uploadingType === 'weekly'}
                    className="w-full inline-flex items-center justify-center gap-2 rounded-xl bg-emerald-500 px-4 py-2.5 text-xs font-bold text-slate-950 hover:bg-emerald-400 transition cursor-pointer"
                  >
                    <Upload className="h-4 w-4" />
                    <span>
                      {uploadingType === 'weekly'
                        ? 'Đang Đọc Kế Hoạch Tuần...'
                        : 'Chọn File Kế Hoạch Tuần'}
                    </span>
                  </button>
                </div>
              </div>
            </div>
          </div>

          {lastSyncLog && (
            <div className="rounded-xl bg-slate-950/90 border border-emerald-500/30 px-4 py-3 text-xs font-mono text-emerald-300 flex items-center justify-between gap-3">
              <div className="flex items-center gap-2">
                <CheckCircle2 className="h-4 w-4 shrink-0 text-emerald-400" />
                <span>{lastSyncLog}</span>
              </div>
              <button
                onClick={onClose}
                className="shrink-0 rounded-lg bg-emerald-500 px-3 py-1.5 text-xs font-sans font-bold text-slate-950 hover:bg-emerald-400 transition cursor-pointer"
              >
                Xem Báo Cáo Đối Soát
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
