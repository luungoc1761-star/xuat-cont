import React, { useRef, useState, useEffect } from 'react';
import {
  Upload,
  Cloud,
  RefreshCw,
  CheckCircle2,
  FileSpreadsheet,
  FolderSync,
  Calendar,
  CalendarDays,
  Database,
  X,
  GitBranch,
  Radio
} from 'lucide-react';

interface DataImportAndOneDriveModalProps {
  isOpen: boolean;
  onClose: () => void;
  onUploadFile: (file: File, type: 'stock' | 'monthly' | 'weekly', autoGitPush: boolean) => Promise<void>;
  liveWatchActive: boolean;
  liveWatchFileName: string | null;
  lastSyncLog: string;
  onStartBrowserLiveWatch: (autoGitPush: boolean) => Promise<void>;
  onStopBrowserLiveWatch: () => void;
  onServerOneDriveSync: (folderPath: string, autoGitPush: boolean) => Promise<void>;
}

export const DataImportAndOneDriveModal: React.FC<DataImportAndOneDriveModalProps> = ({
  isOpen,
  onClose,
  onUploadFile,
  liveWatchActive,
  liveWatchFileName,
  lastSyncLog,
  onStartBrowserLiveWatch,
  onStopBrowserLiveWatch,
  onServerOneDriveSync
}) => {
  const stockInputRef = useRef<HTMLInputElement | null>(null);
  const monthlyInputRef = useRef<HTMLInputElement | null>(null);
  const weeklyInputRef = useRef<HTMLInputElement | null>(null);

  const [autoGitPush, setAutoGitPush] = useState(true);
  const [oneDrivePath, setOneDrivePath] = useState('OneDrive/stock list');
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
              <Upload className="h-5 w-5" />
            </div>
            <div>
              <h2 className="text-lg font-bold text-white">
                Trung Tâm Nhập Dữ Liệu Nguồn & Đồng Bộ OneDrive Realtime
              </h2>
              <p className="text-xs text-slate-400">
                Nhập trực tiếp 3 nguồn dữ liệu (File Tồn Kho, Kế Hoạch Tháng, Kế Hoạch Tuần) hoặc kết nối Realtime thư mục OneDrive `Stock list`
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
                  Tự Động Cập Nhật Vào Repo & Đẩy Lên Nhánh `main` Khi Dữ Liệu Thay Đổi
                </div>
                <div className="text-[11px] text-slate-300">
                  Mỗi khi nhập file mới hoặc OneDrive cập nhật `Stock list.xlsx`, hệ thống tự động ghi vào repo, chạy đối soát 4 Agent và cập nhật `main`
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
                Auto Sync & Push `main`
              </span>
            </label>
          </div>

          {/* 3 Source File Upload Cards */}
          <div>
            <h3 className="text-xs font-bold uppercase tracking-wider text-cyan-300 mb-3">
              1. Nhập 3 Nguồn Dữ Liệu Excel (.xlsx / .xls / .csv)
            </h3>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              {/* Card 1: File Tồn Kho */}
              <div className="rounded-xl border border-cyan-500/30 bg-slate-950/70 p-4 flex flex-col justify-between hover:border-cyan-400 transition">
                <div>
                  <div className="flex items-center justify-between">
                    <span className="inline-flex items-center gap-1.5 text-xs font-bold text-cyan-300">
                      <Database className="h-4 w-4" /> 1. FILE TỒN KHO
                    </span>
                    <span className="rounded bg-cyan-500/15 px-2 py-0.5 text-[10px] font-mono text-cyan-300">
                      Kho 01 & 60
                    </span>
                  </div>
                  <p className="mt-2 text-xs text-slate-300 font-semibold">
                    Stock list.xlsx (Sheet: Stock Balance With Batch)
                  </p>
                  <p className="mt-1 text-[11px] text-slate-400">
                    Tự động lọc chuẩn xác dữ liệu ở Kho 01 và Kho 60, ghép cặp Trái (LHF) / Phải (RHF) và cập nhật Dashboard 3D.
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
                    className="w-full inline-flex items-center justify-center gap-2 rounded-xl bg-cyan-500 px-4 py-2.5 text-xs font-bold text-slate-950 hover:bg-cyan-400 transition cursor-pointer"
                  >
                    <Upload className="h-4 w-4" />
                    <span>
                      {uploadingType === 'stock' ? 'Đang Đọc & Đối Soát...' : 'Chọn File Tồn Kho'}
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
                    Đọc kế hoạch xuất hàng theo tháng (Sheet NOV-full, Data OPENING, booking follow) và đối soát đơn đủ / đơn thiếu.
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

          {/* OneDrive Realtime Sync Section */}
          <div className="rounded-2xl border border-sky-500/35 bg-gradient-to-br from-sky-950/30 via-slate-950 to-slate-900 p-5">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div className="flex items-center gap-2.5">
                <Cloud className="h-5 w-5 text-sky-400" />
                <div>
                  <h3 className="text-sm font-bold text-white flex items-center gap-2">
                    2. Kết Nối Realtime Thư Mục OneDrive (`Stock list.xlsx`)
                    {liveWatchActive && (
                      <span className="inline-flex items-center gap-1 rounded-full bg-emerald-500/20 border border-emerald-500/40 px-2.5 py-0.5 text-[11px] font-semibold text-emerald-300">
                        <Radio className="h-3 w-3 animate-pulse" /> Đang Chạy Realtime ({liveWatchFileName})
                      </span>
                    )}
                  </h3>
                  <p className="text-xs text-slate-300">
                    Tự động theo dõi file <code className="text-cyan-300">Stock list.xlsx</code> trong thư mục OneDrive của anh: mỗi khi file thay đổi sẽ tự động đọc lại, cập nhật vào Repo và làm mới báo cáo tức thì!
                  </p>
                </div>
              </div>

              {!liveWatchActive ? (
                <button
                  onClick={() => onStartBrowserLiveWatch(autoGitPush)}
                  className="inline-flex items-center gap-2 rounded-xl bg-gradient-to-r from-sky-500 to-cyan-400 px-4 py-2.5 text-xs font-bold text-slate-950 shadow-lg shadow-sky-500/20 hover:opacity-95 transition cursor-pointer"
                >
                  <FolderSync className="h-4 w-4" />
                  <span>Kết Nối File `Stock list.xlsx` Trong OneDrive (Realtime)</span>
                </button>
              ) : (
                <button
                  onClick={onStopBrowserLiveWatch}
                  className="inline-flex items-center gap-2 rounded-xl bg-rose-500/20 border border-rose-500/40 px-4 py-2 text-xs font-bold text-rose-300 hover:bg-rose-500/30 transition cursor-pointer"
                >
                  <span>Dừng Theo Dõi Realtime</span>
                </button>
              )}
            </div>

            {/* Server/Local OneDrive Path Sync */}
            <div className="mt-4 pt-4 border-t border-slate-800/80 flex flex-col sm:flex-row items-stretch sm:items-center gap-2.5">
              <div className="flex-1">
                <label className="block text-[11px] text-slate-400 mb-1">
                  Đường dẫn thư mục OneDrive chứa file `Stock list.xlsx` (hoặc chạy `npm run watch-onedrive`):
                </label>
                <input
                  type="text"
                  value={oneDrivePath}
                  onChange={(e) => setOneDrivePath(e.target.value)}
                  placeholder="VD: C:/Users/Admin/OneDrive/stock list hoặc /workspaces/xuat-cont"
                  className="w-full rounded-xl border border-slate-700 bg-slate-900 px-3 py-2 text-xs font-mono text-white focus:border-sky-400 focus:outline-none"
                />
              </div>
              <button
                onClick={() => onServerOneDriveSync(oneDrivePath, autoGitPush)}
                className="sm:mt-5 inline-flex items-center justify-center gap-1.5 rounded-xl bg-slate-800 border border-slate-700 px-4 py-2 text-xs font-semibold text-sky-300 hover:bg-slate-700 transition cursor-pointer"
              >
                <RefreshCw className="h-3.5 w-3.5" />
                <span>Đồng Bộ Từ Thư Mục Ngay</span>
              </button>
            </div>

            {lastSyncLog && (
              <div className="mt-3 rounded-lg bg-slate-900/90 border border-slate-800 px-3 py-2 text-xs font-mono text-emerald-300 flex items-center gap-2">
                <CheckCircle2 className="h-4 w-4 shrink-0 text-emerald-400" />
                <span>{lastSyncLog}</span>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};
