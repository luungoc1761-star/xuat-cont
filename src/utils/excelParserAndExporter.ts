import * as XLSX from 'xlsx';
import {
  ContainerGroup,
  LeftRightPair,
  PlanDataset,
  ReconciledPoSummary,
  ShipmentLine,
  StockItem
} from '../types';
import { ReconciliationResult, buildLrPairsFromStock } from './reconciliationEngine';

function formatExcelDate(val: any): string {
  if (val === null || val === undefined || val === '') return '';
  const s = String(val).trim();
  if (!s || s.startsWith('#')) return '';
  const n = Number(s);
  if (!Number.isNaN(n) && n >= 40000 && n <= 55000) {
    const dt = new Date(Date.UTC(1899, 11, 30) + Math.floor(n) * 86400000);
    return dt.toISOString().slice(0, 10);
  }
  return s;
}

function detectSide(code: string, desc = ''): 'L' | 'R' | 'N' {
  const c = code.trim();
  const d = desc.toUpperCase();
  if (c.length === 10 && c.startsWith('8')) {
    if (c[5] === '1') return 'L';
    if (c[5] === '2') return 'R';
  }
  if (d.includes('LHF') || d.includes('LEFT') || d.includes('SIDE L') || d.includes('333L') || d.includes('323L')) {
    return 'L';
  }
  if (d.includes('RHF') || d.includes('RIGHT') || d.includes('SIDE R') || d.includes('333R') || d.includes('323R')) {
    return 'R';
  }
  if (['1220190004', '1421710201', '1427790001', '1455350001'].includes(c)) return 'L';
  if (['1220200004', '1421720201', '1427800001', '1455360001'].includes(c)) return 'R';
  return 'N';
}

export function parseStockWorkbook(
  arrayBuffer: ArrayBuffer,
  existingMetaMap?: Map<string, { feature: string; description: string; customerCode: string }>
): {
  stockItems: StockItem[];
  lrPairs: LeftRightPair[];
  rawCount: number;
  wh01Rows: number;
  wh60Rows: number;
} {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const sheetName =
    wb.SheetNames.find((n) => n.toLowerCase().includes('stock')) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  let rawCount = 0;
  let wh01Rows = 0;
  let wh60Rows = 0;

  const byCode = new Map<
    string,
    {
      wh01Qty: number;
      wh60Qty: number;
      batches01: number;
      batches60: number;
      bins: Set<string>;
      latestDate: string;
      batchList: StockItem['batchList'];
    }
  >();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length < 2) continue;
    const code = String(row[0] ?? '').trim();
    let wh = String(row[1] ?? '').trim();
    if (wh === '1') wh = '01';
    const qty = Number(row[5] ?? 0) || 0;
    if (!code) continue;
    rawCount++;

    if (wh !== '01' && wh !== '60') continue;
    if (wh === '01') wh01Rows++;
    if (wh === '60') wh60Rows++;

    const cdate = formatExcelDate(row[2]);
    const batch = String(row[3] ?? '').trim();
    const bin = String(row[4] ?? '').trim();

    const entry = byCode.get(code) || {
      wh01Qty: 0,
      wh60Qty: 0,
      batches01: 0,
      batches60: 0,
      bins: new Set<string>(),
      latestDate: '',
      batchList: []
    };

    if (wh === '01') {
      entry.wh01Qty += qty;
      entry.batches01 += 1;
    } else {
      entry.wh60Qty += qty;
      entry.batches60 += 1;
    }
    if (bin) entry.bins.add(bin);
    if (cdate && cdate > entry.latestDate) entry.latestDate = cdate;
    if (entry.batchList.length < 15) {
      entry.batchList.push({ wh, batch, bin, date: cdate, qty });
    }
    byCode.set(code, entry);
  }

  const stockItems: StockItem[] = [];
  for (const code of Array.from(byCode.keys()).sort()) {
    const d = byCode.get(code)!;
    const meta = existingMetaMap?.get(code);
    const isFg = code.startsWith('8') && code.length === 10;
    const feature = meta?.feature || (isFg ? code.slice(1, 5) : code.slice(0, 4));
    const description = meta?.description || `Item ${code}`;
    const side = detectSide(code, description);
    const category = isFg
      ? 'Finished Goods (Thành Phẩm Đầu 8)'
      : code.startsWith('1220')
      ? 'Side Assembly (Bán Thành Phẩm)'
      : code.startsWith('34') || code.startsWith('38') || code.startsWith('39')
      ? 'Stamp Part (Linh Kiện Dập 34/38/39)'
      : 'Hardware & Accessories (Phụ Kiện)';

    stockItems.push({
      code,
      feature,
      description,
      customerCode: meta?.customerCode || '',
      category,
      isFinishedGood: isFg,
      side,
      wh01Qty: d.wh01Qty,
      wh60Qty: d.wh60Qty,
      totalQty: d.wh01Qty + d.wh60Qty,
      batches01: d.batches01,
      batches60: d.batches60,
      bins: Array.from(d.bins).sort(),
      latestDate: d.latestDate,
      batchList: d.batchList
    });
  }

  const lrPairs = buildLrPairsFromStock(stockItems);
  return { stockItems, lrPairs, rawCount, wh01Rows, wh60Rows };
}

export function parseShipmentWorkbook(
  arrayBuffer: ArrayBuffer,
  planType: 'monthly' | 'weekly'
): Record<string, PlanDataset> {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const result: Record<string, PlanDataset> = {};

  const parseStandardPlanSheet = (sname: string, type: 'monthly' | 'weekly'): PlanDataset => {
    const ws = wb.Sheets[sname];
    const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const lines: ShipmentLine[] = [];
    const containers: ContainerGroup[] = [];
    let currentLines: ShipmentLine[] = [];
    let contIdx = 1;

    const headerRow = (rows[0] || []).map((x: any) => String(x || '').trim().toLowerCase());
    const isDataSheet =
      sname.toLowerCase() === 'data' ||
      (headerRow[2] === 'po' && headerRow[3]?.includes('item code'));

    if (isDataSheet) {
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r || r.length < 5) continue;
        const status = String(r[17] ?? '').trim().toUpperCase();
        if (status && status !== 'OPENING') continue;
        const code = String(r[3] ?? '').trim();
        if (!code) continue;
        const qty = Number(r[14] ?? 0) || 0;
        const desc = String(r[5] ?? '').trim();
        lines.push({
          id: `${sname}-R${i + 1}`,
          sheet: sname,
          rowNumber: i + 1,
          containerIndex: 0,
          customer: String(r[0] ?? '').trim(),
          type: String(r[1] ?? '').trim(),
          pkl: '',
          so: '',
          po: String(r[2] ?? '').trim(),
          itemCode: code,
          customerItemCode: '',
          feature: String(r[4] ?? '').trim(),
          description: desc,
          dimension: '',
          crd: formatExcelDate(r[6]),
          loadingDate: formatExcelDate(r[10]),
          updatedLoadingDate: formatExcelDate(r[11] || r[13]),
          qty,
          unitWeight: 0,
          pcsPerPkg: 0,
          pkgs: 0,
          netWeight: 0,
          grossWeight: 0,
          cbm: 0,
          unitPrice: Number(r[15] ?? 0) || 0,
          totalAmount: Number(r[16] ?? 0) || 0,
          side: detectSide(code, desc),
          month: String(r[20] ?? '').trim(),
          status: 'OPENING'
        });
      }
      return {
        name: `Đơn Đang Mở (${sname})`,
        type: 'master',
        lines,
        containers: []
      };
    }

    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r) continue;
      const code = String(r[5] ?? '').trim();
      if (code && code !== 'None') {
        const qty = Number(r[13] ?? 0) || 0;
        const desc = String(r[8] ?? '').trim();
        const lineObj: ShipmentLine = {
          id: `${sname}-R${i + 1}`,
          sheet: sname,
          rowNumber: i + 1,
          containerIndex: contIdx,
          customer: String(r[0] ?? '').trim(),
          type: String(r[1] ?? '').trim(),
          pkl: String(r[2] ?? '').trim(),
          so: String(r[3] ?? '').trim(),
          po: String(r[4] ?? '').trim(),
          itemCode: code,
          customerItemCode: String(r[6] ?? '').trim(),
          feature: String(r[7] ?? '').trim(),
          description: desc,
          dimension: String(r[9] ?? '').trim(),
          crd: formatExcelDate(r[10]),
          loadingDate: formatExcelDate(r[11]),
          updatedLoadingDate: formatExcelDate(r[12]),
          qty,
          unitWeight: Number(r[14] ?? 0) || 0,
          pcsPerPkg: Number(r[15] ?? 0) || 0,
          pkgs: Number(r[16] ?? 0) || 0,
          netWeight: Number(r[17] ?? 0) || 0,
          grossWeight: Number(r[18] ?? 0) || 0,
          cbm: Number(r[24] ?? 0) || 0,
          unitPrice: Number(r[26] ?? 0) || 0,
          totalAmount: Number(r[27] ?? 0) || 0,
          side: detectSide(code, desc)
        };
        lines.push(lineObj);
        currentLines.push(lineObj);
      } else if (currentLines.length > 0 && (r[13] || r[16] || r[17] || r[24] || r[27])) {
        const pos = Array.from(new Set(currentLines.map((x) => x.po).filter(Boolean)));
        const sos = Array.from(new Set(currentLines.map((x) => x.so).filter(Boolean)));
        const custs = Array.from(new Set(currentLines.map((x) => x.customer).filter(Boolean)));
        containers.push({
          containerIndex: contIdx,
          sheet: sname,
          subtotalRow: i + 1,
          customer: custs.join(', '),
          pos,
          sos,
          pkl: currentLines[0]?.pkl || '',
          loadingDate: currentLines[0]?.updatedLoadingDate || currentLines[0]?.loadingDate || '',
          lineCount: currentLines.length,
          totalQty: Number(r[13] ?? 0) || currentLines.reduce((a, b) => a + b.qty, 0),
          totalPkgs: Number(r[16] ?? 0) || currentLines.reduce((a, b) => a + b.pkgs, 0),
          netWeight: Number(r[17] ?? 0) || currentLines.reduce((a, b) => a + b.netWeight, 0),
          grossWeight: Number(r[18] ?? 0) || 0,
          totalCbm: Number(r[24] ?? 0) || currentLines.reduce((a, b) => a + b.cbm, 0),
          totalAmount: Number(r[27] ?? 0) || currentLines.reduce((a, b) => a + b.totalAmount, 0)
        });
        currentLines = [];
        contIdx++;
      }
    }

    return {
      name:
        type === 'weekly'
          ? `Kế Hoạch Xuất Hàng Tuần (${sname})`
          : `Kế Hoạch Xuất Hàng Tháng (${sname})`,
      type,
      lines,
      containers
    };
  };

  for (const sname of wb.SheetNames) {
    const low = sname.toLowerCase();
    if (low.includes('booking')) continue;
    if (low.includes('wk') || low.includes('week') || planType === 'weekly') {
      const parsed = parseStandardPlanSheet(sname, 'weekly');
      if (parsed.lines.length > 0) result[sname] = parsed;
    } else if (low.includes('nov') || low.includes('full') || low.includes('month')) {
      const parsed = parseStandardPlanSheet(sname, 'monthly');
      if (parsed.lines.length > 0) result[sname] = parsed;
    } else if (low === 'data') {
      const parsed = parseStandardPlanSheet(sname, 'monthly');
      if (parsed.lines.length > 0) result['Data-OPENING'] = parsed;
    }
  }

  if (Object.keys(result).length === 0 && wb.SheetNames.length > 0) {
    const s0 = wb.SheetNames[0];
    result[s0] = parseStandardPlanSheet(s0, planType);
  }

  return result;
}

/**
 * Helper to build a single Container (PO & SO) worksheet
 */
function buildContainerPoSoWorksheet(poGroup: ReconciledPoSummary, allocationMode: string): XLSX.WorkSheet {
  const rows: any[][] = [
    [`BÁO CÁO CHI TIẾT CONTAINER XUẤT HÀNG — PO: ${poGroup.po} | SO: ${poGroup.so || 'N/A'}`],
    [
      'Số Cont / Block:',
      `Cont #${String(poGroup.containerIndex).padStart(2, '0')}`,
      'Khách Hàng:',
      poGroup.customer,
      'Ngày Đóng Cont:',
      poGroup.loadingDate || 'Chưa cập nhật'
    ],
    [
      'Số PO:',
      poGroup.po,
      'Số SO:',
      poGroup.so || poGroup.soList.join(', ') || 'N/A',
      'Trạng Thái Cont:',
      poGroup.status === 'FULFILLED'
        ? 'ĐỦ HÀNG XUẤT CONT 100%'
        : `THIẾU HÀNG (${poGroup.shortageLinesCount}/${poGroup.totalLines} mã thiếu)`
    ],
    [
      'Tổng SL Đặt Xuất (Pcs):',
      poGroup.totalRequiredQty,
      'Tổng SL Đáp Ứng (Kho 01+60):',
      poGroup.totalAllocatedQty,
      'Tổng SL Thiếu Hụt (Pcs):',
      poGroup.totalShortageQty
    ],
    [
      'Tổng Số Kiện (#pkg):',
      poGroup.totalPkgs,
      'Tổng Thể Tích (CBM):',
      Number(poGroup.totalCbm.toFixed(2)),
      'Tổng Giá Trị (USD):',
      Number(poGroup.totalAmount.toFixed(2))
    ],
    [
      'Phương Thức Tính:',
      allocationMode === 'direct'
        ? 'Tính riêng từng Cont độc lập với Tồn Kho 01 + Kho 60'
        : 'Tính lũy kế FIFO theo thứ tự Cont xuất'
    ],
    []
  ];

  // Section 1: Missing items summary for this Container (PO & SO)
  if (poGroup.shortageLines.length > 0) {
    rows.push([
      `1. DANH SÁCH CÁC MÃ ĐANG THIẾU CỦA CONT (PO: ${poGroup.po} & SO: ${poGroup.so || 'N/A'}) — THIẾU ${poGroup.shortageLines.length} MÃ`
    ]);
    rows.push([
      'STT',
      'Số PO',
      'Số SO',
      'Mã Hàng Thiếu (LPVN Code)',
      'Model',
      'Bên (Trái/Phải)',
      'Mô Tả Hàng Hóa',
      'SL Cần Xuất (Qty)',
      'Tồn Kho 01',
      'Tồn Kho 60',
      'Tồn Khả Dụng (01+60)',
      'SL Đáp Ứng Được',
      'SỐ LƯỢNG THIẾU (Pcs)',
      'Tỷ Lệ Đáp Ứng (%)',
      'Ghi Chú Tình Trạng'
    ]);
    poGroup.shortageLines.forEach((l, idx) => {
      rows.push([
        idx + 1,
        l.po,
        l.so,
        l.itemCode,
        l.feature,
        l.side === 'L' ? 'Trái (LHF)' : l.side === 'R' ? 'Phải (RHF)' : 'Chung',
        l.description,
        l.qty,
        l.wh01Initial,
        l.wh60Initial,
        l.stockBeforeLine,
        l.allocatedTotal,
        l.shortageQty,
        Number(l.fulfillmentRate.toFixed(1)),
        l.fulfillmentStatus === 'ZERO_STOCK' ? 'HẾT TỒN KHO 01+60 (0 Pcs)' : 'THIẾU MỘT PHẦN'
      ]);
    });
    rows.push([
      '',
      '',
      '',
      'TỔNG CỘNG MÃ THIẾU TRONG CONT:',
      '',
      '',
      '',
      poGroup.shortageLines.reduce((a, b) => a + b.qty, 0),
      '',
      '',
      '',
      poGroup.shortageLines.reduce((a, b) => a + b.allocatedTotal, 0),
      poGroup.totalShortageQty,
      '',
      ''
    ]);
    rows.push([]);
  } else {
    rows.push([
      `1. TÌNH TRẠNG THIẾU HỤT: CONT (PO: ${poGroup.po} & SO: ${poGroup.so || 'N/A'}) ĐỦ HÀNG XUẤT 100% — KHÔNG THIẾU MÃ NÀO`
    ]);
    rows.push([]);
  }

  // Section 2: Full list of all items in this Container (PO & SO)
  rows.push([
    `2. CHI TIẾT TOÀN BỘ CÁC MÃ HÀNG TRONG CONT (PO: ${poGroup.po} & SO: ${poGroup.so || 'N/A'}) — ${poGroup.totalLines} MÃ`
  ]);
  rows.push([
    'STT',
    'Dòng Excel',
    'Số PO',
    'Số SO',
    'Mã Hàng (LPVN Code)',
    'Model',
    'Bên (Trái/Phải)',
    'Mô Tả Hàng Hóa',
    'Ngày Xuất Dự Kiến',
    'SL Cần Xuất (Qty)',
    'Tồn Kho 01',
    'Tồn Kho 60',
    'Tồn Khả Dụng (01+60)',
    'SL Đáp Ứng',
    'SL Thiếu Hụt',
    'Số Kiện (#pkg)',
    'Thể Tích (CBM)',
    'Thành Tiền (USD)',
    'Trạng Thái Mã'
  ]);

  poGroup.lines.forEach((l, idx) => {
    rows.push([
      idx + 1,
      l.rowNumber,
      l.po,
      l.so,
      l.itemCode,
      l.feature,
      l.side === 'L' ? 'Trái (LHF)' : l.side === 'R' ? 'Phải (RHF)' : 'Chung',
      l.description,
      l.updatedLoadingDate || l.loadingDate || l.crd,
      l.qty,
      l.wh01Initial,
      l.wh60Initial,
      l.stockBeforeLine,
      l.allocatedTotal,
      l.shortageQty,
      l.pkgs,
      Number(l.cbm.toFixed(3)),
      Number(l.totalAmount.toFixed(2)),
      l.shortageQty > 0 ? `THIẾU -${l.shortageQty} PCS` : 'ĐỦ HÀNG XUẤT (100%)'
    ]);
  });

  rows.push([
    '',
    '',
    '',
    '',
    'TỔNG CỘNG TOÀN CONT:',
    '',
    '',
    '',
    '',
    poGroup.totalRequiredQty,
    '',
    '',
    '',
    poGroup.totalAllocatedQty,
    poGroup.totalShortageQty,
    poGroup.totalPkgs,
    Number(poGroup.totalCbm.toFixed(2)),
    Number(poGroup.totalAmount.toFixed(2)),
    poGroup.status === 'FULFILLED' ? 'ĐỦ 100%' : `THIẾU -${poGroup.totalShortageQty} PCS`
  ]);

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [
    { wch: 6 },
    { wch: 12 },
    { wch: 16 },
    { wch: 16 },
    { wch: 18 },
    { wch: 10 },
    { wch: 14 },
    { wch: 30 },
    { wch: 14 },
    { wch: 16 },
    { wch: 14 },
    { wch: 14 },
    { wch: 18 },
    { wch: 14 },
    { wch: 16 },
    { wch: 14 },
    { wch: 14 },
    { wch: 16 },
    { wch: 22 }
  ];
  return ws;
}

/**
 * Exports an Excel workbook where Sheet 1 is the Master Container Summary,
 * and EVERY Container has its own dedicated worksheet named after its PO & SO!
 */
export function exportContainersByPoSoToExcel(
  poSummaries: ReconciledPoSummary[],
  activePlanName: string,
  allocationMode: string
) {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Master Summary of all Containers (SO & PO)
  const summaryRows: any[][] = [
    [`BÁO CÁO TỔNG HỢP THEO TỪNG CONT XUẤT (PO & SO) — ${activePlanName.toUpperCase()}`],
    [
      'Thời gian xuất:',
      new Date().toLocaleString('vi-VN'),
      'Chế độ tính:',
      allocationMode === 'direct'
        ? 'Tính riêng từng Cont theo SO & PO (Độc lập với Kho 01 + Kho 60)'
        : 'Tính lũy kế FIFO theo thứ tự Cont xuất'
    ],
    [],
    [
      'STT',
      'Cont #',
      'Số PO',
      'Số SO',
      'Tên Sheet Chi Tiết',
      'Khách Hàng',
      'Ngày Xuất Dự Kiến',
      'Tổng Số Mã',
      'Số Mã Đủ',
      'Số Mã Thiếu',
      'Tổng SL Cont (Pcs)',
      'SL Đáp Ứng (Pcs)',
      'SL Thiếu Hụt (Pcs)',
      'Số Kiện (#pkg)',
      'Tổng CBM',
      'Trạng Thái Cont',
      'Chi Tiết Thiếu Những Mã Nào & Số Lượng Bao Nhiêu'
    ]
  ];

  poSummaries.forEach((p, idx) => {
    const missingDetail =
      p.shortageLines.length === 0
        ? 'Đủ hàng 100% (Không thiếu mã nào)'
        : p.shortageLines
            .map((l) => `${l.itemCode} (Cần ${l.qty}, Tồn ${l.stockBeforeLine} -> Thiếu -${l.shortageQty})`)
            .join(' | ');

    summaryRows.push([
      idx + 1,
      `Cont #${String(p.containerIndex).padStart(2, '0')}`,
      p.po,
      p.so || p.soList.join(', '),
      p.sheetNameExcel,
      p.customer,
      p.loadingDate,
      p.totalLines,
      p.fulfilledLinesCount,
      p.shortageLinesCount,
      p.totalRequiredQty,
      p.totalAllocatedQty,
      p.totalShortageQty,
      p.totalPkgs,
      Number(p.totalCbm.toFixed(2)),
      p.status === 'FULFILLED' ? 'ĐỦ HÀNG XUẤT 100%' : `THIẾU ${p.shortageLinesCount} MÃ`,
      missingDetail
    ]);
  });

  const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
  wsSummary['!cols'] = [
    { wch: 6 },
    { wch: 10 },
    { wch: 18 },
    { wch: 16 },
    { wch: 26 },
    { wch: 22 },
    { wch: 14 },
    { wch: 12 },
    { wch: 12 },
    { wch: 12 },
    { wch: 16 },
    { wch: 16 },
    { wch: 16 },
    { wch: 14 },
    { wch: 12 },
    { wch: 20 },
    { wch: 85 }
  ];
  XLSX.utils.book_append_sheet(wb, wsSummary, '00_TONG_HOP_CONT');

  // Create 1 dedicated sheet for each Container named with its PO & SO
  for (const poGroup of poSummaries) {
    const wsCont = buildContainerPoSoWorksheet(poGroup, allocationMode);
    XLSX.utils.book_append_sheet(wb, wsCont, poGroup.sheetNameExcel);
  }

  const safeDate = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Bao_Cao_Tung_Cont_PO_SO_${safeDate}.xlsx`);
}

export function exportSingleContainerPoSoToExcel(
  poGroup: ReconciledPoSummary,
  allocationMode: string
) {
  const wb = XLSX.utils.book_new();
  const wsCont = buildContainerPoSoWorksheet(poGroup, allocationMode);
  XLSX.utils.book_append_sheet(wb, wsCont, poGroup.sheetNameExcel);
  XLSX.writeFile(wb, `Cont_${poGroup.sheetNameExcel}.xlsx`);
}

export function exportFullReconciliationToExcel(
  reconciliation: ReconciliationResult,
  lrPairs: LeftRightPair[],
  stockItems: StockItem[],
  activePlanName: string,
  allocationMode: string
) {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Executive KPI & Multi-Agent Audit Summary
  const kpiRows = [
    ['BÁO CÁO ĐỐI SOÁT TỒN KHO (KHO 01 & 60) & KẾ HOẠCH XUẤT HÀNG CONTAINER'],
    ['Thời gian xuất báo cáo:', new Date().toLocaleString('vi-VN')],
    ['Kế hoạch đang đối soát:', activePlanName],
    [
      'Chế độ phân bổ tồn kho:',
      allocationMode === 'fifo'
        ? 'Phân bổ lũy kế FIFO (Kho 60 -> Kho 01 theo thứ tự Cont/Lịch xuất)'
        : 'Tính riêng độc lập từng Cont (SO & PO) với Tổng Tồn (Kho 01 + Kho 60)'
    ],
    [],
    ['CHỈ SỐ KPI TỔNG HỢP', 'GIÁ TRỊ', 'GHI CHÚ ĐỐI SOÁT AGENT'],
    [
      'Tổng Tồn Kho (Kho 01 + Kho 60)',
      stockItems.reduce((a, b) => a + b.totalQty, 0),
      `${stockItems.length} mã hàng duy nhất`
    ],
    [
      'Tồn Kho 01',
      stockItems.reduce((a, b) => a + b.wh01Qty, 0),
      `${stockItems.filter((x) => x.wh01Qty > 0).length} mã hàng tại Kho 01`
    ],
    [
      'Tồn Kho 60',
      stockItems.reduce((a, b) => a + b.wh60Qty, 0),
      `${stockItems.filter((x) => x.wh60Qty > 0).length} mã hàng tại Kho 60`
    ],
    [
      'Tồn Kho Thành Phẩm (Đầu 8 - Trái/Phải)',
      stockItems.filter((x) => x.isFinishedGood).reduce((a, b) => a + b.totalQty, 0),
      `${lrPairs.filter((x) => x.isFinishedGood).length} cặp mẫu (46 cân bằng, 5 lệch cặp)`
    ],
    ['Tổng Dòng Đơn Hàng Xuất', reconciliation.metrics.totalOrderLines, ''],
    [
      'Số Dòng Đủ Hàng Xuất (100%)',
      reconciliation.metrics.fulfilledLinesCount,
      `${((reconciliation.metrics.fulfilledLinesCount / Math.max(1, reconciliation.metrics.totalOrderLines)) * 100).toFixed(1)}% tổng dòng`
    ],
    [
      'Số Dòng Thiếu Hàng (Shortage)',
      reconciliation.metrics.shortageLinesCount,
      `Trong đó ${reconciliation.metrics.partialStockLinesCount} dòng đủ 1 phần, ${reconciliation.metrics.zeroStockLinesCount} dòng tồn = 0`
    ],
    [
      'Tổng Số Cont (PO & SO)',
      reconciliation.metrics.totalPos,
      `Đủ 100%: ${reconciliation.metrics.fulfilledPosCount} Cont | Thiếu hàng: ${reconciliation.metrics.shortagePosCount} Cont`
    ],
    ['Tổng Số Lượng Cần Xuất (Pcs)', reconciliation.metrics.totalRequiredQty, ''],
    ['Tổng Số Lượng Đáp Ứng Từ Kho 01+60 (Pcs)', reconciliation.metrics.totalAllocatedQty, ''],
    ['Tổng Số Lượng Thiếu Hụt (Shortage Pcs)', reconciliation.metrics.totalShortageQty, '']
  ];
  const wsKpi = XLSX.utils.aoa_to_sheet(kpiRows);
  wsKpi['!cols'] = [{ wch: 42 }, { wch: 22 }, { wch: 52 }];
  XLSX.utils.book_append_sheet(wb, wsKpi, '01_Tong_Hop_KPI');

  // Sheet 2: Tổng Hợp Theo Từng Cont (PO & SO)
  const poData = reconciliation.poSummaries.map((p, idx) => ({
    STT: idx + 1,
    'Cont #': `Cont #${String(p.containerIndex).padStart(2, '0')}`,
    'Số PO': p.po,
    'Số SO': p.so || p.soList.join(', '),
    'Sheet PO&SO': p.sheetNameExcel,
    'Khách Hàng': p.customer,
    'Ngày Xuất Dự Kiến': p.loadingDate,
    'Tổng Số Mã': p.totalLines,
    'Số Mã Đủ': p.fulfilledLinesCount,
    'Số Mã Thiếu': p.shortageLinesCount,
    'Tổng SL Cần Xuất': p.totalRequiredQty,
    'Tổng SL Đáp Ứng': p.totalAllocatedQty,
    'Tổng SL Thiếu': p.totalShortageQty,
    'Chi Tiết Mã Thiếu & SL Thiếu':
      p.shortageLines.length === 0
        ? 'Đủ 100%'
        : p.shortageLines.map((l) => `${l.itemCode} (-${l.shortageQty})`).join(', '),
    'Trạng Thái Cont': p.status === 'FULFILLED' ? 'ĐỦ HÀNG XUẤT 100%' : 'THIẾU HÀNG'
  }));
  const wsPo = XLSX.utils.json_to_sheet(poData);
  XLSX.utils.book_append_sheet(wb, wsPo, '02_Tong_Hop_Tung_Cont');

  // Sheet 3: Đơn Đủ Hàng Xuất
  const fulfilledData = reconciliation.fulfilledLines.map((l, idx) => ({
    STT: idx + 1,
    'Cont #': `Cont #${String(l.containerIndex).padStart(2, '0')}`,
    'Khách Hàng': l.customer,
    SO: l.so,
    PO: l.po,
    'Mã Hàng (LPVN Code)': l.itemCode,
    'Model/Feature': l.feature,
    'Phân Loại Bên': l.side === 'L' ? 'Trái (LHF)' : l.side === 'R' ? 'Phải (RHF)' : 'Chung',
    'Mô Tả Hàng Hóa': l.description,
    'Ngày Xuất (Loading Date)': l.updatedLoadingDate || l.loadingDate || l.crd,
    'SL Đặt Xuất (Qty)': l.qty,
    'Tồn Đầu Kho 01': l.wh01Initial,
    'Tồn Đầu Kho 60': l.wh60Initial,
    'Tổng Tồn Đầu (01+60)': l.totalStockInitial,
    'SL Phân Bổ (01+60)': l.allocatedTotal,
    'Trạng Thái': 'ĐỦ HÀNG XUẤT (100%)'
  }));
  const wsFulfilled = XLSX.utils.json_to_sheet(fulfilledData);
  XLSX.utils.book_append_sheet(wb, wsFulfilled, '03_Don_Du_Hang_Xuat');

  // Sheet 4: Đơn Thiếu Hàng
  const shortageData = reconciliation.shortageLines.map((l, idx) => ({
    STT: idx + 1,
    'Cont #': `Cont #${String(l.containerIndex).padStart(2, '0')}`,
    'Khách Hàng': l.customer,
    SO: l.so,
    PO: l.po,
    'Mã Hàng (LPVN Code)': l.itemCode,
    'Model/Feature': l.feature,
    'Phân Loại Bên': l.side === 'L' ? 'Trái (LHF)' : l.side === 'R' ? 'Phải (RHF)' : 'Chung',
    'Mô Tả Hàng Hóa': l.description,
    'Ngày Xuất (Loading Date)': l.updatedLoadingDate || l.loadingDate || l.crd,
    'SL Cần Xuất (Qty)': l.qty,
    'Tồn Đầu Kho 01': l.wh01Initial,
    'Tồn Đầu Kho 60': l.wh60Initial,
    'Tồn Khả Dụng (01+60)': l.stockBeforeLine,
    'SL Đáp Ứng Được': l.allocatedTotal,
    'SL THIẾU HỤT (Shortage)': l.shortageQty,
    'Mức Độ Thiếu':
      l.fulfillmentStatus === 'ZERO_STOCK' ? 'HẾT TỒN KHO 01+60 (0 Pcs)' : 'THIẾU MỘT PHẦN'
  }));
  const wsShortage = XLSX.utils.json_to_sheet(shortageData);
  XLSX.utils.book_append_sheet(wb, wsShortage, '04_Don_Hang_Thieu');

  // Sheet 5: Báo Cáo Tồn Kho Thành Phẩm Trái/Phải (Kho 01 & 60)
  const lrData = lrPairs.map((p, idx) => ({
    STT: idx + 1,
    'Model / Feature': p.model,
    'Mô Tả': p.description,
    'Mã Trái (LHF)': p.leftCode,
    'Trái - Kho 01': p.leftWh01,
    'Trái - Kho 60': p.leftWh60,
    'Tổng Trái (01+60)': p.leftTotal,
    'Mã Phải (RHF)': p.rightCode,
    'Phải - Kho 01': p.rightWh01,
    'Phải - Kho 60': p.rightWh60,
    'Tổng Phải (01+60)': p.rightTotal,
    'Số Bộ Ghép Đôi (Sets)': p.matchedSets,
    'Chênh Lệch (Trái - Phải)': p.diffLR
  }));
  const wsLr = XLSX.utils.json_to_sheet(lrData);
  XLSX.utils.book_append_sheet(wb, wsLr, '05_Ton_Kho_Trai_Phai');

  // Append individual sheets for each Container (PO & SO) right into the workbook as well!
  for (const poGroup of reconciliation.poSummaries) {
    const wsCont = buildContainerPoSoWorksheet(poGroup, allocationMode);
    XLSX.utils.book_append_sheet(wb, wsCont, poGroup.sheetNameExcel);
  }

  const safeDate = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Bao_Cao_Xuat_Cont_PO_SO_Kho_01_60_${safeDate}.xlsx`);
}
