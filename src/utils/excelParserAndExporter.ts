import XLSXStyle from 'xlsx-js-style';
import {
  ContainerGroup,
  LeftRightPair,
  PlanDataset,
  ReconciledPoSummary,
  ShipmentLine,
  StockItem
} from '../types';
import {
  ReconciliationResult,
  buildLrPairsFromStock,
  isAccessoryModel
} from './reconciliationEngine';

const XLSX: any = (XLSXStyle as any).default || XLSXStyle;

// Known 3* Accessories from historical Data sheet so even standalone Stock list uploads know their model
const KNOWN_3_ACCESSORIES: Record<string, { feature: string; description: string }> = {
  '3456810161': { feature: 'Accessories', description: "MOD CONNECTOR BRK'T" },
  '3432250061': { feature: 'Accessories', description: 'Seat Mounting Bracket LHF' },
  '3432260061': { feature: 'Accessories', description: 'Seat Mounting Bracket RHF' }
};

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
    wb.SheetNames.find((n: string) => n.toLowerCase().includes('stock')) || wb.SheetNames[0];
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

    // Quy tắc: Ở kho 01 nếu đầu 3* có model là phụ kiện thì mới lấy, còn không sẽ bỏ qua
    if (wh === '01' && code.startsWith('3')) {
      const metaFeat =
        existingMetaMap?.get(code)?.feature ||
        KNOWN_3_ACCESSORIES[code]?.feature ||
        String(row[6] ?? row[7] ?? '').trim();
      if (!isAccessoryModel(metaFeat)) {
        continue;
      }
    }

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
    entry.batchList.push({ wh, batch, bin, date: cdate, qty });
    byCode.set(code, entry);
  }

  const stockItems: StockItem[] = [];
  for (const code of Array.from(byCode.keys()).sort()) {
    const d = byCode.get(code)!;
    const meta = existingMetaMap?.get(code) || KNOWN_3_ACCESSORIES[code];
    const isFg = code.startsWith('8') && code.length === 10;
    const feature = meta?.feature || (isFg ? code.slice(1, 5) : code.slice(0, 4));
    const description = meta?.description || `Item ${code}`;
    const side = detectSide(code, description);
    const category = isFg
      ? 'Finished Goods (Thành Phẩm Đầu 8)'
      : code.startsWith('1220')
      ? 'Side Assembly (Bán Thành Phẩm)'
      : isAccessoryModel(feature)
      ? 'Hardware & Accessories (Phụ Kiện)'
      : code.startsWith('34') || code.startsWith('36') || code.startsWith('38') || code.startsWith('39')
      ? 'Stamp Part (Linh Kiện Dập 34/38/39)'
      : 'Hardware & Accessories (Phụ Kiện)';

    stockItems.push({
      code,
      feature,
      description,
      customerCode: (meta as any)?.customerCode || '',
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

// ============================================================================
// EXCEL STYLING HELPERS (xlsx-js-style) FOR CLEAR VISUAL LAYOUT & BIN PICKING
// ============================================================================

const THIN_BORDER = {
  top: { style: 'thin', color: { rgb: 'CBD5E1' } },
  bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
  left: { style: 'thin', color: { rgb: 'CBD5E1' } },
  right: { style: 'thin', color: { rgb: 'CBD5E1' } }
};

const MEDIUM_BOX_BORDER = {
  top: { style: 'medium', color: { rgb: '475569' } },
  bottom: { style: 'medium', color: { rgb: '475569' } },
  left: { style: 'thin', color: { rgb: '94A3B8' } },
  right: { style: 'thin', color: { rgb: '94A3B8' } }
};

function setCellStyle(ws: any, r: number, c: number, style: any) {
  const ref = XLSX.utils.encode_cell({ r, c });
  if (!ws[ref]) {
    ws[ref] = { t: 's', v: '' };
  }
  ws[ref].s = style;
  if (typeof ws[ref].v === 'number' && style.numFmt) {
    ws[ref].z = style.numFmt;
  }
}

function styleRange(ws: any, sRow: number, sCol: number, eRow: number, eCol: number, style: any) {
  for (let r = sRow; r <= eRow; r++) {
    for (let c = sCol; c <= eCol; c++) {
      setCellStyle(ws, r, c, style);
    }
  }
}

/**
 * Helper to build a rich, color-coded, print-ready Container (PO & SO) worksheet
 * with Bin picking locations for warehouse staff.
 */
function buildContainerPoSoWorksheet(
  poGroup: ReconciledPoSummary,
  allocationMode: string
): any {
  const contLabel = `Cont #${String(poGroup.containerIndex).padStart(2, '0')}`;
  const isFulfilled = poGroup.status === 'FULFILLED';
  const totalCols = 20; // Columns A (0) through T (19)
  const lastColIdx = totalCols - 1;

  const rows: any[][] = [];
  const merges: any[] = [];
  const rowHeights: Array<{ hpt: number }> = [];

  // Row 0: Main Title Banner
  rows.push([
    `PHIẾU SOẠN ĐƠN & BÁO CÁO XUẤT CONTAINER — ${contLabel.toUpperCase()} | PO: ${poGroup.po} | SO: ${poGroup.so || 'N/A'}`
  ]);
  merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: lastColIdx } });
  rowHeights.push({ hpt: 32 });

  // Row 1: Subtitle Banner for Printing & Picking
  rows.push([
    `TÀI LIỆU IN ĐI SOẠN HÀNG TẠI KHO (KHO 60 & KHO 01) — KÈM VỊ TRÍ BIN (KỆ / Ô) LẤY HÀNG TRỰC QUAN`
  ]);
  merges.push({ s: { r: 1, c: 0 }, e: { r: 1, c: lastColIdx } });
  rowHeights.push({ hpt: 22 });

  // Rows 2..5: Structured Container Metadata Card Grid
  const infoRow1 = new Array(totalCols).fill('');
  infoRow1[0] = 'Số Cont / Block:';
  infoRow1[2] = contLabel;
  infoRow1[5] = 'Khách Hàng (Customer):';
  infoRow1[7] = poGroup.customer || 'N/A';
  infoRow1[11] = 'Ngày Đóng Cont:';
  infoRow1[13] = poGroup.loadingDate || 'Chưa cập nhật';
  infoRow1[16] = 'Trạng Thái Cont:';
  infoRow1[18] = isFulfilled
    ? 'ĐỦ HÀNG XUẤT CONT 100%'
    : `THIẾU HÀNG (${poGroup.shortageLinesCount}/${poGroup.totalLines} mã)`;
  rows.push(infoRow1);
  rowHeights.push({ hpt: 22 });

  const infoRow2 = new Array(totalCols).fill('');
  infoRow2[0] = 'Số PO (Purchase Order):';
  infoRow2[2] = poGroup.po;
  infoRow2[5] = 'Số SO (Sales Order):';
  infoRow2[7] = poGroup.so || poGroup.soList.join(', ') || 'N/A';
  infoRow2[11] = 'Tổng Số Mã Hàng:';
  infoRow2[13] = `${poGroup.totalLines} mã (${poGroup.fulfilledLinesCount} đủ / ${poGroup.shortageLinesCount} thiếu)`;
  infoRow2[16] = 'Tỷ Lệ Đáp Ứng:';
  infoRow2[18] = `${poGroup.fulfillmentRate.toFixed(1)}%`;
  rows.push(infoRow2);
  rowHeights.push({ hpt: 22 });

  const infoRow3 = new Array(totalCols).fill('');
  infoRow3[0] = 'Tổng SL Đặt Xuất (Pcs):';
  infoRow3[2] = poGroup.totalRequiredQty;
  infoRow3[5] = 'SL Đáp Ứng (Kho 60+01):';
  infoRow3[7] = poGroup.totalAllocatedQty;
  infoRow3[11] = 'Tổng SL Thiếu Hụt:';
  infoRow3[13] = poGroup.totalShortageQty;
  infoRow3[16] = 'Tổng Kiện / CBM:';
  infoRow3[18] = `${poGroup.totalPkgs.toLocaleString('en-US')} pkg | ${poGroup.totalCbm.toFixed(2)} CBM`;
  rows.push(infoRow3);
  rowHeights.push({ hpt: 22 });

  const infoRow4 = new Array(totalCols).fill('');
  infoRow4[0] = 'Phương Thức Tính Tồn:';
  infoRow4[2] =
    allocationMode === 'direct'
      ? 'Tính riêng từng Cont độc lập với Tồn Kho 01 + Kho 60'
      : 'Tính lũy kế FIFO theo thứ tự Cont xuất';
  infoRow4[11] = 'Tổng Giá Trị (USD):';
  infoRow4[13] = Number(poGroup.totalAmount.toFixed(2));
  infoRow4[16] = 'Ngày In Phiếu:';
  infoRow4[18] = new Date().toLocaleString('vi-VN');
  rows.push(infoRow4);
  rowHeights.push({ hpt: 22 });

  for (let r = 2; r <= 5; r++) {
    merges.push(
      { s: { r, c: 0 }, e: { r, c: 1 } },
      { s: { r, c: 2 }, e: { r, c: r === 5 ? 10 : 4 } },
      ...(r === 5
        ? []
        : [
            { s: { r, c: 5 }, e: { r, c: 6 } },
            { s: { r, c: 7 }, e: { r, c: 10 } }
          ]),
      { s: { r, c: 11 }, e: { r, c: 12 } },
      { s: { r, c: 13 }, e: { r, c: 15 } },
      { s: { r, c: 16 }, e: { r, c: 17 } },
      { s: { r, c: 18 }, e: { r, c: 19 } }
    );
  }

  // Row 6: Blank separator
  rows.push([]);
  rowHeights.push({ hpt: 10 });

  // Track section indices for styling after sheet creation
  const section1BannerRow = rows.length;
  let section1HeaderRow = -1;
  let section1DataStart = -1;
  let section1DataEnd = -1;
  let section1TotalRow = -1;

  if (poGroup.shortageLines.length > 0) {
    rows.push([
      `I. DANH SÁCH CÁC MÃ ĐANG THIẾU CỦA CONT (PO: ${poGroup.po} & SO: ${poGroup.so || 'N/A'}) — THIẾU ${poGroup.shortageLines.length} MÃ (TỔNG THIẾU: -${poGroup.totalShortageQty.toLocaleString('en-US')} PCS)`
    ]);
    merges.push({ s: { r: section1BannerRow, c: 0 }, e: { r: section1BannerRow, c: lastColIdx } });
    rowHeights.push({ hpt: 25 });

    section1HeaderRow = rows.length;
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
      'SL Lấy Kho 60',
      'SL Lấy Kho 01',
      'Tổng Đáp Ứng',
      'SỐ LƯỢNG THIẾU (Pcs)',
      'Tỷ Lệ (%)',
      'VỊ TRÍ BIN LẤY PHẦN TỒN HIỆN CÓ (KHO 60 / KHO 01)',
      'TOÀN BỘ VỊ TRÍ BIN CỦA MÃ',
      'Mức Độ Thiếu',
      'Ghi Chú Soạn Đơn'
    ]);
    rowHeights.push({ hpt: 28 });

    section1DataStart = rows.length;
    poGroup.shortageLines.forEach((l, idx) => {
      const binPick = l.pickingBinsText || 'Hết tồn kho (Không có Bin)';
      const binAll = l.allBinsText || (l.bins?.length ? l.bins.join(', ') : 'Không có Bin');
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
        l.allocatedWh60,
        l.allocatedWh01,
        l.allocatedTotal,
        l.shortageQty,
        Number(l.fulfillmentRate.toFixed(1)),
        binPick,
        binAll,
        l.fulfillmentStatus === 'ZERO_STOCK' ? 'HẾT TỒN KHO (0 Pcs)' : 'THIẾU MỘT PHẦN',
        `Cần bổ sung ${l.shortageQty} pcs`
      ]);
      const estLines = Math.max(1, Math.ceil(Math.max(binPick.length, binAll.length) / 36));
      rowHeights.push({ hpt: Math.max(22, Math.min(60, estLines * 16 + 6)) });
    });
    section1DataEnd = rows.length - 1;

    section1TotalRow = rows.length;
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
      poGroup.shortageLines.reduce((a, b) => a + b.allocatedWh60, 0),
      poGroup.shortageLines.reduce((a, b) => a + b.allocatedWh01, 0),
      poGroup.shortageLines.reduce((a, b) => a + b.allocatedTotal, 0),
      poGroup.totalShortageQty,
      '',
      '',
      '',
      `THIẾU ${poGroup.shortageLines.length} MÃ`,
      ''
    ]);
    rowHeights.push({ hpt: 24 });

    rows.push([]);
    rowHeights.push({ hpt: 10 });
  } else {
    rows.push([
      `I. TÌNH TRẠNG THIẾU HỤT: CONT (PO: ${poGroup.po} & SO: ${poGroup.so || 'N/A'}) ĐỦ HÀNG XUẤT 100% — KHÔNG THIẾU MÃ NÀO (SẴN SÀNG ĐI SOẠN ĐƠN THEO BIN)`
    ]);
    merges.push({ s: { r: section1BannerRow, c: 0 }, e: { r: section1BannerRow, c: lastColIdx } });
    rowHeights.push({ hpt: 25 });

    rows.push([]);
    rowHeights.push({ hpt: 10 });
  }

  // Section 2: Full Picking & Container Line List with Bin Locations
  const section2BannerRow = rows.length;
  rows.push([
    `II. BẢNG SOẠN ĐƠN CHI TIẾT TOÀN BỘ MÃ HÀNG TRONG CONT & VỊ TRÍ BIN LẤY HÀNG (PO: ${poGroup.po} & SO: ${poGroup.so || 'N/A'}) — ${poGroup.totalLines} MÃ`
  ]);
  merges.push({ s: { r: section2BannerRow, c: 0 }, e: { r: section2BannerRow, c: lastColIdx } });
  rowHeights.push({ hpt: 26 });

  const section2HeaderRow = rows.length;
  rows.push([
    'STT',
    'Số PO',
    'Số SO',
    'Mã Hàng (LPVN Code)',
    'Model',
    'Bên (Trái/Phải)',
    'Mô Tả Hàng Hóa',
    'Ngày Xuất Dự Kiến',
    'SL Cần Xuất (Qty)',
    'Số Kiện (#Pkg)',
    'Tồn Kho 01',
    'Tồn Kho 60',
    'Tồn Khả Dụng (01+60)',
    'Lấy Kho 60 (Pcs)',
    'Lấy Kho 01 (Pcs)',
    'SL Đáp Ứng',
    'SL Thiếu Hụt',
    'VỊ TRÍ BIN GỢI Ý SOẠN ĐƠN (KỆ / Ô & SỐ LƯỢNG LẤY)',
    'TOÀN BỘ VỊ TRÍ BIN ĐANG TỒN (KHO 60 & KHO 01)',
    'Trạng Thái / Tick Soạn'
  ]);
  rowHeights.push({ hpt: 30 });

  const section2DataStart = rows.length;
  poGroup.lines.forEach((l, idx) => {
    const binPick = l.pickingBinsText || (l.bins?.length ? l.bins.join(', ') : 'Không có Bin');
    const binAll = l.allBinsText || (l.bins?.length ? l.bins.join(', ') : 'Không có Bin');
    rows.push([
      idx + 1,
      l.po,
      l.so,
      l.itemCode,
      l.feature,
      l.side === 'L' ? 'Trái (LHF)' : l.side === 'R' ? 'Phải (RHF)' : 'Chung',
      l.description,
      l.updatedLoadingDate || l.loadingDate || l.crd,
      l.qty,
      l.pkgs,
      l.wh01Initial,
      l.wh60Initial,
      l.stockBeforeLine,
      l.allocatedWh60,
      l.allocatedWh01,
      l.allocatedTotal,
      l.shortageQty,
      binPick,
      binAll,
      l.shortageQty > 0 ? `THIẾU -${l.shortageQty} PCS` : '☐ ĐỦ HÀNG (100%)'
    ]);
    const estLines = Math.max(1, Math.ceil(Math.max(binPick.length, binAll.length) / 38));
    rowHeights.push({ hpt: Math.max(23, Math.min(68, estLines * 15 + 7)) });
  });
  const section2DataEnd = rows.length - 1;

  const section2TotalRow = rows.length;
  rows.push([
    '',
    '',
    '',
    'TỔNG CỘNG TOÀN CONT:',
    '',
    '',
    '',
    '',
    poGroup.totalRequiredQty,
    poGroup.totalPkgs,
    '',
    '',
    '',
    poGroup.lines.reduce((a, b) => a + b.allocatedWh60, 0),
    poGroup.lines.reduce((a, b) => a + b.allocatedWh01, 0),
    poGroup.totalAllocatedQty,
    poGroup.totalShortageQty,
    `Tổng thể tích: ${poGroup.totalCbm.toFixed(2)} CBM`,
    `Tổng giá trị: $${poGroup.totalAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
    isFulfilled ? 'ĐỦ 100%' : `THIẾU -${poGroup.totalShortageQty} PCS`
  ]);
  rowHeights.push({ hpt: 26 });

  // Blank row + Signature Block for Printed Picking Sheet
  rows.push([]);
  rowHeights.push({ hpt: 12 });

  const sigRow = rows.length;
  const sigLine = new Array(totalCols).fill('');
  sigLine[1] = 'NGƯỜI LẬP PHIẾU';
  sigLine[6] = 'NHÂN VIÊN SOẠN ĐƠN (LẤY HÀNG THEO BIN)';
  sigLine[12] = 'THỦ KHO XUẤT (KHO 01 / KHO 60)';
  sigLine[17] = 'GIÁM SÁT ĐÓNG CONT';
  rows.push(sigLine);
  rowHeights.push({ hpt: 22 });

  const sigNoteRow = rows.length;
  const sigNote = new Array(totalCols).fill('');
  sigNote[1] = '(Ký & ghi rõ họ tên)';
  sigNote[6] = '(Kiểm tra đúng mã, bên Trái/Phải & vị trí Bin)';
  sigNote[12] = '(Xác nhận trừ tồn kho thực tế)';
  sigNote[17] = '(Xác nhận hoàn thành đóng Cont)';
  rows.push(sigNote);
  rowHeights.push({ hpt: 18 });

  merges.push(
    { s: { r: sigRow, c: 1 }, e: { r: sigRow, c: 4 } },
    { s: { r: sigRow, c: 6 }, e: { r: sigRow, c: 10 } },
    { s: { r: sigRow, c: 12 }, e: { r: sigRow, c: 15 } },
    { s: { r: sigRow, c: 17 }, e: { r: sigRow, c: 19 } },
    { s: { r: sigNoteRow, c: 1 }, e: { r: sigNoteRow, c: 4 } },
    { s: { r: sigNoteRow, c: 6 }, e: { r: sigNoteRow, c: 10 } },
    { s: { r: sigNoteRow, c: 12 }, e: { r: sigNoteRow, c: 15 } },
    { s: { r: sigNoteRow, c: 17 }, e: { r: sigNoteRow, c: 19 } }
  );

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!merges'] = merges;
  ws['!rows'] = rowHeights;
  ws['!cols'] = [
    { wch: 6 },  // A: STT
    { wch: 15 }, // B: PO
    { wch: 15 }, // C: SO
    { wch: 16 }, // D: Mã Hàng
    { wch: 11 }, // E: Model
    { wch: 13 }, // F: Bên
    { wch: 28 }, // G: Mô Tả
    { wch: 13 }, // H: Ngày Xuất
    { wch: 14 }, // I: SL Cần Xuất
    { wch: 11 }, // J: Số Kiện
    { wch: 12 }, // K: Tồn Kho 01
    { wch: 12 }, // L: Tồn Kho 60
    { wch: 15 }, // M: Tồn Khả Dụng
    { wch: 13 }, // N: Lấy Kho 60
    { wch: 13 }, // O: Lấy Kho 01
    { wch: 13 }, // P: SL Đáp Ứng
    { wch: 13 }, // Q: SL Thiếu
    { wch: 38 }, // R: Vị Trí Bin Gợi Ý Soạn Đơn
    { wch: 38 }, // S: Toàn Bộ Vị Trí Bin Đang Tồn
    { wch: 20 }  // T: Trạng Thái / Tick Soạn
  ];

  // Page Setup for Landscape Printing
  ws['!pageSetup'] = {
    orientation: 'landscape',
    paperSize: 9, // A4
    fitToWidth: 1,
    fitToHeight: 0
  };

  // Apply Styles to Row 0 (Main Banner) & Row 1 (Subtitle Banner)
  styleRange(ws, 0, 0, 0, lastColIdx, {
    fill: { fgColor: { rgb: '0F172A' } },
    font: { name: 'Arial', sz: 13, bold: true, color: { rgb: 'FFFFFF' } },
    alignment: { horizontal: 'center', vertical: 'center' }
  });
  styleRange(ws, 1, 0, 1, lastColIdx, {
    fill: { fgColor: { rgb: '1E3A8A' } },
    font: { name: 'Arial', sz: 10, bold: true, color: { rgb: 'FDE047' } },
    alignment: { horizontal: 'center', vertical: 'center' }
  });

  // Apply Styles to Metadata Info Box (Rows 2..5)
  for (let r = 2; r <= 5; r++) {
    styleRange(ws, r, 0, r, lastColIdx, {
      fill: { fgColor: { rgb: 'F8FAFC' } },
      font: { name: 'Arial', sz: 10, bold: true, color: { rgb: '0F172A' } },
      border: THIN_BORDER,
      alignment: { horizontal: 'left', vertical: 'center' }
    });
    for (const labelCol of [0, 1, 5, 6, 11, 12, 16, 17]) {
      if (r === 5 && (labelCol === 5 || labelCol === 6)) continue;
      setCellStyle(ws, r, labelCol, {
        fill: { fgColor: { rgb: 'E2E8F0' } },
        font: { name: 'Arial', sz: 10, bold: true, color: { rgb: '1E293B' } },
        border: THIN_BORDER,
        alignment: { horizontal: 'left', vertical: 'center' }
      });
    }
  }
  // Highlight PO, SO, Status in Metadata Box
  setCellStyle(ws, 3, 2, {
    fill: { fgColor: { rgb: 'FEF3C7' } },
    font: { name: 'Arial', sz: 11, bold: true, color: { rgb: 'B45309' } },
    border: THIN_BORDER,
    alignment: { horizontal: 'left', vertical: 'center' }
  });
  setCellStyle(ws, 3, 7, {
    fill: { fgColor: { rgb: 'E0F2FE' } },
    font: { name: 'Arial', sz: 11, bold: true, color: { rgb: '0369A1' } },
    border: THIN_BORDER,
    alignment: { horizontal: 'left', vertical: 'center' }
  });
  setCellStyle(ws, 2, 18, {
    fill: { fgColor: { rgb: isFulfilled ? 'DCFCE7' : 'FEE2E2' } },
    font: { name: 'Arial', sz: 10, bold: true, color: { rgb: isFulfilled ? '15803D' : 'B91C1C' } },
    border: THIN_BORDER,
    alignment: { horizontal: 'center', vertical: 'center' }
  });
  setCellStyle(ws, 4, 2, {
    fill: { fgColor: { rgb: 'F1F5F9' } },
    font: { name: 'Arial', sz: 11, bold: true, color: { rgb: '0F172A' } },
    border: THIN_BORDER,
    alignment: { horizontal: 'right', vertical: 'center' },
    numFmt: '#,##0'
  });
  setCellStyle(ws, 4, 7, {
    fill: { fgColor: { rgb: 'DCFCE7' } },
    font: { name: 'Arial', sz: 11, bold: true, color: { rgb: '15803D' } },
    border: THIN_BORDER,
    alignment: { horizontal: 'right', vertical: 'center' },
    numFmt: '#,##0'
  });
  setCellStyle(ws, 4, 13, {
    fill: { fgColor: { rgb: poGroup.totalShortageQty > 0 ? 'FEE2E2' : 'F8FAFC' } },
    font: {
      name: 'Arial',
      sz: 11,
      bold: true,
      color: { rgb: poGroup.totalShortageQty > 0 ? 'B91C1C' : '64748B' }
    },
    border: THIN_BORDER,
    alignment: { horizontal: 'right', vertical: 'center' },
    numFmt: '#,##0'
  });

  // Style Section 1
  if (poGroup.shortageLines.length > 0) {
    styleRange(ws, section1BannerRow, 0, section1BannerRow, lastColIdx, {
      fill: { fgColor: { rgb: '991B1B' } },
      font: { name: 'Arial', sz: 11, bold: true, color: { rgb: 'FFFFFF' } },
      alignment: { horizontal: 'left', vertical: 'center' }
    });
    styleRange(ws, section1HeaderRow, 0, section1HeaderRow, lastColIdx, {
      fill: { fgColor: { rgb: 'B91C1C' } },
      font: { name: 'Arial', sz: 9, bold: true, color: { rgb: 'FFFFFF' } },
      border: MEDIUM_BOX_BORDER,
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
    });
    for (let r = section1DataStart; r <= section1DataEnd; r++) {
      for (let c = 0; c <= lastColIdx; c++) {
        const isNum = c >= 7 && c <= 15;
        const isShortCol = c === 14;
        const isBinCol = c === 16 || c === 17;
        setCellStyle(ws, r, c, {
          fill: { fgColor: { rgb: isBinCol ? 'FEF3C7' : isShortCol ? 'FEE2E2' : 'FFF1F2' } },
          font: {
            name: 'Arial',
            sz: 9,
            bold: c === 3 || isShortCol || isBinCol,
            color: { rgb: isShortCol ? 'B91C1C' : isBinCol ? '92400E' : '0F172A' }
          },
          border: THIN_BORDER,
          alignment: {
            horizontal: c === 0 || c === 4 || c === 5 || c === 18 ? 'center' : isNum ? 'right' : 'left',
            vertical: 'center',
            wrapText: isBinCol || c === 6
          },
          ...(isNum && c !== 15 ? { numFmt: '#,##0' } : {})
        });
      }
    }
    styleRange(ws, section1TotalRow, 0, section1TotalRow, lastColIdx, {
      fill: { fgColor: { rgb: 'FECACA' } },
      font: { name: 'Arial', sz: 10, bold: true, color: { rgb: '7F1D1D' } },
      border: MEDIUM_BOX_BORDER,
      alignment: { horizontal: 'right', vertical: 'center' },
      numFmt: '#,##0'
    });
    setCellStyle(ws, section1TotalRow, 3, {
      fill: { fgColor: { rgb: 'FECACA' } },
      font: { name: 'Arial', sz: 10, bold: true, color: { rgb: '7F1D1D' } },
      border: MEDIUM_BOX_BORDER,
      alignment: { horizontal: 'left', vertical: 'center' }
    });
  } else {
    styleRange(ws, section1BannerRow, 0, section1BannerRow, lastColIdx, {
      fill: { fgColor: { rgb: '065F46' } },
      font: { name: 'Arial', sz: 10, bold: true, color: { rgb: 'FFFFFF' } },
      alignment: { horizontal: 'left', vertical: 'center' }
    });
  }

  // Style Section 2 (Main Picking Table with Bin Locations)
  styleRange(ws, section2BannerRow, 0, section2BannerRow, lastColIdx, {
    fill: { fgColor: { rgb: '0F766E' } },
    font: { name: 'Arial', sz: 11, bold: true, color: { rgb: 'FFFFFF' } },
    alignment: { horizontal: 'left', vertical: 'center' }
  });

  for (let c = 0; c <= lastColIdx; c++) {
    const headerBg = c <= 9 ? '1E3A8A' : c <= 16 ? '0E7490' : 'B45309';
    setCellStyle(ws, section2HeaderRow, c, {
      fill: { fgColor: { rgb: headerBg } },
      font: { name: 'Arial', sz: 9, bold: true, color: { rgb: 'FFFFFF' } },
      border: MEDIUM_BOX_BORDER,
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
    });
  }

  for (let r = section2DataStart; r <= section2DataEnd; r++) {
    const lineIdx = r - section2DataStart;
    const lineObj = poGroup.lines[lineIdx];
    const isShort = lineObj && lineObj.shortageQty > 0;
    const zebraBg = isShort ? 'FFF1F2' : lineIdx % 2 === 0 ? 'FFFFFF' : 'F8FAFC';

    for (let c = 0; c <= lastColIdx; c++) {
      const isNum = c >= 8 && c <= 16;
      const isBinSuggest = c === 17;
      const isBinAll = c === 18;
      const isStatusCol = c === 19;
      const isSideCol = c === 5;

      let cellBg = zebraBg;
      let textColor = '0F172A';
      let isBold = c === 3 || c === 8;

      if (isSideCol && lineObj) {
        if (lineObj.side === 'L') {
          cellBg = 'E0F2FE';
          textColor = '0369A1';
          isBold = true;
        } else if (lineObj.side === 'R') {
          cellBg = 'D1FAE5';
          textColor = '047857';
          isBold = true;
        }
      } else if (isBinSuggest) {
        cellBg = 'FEF3C7';
        textColor = '92400E';
        isBold = true;
      } else if (isBinAll) {
        cellBg = 'FFFBEB';
        textColor = '78350F';
      } else if (isStatusCol) {
        cellBg = isShort ? 'FEE2E2' : 'DCFCE7';
        textColor = isShort ? 'B91C1C' : '15803D';
        isBold = true;
      } else if (c === 16 && isShort) {
        cellBg = 'FEE2E2';
        textColor = 'B91C1C';
        isBold = true;
      } else if (c === 13 || c === 14 || c === 15) {
        cellBg = 'F0FDFA';
        textColor = '0F766E';
        isBold = c === 15;
      }

      setCellStyle(ws, r, c, {
        fill: { fgColor: { rgb: cellBg } },
        font: { name: 'Arial', sz: 9, bold: isBold, color: { rgb: textColor } },
        border: THIN_BORDER,
        alignment: {
          horizontal:
            c === 0 || c === 4 || c === 5 || c === 7 || isStatusCol
              ? 'center'
              : isNum
              ? 'right'
              : 'left',
          vertical: 'center',
          wrapText: isBinSuggest || isBinAll || c === 6
        },
        ...(isNum ? { numFmt: '#,##0' } : {})
      });
    }
  }

  // Style Section 2 Total Row
  for (let c = 0; c <= lastColIdx; c++) {
    const isNum = c >= 8 && c <= 16;
    setCellStyle(ws, section2TotalRow, c, {
      fill: { fgColor: { rgb: 'FEF08A' } },
      font: { name: 'Arial', sz: 10, bold: true, color: { rgb: '0F172A' } },
      border: MEDIUM_BOX_BORDER,
      alignment: {
        horizontal: c === 3 || c === 17 || c === 18 ? 'left' : c === 19 ? 'center' : 'right',
        vertical: 'center'
      },
      ...(isNum ? { numFmt: '#,##0' } : {})
    });
  }

  // Style Signature Rows
  for (const colStart of [1, 6, 12, 17]) {
    setCellStyle(ws, sigRow, colStart, {
      font: { name: 'Arial', sz: 10, bold: true, color: { rgb: '1E293B' } },
      alignment: { horizontal: 'center', vertical: 'center' }
    });
    setCellStyle(ws, sigNoteRow, colStart, {
      font: { name: 'Arial', sz: 9, italic: true, color: { rgb: '475569' } },
      alignment: { horizontal: 'center', vertical: 'center' }
    });
  }

  return ws;
}

/**
 * Universal helper to apply clean header, zebra striping, status badge colors, and borders
 * to standard summary tables in Excel workbooks.
 */
function styleStandardDataSheet(
  ws: any,
  headerRowIdx: number,
  totalCols: number,
  totalRows: number,
  options?: {
    titleBanner?: boolean;
    bannerBg?: string;
    headerBg?: string;
    binColIndices?: number[];
    statusColIndices?: number[];
    shortageColIndices?: number[];
    numberColIndices?: number[];
  }
) {
  const bannerBg = options?.bannerBg || '0F172A';
  const headerBg = options?.headerBg || '1E3A8A';
  const binCols = new Set(options?.binColIndices || []);
  const statusCols = new Set(options?.statusColIndices || []);
  const shortageCols = new Set(options?.shortageColIndices || []);
  const numCols = new Set(options?.numberColIndices || []);

  if (options?.titleBanner && headerRowIdx > 0) {
    styleRange(ws, 0, 0, 0, totalCols - 1, {
      fill: { fgColor: { rgb: bannerBg } },
      font: { name: 'Arial', sz: 13, bold: true, color: { rgb: 'FFFFFF' } },
      alignment: { horizontal: 'left', vertical: 'center' }
    });
  }

  for (let c = 0; c < totalCols; c++) {
    const isBin = binCols.has(c);
    setCellStyle(ws, headerRowIdx, c, {
      fill: { fgColor: { rgb: isBin ? 'B45309' : headerBg } },
      font: { name: 'Arial', sz: 9, bold: true, color: { rgb: 'FFFFFF' } },
      border: MEDIUM_BOX_BORDER,
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
    });
  }

  for (let r = headerRowIdx + 1; r < totalRows; r++) {
    const zebra = (r - headerRowIdx) % 2 === 1 ? 'FFFFFF' : 'F8FAFC';
    for (let c = 0; c < totalCols; c++) {
      const ref = XLSX.utils.encode_cell({ r, c });
      const val = String(ws[ref]?.v ?? '');
      const isBin = binCols.has(c);
      const isStatus = statusCols.has(c);
      const isShort = shortageCols.has(c) && Number(ws[ref]?.v ?? 0) > 0;
      const isNum = numCols.has(c) || typeof ws[ref]?.v === 'number';

      let bg = zebra;
      let fg = '0F172A';
      let bold = false;

      if (isBin) {
        bg = 'FEF3C7';
        fg = '92400E';
        bold = true;
      } else if (isStatus) {
        if (val.includes('THIẾU') || val.includes('HẾT') || val.includes('LỆCH')) {
          bg = 'FEE2E2';
          fg = 'B91C1C';
          bold = true;
        } else if (val.includes('ĐỦ') || val.includes('CÂN BẰNG')) {
          bg = 'DCFCE7';
          fg = '15803D';
          bold = true;
        }
      } else if (isShort) {
        bg = 'FEE2E2';
        fg = 'B91C1C';
        bold = true;
      }

      setCellStyle(ws, r, c, {
        fill: { fgColor: { rgb: bg } },
        font: { name: 'Arial', sz: 9, bold, color: { rgb: fg } },
        border: THIN_BORDER,
        alignment: {
          horizontal: isStatus || c === 0 ? 'center' : isNum ? 'right' : 'left',
          vertical: 'center',
          wrapText: isBin
        },
        ...(isNum && c !== 0 ? { numFmt: '#,##0' } : {})
      });
    }
  }

  ws['!pageSetup'] = {
    orientation: 'landscape',
    paperSize: 9,
    fitToWidth: 1,
    fitToHeight: 0
  };
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
  wsSummary['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 16 } }];
  wsSummary['!rows'] = [{ hpt: 30 }, { hpt: 20 }, { hpt: 10 }, { hpt: 28 }];
  wsSummary['!cols'] = [
    { wch: 6 },
    { wch: 11 },
    { wch: 18 },
    { wch: 16 },
    { wch: 24 },
    { wch: 20 },
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
    { wch: 80 }
  ];
  styleStandardDataSheet(wsSummary, 3, 17, summaryRows.length, {
    titleBanner: true,
    bannerBg: '0F172A',
    headerBg: '1E3A8A',
    statusColIndices: [15],
    shortageColIndices: [9, 12],
    numberColIndices: [7, 8, 9, 10, 11, 12, 13]
  });
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
  const stockByCode = new Map(stockItems.map((s) => [s.code.trim(), s]));

  // Sheet 1: Executive KPI & Multi-Agent Audit Summary
  const kpiRows = [
    ['BÁO CÁO ĐỐI SOÁT TỒN KHO (KHO 01 & 60) & KẾ HOẠCH XUẤT HÀNG CONTAINER'],
    ['Thời gian xuất báo cáo:', new Date().toLocaleString('vi-VN'), ''],
    ['Kế hoạch đang đối soát:', activePlanName, ''],
    [
      'Chế độ phân bổ tồn kho:',
      allocationMode === 'fifo'
        ? 'Phân bổ lũy kế FIFO (Kho 60 -> Kho 01 theo thứ tự Cont/Lịch xuất)'
        : 'Tính riêng độc lập từng Cont (SO & PO) với Tổng Tồn (Kho 01 + Kho 60)',
      ''
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
      `${stockItems.filter((x) => x.wh01Qty > 0).length} mã hàng tại Kho 01 (Đầu 3* chỉ lấy Phụ kiện)`
    ],
    [
      'Tồn Kho 60',
      stockItems.reduce((a, b) => a + b.wh60Qty, 0),
      `${stockItems.filter((x) => x.wh60Qty > 0).length} mã hàng tại Kho 60`
    ],
    [
      'Tồn Kho Thành Phẩm (Đầu 8 - Trái/Phải)',
      stockItems.filter((x) => x.isFinishedGood).reduce((a, b) => a + b.totalQty, 0),
      `${lrPairs.filter((x) => x.isFinishedGood).length} cặp mẫu (${lrPairs.filter((x) => x.isFinishedGood && x.isBalanced).length} cân bằng, ${lrPairs.filter((x) => x.isFinishedGood && !x.isBalanced).length} lệch cặp)`
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
  wsKpi['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }];
  wsKpi['!cols'] = [{ wch: 44 }, { wch: 24 }, { wch: 58 }];
  styleStandardDataSheet(wsKpi, 5, 3, kpiRows.length, {
    titleBanner: true,
    bannerBg: '0F172A',
    headerBg: '0F766E',
    numberColIndices: [1]
  });
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
  wsPo['!cols'] = [
    { wch: 6 },
    { wch: 11 },
    { wch: 16 },
    { wch: 15 },
    { wch: 22 },
    { wch: 18 },
    { wch: 14 },
    { wch: 12 },
    { wch: 11 },
    { wch: 12 },
    { wch: 16 },
    { wch: 16 },
    { wch: 14 },
    { wch: 48 },
    { wch: 20 }
  ];
  styleStandardDataSheet(wsPo, 0, 15, poData.length + 1, {
    headerBg: '1E3A8A',
    statusColIndices: [14],
    shortageColIndices: [9, 12],
    numberColIndices: [7, 8, 9, 10, 11, 12]
  });
  XLSX.utils.book_append_sheet(wb, wsPo, '02_Tong_Hop_Tung_Cont');

  // Sheet 3: Đơn Đủ Hàng Xuất (Kèm Vị Trí Bin Lấy Hàng)
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
    'Lấy Tại Kho 60': l.allocatedWh60,
    'Lấy Tại Kho 01': l.allocatedWh01,
    'SL Phân Bổ (01+60)': l.allocatedTotal,
    'Vị Trí Bin Gợi Ý Lấy Hàng (Kho 60 / Kho 01)':
      l.pickingBinsText || (l.bins?.length ? l.bins.join(', ') : 'Không có Bin'),
    'Toàn Bộ Vị Trí Bin Đang Tồn':
      l.allBinsText || (l.bins?.length ? l.bins.join(', ') : 'Không có Bin'),
    'Trạng Thái': 'ĐỦ HÀNG XUẤT (100%)'
  }));
  const wsFulfilled = XLSX.utils.json_to_sheet(fulfilledData);
  wsFulfilled['!cols'] = [
    { wch: 6 },
    { wch: 10 },
    { wch: 16 },
    { wch: 14 },
    { wch: 15 },
    { wch: 16 },
    { wch: 12 },
    { wch: 13 },
    { wch: 26 },
    { wch: 13 },
    { wch: 14 },
    { wch: 13 },
    { wch: 13 },
    { wch: 13 },
    { wch: 13 },
    { wch: 14 },
    { wch: 36 },
    { wch: 36 },
    { wch: 20 }
  ];
  styleStandardDataSheet(wsFulfilled, 0, 19, fulfilledData.length + 1, {
    headerBg: '065F46',
    binColIndices: [16, 17],
    statusColIndices: [18],
    numberColIndices: [10, 11, 12, 13, 14, 15]
  });
  XLSX.utils.book_append_sheet(wb, wsFulfilled, '03_Don_Du_Hang_Xuat');

  // Sheet 4: Đơn Thiếu Hàng (Kèm Vị Trí Bin Tồn Hiện Có)
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
    'Vị Trí Bin Phần Tồn Hiện Có':
      l.pickingBinsText || l.allBinsText || 'Hết tồn kho (Không có Bin)',
    'Mức Độ Thiếu':
      l.fulfillmentStatus === 'ZERO_STOCK' ? 'HẾT TỒN KHO 01+60 (0 Pcs)' : 'THIẾU MỘT PHẦN'
  }));
  const wsShortage = XLSX.utils.json_to_sheet(shortageData);
  wsShortage['!cols'] = [
    { wch: 6 },
    { wch: 10 },
    { wch: 16 },
    { wch: 14 },
    { wch: 15 },
    { wch: 16 },
    { wch: 12 },
    { wch: 13 },
    { wch: 26 },
    { wch: 13 },
    { wch: 14 },
    { wch: 13 },
    { wch: 13 },
    { wch: 15 },
    { wch: 14 },
    { wch: 16 },
    { wch: 36 },
    { wch: 22 }
  ];
  styleStandardDataSheet(wsShortage, 0, 18, shortageData.length + 1, {
    headerBg: '991B1B',
    binColIndices: [16],
    statusColIndices: [17],
    shortageColIndices: [15],
    numberColIndices: [10, 11, 12, 13, 14, 15]
  });
  XLSX.utils.book_append_sheet(wb, wsShortage, '04_Don_Hang_Thieu');

  // Sheet 5: Báo Cáo Tồn Kho Thành Phẩm Trái/Phải (Kho 01 & 60) Kèm Vị Trí Bin
  const lrData = lrPairs.map((p, idx) => {
    const leftBins = stockByCode.get(p.leftCode)?.bins?.join(', ') || '';
    const rightBins = stockByCode.get(p.rightCode)?.bins?.join(', ') || '';
    return {
      STT: idx + 1,
      'Model / Feature': p.model,
      'Mô Tả': p.description,
      'Mã Trái (LHF)': p.leftCode,
      'Trái - Kho 01': p.leftWh01,
      'Trái - Kho 60': p.leftWh60,
      'Tổng Trái (01+60)': p.leftTotal,
      'Vị Trí Bin Bên Trái (LHF)': leftBins,
      'Mã Phải (RHF)': p.rightCode,
      'Phải - Kho 01': p.rightWh01,
      'Phải - Kho 60': p.rightWh60,
      'Tổng Phải (01+60)': p.rightTotal,
      'Vị Trí Bin Bên Phải (RHF)': rightBins,
      'Số Bộ Ghép Đôi (Sets)': p.matchedSets,
      'Chênh Lệch (Trái - Phải)': p.diffLR,
      'Trạng Thái Cặp': p.isBalanced ? 'CÂN BẰNG 100%' : `LỆCH CẶP (${p.diffLR > 0 ? '+' : ''}${p.diffLR})`
    };
  });
  const wsLr = XLSX.utils.json_to_sheet(lrData);
  wsLr['!cols'] = [
    { wch: 6 },
    { wch: 12 },
    { wch: 24 },
    { wch: 15 },
    { wch: 12 },
    { wch: 12 },
    { wch: 14 },
    { wch: 32 },
    { wch: 15 },
    { wch: 12 },
    { wch: 12 },
    { wch: 14 },
    { wch: 32 },
    { wch: 14 },
    { wch: 15 },
    { wch: 18 }
  ];
  styleStandardDataSheet(wsLr, 0, 16, lrData.length + 1, {
    headerBg: '1E3A8A',
    binColIndices: [7, 12],
    statusColIndices: [15],
    numberColIndices: [4, 5, 6, 9, 10, 11, 13, 14]
  });
  XLSX.utils.book_append_sheet(wb, wsLr, '05_Ton_Kho_Trai_Phai');

  // Append individual sheets for each Container (PO & SO) right into the workbook as well!
  for (const poGroup of reconciliation.poSummaries) {
    const wsCont = buildContainerPoSoWorksheet(poGroup, allocationMode);
    XLSX.utils.book_append_sheet(wb, wsCont, poGroup.sheetNameExcel);
  }

  const safeDate = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Bao_Cao_Xuat_Cont_PO_SO_Kho_01_60_${safeDate}.xlsx`);
}
