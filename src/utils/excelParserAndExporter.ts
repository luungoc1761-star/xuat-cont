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
        const rawPkg = Number(r[16] ?? 0) || 0;
        const poVal = String(r[4] ?? '').trim();
        const featVal = String(r[7] ?? '').trim();
        const desc = String(r[8] ?? '').trim();
        const sideVal = detectSide(code, desc);

        let isPairedSecond = false;
        let pairPkg = rawPkg;
        let pkgsVal = rawPkg;

        // Số #pkg trong file xuất hàng là số kiện tính theo cặp Trái/Phải -> giữ nguyên theo cặp, không cộng riêng Trái + Phải
        const prevLine = currentLines[currentLines.length - 1];
        if (
          prevLine &&
          (prevLine.side === 'L' || prevLine.side === 'R') &&
          (sideVal === 'L' || sideVal === 'R') &&
          prevLine.side !== sideVal &&
          prevLine.po === poVal &&
          prevLine.feature === featVal &&
          !prevLine.isPairedPkgHead &&
          !prevLine.isPairedPkgSecondRow
        ) {
          pairPkg = prevLine.pkgs > 0 ? prevLine.pkgs : rawPkg;
          prevLine.pkgs = pairPkg;
          prevLine.pairPkgs = pairPkg;
          prevLine.isPairedPkgHead = true;
          pkgsVal = 0;
          isPairedSecond = true;
        }

        const lineObj: ShipmentLine = {
          id: `${sname}-R${i + 1}`,
          sheet: sname,
          rowNumber: i + 1,
          containerIndex: contIdx,
          customer: String(r[0] ?? '').trim(),
          type: String(r[1] ?? '').trim(),
          pkl: String(r[2] ?? '').trim(),
          so: String(r[3] ?? '').trim(),
          po: poVal,
          itemCode: code,
          customerItemCode: String(r[6] ?? '').trim(),
          feature: featVal,
          description: desc,
          dimension: String(r[9] ?? '').trim(),
          crd: formatExcelDate(r[10]),
          loadingDate: formatExcelDate(r[11]),
          updatedLoadingDate: formatExcelDate(r[12]),
          qty,
          unitWeight: Number(r[14] ?? 0) || 0,
          pcsPerPkg: Number(r[15] ?? 0) || 0,
          pkgs: pkgsVal,
          pairPkgs: pairPkg,
          isPairedPkgHead: false,
          isPairedPkgSecondRow: isPairedSecond,
          netWeight: Number(r[17] ?? 0) || 0,
          grossWeight: Number(r[18] ?? 0) || 0,
          cbm: Number(r[24] ?? 0) || 0,
          unitPrice: Number(r[26] ?? 0) || 0,
          totalAmount: Number(r[27] ?? 0) || 0,
          side: sideVal
        };
        lines.push(lineObj);
        currentLines.push(lineObj);
      } else if (currentLines.length > 0 && (r[13] || r[16] || r[17] || r[24] || r[27])) {
        const pos = Array.from(new Set(currentLines.map((x) => x.po).filter(Boolean)));
        const sos = Array.from(new Set(currentLines.map((x) => x.so).filter(Boolean)));
        const custs = Array.from(new Set(currentLines.map((x) => x.customer).filter(Boolean)));
        const pairBasedPkgs = currentLines.reduce((a, b) => a + b.pkgs, 0);
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
          totalPkgs: pairBasedPkgs || Number(r[16] ?? 0),
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
 * Helper to build a clean, easy-to-read, print-ready Container (PO & SO) worksheet.
 * - Minimalist 3-tone palette (Navy header, Clean White/Slate rows, Soft Red only for shortages)
 * - Essential columns only (no redundant metadata or duplicate tables)
 * - Merges "Số Kiện (#Pkg - Cặp)" vertically across Left/Right pairs so #pkg is kept by pair
 */
function buildContainerPoSoWorksheet(
  poGroup: ReconciledPoSummary,
  allocationMode: string
): any {
  const contLabel = `Cont #${String(poGroup.containerIndex).padStart(2, '0')}`;
  const isFulfilled = poGroup.status === 'FULFILLED';
  const totalCols = 13; // Columns A (0) through M (12)
  const lastColIdx = totalCols - 1;

  const rows: any[][] = [];
  const merges: any[] = [];
  const rowHeights: Array<{ hpt: number }> = [];

  // Row 0: Clean Title Banner
  rows.push([
    `PHIẾU SOẠN HÀNG & BÁO CÁO XUẤT CONTAINER — ${contLabel.toUpperCase()} | PO: ${poGroup.po} | SO: ${poGroup.so || poGroup.soList.join(', ') || 'N/A'}`
  ]);
  merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: lastColIdx } });
  rowHeights.push({ hpt: 30 });

  // Row 1: Compact Key Info Line 1
  const infoRow1 = new Array(totalCols).fill('');
  infoRow1[0] = `Khách hàng: ${poGroup.customer || 'N/A'}   |   Ngày đóng Cont: ${poGroup.loadingDate || 'Chưa cập nhật'}   |   Số mã hàng: ${poGroup.totalLines} mã (${poGroup.fulfilledLinesCount} đủ / ${poGroup.shortageLinesCount} thiếu)`;
  infoRow1[8] = `Tổng số kiện (theo cặp Trái/Phải): ${poGroup.totalPkgs.toLocaleString('en-US')} kiện   |   Thể tích: ${poGroup.totalCbm.toFixed(2)} CBM`;
  rows.push(infoRow1);
  merges.push(
    { s: { r: 1, c: 0 }, e: { r: 1, c: 7 } },
    { s: { r: 1, c: 8 }, e: { r: 1, c: lastColIdx } }
  );
  rowHeights.push({ hpt: 22 });

  // Row 2: Compact Key Info Line 2 (Quantities & Status Summary)
  const infoRow2 = new Array(totalCols).fill('');
  const modeLabel =
    allocationMode === 'direct'
      ? 'Tính độc lập từng Cont (Kho 01 + Kho 60)'
      : 'Tính lũy kế FIFO theo thứ tự Cont';
  infoRow2[0] = `Tổng SL xuất: ${poGroup.totalRequiredQty.toLocaleString('en-US')} pcs   |   Đáp ứng: ${poGroup.totalAllocatedQty.toLocaleString('en-US')} pcs (${poGroup.fulfillmentRate.toFixed(1)}%)   |   SL thiếu: ${poGroup.totalShortageQty.toLocaleString('en-US')} pcs   |   Chế độ: ${modeLabel}`;
  infoRow2[8] = isFulfilled
    ? 'TRẠNG THÁI: ĐỦ HÀNG XUẤT 100%'
    : `TRẠNG THÁI: THIẾU ${poGroup.shortageLinesCount} MÃ (-${poGroup.totalShortageQty.toLocaleString('en-US')} pcs): ${poGroup.shortageLines.map((x) => `${x.itemCode} (-${x.shortageQty})`).join(', ')}`;
  rows.push(infoRow2);
  merges.push(
    { s: { r: 2, c: 0 }, e: { r: 2, c: 7 } },
    { s: { r: 2, c: 8 }, e: { r: 2, c: lastColIdx } }
  );
  rowHeights.push({ hpt: 22 });

  // Row 3: Blank separator
  rows.push([]);
  rowHeights.push({ hpt: 8 });

  // Row 4: Unified Table Header (13 Essential Columns)
  const headerRowIdx = rows.length;
  rows.push([
    'STT',
    'Mã Hàng (Item Code)',
    'Model',
    'Bên (Trái/Phải)',
    'Mô Tả Hàng Hóa',
    'SL Xuất (Pcs)',
    'Số Kiện (#Pkg - Cặp)',
    'Tồn Kho 01',
    'Tồn Kho 60',
    'Tổng Tồn (01+60)',
    'SL Thiếu (Pcs)',
    'Trạng Thái',
    'Vị Trí Bin Lấy Hàng (Kho 60 / Kho 01)'
  ]);
  rowHeights.push({ hpt: 28 });

  // Rows 5..N: Container Lines (with vertical merge on Col 6 for Left/Right pairs)
  const dataStartRow = rows.length;
  poGroup.lines.forEach((l, idx) => {
    const rIdx = rows.length;
    const binPick =
      l.pickingBinsText && !l.pickingBinsText.startsWith('Hết tồn')
        ? l.pickingBinsText
        : l.allBinsText || (l.bins?.length ? l.bins.join(', ') : 'Không có Bin');

    // If this row is the first row of a Left/Right pair, merge Col 6 (Số Kiện) with the next row
    const pkgDisplay = l.isPairedPkgSecondRow ? '' : l.pairPkgs ?? l.pkgs;
    if (l.isPairedPkgHead && idx + 1 < poGroup.lines.length) {
      merges.push({ s: { r: rIdx, c: 6 }, e: { r: rIdx + 1, c: 6 } });
    }

    rows.push([
      idx + 1,
      l.itemCode,
      l.feature,
      l.side === 'L' ? 'Trái (L)' : l.side === 'R' ? 'Phải (R)' : 'Chung',
      l.description,
      l.qty,
      pkgDisplay,
      l.wh01Initial,
      l.wh60Initial,
      l.stockBeforeLine,
      l.shortageQty > 0 ? l.shortageQty : 0,
      l.shortageQty > 0 ? `THIẾU -${l.shortageQty.toLocaleString('en-US')}` : 'Đủ hàng',
      binPick
    ]);

    const estLines = Math.max(1, Math.ceil(binPick.length / 46));
    rowHeights.push({ hpt: Math.max(22, Math.min(54, estLines * 15 + 6)) });
  });
  const dataEndRow = rows.length - 1;

  // Total Row
  const totalRowIdx = rows.length;
  rows.push([
    '',
    'TỔNG CỘNG CONTAINER',
    '',
    '',
    `${poGroup.totalLines} mã hàng`,
    poGroup.totalRequiredQty,
    poGroup.totalPkgs,
    '',
    '',
    '',
    poGroup.totalShortageQty,
    isFulfilled ? 'ĐỦ 100%' : `THIẾU -${poGroup.totalShortageQty.toLocaleString('en-US')}`,
    `Tổng thể tích: ${poGroup.totalCbm.toFixed(2)} CBM   |   Trị giá: $${poGroup.totalAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
  ]);
  merges.push({ s: { r: totalRowIdx, c: 1 }, e: { r: totalRowIdx, c: 3 } });
  rowHeights.push({ hpt: 25 });

  // Blank row + Clean Signature Footer
  rows.push([]);
  rowHeights.push({ hpt: 12 });

  const sigRow = rows.length;
  const sigLine = new Array(totalCols).fill('');
  sigLine[1] = 'NGƯỜI LẬP PHIẾU';
  sigLine[5] = 'NHÂN VIÊN SOẠN HÀNG (THEO BIN)';
  sigLine[9] = 'THỦ KHO XUẤT (KHO 01 / 60)';
  sigLine[12] = 'GIÁM SÁT ĐÓNG CONT';
  rows.push(sigLine);
  rowHeights.push({ hpt: 20 });

  merges.push(
    { s: { r: sigRow, c: 1 }, e: { r: sigRow, c: 3 } },
    { s: { r: sigRow, c: 5 }, e: { r: sigRow, c: 7 } },
    { s: { r: sigRow, c: 9 }, e: { r: sigRow, c: 11 } }
  );

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!merges'] = merges;
  ws['!rows'] = rowHeights;
  ws['!cols'] = [
    { wch: 6 },  // A: STT
    { wch: 15 }, // B: Mã Hàng
    { wch: 10 }, // C: Model
    { wch: 12 }, // D: Bên (Trái/Phải)
    { wch: 28 }, // E: Mô Tả Hàng Hóa
    { wch: 13 }, // F: SL Xuất (Pcs)
    { wch: 15 }, // G: Số Kiện (#Pkg - Cặp)
    { wch: 12 }, // H: Tồn Kho 01
    { wch: 12 }, // I: Tồn Kho 60
    { wch: 14 }, // J: Tổng Tồn (01+60)
    { wch: 13 }, // K: SL Thiếu (Pcs)
    { wch: 16 }, // L: Trạng Thái
    { wch: 46 }  // M: Vị Trí Bin Lấy Hàng
  ];

  ws['!pageSetup'] = {
    orientation: 'landscape',
    paperSize: 9,
    fitToWidth: 1,
    fitToHeight: 0
  };

  // 1. Style Title Banner (Row 0) - Clean Navy
  styleRange(ws, 0, 0, 0, lastColIdx, {
    fill: { fgColor: { rgb: '1E3A8A' } },
    font: { name: 'Arial', sz: 12, bold: true, color: { rgb: 'FFFFFF' } },
    alignment: { horizontal: 'left', vertical: 'center' }
  });

  // 2. Style Compact Summary Rows (Rows 1..2) - Subtle Slate
  for (const r of [1, 2]) {
    styleRange(ws, r, 0, r, lastColIdx, {
      fill: { fgColor: { rgb: 'F8FAFC' } },
      font: { name: 'Arial', sz: 9.5, bold: false, color: { rgb: '1E293B' } },
      border: THIN_BORDER,
      alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
    });
  }
  setCellStyle(ws, 1, 8, {
    fill: { fgColor: { rgb: 'F8FAFC' } },
    font: { name: 'Arial', sz: 9.5, bold: true, color: { rgb: '0F172A' } },
    border: THIN_BORDER,
    alignment: { horizontal: 'left', vertical: 'center' }
  });
  setCellStyle(ws, 2, 8, {
    fill: { fgColor: { rgb: isFulfilled ? 'F0FDF4' : 'FEF2F2' } },
    font: {
      name: 'Arial',
      sz: 9.5,
      bold: true,
      color: { rgb: isFulfilled ? '15803D' : 'B91C1C' }
    },
    border: THIN_BORDER,
    alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
  });

  // 3. Style Table Header (Row 4) - Single Uniform Navy
  for (let c = 0; c <= lastColIdx; c++) {
    setCellStyle(ws, headerRowIdx, c, {
      fill: { fgColor: { rgb: '1E3A8A' } },
      font: { name: 'Arial', sz: 9.5, bold: true, color: { rgb: 'FFFFFF' } },
      border: MEDIUM_BOX_BORDER,
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
    });
  }

  // 4. Style Data Rows - Clean White/Slate for Normal, Soft Rose Only for Shortage
  for (let r = dataStartRow; r <= dataEndRow; r++) {
    const lineIdx = r - dataStartRow;
    const lineObj = poGroup.lines[lineIdx];
    const isShort = lineObj && lineObj.shortageQty > 0;
    const rowBg = isShort ? 'FFF1F2' : lineIdx % 2 === 0 ? 'FFFFFF' : 'F8FAFC';

    for (let c = 0; c <= lastColIdx; c++) {
      const isNum = c >= 5 && c <= 10;
      const isPkgCol = c === 6;
      const isShortCol = c === 10;
      const isStatusCol = c === 11;
      const isSideCol = c === 3;

      let cellBg = rowBg;
      let textColor = '0F172A';
      let isBold = c === 1 || c === 5 || isPkgCol;

      if (isPkgCol) {
        // Keep paired #pkg cell clean and clearly readable across the L/R pair
        cellBg = isShort ? 'FFF1F2' : 'F8FAFC';
        textColor = '0F172A';
      } else if (isSideCol && lineObj) {
        textColor = lineObj.side === 'L' ? '0369A1' : lineObj.side === 'R' ? '047857' : '475569';
        isBold = true;
      } else if (isShortCol) {
        if (isShort) {
          cellBg = 'FEE2E2';
          textColor = 'B91C1C';
          isBold = true;
        } else {
          textColor = '94A3B8';
        }
      } else if (isStatusCol) {
        if (isShort) {
          cellBg = 'FEE2E2';
          textColor = 'B91C1C';
          isBold = true;
        } else {
          textColor = '15803D';
          isBold = true;
        }
      }

      setCellStyle(ws, r, c, {
        fill: { fgColor: { rgb: cellBg } },
        font: { name: 'Arial', sz: 9.5, bold: isBold, color: { rgb: textColor } },
        border: THIN_BORDER,
        alignment: {
          horizontal:
            c === 0 || c === 2 || isSideCol || isPkgCol || isStatusCol
              ? 'center'
              : isNum
              ? 'right'
              : 'left',
          vertical: 'center',
          wrapText: c === 4 || c === 12
        },
        ...(isNum ? { numFmt: '#,##0' } : {})
      });
    }
  }

  // 5. Style Total Row - Clean Soft Slate (No bright neon yellow)
  for (let c = 0; c <= lastColIdx; c++) {
    const isNum = c === 5 || c === 6 || c === 10;
    const isShortTotal = (c === 10 || c === 11) && !isFulfilled;
    setCellStyle(ws, totalRowIdx, c, {
      fill: { fgColor: { rgb: isShortTotal ? 'FEE2E2' : 'E2E8F0' } },
      font: {
        name: 'Arial',
        sz: 10,
        bold: true,
        color: { rgb: isShortTotal ? 'B91C1C' : '0F172A' }
      },
      border: MEDIUM_BOX_BORDER,
      alignment: {
        horizontal:
          c === 6 || c === 11 ? 'center' : isNum ? 'right' : 'left',
        vertical: 'center'
      },
      ...(isNum ? { numFmt: '#,##0' } : {})
    });
  }

  // 6. Style Signature Row
  for (const colStart of [1, 5, 9, 12]) {
    setCellStyle(ws, sigRow, colStart, {
      font: { name: 'Arial', sz: 9.5, bold: true, color: { rgb: '334155' } },
      alignment: { horizontal: 'center', vertical: 'center' }
    });
  }

  return ws;
}

/**
 * Universal helper to apply a clean, minimalist style to summary worksheets.
 * Highlights only headers and actual shortage/imbalance alerts.
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
  const bannerBg = options?.bannerBg || '1E3A8A';
  const headerBg = options?.headerBg || '1E3A8A';
  const binCols = new Set(options?.binColIndices || []);
  const statusCols = new Set(options?.statusColIndices || []);
  const shortageCols = new Set(options?.shortageColIndices || []);
  const numCols = new Set(options?.numberColIndices || []);

  if (options?.titleBanner && headerRowIdx > 0) {
    styleRange(ws, 0, 0, 0, totalCols - 1, {
      fill: { fgColor: { rgb: bannerBg } },
      font: { name: 'Arial', sz: 12, bold: true, color: { rgb: 'FFFFFF' } },
      alignment: { horizontal: 'left', vertical: 'center' }
    });
  }

  for (let c = 0; c < totalCols; c++) {
    setCellStyle(ws, headerRowIdx, c, {
      fill: { fgColor: { rgb: headerBg } },
      font: { name: 'Arial', sz: 9.5, bold: true, color: { rgb: 'FFFFFF' } },
      border: MEDIUM_BOX_BORDER,
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
    });
  }

  for (let r = headerRowIdx + 1; r < totalRows; r++) {
    // Check if this row has a shortage status so we can give it a subtle rose tint
    let rowHasShortage = false;
    for (const sc of statusCols) {
      const sRef = XLSX.utils.encode_cell({ r, c: sc });
      const sVal = String(ws[sRef]?.v ?? '');
      if (sVal.includes('THIẾU') || sVal.includes('HẾT') || sVal.includes('LỆCH')) {
        rowHasShortage = true;
      }
    }
    const zebra = rowHasShortage
      ? 'FFF1F2'
      : (r - headerRowIdx) % 2 === 1
      ? 'FFFFFF'
      : 'F8FAFC';

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

      if (isStatus) {
        if (val.includes('THIẾU') || val.includes('HẾT') || val.includes('LỆCH')) {
          bg = 'FEE2E2';
          fg = 'B91C1C';
          bold = true;
        } else if (val.includes('ĐỦ') || val.includes('CÂN BẰNG')) {
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
        font: { name: 'Arial', sz: 9.5, bold, color: { rgb: fg } },
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
 * Exports a clean Excel workbook where Sheet 1 is the Master Container Summary,
 * and EVERY Container has its own dedicated worksheet named after its PO & SO.
 */
export function exportContainersByPoSoToExcel(
  poSummaries: ReconciledPoSummary[],
  activePlanName: string,
  allocationMode: string
) {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Clean Master Summary of all Containers (SO & PO)
  const summaryRows: any[][] = [
    [`BÁO CÁO TỔNG HỢP THEO TỪNG CONT XUẤT (PO & SO) — ${activePlanName.toUpperCase()}`],
    [
      'Thời gian xuất:',
      new Date().toLocaleString('vi-VN'),
      'Chế độ tính:',
      allocationMode === 'direct'
        ? 'Tính độc lập từng Cont (Kho 01 + Kho 60)'
        : 'Tính lũy kế FIFO theo thứ tự Cont'
    ],
    [],
    [
      'STT',
      'Cont #',
      'Số PO',
      'Số SO',
      'Khách Hàng',
      'Ngày Đóng Cont',
      'Số Mã Hàng',
      'Số Mã Thiếu',
      'Tổng SL Xuất (Pcs)',
      'SL Đáp Ứng (Pcs)',
      'SL Thiếu (Pcs)',
      'Số Kiện (#Pkg - Cặp)',
      'Tổng CBM',
      'Trạng Thái Cont',
      'Chi Tiết Mã Thiếu & Số Lượng Thiếu'
    ]
  ];

  poSummaries.forEach((p, idx) => {
    const missingDetail =
      p.shortageLines.length === 0
        ? 'Đủ hàng 100%'
        : p.shortageLines
            .map((l) => `${l.itemCode} (Thiếu -${l.shortageQty.toLocaleString('en-US')} pcs)`)
            .join(' | ');

    summaryRows.push([
      idx + 1,
      `Cont #${String(p.containerIndex).padStart(2, '0')}`,
      p.po,
      p.so || p.soList.join(', '),
      p.customer,
      p.loadingDate,
      p.totalLines,
      p.shortageLinesCount,
      p.totalRequiredQty,
      p.totalAllocatedQty,
      p.totalShortageQty,
      p.totalPkgs,
      Number(p.totalCbm.toFixed(2)),
      p.status === 'FULFILLED' ? 'ĐỦ HÀNG 100%' : `THIẾU ${p.shortageLinesCount} MÃ`,
      missingDetail
    ]);
  });

  const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
  wsSummary['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 14 } }];
  wsSummary['!rows'] = [{ hpt: 28 }, { hpt: 20 }, { hpt: 8 }, { hpt: 26 }];
  wsSummary['!cols'] = [
    { wch: 6 },
    { wch: 11 },
    { wch: 16 },
    { wch: 15 },
    { wch: 18 },
    { wch: 14 },
    { wch: 12 },
    { wch: 12 },
    { wch: 16 },
    { wch: 16 },
    { wch: 14 },
    { wch: 16 },
    { wch: 12 },
    { wch: 18 },
    { wch: 64 }
  ];
  styleStandardDataSheet(wsSummary, 3, 15, summaryRows.length, {
    titleBanner: true,
    bannerBg: '1E3A8A',
    headerBg: '1E3A8A',
    statusColIndices: [13],
    shortageColIndices: [7, 10],
    numberColIndices: [6, 7, 8, 9, 10, 11]
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

  // Sheet 1: Clean Executive KPI Summary
  const kpiRows = [
    ['BÁO CÁO ĐỐI SOÁT TỒN KHO (KHO 01 & 60) & KẾ HOẠCH XUẤT CONTAINER'],
    ['Thời gian xuất báo cáo:', new Date().toLocaleString('vi-VN'), ''],
    ['Kế hoạch đang đối soát:', activePlanName, ''],
    [
      'Chế độ phân bổ tồn kho:',
      allocationMode === 'fifo'
        ? 'Phân bổ lũy kế FIFO theo thứ tự Cont xuất'
        : 'Tính độc lập từng Cont (SO & PO) với Tổng Tồn (Kho 01 + Kho 60)',
      ''
    ],
    [],
    ['CHỈ SỐ TỔNG HỢP', 'GIÁ TRỊ', 'GHI CHÚ CHI TIẾT'],
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
      'Tổng Số Cont (PO & SO)',
      reconciliation.metrics.totalPos,
      `Đủ 100%: ${reconciliation.metrics.fulfilledPosCount} Cont | Thiếu hàng: ${reconciliation.metrics.shortagePosCount} Cont`
    ],
    [
      'Tổng Số Kiện Xuất (#Pkg - Theo Cặp Trái/Phải)',
      reconciliation.metrics.totalPkgs,
      'Đã giữ nguyên số kiện theo cặp Trái/Phải (không cộng đúp)'
    ],
    ['Tổng Số Lượng Cần Xuất (Pcs)', reconciliation.metrics.totalRequiredQty, ''],
    ['Tổng Số Lượng Đáp Ứng (Pcs)', reconciliation.metrics.totalAllocatedQty, ''],
    ['Tổng Số Lượng Thiếu Hụt (Pcs)', reconciliation.metrics.totalShortageQty, '']
  ];
  const wsKpi = XLSX.utils.aoa_to_sheet(kpiRows);
  wsKpi['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }];
  wsKpi['!cols'] = [{ wch: 44 }, { wch: 22 }, { wch: 56 }];
  styleStandardDataSheet(wsKpi, 5, 3, kpiRows.length, {
    titleBanner: true,
    bannerBg: '1E3A8A',
    headerBg: '1E3A8A',
    numberColIndices: [1]
  });
  XLSX.utils.book_append_sheet(wb, wsKpi, '01_Tong_Hop_KPI');

  // Sheet 2: Tổng Hợp Theo Từng Cont (PO & SO)
  const poData = reconciliation.poSummaries.map((p, idx) => ({
    STT: idx + 1,
    'Cont #': `Cont #${String(p.containerIndex).padStart(2, '0')}`,
    'Số PO': p.po,
    'Số SO': p.so || p.soList.join(', '),
    'Khách Hàng': p.customer,
    'Ngày Đóng Cont': p.loadingDate,
    'Số Mã Hàng': p.totalLines,
    'Số Mã Thiếu': p.shortageLinesCount,
    'Tổng SL Xuất (Pcs)': p.totalRequiredQty,
    'SL Đáp Ứng (Pcs)': p.totalAllocatedQty,
    'SL Thiếu (Pcs)': p.totalShortageQty,
    'Số Kiện (#Pkg - Cặp)': p.totalPkgs,
    'Tổng CBM': Number(p.totalCbm.toFixed(2)),
    'Trạng Thái Cont': p.status === 'FULFILLED' ? 'ĐỦ HÀNG 100%' : `THIẾU ${p.shortageLinesCount} MÃ`,
    'Chi Tiết Mã Thiếu & SL Thiếu':
      p.shortageLines.length === 0
        ? 'Đủ 100%'
        : p.shortageLines.map((l) => `${l.itemCode} (-${l.shortageQty})`).join(', ')
  }));
  const wsPo = XLSX.utils.json_to_sheet(poData);
  wsPo['!cols'] = [
    { wch: 6 },
    { wch: 11 },
    { wch: 16 },
    { wch: 15 },
    { wch: 18 },
    { wch: 14 },
    { wch: 12 },
    { wch: 12 },
    { wch: 16 },
    { wch: 16 },
    { wch: 14 },
    { wch: 16 },
    { wch: 12 },
    { wch: 18 },
    { wch: 52 }
  ];
  styleStandardDataSheet(wsPo, 0, 15, poData.length + 1, {
    headerBg: '1E3A8A',
    statusColIndices: [13],
    shortageColIndices: [7, 10],
    numberColIndices: [6, 7, 8, 9, 10, 11]
  });
  XLSX.utils.book_append_sheet(wb, wsPo, '02_Tong_Hop_Tung_Cont');

  // Sheet 3: Đơn Đủ Hàng Xuất (Gọn, rõ ràng)
  const fulfilledData = reconciliation.fulfilledLines.map((l, idx) => ({
    STT: idx + 1,
    'Cont #': `Cont #${String(l.containerIndex).padStart(2, '0')}`,
    SO: l.so,
    PO: l.po,
    'Mã Hàng (Item Code)': l.itemCode,
    Model: l.feature,
    'Bên (Trái/Phải)': l.side === 'L' ? 'Trái (L)' : l.side === 'R' ? 'Phải (R)' : 'Chung',
    'Mô Tả Hàng Hóa': l.description,
    'SL Xuất (Pcs)': l.qty,
    'Số Kiện (#Pkg - Cặp)': l.isPairedPkgSecondRow ? '' : l.pairPkgs ?? l.pkgs,
    'Tồn Kho 01': l.wh01Initial,
    'Tồn Kho 60': l.wh60Initial,
    'Tổng Tồn (01+60)': l.stockBeforeLine,
    'Vị Trí Bin Lấy Hàng (Kho 60 / Kho 01)':
      l.pickingBinsText || (l.bins?.length ? l.bins.join(', ') : 'Không có Bin'),
    'Trạng Thái': 'ĐỦ HÀNG'
  }));
  const wsFulfilled = XLSX.utils.json_to_sheet(fulfilledData);
  wsFulfilled['!cols'] = [
    { wch: 6 },
    { wch: 10 },
    { wch: 14 },
    { wch: 15 },
    { wch: 15 },
    { wch: 10 },
    { wch: 12 },
    { wch: 26 },
    { wch: 13 },
    { wch: 15 },
    { wch: 12 },
    { wch: 12 },
    { wch: 14 },
    { wch: 42 },
    { wch: 14 }
  ];
  styleStandardDataSheet(wsFulfilled, 0, 15, fulfilledData.length + 1, {
    headerBg: '1E3A8A',
    binColIndices: [13],
    statusColIndices: [14],
    numberColIndices: [8, 9, 10, 11, 12]
  });
  XLSX.utils.book_append_sheet(wb, wsFulfilled, '03_Don_Du_Hang_Xuat');

  // Sheet 4: Đơn Thiếu Hàng (Tập trung vào số lượng thiếu & tồn hiện có)
  const shortageData = reconciliation.shortageLines.map((l, idx) => ({
    STT: idx + 1,
    'Cont #': `Cont #${String(l.containerIndex).padStart(2, '0')}`,
    SO: l.so,
    PO: l.po,
    'Mã Hàng (Item Code)': l.itemCode,
    Model: l.feature,
    'Bên (Trái/Phải)': l.side === 'L' ? 'Trái (L)' : l.side === 'R' ? 'Phải (R)' : 'Chung',
    'Mô Tả Hàng Hóa': l.description,
    'SL Xuất (Pcs)': l.qty,
    'Số Kiện (#Pkg - Cặp)': l.pairPkgs ?? l.pkgs,
    'Tồn Kho 01': l.wh01Initial,
    'Tồn Kho 60': l.wh60Initial,
    'Tổng Tồn (01+60)': l.stockBeforeLine,
    'SL THIẾU (Pcs)': l.shortageQty,
    'Vị Trí Bin Phần Tồn Hiện Có':
      l.pickingBinsText || l.allBinsText || 'Hết tồn kho (Không có Bin)',
    'Trạng Thái':
      l.fulfillmentStatus === 'ZERO_STOCK' ? 'HẾT TỒN KHO (0 Pcs)' : `THIẾU -${l.shortageQty}`
  }));
  const wsShortage = XLSX.utils.json_to_sheet(shortageData);
  wsShortage['!cols'] = [
    { wch: 6 },
    { wch: 10 },
    { wch: 14 },
    { wch: 15 },
    { wch: 15 },
    { wch: 10 },
    { wch: 12 },
    { wch: 26 },
    { wch: 13 },
    { wch: 15 },
    { wch: 12 },
    { wch: 12 },
    { wch: 14 },
    { wch: 14 },
    { wch: 40 },
    { wch: 18 }
  ];
  styleStandardDataSheet(wsShortage, 0, 16, shortageData.length + 1, {
    headerBg: '1E3A8A',
    binColIndices: [14],
    statusColIndices: [15],
    shortageColIndices: [13],
    numberColIndices: [8, 9, 10, 11, 12, 13]
  });
  XLSX.utils.book_append_sheet(wb, wsShortage, '04_Don_Hang_Thieu');

  // Sheet 5: Báo Cáo Tồn Kho Thành Phẩm Trái/Phải (Kho 01 & 60)
  const lrData = lrPairs.map((p, idx) => {
    const leftBins = stockByCode.get(p.leftCode)?.bins?.join(', ') || '';
    const rightBins = stockByCode.get(p.rightCode)?.bins?.join(', ') || '';
    return {
      STT: idx + 1,
      Model: p.model,
      'Mô Tả': p.description,
      'Mã Trái (LHF)': p.leftCode,
      'Tổng Trái (01+60)': p.leftTotal,
      'Bin Bên Trái': leftBins,
      'Mã Phải (RHF)': p.rightCode,
      'Tổng Phải (01+60)': p.rightTotal,
      'Bin Bên Phải': rightBins,
      'Số Bộ Ghép Đôi (Sets)': p.matchedSets,
      'Lệch (Trái - Phải)': p.diffLR,
      'Trạng Thái Cặp': p.isBalanced ? 'CÂN BẰNG' : `LỆCH (${p.diffLR > 0 ? '+' : ''}${p.diffLR})`
    };
  });
  const wsLr = XLSX.utils.json_to_sheet(lrData);
  wsLr['!cols'] = [
    { wch: 6 },
    { wch: 10 },
    { wch: 24 },
    { wch: 15 },
    { wch: 15 },
    { wch: 30 },
    { wch: 15 },
    { wch: 15 },
    { wch: 30 },
    { wch: 16 },
    { wch: 15 },
    { wch: 16 }
  ];
  styleStandardDataSheet(wsLr, 0, 12, lrData.length + 1, {
    headerBg: '1E3A8A',
    binColIndices: [5, 8],
    statusColIndices: [11],
    numberColIndices: [4, 7, 9, 10]
  });
  XLSX.utils.book_append_sheet(wb, wsLr, '05_Ton_Kho_Trai_Phai');

  // Append individual sheets for each Container (PO & SO)
  for (const poGroup of reconciliation.poSummaries) {
    const wsCont = buildContainerPoSoWorksheet(poGroup, allocationMode);
    XLSX.utils.book_append_sheet(wb, wsCont, poGroup.sheetNameExcel);
  }

  const safeDate = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Bao_Cao_Xuat_Cont_PO_SO_Kho_01_60_${safeDate}.xlsx`);
}
