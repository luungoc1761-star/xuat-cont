import {
  AllocationMode,
  LeftRightPair,
  ReconciledItemDemand,
  ReconciledOrderLine,
  ReconciledPoSummary,
  ShipmentLine,
  StockItem
} from '../types';

export interface ReconciliationResult {
  reconciledLines: ReconciledOrderLine[];
  fulfilledLines: ReconciledOrderLine[];
  shortageLines: ReconciledOrderLine[];
  poSummaries: ReconciledPoSummary[];
  fulfilledPos: ReconciledPoSummary[];
  shortagePos: ReconciledPoSummary[];
  itemDemands: ReconciledItemDemand[];
  enoughItems: ReconciledItemDemand[];
  shortageItems: ReconciledItemDemand[];
  metrics: {
    totalOrderLines: number;
    fulfilledLinesCount: number;
    shortageLinesCount: number;
    zeroStockLinesCount: number;
    partialStockLinesCount: number;
    totalPos: number;
    fulfilledPosCount: number;
    shortagePosCount: number;
    totalRequiredQty: number;
    totalAllocatedQty: number;
    totalShortageQty: number;
    overallFulfillmentRate: number;
    totalCbm: number;
    totalAmountUsd: number;
    totalPkgs: number;
  };
}

function buildSafeExcelSheetName(
  po: string,
  so: string,
  contIdx: number,
  usedNames: Set<string>
): string {
  const clean = (s: string) => s.replace(/[:\\/?*\[\]\s]+/g, '').trim();
  const cPo = clean(po) || `C${contIdx}`;
  const cSo = clean(so);

  // Shorten SO if needed so both PO and SO fit inside 31 chars
  const soShort = cSo.length > 10 ? cSo.slice(-6) : cSo;
  let candidate = cSo ? `PO${cPo}_SO${cSo}` : `PO_${cPo}`;
  if (candidate.length > 31) {
    candidate = `PO${cPo}_SO${soShort}`;
  }
  if (candidate.length > 31) {
    candidate = `${cPo}_${soShort}`.slice(0, 31);
  }

  let finalName = candidate;
  if (usedNames.has(finalName)) {
    const suffix = `_C${contIdx}`;
    finalName = `${candidate.slice(0, 31 - suffix.length)}${suffix}`;
  }
  let counter = 2;
  while (usedNames.has(finalName)) {
    const suffix = `_${counter}`;
    finalName = `${candidate.slice(0, 31 - suffix.length)}${suffix}`;
    counter++;
  }
  usedNames.add(finalName);
  return finalName;
}

export function isAccessoryModel(feature?: string, category?: string): boolean {
  const f = (feature || '').trim().toLowerCase();
  const c = (category || '').trim().toLowerCase();
  return (
    f === 'accessories' ||
    f === 'accessory' ||
    f.includes('access') ||
    f.includes('phụ kiện') ||
    f.includes('phu kien') ||
    c.includes('accessories') ||
    c.includes('phụ kiện')
  );
}

function summarizeBinsForStockItem(s?: StockItem): {
  bins: string[];
  wh60BinsMap: Array<{ bin: string; qty: number }>;
  wh01BinsMap: Array<{ bin: string; qty: number }>;
  wh60BinsText: string;
  wh01BinsText: string;
  allBinsText: string;
} {
  if (!s || s.totalQty <= 0) {
    return {
      bins: [],
      wh60BinsMap: [],
      wh01BinsMap: [],
      wh60BinsText: '',
      wh01BinsText: '',
      allBinsText: 'Hết tồn (Không có Bin)'
    };
  }

  const map60 = new Map<string, number>();
  const map01 = new Map<string, number>();

  for (const b of s.batchList || []) {
    if (b.qty <= 0) continue;
    const binLabel = (b.bin || '').trim() || 'Chưa gán Bin';
    if (b.wh === '60') {
      map60.set(binLabel, (map60.get(binLabel) || 0) + b.qty);
    } else if (b.wh === '01') {
      map01.set(binLabel, (map01.get(binLabel) || 0) + b.qty);
    }
  }

  const wh60BinsMap = Array.from(map60.entries()).map(([bin, qty]) => ({ bin, qty }));
  const wh01BinsMap = Array.from(map01.entries()).map(([bin, qty]) => ({ bin, qty }));

  const wh60BinsText = wh60BinsMap.map((x) => `${x.bin} (${x.qty})`).join(', ');
  const wh01BinsText = wh01BinsMap.map((x) => `${x.bin} (${x.qty})`).join(', ');

  const parts: string[] = [];
  if (wh60BinsText) parts.push(`[Kho 60] ${wh60BinsText}`);
  if (wh01BinsText) parts.push(`[Kho 01] ${wh01BinsText}`);
  if (parts.length === 0 && s.bins?.length) {
    parts.push(s.bins.join(', '));
  }

  return {
    bins: s.bins || [],
    wh60BinsMap,
    wh01BinsMap,
    wh60BinsText,
    wh01BinsText,
    allBinsText: parts.join(' | ') || 'Chưa gán vị trí Bin'
  };
}

function buildPickingBinsSuggestion(
  binSummary: ReturnType<typeof summarizeBinsForStockItem>,
  alloc60: number,
  alloc01: number
): string {
  if (alloc60 <= 0 && alloc01 <= 0) {
    return 'Hết tồn kho (Không có Bin lấy hàng)';
  }
  const pickParts: string[] = [];

  if (alloc60 > 0) {
    let rem = alloc60;
    const picked60: string[] = [];
    for (const entry of binSummary.wh60BinsMap) {
      if (rem <= 0) break;
      const take = Math.min(entry.qty, rem);
      if (take > 0) {
        picked60.push(`${entry.bin} (${take})`);
        rem -= take;
      }
    }
    if (picked60.length > 0) {
      pickParts.push(`Kho 60: ${picked60.join(', ')}`);
    } else if (binSummary.bins.length > 0) {
      pickParts.push(`Kho 60 (${alloc60} pcs): ${binSummary.bins.join(', ')}`);
    } else {
      pickParts.push(`Kho 60: ${alloc60} pcs`);
    }
  }

  if (alloc01 > 0) {
    let rem = alloc01;
    const picked01: string[] = [];
    for (const entry of binSummary.wh01BinsMap) {
      if (rem <= 0) break;
      const take = Math.min(entry.qty, rem);
      if (take > 0) {
        picked01.push(`${entry.bin} (${take})`);
        rem -= take;
      }
    }
    if (picked01.length > 0) {
      pickParts.push(`Kho 01: ${picked01.join(', ')}`);
    } else if (binSummary.bins.length > 0) {
      pickParts.push(`Kho 01 (${alloc01} pcs): ${binSummary.bins.join(', ')}`);
    } else {
      pickParts.push(`Kho 01: ${alloc01} pcs`);
    }
  }

  return pickParts.join(' | ');
}

export function deductShippedContainerFromStock(
  stockItems: StockItem[],
  shippedGroup: ReconciledPoSummary
): StockItem[] {
  const nextItems: StockItem[] = stockItems.map((s) => ({
    ...s,
    bins: [...(s.bins || [])],
    batchList: (s.batchList || []).map((b) => ({ ...b }))
  }));

  const byCode = new Map<string, StockItem>();
  for (const s of nextItems) {
    byCode.set(s.code.trim(), s);
  }

  for (const line of shippedGroup.lines) {
    const code = line.itemCode.trim();
    const item = byCode.get(code);
    if (!item) continue;

    const reqQty = line.qty;
    const deduct60 = Math.min(item.wh60Qty, reqQty);
    const remAfter60 = reqQty - deduct60;
    const deduct01 = Math.min(item.wh01Qty, remAfter60);

    item.wh60Qty = Math.max(0, item.wh60Qty - deduct60);
    item.wh01Qty = Math.max(0, item.wh01Qty - deduct01);
    item.totalQty = item.wh60Qty + item.wh01Qty;

    // Deduct FIFO from batches in WH 60
    let remBatch60 = deduct60;
    for (const b of item.batchList) {
      if (b.wh === '60' && b.qty > 0 && remBatch60 > 0) {
        const take = Math.min(b.qty, remBatch60);
        b.qty -= take;
        remBatch60 -= take;
      }
    }

    // Deduct FIFO from batches in WH 01
    let remBatch01 = deduct01;
    for (const b of item.batchList) {
      if (b.wh === '01' && b.qty > 0 && remBatch01 > 0) {
        const take = Math.min(b.qty, remBatch01);
        b.qty -= take;
        remBatch01 -= take;
      }
    }

    item.batchList = item.batchList.filter((b) => b.qty > 0);
    item.batches60 = item.batchList.filter((b) => b.wh === '60').length;
    item.batches01 = item.batchList.filter((b) => b.wh === '01').length;
    item.bins = Array.from(
      new Set(item.batchList.map((b) => (b.bin || '').trim()).filter(Boolean))
    ).sort();
  }

  return nextItems;
}

export function reconcileShipmentPlan(
  stockItems: StockItem[],
  planLines: ShipmentLine[],
  mode: AllocationMode
): ReconciliationResult {
  const stockMap = new Map<string, { wh01: number; wh60: number; total: number; item: StockItem }>();
  const binSummaryMap = new Map<string, ReturnType<typeof summarizeBinsForStockItem>>();
  for (const s of stockItems) {
    const c = s.code.trim();
    stockMap.set(c, {
      wh01: s.wh01Qty,
      wh60: s.wh60Qty,
      total: s.totalQty,
      item: s
    });
    binSummaryMap.set(c, summarizeBinsForStockItem(s));
  }

  // Global remaining inventory for FIFO cumulative mode across containers
  const globalRem60 = new Map<string, number>();
  const globalRem01 = new Map<string, number>();
  for (const s of stockItems) {
    globalRem60.set(s.code.trim(), s.wh60Qty);
    globalRem01.set(s.code.trim(), s.wh01Qty);
  }

  // Group lines by Container + PO + SO first so per-container mode ("direct") calculates each container independently
  const contGroupOrder: string[] = [];
  const rawContGroups = new Map<string, ShipmentLine[]>();
  for (const line of planLines) {
    const cIdx = line.containerIndex || 0;
    const po = line.po || 'NO-PO';
    const so = line.so || '';
    const gKey = `${cIdx}__${po}__${so}`;
    if (!rawContGroups.has(gKey)) {
      contGroupOrder.push(gKey);
      rawContGroups.set(gKey, []);
    }
    rawContGroups.get(gKey)!.push(line);
  }

  const reconciledLines: ReconciledOrderLine[] = [];
  const poSummaries: ReconciledPoSummary[] = [];
  const usedSheetNames = new Set<string>();

  for (const gKey of contGroupOrder) {
    const groupLines = rawContGroups.get(gKey)!;

    // In 'direct' mode (Tính riêng từng Cont theo SO & PO), each container starts with full Warehouse 01 + 60 stock
    const contRem60 = new Map<string, number>();
    const contRem01 = new Map<string, number>();
    if (mode === 'direct') {
      for (const s of stockItems) {
        contRem60.set(s.code.trim(), s.wh60Qty);
        contRem01.set(s.code.trim(), s.wh01Qty);
      }
    }

    const activeRem60 = mode === 'fifo' ? globalRem60 : contRem60;
    const activeRem01 = mode === 'fifo' ? globalRem01 : contRem01;

    const reconciledGroupLines: ReconciledOrderLine[] = [];

    for (let idx = 0; idx < groupLines.length; idx++) {
      const rawLine = groupLines[idx];
      let line = rawLine;

      // Đảm bảo số kiện (#pkg) luôn giữ nguyên theo cặp Trái/Phải, không cộng đúp Trái riêng Phải riêng
      if (
        rawLine.pairPkgs === undefined &&
        rawLine.isPairedPkgHead === undefined &&
        rawLine.isPairedPkgSecondRow === undefined
      ) {
        const nextLine = groupLines[idx + 1];
        const prevRec = reconciledGroupLines[reconciledGroupLines.length - 1];
        if (
          prevRec &&
          prevRec.isPairedPkgHead &&
          prevRec.po === rawLine.po &&
          prevRec.feature === rawLine.feature &&
          (rawLine.side === 'L' || rawLine.side === 'R') &&
          prevRec.side !== rawLine.side
        ) {
          line = {
            ...rawLine,
            pkgs: 0,
            pairPkgs: prevRec.pairPkgs ?? prevRec.pkgs,
            isPairedPkgHead: false,
            isPairedPkgSecondRow: true
          };
        } else if (
          nextLine &&
          (rawLine.side === 'L' || rawLine.side === 'R') &&
          (nextLine.side === 'L' || nextLine.side === 'R') &&
          rawLine.side !== nextLine.side &&
          rawLine.po === nextLine.po &&
          rawLine.feature === nextLine.feature
        ) {
          const pairPkg = rawLine.pkgs > 0 ? rawLine.pkgs : nextLine.pkgs;
          line = {
            ...rawLine,
            pkgs: pairPkg,
            pairPkgs: pairPkg,
            isPairedPkgHead: true,
            isPairedPkgSecondRow: false
          };
        } else {
          line = {
            ...rawLine,
            pairPkgs: rawLine.pkgs,
            isPairedPkgHead: false,
            isPairedPkgSecondRow: false
          };
        }
      }

      const code = line.itemCode.trim();
      const init = stockMap.get(code) || { wh01: 0, wh60: 0, total: 0 };
      const reqQty = line.qty;

      const directShortageQty = Math.max(0, reqQty - init.total);
      const directStatus: 'FULFILLED' | 'SHORTAGE' =
        directShortageQty === 0 ? 'FULFILLED' : 'SHORTAGE';

      const cur60 = activeRem60.get(code) ?? 0;
      const cur01 = activeRem01.get(code) ?? 0;
      const availBefore = cur60 + cur01;

      const alloc60 = Math.min(cur60, reqQty);
      const remNeeded = reqQty - alloc60;
      const alloc01 = Math.min(cur01, remNeeded);

      activeRem60.set(code, cur60 - alloc60);
      activeRem01.set(code, cur01 - alloc01);

      const allocatedTotal = alloc60 + alloc01;
      const shortageQty = Math.max(0, reqQty - allocatedTotal);
      const stockAfterLine = Math.max(0, availBefore - allocatedTotal);
      const fulfillmentRate = reqQty > 0 ? (allocatedTotal / reqQty) * 100 : 100;

      let fulfillmentStatus: 'FULFILLED' | 'PARTIAL_SHORTAGE' | 'ZERO_STOCK' = 'FULFILLED';
      if (shortageQty > 0) {
        fulfillmentStatus = allocatedTotal > 0 ? 'PARTIAL_SHORTAGE' : 'ZERO_STOCK';
      }

      const binInfo = binSummaryMap.get(code) || summarizeBinsForStockItem(undefined);
      const pickingBinsText = buildPickingBinsSuggestion(binInfo, alloc60, alloc01);

      const recLine: ReconciledOrderLine = {
        ...line,
        wh01Initial: init.wh01,
        wh60Initial: init.wh60,
        totalStockInitial: init.total,
        stockBeforeLine: availBefore,
        allocatedWh60: alloc60,
        allocatedWh01: alloc01,
        allocatedTotal,
        shortageQty,
        stockAfterLine,
        fulfillmentRate,
        fulfillmentStatus,
        directStatus,
        directShortageQty,
        bins: binInfo.bins,
        wh01BinsText: binInfo.wh01BinsText,
        wh60BinsText: binInfo.wh60BinsText,
        pickingBinsText,
        allBinsText: binInfo.allBinsText
      };

      reconciledLines.push(recLine);
      reconciledGroupLines.push(recLine);
    }

    const first = reconciledGroupLines[0];
    const totalReq = reconciledGroupLines.reduce((acc, x) => acc + x.qty, 0);
    const totalAlloc = reconciledGroupLines.reduce((acc, x) => acc + x.allocatedTotal, 0);
    const totalShort = reconciledGroupLines.reduce((acc, x) => acc + x.shortageQty, 0);
    const fulfilledCnt = reconciledGroupLines.filter((x) => x.shortageQty === 0).length;
    const shortLines = reconciledGroupLines.filter((x) => x.shortageQty > 0);
    const shortCnt = shortLines.length;
    const soSet = Array.from(new Set(reconciledGroupLines.map((x) => x.so).filter(Boolean)));
    const sheetNameExcel = buildSafeExcelSheetName(
      first.po || 'NOPO',
      first.so || '',
      first.containerIndex || poSummaries.length + 1,
      usedSheetNames
    );

    poSummaries.push({
      id: gKey,
      po: first.po || `CONT-${first.containerIndex}`,
      so: first.so || '',
      soList: soSet,
      customer: first.customer,
      pkl: first.pkl || '',
      loadingDate: first.updatedLoadingDate || first.loadingDate || first.crd,
      containerIndex: first.containerIndex || poSummaries.length + 1,
      sheetNameExcel,
      totalLines: reconciledGroupLines.length,
      fulfilledLinesCount: fulfilledCnt,
      shortageLinesCount: shortCnt,
      totalRequiredQty: totalReq,
      totalAllocatedQty: totalAlloc,
      totalShortageQty: totalShort,
      totalPkgs: reconciledGroupLines.reduce((acc, x) => acc + x.pkgs, 0),
      totalCbm: reconciledGroupLines.reduce((acc, x) => acc + x.cbm, 0),
      totalAmount: reconciledGroupLines.reduce((acc, x) => acc + x.totalAmount, 0),
      fulfillmentRate: totalReq > 0 ? (totalAlloc / totalReq) * 100 : 100,
      status: shortCnt === 0 ? 'FULFILLED' : 'SHORTAGE',
      lines: reconciledGroupLines,
      shortageLines: shortLines
    });
  }

  const fulfilledLines = reconciledLines.filter((l) => l.shortageQty === 0);
  const shortageLines = reconciledLines.filter((l) => l.shortageQty > 0);
  const fulfilledPos = poSummaries.filter((p) => p.status === 'FULFILLED');
  const shortagePos = poSummaries.filter((p) => p.status === 'SHORTAGE');

  // Aggregate by Item Code across the entire plan
  const itemGroupMap = new Map<string, ReconciledOrderLine[]>();
  for (const l of reconciledLines) {
    const arr = itemGroupMap.get(l.itemCode) || [];
    arr.push(l);
    itemGroupMap.set(l.itemCode, arr);
  }

  const itemDemands: ReconciledItemDemand[] = [];
  for (const [code, lines] of itemGroupMap.entries()) {
    const first = lines[0];
    const init = stockMap.get(code) || { wh01: 0, wh60: 0, total: 0 };
    const totalReq = lines.reduce((acc, x) => acc + x.qty, 0);
    const alloc = Math.min(init.total, totalReq);
    const short = Math.max(0, totalReq - init.total);
    const surplus = Math.max(0, init.total - totalReq);

    const binInfo = binSummaryMap.get(code) || summarizeBinsForStockItem(undefined);

    itemDemands.push({
      itemCode: code,
      feature: first.feature,
      description: first.description,
      side: first.side,
      wh01Qty: init.wh01,
      wh60Qty: init.wh60,
      totalStock: init.total,
      totalRequiredQty: totalReq,
      orderLinesCount: lines.length,
      posList: Array.from(new Set(lines.map((x) => x.po).filter(Boolean))),
      customersList: Array.from(new Set(lines.map((x) => x.customer).filter(Boolean))),
      allocatedQty: alloc,
      shortageQty: short,
      surplusQty: surplus,
      status: short === 0 ? 'ENOUGH' : 'SHORTAGE',
      bins: binInfo.bins,
      allBinsText: binInfo.allBinsText
    });
  }

  itemDemands.sort(
    (a, b) => b.shortageQty - a.shortageQty || b.totalRequiredQty - a.totalRequiredQty
  );

  const enoughItems = itemDemands.filter((i) => i.status === 'ENOUGH');
  const shortageItems = itemDemands.filter((i) => i.status === 'SHORTAGE');

  const totalRequiredQty = reconciledLines.reduce((acc, x) => acc + x.qty, 0);
  const totalAllocatedQty = reconciledLines.reduce((acc, x) => acc + x.allocatedTotal, 0);
  const totalShortageQty = reconciledLines.reduce((acc, x) => acc + x.shortageQty, 0);

  return {
    reconciledLines,
    fulfilledLines,
    shortageLines,
    poSummaries,
    fulfilledPos,
    shortagePos,
    itemDemands,
    enoughItems,
    shortageItems,
    metrics: {
      totalOrderLines: reconciledLines.length,
      fulfilledLinesCount: fulfilledLines.length,
      shortageLinesCount: shortageLines.length,
      zeroStockLinesCount: reconciledLines.filter((x) => x.fulfillmentStatus === 'ZERO_STOCK')
        .length,
      partialStockLinesCount: reconciledLines.filter(
        (x) => x.fulfillmentStatus === 'PARTIAL_SHORTAGE'
      ).length,
      totalPos: poSummaries.length,
      fulfilledPosCount: fulfilledPos.length,
      shortagePosCount: shortagePos.length,
      totalRequiredQty,
      totalAllocatedQty,
      totalShortageQty,
      overallFulfillmentRate: totalRequiredQty > 0 ? (totalAllocatedQty / totalRequiredQty) * 100 : 100,
      totalCbm: reconciledLines.reduce((acc, x) => acc + x.cbm, 0),
      totalAmountUsd: reconciledLines.reduce((acc, x) => acc + x.totalAmount, 0),
      totalPkgs: reconciledLines.reduce((acc, x) => acc + x.pkgs, 0)
    }
  };
}

export function buildLrPairsFromStock(stockItems: StockItem[]): LeftRightPair[] {
  const fgPairsMap = new Map<
    string,
    { model: string; pairKey: string; isFinishedGood: boolean; L?: StockItem; R?: StockItem }
  >();

  for (const item of stockItems) {
    const code = item.code.trim();
    if (item.isFinishedGood && code.length === 10 && code.startsWith('8')) {
      const pairKey = code.slice(0, 5) + 'X' + code.slice(6);
      const model = code.slice(1, 5);
      const side = code[5] === '1' ? 'L' : code[5] === '2' ? 'R' : undefined;
      const entry = fgPairsMap.get(pairKey) || { model, pairKey, isFinishedGood: true };
      if (side === 'L') entry.L = item;
      if (side === 'R') entry.R = item;
      fgPairsMap.set(pairKey, entry);
    }
  }

  const codeLookup = new Map(stockItems.map((x) => [x.code.trim(), x]));
  const explicitPairs: Array<[string, string, string, string, boolean]> = [
    ['1220XX0004', '1220', '1220190004', '1220200004', false],
    ['14217X0201', '1421', '1421710201', '1421720201', false],
    ['1427XX0001', '1427', '1427790001', '1427800001', false],
    ['14553X0001', '1455', '1455350001', '1455360001', false]
  ];

  for (const [pk, mid, lcode, rcode, isFg] of explicitPairs) {
    if (codeLookup.has(lcode) || codeLookup.has(rcode)) {
      fgPairsMap.set(pk, {
        model: mid,
        pairKey: pk,
        isFinishedGood: isFg,
        L: codeLookup.get(lcode),
        R: codeLookup.get(rcode)
      });
    }
  }

  const result: LeftRightPair[] = [];
  for (const pk of Array.from(fgPairsMap.keys()).sort()) {
    const entry = fgPairsMap.get(pk)!;
    const l = entry.L;
    const r = entry.R;
    const l01 = l?.wh01Qty ?? 0;
    const l60 = l?.wh60Qty ?? 0;
    const lTot = l?.totalQty ?? 0;
    const r01 = r?.wh01Qty ?? 0;
    const r60 = r?.wh60Qty ?? 0;
    const rTot = r?.totalQty ?? 0;
    const diff = lTot - rTot;
    const whInternalImbalance = l01 !== r01 || l60 !== r60;

    result.push({
      pairKey: pk,
      model: entry.model,
      isFinishedGood: entry.isFinishedGood,
      description: l?.description || r?.description || `Model ${entry.model}`,
      leftCode: l?.code || `${pk.slice(0, 5)}1${pk.slice(6)}`,
      rightCode: r?.code || `${pk.slice(0, 5)}2${pk.slice(6)}`,
      leftWh01: l01,
      leftWh60: l60,
      leftTotal: lTot,
      rightWh01: r01,
      rightWh60: r60,
      rightTotal: rTot,
      combinedTotal: lTot + rTot,
      matchedSets: Math.min(lTot, rTot),
      diffLR: diff,
      isBalanced: diff === 0,
      hasWarehouseCrossBalance: diff === 0 && whInternalImbalance
    });
  }

  return result;
}
