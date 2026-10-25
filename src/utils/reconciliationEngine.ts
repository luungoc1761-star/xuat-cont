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

export function reconcileShipmentPlan(
  stockItems: StockItem[],
  planLines: ShipmentLine[],
  mode: AllocationMode
): ReconciliationResult {
  const stockMap = new Map<string, { wh01: number; wh60: number; total: number }>();
  for (const s of stockItems) {
    stockMap.set(s.code.trim(), {
      wh01: s.wh01Qty,
      wh60: s.wh60Qty,
      total: s.totalQty
    });
  }

  // Remaining inventory for FIFO allocation (prioritizing Warehouse 60 finished goods, then Warehouse 01)
  const rem60 = new Map<string, number>();
  const rem01 = new Map<string, number>();
  for (const s of stockItems) {
    rem60.set(s.code.trim(), s.wh60Qty);
    rem01.set(s.code.trim(), s.wh01Qty);
  }

  const reconciledLines: ReconciledOrderLine[] = [];

  for (const line of planLines) {
    const code = line.itemCode.trim();
    const init = stockMap.get(code) || { wh01: 0, wh60: 0, total: 0 };
    const reqQty = line.qty;

    // Direct check against initial total stock (WH 01 + WH 60)
    const directShortageQty = Math.max(0, reqQty - init.total);
    const directStatus: 'FULFILLED' | 'SHORTAGE' = directShortageQty === 0 ? 'FULFILLED' : 'SHORTAGE';

    // FIFO cumulative check
    const cur60 = rem60.get(code) ?? 0;
    const cur01 = rem01.get(code) ?? 0;
    const availBefore = mode === 'fifo' ? cur60 + cur01 : init.total;

    let alloc60 = 0;
    let alloc01 = 0;

    if (mode === 'fifo') {
      alloc60 = Math.min(cur60, reqQty);
      const remNeeded = reqQty - alloc60;
      alloc01 = Math.min(cur01, remNeeded);
      rem60.set(code, cur60 - alloc60);
      rem01.set(code, cur01 - alloc01);
    } else {
      alloc60 = Math.min(init.wh60, reqQty);
      const remNeeded = reqQty - alloc60;
      alloc01 = Math.min(init.wh01, remNeeded);
    }

    const allocatedTotal = alloc60 + alloc01;
    const shortageQty = Math.max(0, reqQty - allocatedTotal);
    const stockAfterLine = Math.max(0, availBefore - allocatedTotal);
    const fulfillmentRate = reqQty > 0 ? (allocatedTotal / reqQty) * 100 : 100;

    let fulfillmentStatus: 'FULFILLED' | 'PARTIAL_SHORTAGE' | 'ZERO_STOCK' = 'FULFILLED';
    if (shortageQty > 0) {
      fulfillmentStatus = allocatedTotal > 0 ? 'PARTIAL_SHORTAGE' : 'ZERO_STOCK';
    }

    reconciledLines.push({
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
      directShortageQty
    });
  }

  const fulfilledLines = reconciledLines.filter((l) => l.shortageQty === 0);
  const shortageLines = reconciledLines.filter((l) => l.shortageQty > 0);

  // Group by PO
  const poGroupMap = new Map<string, ReconciledOrderLine[]>();
  for (const l of reconciledLines) {
    const key = l.po || `NO-PO-${l.containerIndex}`;
    const arr = poGroupMap.get(key) || [];
    arr.push(l);
    poGroupMap.set(key, arr);
  }

  const poSummaries: ReconciledPoSummary[] = [];
  for (const [poKey, lines] of poGroupMap.entries()) {
    const totalReq = lines.reduce((acc, x) => acc + x.qty, 0);
    const totalAlloc = lines.reduce((acc, x) => acc + x.allocatedTotal, 0);
    const totalShort = lines.reduce((acc, x) => acc + x.shortageQty, 0);
    const fulfilledCnt = lines.filter((x) => x.shortageQty === 0).length;
    const shortCnt = lines.length - fulfilledCnt;
    const soSet = Array.from(new Set(lines.map((x) => x.so).filter(Boolean)));
    const first = lines[0];

    poSummaries.push({
      po: poKey,
      soList: soSet,
      customer: first.customer,
      loadingDate: first.updatedLoadingDate || first.loadingDate || first.crd,
      containerIndex: first.containerIndex,
      totalLines: lines.length,
      fulfilledLinesCount: fulfilledCnt,
      shortageLinesCount: shortCnt,
      totalRequiredQty: totalReq,
      totalAllocatedQty: totalAlloc,
      totalShortageQty: totalShort,
      totalPkgs: lines.reduce((acc, x) => acc + x.pkgs, 0),
      totalCbm: lines.reduce((acc, x) => acc + x.cbm, 0),
      totalAmount: lines.reduce((acc, x) => acc + x.totalAmount, 0),
      fulfillmentRate: totalReq > 0 ? (totalAlloc / totalReq) * 100 : 100,
      status: shortCnt === 0 ? 'FULFILLED' : 'SHORTAGE',
      lines
    });
  }

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
      status: short === 0 ? 'ENOUGH' : 'SHORTAGE'
    });
  }

  itemDemands.sort((a, b) => b.shortageQty - a.shortageQty || b.totalRequiredQty - a.totalRequiredQty);

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
      zeroStockLinesCount: reconciledLines.filter((x) => x.fulfillmentStatus === 'ZERO_STOCK').length,
      partialStockLinesCount: reconciledLines.filter((x) => x.fulfillmentStatus === 'PARTIAL_SHORTAGE').length,
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
