export interface BatchRecord {
  wh: string;
  batch: string;
  bin: string;
  date: string;
  qty: number;
}

export interface StockItem {
  code: string;
  feature: string;
  description: string;
  customerCode: string;
  category: string;
  isFinishedGood: boolean;
  side: 'L' | 'R' | 'N';
  wh01Qty: number;
  wh60Qty: number;
  totalQty: number;
  batches01: number;
  batches60: number;
  bins: string[];
  latestDate: string;
  batchList: BatchRecord[];
}

export interface LeftRightPair {
  pairKey: string;
  model: string;
  isFinishedGood: boolean;
  description: string;
  leftCode: string;
  rightCode: string;
  leftWh01: number;
  leftWh60: number;
  leftTotal: number;
  rightWh01: number;
  rightWh60: number;
  rightTotal: number;
  combinedTotal: number;
  matchedSets: number;
  diffLR: number;
  isBalanced: boolean;
  hasWarehouseCrossBalance: boolean;
}

export interface ShipmentLine {
  id: string;
  sheet: string;
  rowNumber: number;
  containerIndex: number;
  customer: string;
  type: string;
  pkl: string;
  so: string;
  po: string;
  itemCode: string;
  customerItemCode: string;
  feature: string;
  description: string;
  dimension: string;
  crd: string;
  loadingDate: string;
  updatedLoadingDate: string;
  qty: number;
  unitWeight: number;
  pcsPerPkg: number;
  pkgs: number;
  pairPkgs?: number;
  isPairedPkgHead?: boolean;
  isPairedPkgSecondRow?: boolean;
  netWeight: number;
  grossWeight: number;
  cbm: number;
  unitPrice: number;
  totalAmount: number;
  side: 'L' | 'R' | 'N';
  month?: string;
  status?: string;
}

export interface ContainerGroup {
  containerIndex: number;
  sheet: string;
  subtotalRow: number;
  customer: string;
  pos: string[];
  sos: string[];
  pkl: string;
  loadingDate: string;
  lineCount: number;
  totalQty: number;
  totalPkgs: number;
  netWeight: number;
  grossWeight: number;
  totalCbm: number;
  totalAmount: number;
}

export interface BookingRecord {
  rowNumber: number;
  contNo: string;
  customer: string;
  po: string;
  grossWeight: number;
  bookingReqDate: string;
  crd: string;
  loadingDate: string;
  vendorBooking: string;
  bookingNo: string;
  etd: string;
  eta: string;
  siCutOff: string;
  cyCutOff: string;
}

export interface PlanDataset {
  name: string;
  type: 'monthly' | 'weekly' | 'master';
  lines: ShipmentLine[];
  containers: ContainerGroup[];
}

export interface AuditSummary {
  totalRawStockRows: number;
  wh01Rows: number;
  wh01UniqueCodes: number;
  wh01TotalQty: number;
  wh60Rows: number;
  wh60UniqueCodes: number;
  wh60TotalQty: number;
  combinedUniqueCodes: number;
  overlappingCodesCount: number;
  combinedTotalQty: number;
  finishedGoodsCodesCount: number;
  finishedGoodsTotalQty: number;
  finishedGoodsModelsCount: number;
  balancedFgModelsCount: number;
  imbalancedFgModelsCount: number;
  novFullLinesCount: number;
  novFullPosCount: number;
  novFullContainersCount: number;
  octWk42LinesCount: number;
  octWk42PosCount: number;
  dataOpeningLinesCount: number;
}

export interface InitialDataPayload {
  generatedAt: string;
  sourceFiles: {
    stockFile: string;
    shipmentFile: string;
    weeklyFile?: string;
  };
  auditSummary: AuditSummary;
  stockItems: StockItem[];
  lrPairs: LeftRightPair[];
  plans: Record<string, PlanDataset>;
  bookings: BookingRecord[];
}

export type AllocationMode = 'fifo' | 'direct';

export interface ReconciledOrderLine extends ShipmentLine {
  wh01Initial: number;
  wh60Initial: number;
  totalStockInitial: number;
  stockBeforeLine: number;
  allocatedWh60: number;
  allocatedWh01: number;
  allocatedTotal: number;
  shortageQty: number;
  stockAfterLine: number;
  fulfillmentRate: number;
  fulfillmentStatus: 'FULFILLED' | 'PARTIAL_SHORTAGE' | 'ZERO_STOCK';
  directStatus: 'FULFILLED' | 'SHORTAGE';
  directShortageQty: number;
  bins?: string[];
  wh01BinsText?: string;
  wh60BinsText?: string;
  pickingBinsText?: string;
  allBinsText?: string;
}

export interface ReconciledPoSummary {
  id: string;
  po: string;
  so: string;
  soList: string[];
  customer: string;
  pkl: string;
  loadingDate: string;
  containerIndex: number;
  sheetNameExcel: string;
  totalLines: number;
  fulfilledLinesCount: number;
  shortageLinesCount: number;
  totalRequiredQty: number;
  totalAllocatedQty: number;
  totalShortageQty: number;
  totalPkgs: number;
  totalCbm: number;
  totalAmount: number;
  fulfillmentRate: number;
  status: 'FULFILLED' | 'SHORTAGE';
  lines: ReconciledOrderLine[];
  shortageLines: ReconciledOrderLine[];
}

export interface ReconciledItemDemand {
  itemCode: string;
  feature: string;
  description: string;
  side: 'L' | 'R' | 'N';
  wh01Qty: number;
  wh60Qty: number;
  totalStock: number;
  totalRequiredQty: number;
  orderLinesCount: number;
  posList: string[];
  customersList: string[];
  allocatedQty: number;
  shortageQty: number;
  surplusQty: number;
  status: 'ENOUGH' | 'SHORTAGE';
  bins?: string[];
  allBinsText?: string;
}
