import * as XLSX from 'xlsx-js-style';
import { ProductionOrder } from './types';

/**
 * Chuẩn hóa mã Model sản phẩm: Bỏ các tiền tố như "RMV-", "VT-", khoảng trắng thừa
 */
export function cleanModelId(rawModel: string): string {
  if (!rawModel) return '';
  let model = String(rawModel).trim();
  // Bỏ tiền tố RMV (ví dụ RMVSHA76222KL -> SHA76222KL, RMV-SHA76222KL -> SHA76222KL)
  model = model.replace(/^RMV[-_.\s]?/i, '');
  model = model.replace(/^VT[-_.\s]?/i, '');
  return model.trim();
}

function formatDateToYYYYMMDD(d: Date): string {
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${yyyy}${mm}${dd}`;
}

function formatDateToDDMMYYYY(d: Date): string {
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

/**
 * Trích xuất ngày từ ô Excel: Chuẩn hóa theo định dạng Ngày/Tháng/Năm (DD/MM/YYYY)
 * Ưu tiên:
 * 1. Excel serial number qua SSF.parse_date_code (chuẩn xác 100%, không bị ảnh hưởng bởi locale)
 * 2. Chuỗi DD/MM/YYYY chuẩn Việt Nam
 * 3. Chuỗi YYYY-MM-DD
 */
export function extractDateFromCell(
  cellObj: XLSX.CellObject | undefined, 
  rawVal: any
): { y: number; m: number; d: number } | null {
  // 1. ƯU TIÊN CAO NHẤT: Nếu ô chứa Excel Serial Number (kiểu number trong khoảng 30000..60000)
  // Khi người dùng nhập ngày trong Excel, Excel lưu dưới dạng số thực serial (ví dụ 46204 = 01/07/2026).
  // Dùng trực tiếp SSF.parse_date_code trên giá trị số thô cellObj.v sẽ ra chính xác {y: 2026, m: 7, d: 1}
  // mà KHÔNG bị lỗi đảo ngược ngày/tháng do SheetJS format hiển thị theo US locale (m/d/yy).
  if (cellObj && typeof cellObj.v === 'number' && cellObj.v > 30000 && cellObj.v < 60000) {
    try {
      const parsed = XLSX.SSF.parse_date_code(cellObj.v);
      if (parsed && parsed.y && parsed.m && parsed.d) {
        return { y: parsed.y, m: parsed.m, d: parsed.d };
      }
    } catch {
      // ignore
    }
  }

  if (typeof rawVal === 'number' && rawVal > 30000 && rawVal < 60000) {
    try {
      const parsed = XLSX.SSF.parse_date_code(rawVal);
      if (parsed && parsed.y && parsed.m && parsed.d) {
        return { y: parsed.y, m: parsed.m, d: parsed.d };
      }
    } catch {
      // ignore
    }
  }

  // 2. Nếu là Date object hợp lệ
  const dateObj = (cellObj && cellObj.v instanceof Date) ? cellObj.v : (rawVal instanceof Date ? rawVal : null);
  if (dateObj && !isNaN(dateObj.getTime())) {
    return { y: dateObj.getFullYear(), m: dateObj.getMonth() + 1, d: dateObj.getDate() };
  }

  // 3. Nếu là chuỗi text (ví dụ: "01/07/2026", "08/10/2026", "1/7/2026")
  // Lấy chuỗi từ cellObj.v (nếu là string) hoặc cellObj.w hoặc rawVal
  const str = String(
    (cellObj && typeof cellObj.v === 'string' && cellObj.v) ||
    (cellObj && cellObj.w) ||
    rawVal ||
    ''
  ).trim();

  if (!str) return null;

  // Bỏ qua các chuỗi ngày trong tuần (T2, T3, T4, T5, T6, T7, CN)
  if (/^(T[2-7]|CN|THỨ\s*[2-7]|CHỦ\s*NHẬT)$/i.test(str)) {
    return null;
  }

  // Dạng chuẩn Việt Nam DD/MM/YYYY: "01/07/2026", "1/7/2026", "01-07-2026", "01.07.2026"
  // match[1] = Ngày (DD), match[2] = Tháng (MM), match[3] = Năm (YYYY)
  const dmyMatch = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})/);
  if (dmyMatch) {
    const day = parseInt(dmyMatch[1], 10);
    const month = parseInt(dmyMatch[2], 10);
    let year = parseInt(dmyMatch[3], 10);
    if (year < 100) year += 2000;
    if (day >= 1 && day <= 31 && month >= 1 && month <= 12 && year >= 2000 && year <= 2100) {
      return { y: year, m: month, d: day };
    }
  }

  // Dạng YYYY-MM-DD hoặc YYYY/MM/DD
  const ymdMatch = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (ymdMatch) {
    const year = parseInt(ymdMatch[1], 10);
    const month = parseInt(ymdMatch[2], 10);
    const day = parseInt(ymdMatch[3], 10);
    if (day >= 1 && day <= 31 && month >= 1 && month <= 12 && year >= 2000 && year <= 2100) {
      return { y: year, m: month, d: day };
    }
  }

  return null;
}

export function parseHeaderToDate(headerVal: any): Date | null {
  if (headerVal === undefined || headerVal === null || headerVal === '') return null;
  const res = extractDateFromCell(undefined, headerVal);
  if (res) {
    return new Date(res.y, res.m - 1, res.d);
  }
  return null;
}

/**
 * Giải mã ngày tháng chính xác 100% cho từng cột của bảng:
 * - Hàng 1 (Row 0): Chứa chuỗi ngày "01/07/2026" hoặc mã serial ngày
 * - Hàng 3 (Row 2): Chứa số ngày viết tắt "1, 2, 3, 4..."
 * - Đối soát chéo để loại bỏ triệt để lỗi đảo ngược ngày/tháng (ví dụ 01/07 bị đổi thành 07/01)
 */
export function resolveColumnDate(
  c: number,
  worksheet: XLSX.WorkSheet,
  rawGrid: any[][]
): Date | null {
  // 1. Tìm số ngày viết tắt ở Hàng 3 (Row index 2) hoặc Row index 1 (ví dụ số 1, 2, 3... 31)
  let explicitDayNum: number | null = null;
  for (let r = 1; r < Math.min(rawGrid.length, 5); r++) {
    const cellVal = rawGrid[r] ? rawGrid[r][c] : undefined;
    if (cellVal !== undefined && cellVal !== null && cellVal !== '') {
      const numStr = String(cellVal).trim();
      if (/^\d{1,2}$/.test(numStr)) {
        const d = parseInt(numStr, 10);
        if (d >= 1 && d <= 31) {
          explicitDayNum = d;
          break;
        }
      }
    }
  }

  // 2. Đọc giá trị ngày ở Hàng 1 (Row index 0) hoặc các hàng tiêu đề
  let datePart: { y: number; m: number; d: number } | null = null;

  for (let r = 0; r < Math.min(rawGrid.length, 3); r++) {
    const cellRef = XLSX.utils.encode_cell({ r, c });
    const cellObj = worksheet[cellRef];
    const rawVal = rawGrid[r] ? rawGrid[r][c] : undefined;
    const parsed = extractDateFromCell(cellObj, rawVal);
    if (parsed) {
      datePart = parsed;
      break;
    }
  }

  if (!datePart) {
    return null;
  }

  let { y, m, d } = datePart;

  // 3. Đối soát với số ngày viết tắt ở Hàng 3 (Row 2) để đảm bảo không bị đảo ngược Ngày / Tháng
  if (explicitDayNum !== null) {
    if (explicitDayNum === d) {
      // Đúng ngày: d là ngày, m là tháng
      return new Date(y, m - 1, d);
    } else if (explicitDayNum === m) {
      // Đã bị đảo ngược vị trí (m chính là ngày, d là tháng) -> Đổi lại cho đúng
      return new Date(y, d - 1, m);
    } else {
      // Nếu không khớp chính xác, tin tưởng số ngày viết tắt explicitDayNum
      return new Date(y, m - 1, explicitDayNum);
    }
  }

  // Nếu không có explicitDayNum, mặc định d là Ngày (DD) và m là Tháng (MM) theo chuẩn Việt Nam
  return new Date(y, m - 1, d);
}

interface DailyPlanPoint {
  date: Date;
  dateKey: string;
  qty: number;
}

export interface MissingBomModelInfo {
  modelId: string;
  name?: string;
  totalQty: number;
  datesCount: number;
  hasModelBOM: boolean;
  issueReason: string;
}

export interface ImportExcelResult {
  orders: ProductionOrder[];
  detectedModels: string[];
  importedModels: string[];
  skippedModels: string[];
  missingBomModels: MissingBomModelInfo[];
  grandTotalFileQty: number;
  validTotalQty: number;
  missingTotalQty: number;
}

/**
 * Hàm phân tích file Excel Kế hoạch Sản xuất:
 * - Cột Model: Lấy ở cột B (index 1)
 * - Hàng 1 đầu tiên: Chứa ngày chi tiết thực tế (01/07/2026, 02/07/2026...)
 * - Lọc theo định mức Model BOM nếu có
 * - Phân tách PO theo từng ngày (không gộp chung nếu chạy nhiều ngày liên tiếp)
 * - Thời gian mong muốn hoàn thành: Lùi 48 tiếng so với 08:00 sáng ngày ghi trong file Excel
 */
export async function importExcelToPO(
  file: File, 
  validModelIds?: Set<string> | string[]
): Promise<ImportExcelResult> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        // Đọc workbook với cellDates: false để giữ nguyên chuỗi text định dạng chuẩn DD/MM/YYYY (ví dụ: "01/07/2026")
        const workbook = XLSX.read(data, { type: 'array', cellDates: false, raw: false });
        
        const firstSheetName = workbook.SheetNames[0];
        if (!firstSheetName) {
          throw new Error('File Excel không có sheet dữ liệu nào!');
        }

        const worksheet = workbook.Sheets[firstSheetName];
        // Đọc dưới dạng mảng 2 chiều 2D Array
        const rawGrid: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

        if (!rawGrid || rawGrid.length === 0) {
          throw new Error('File Excel rỗng!');
        }

        // Tạo tập hợp các model hợp lệ từ hệ thống
        const validModelsSet = new Set<string>();
        if (validModelIds) {
          const list = Array.isArray(validModelIds) ? validModelIds : Array.from(validModelIds);
          list.forEach(id => {
            if (id) {
              validModelsSet.add(cleanModelId(id).toUpperCase());
              validModelsSet.add(String(id).trim().toUpperCase());
            }
          });
        }

        // 1. Xác định Cột Model: Ưu tiên cột B (index 1)
        let modelColIndex = 1;
        let modelNameColIndex = 2; // Cột C: Tên hàng
        for (let r = 0; r < Math.min(rawGrid.length, 5); r++) {
          const row = rawGrid[r] || [];
          for (let c = 0; c < Math.min(row.length, 4); c++) {
            const cell = String(row[c] || '').trim().toLowerCase();
            if (cell === 'model' || cell === 'mã model' || cell === 'ma model') {
              modelColIndex = c;
              break;
            }
          }
        }

        // 2. Xác định các cột ngày từ Hàng 1 (Row index 0) và các hàng tiêu đề
        const dateColumns: { colIndex: number; date: Date }[] = [];
        const maxCols = Math.max(...rawGrid.map(r => (r ? r.length : 0)));

        for (let c = 0; c < maxCols; c++) {
          // Bỏ qua cột Model, cột Mã hàng, và cột Tên hàng
          if (c === modelColIndex || c === 0 || c === modelNameColIndex) continue;

          // Sử dụng resolveColumnDate để giải mã ngày chính xác 100% (chống đảo ngược ngày/tháng)
          const foundDate = resolveColumnDate(c, worksheet, rawGrid);

          if (foundDate) {
            dateColumns.push({ colIndex: c, date: foundDate });
          }
        }

        if (dateColumns.length === 0) {
          throw new Error('Không tìm thấy các cột ngày sản xuất ở hàng 1 của file Excel!');
        }

        // 3. Quét các dòng dữ liệu Model (từ hàng có dữ liệu sản phẩm)
        const modelPointsMap = new Map<string, DailyPlanPoint[]>();
        const detectedModelsSet = new Set<string>();
        const skippedModelsSet = new Set<string>();
        const importedModelsSet = new Set<string>();
        const modelNamesMap = new Map<string, string>();
        const modelTotalQtyMap = new Map<string, number>();

        for (let r = 0; r < rawGrid.length; r++) {
          const row = rawGrid[r] || [];
          const rawModel = row[modelColIndex];
          if (!rawModel) continue;

          const modelId = cleanModelId(rawModel);
          if (!modelId) continue;

          const lower = modelId.toLowerCase();
          // Bỏ qua các dòng tiêu đề hoặc dòng tổng
          if (
            lower === 'model' || 
            lower === 'mã model' || 
            lower === 'tên hàng' || 
            lower === 'mã hàng' || 
            lower.includes('tổng') || 
            lower.includes('total') || 
            lower.includes('cộng') ||
            lower.length < 2
          ) {
            continue;
          }

          detectedModelsSet.add(modelId);
          if (row[modelNameColIndex]) {
            modelNamesMap.set(modelId, String(row[modelNameColIndex]).trim());
          }

          // Tính tổng sản lượng trong file của model này
          let modelSum = 0;
          for (const dateCol of dateColumns) {
            const rawVal = row[dateCol.colIndex];
            if (rawVal === undefined || rawVal === null || rawVal === '') continue;
            const numStr = String(rawVal).replace(/,/g, '').replace(/[^\d.-]/g, '');
            const qty = parseFloat(numStr);
            if (!isNaN(qty) && qty > 0) {
              modelSum += qty;
            }
          }
          modelTotalQtyMap.set(modelId, (modelTotalQtyMap.get(modelId) || 0) + modelSum);

          // Kiểm tra xem Model có trong Định mức Model của hệ thống không
          const isModelValidInSystem = validModelsSet.size === 0 || 
            validModelsSet.has(modelId.toUpperCase()) ||
            validModelsSet.has(cleanModelId(rawModel).toUpperCase()) ||
            Array.from(validModelsSet).some(vm => vm.includes(modelId.toUpperCase()) || modelId.toUpperCase().includes(vm));

          if (!isModelValidInSystem) {
            skippedModelsSet.add(modelId);
            continue;
          }

          importedModelsSet.add(modelId);

          // Lấy sản lượng tại các cột ngày
          for (const dateCol of dateColumns) {
            const rawVal = row[dateCol.colIndex];
            if (rawVal === undefined || rawVal === null || rawVal === '') continue;

            const numStr = String(rawVal).replace(/,/g, '').replace(/[^\d.-]/g, '');
            const qty = parseFloat(numStr);

            if (!isNaN(qty) && qty > 0) {
              if (!modelPointsMap.has(modelId)) {
                modelPointsMap.set(modelId, []);
              }
              const d = dateCol.date;
              const dateKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
              
              // Nếu cùng model có nhiều dòng trong cùng ngày thì cộng dồn số lượng ngày đó
              const existingPoint = modelPointsMap.get(modelId)!.find(p => p.dateKey === dateKey);
              if (existingPoint) {
                existingPoint.qty += qty;
              } else {
                modelPointsMap.get(modelId)!.push({
                  date: d,
                  dateKey,
                  qty
                });
              }
            }
          }
        }

        // Tính toán tổng sản lượng toàn file và danh sách model thiếu định mức
        let grandTotalFileQty = 0;
        for (const qty of modelTotalQtyMap.values()) {
          grandTotalFileQty += qty;
        }

        const missingBomModels: MissingBomModelInfo[] = [];
        let missingTotalQty = 0;

        for (const mId of detectedModelsSet) {
          const totalQty = modelTotalQtyMap.get(mId) || 0;
          // QUAN TRỌNG: Chỉ cảnh báo thiếu BOM nếu Model đó CÓ SẢN LƯỢNG (> 0) trong file Excel!
          if (totalQty <= 0) continue;

          const hasBOM = validModelsSet.size > 0 ? (
            validModelsSet.has(mId.toUpperCase()) ||
            Array.from(validModelsSet).some(vm => vm.includes(mId.toUpperCase()) || mId.toUpperCase().includes(vm))
          ) : true;

          if (!hasBOM) {
            missingTotalQty += totalQty;
            missingBomModels.push({
              modelId: mId,
              name: modelNamesMap.get(mId) || '',
              totalQty,
              datesCount: 0,
              hasModelBOM: false,
              issueReason: 'Có lịch sản xuất trong file nhưng chưa khai báo Định mức Model (Model BOM)'
            });
          }
        }

        // Sắp xếp danh sách thiếu BOM giảm dần theo số lượng PCS bị ảnh hưởng
        missingBomModels.sort((a, b) => b.totalQty - a.totalQty);

        // Nếu tất cả model bị lọc do chưa có BOM, nhưng có model trong file
        if (modelPointsMap.size === 0 && detectedModelsSet.size > 0 && validModelsSet.size > 0) {
          resolve({
            orders: [],
            detectedModels: Array.from(detectedModelsSet),
            importedModels: [],
            skippedModels: Array.from(skippedModelsSet),
            missingBomModels,
            grandTotalFileQty,
            validTotalQty: 0,
            missingTotalQty
          });
          return;
        }

        // 4. Phân tách Lệnh PO theo từng ngày (không gộp chung nếu chạy nhiều ngày liên tục)
        const productionOrders: ProductionOrder[] = [];
        const usedIds = new Set<string>();

        for (const [modelId, points] of modelPointsMap.entries()) {
          if (points.length === 0) continue;

          // Sắp xếp các mốc ngày tăng dần theo thời gian
          points.sort((a, b) => a.date.getTime() - b.date.getTime());

          for (const point of points) {
            const po = buildPOForSingleDay(modelId, point, usedIds);
            productionOrders.push(po);
          }
        }

        // Sắp xếp danh sách Lệnh PO theo thời gian mong muốn hoàn thành
        productionOrders.sort((a, b) => (a.expectedCompletionTime || 0) - (b.expectedCompletionTime || 0));

        const validTotalQty = productionOrders.reduce((sum, o) => sum + o.targetQuantity, 0);

        resolve({
          orders: productionOrders,
          detectedModels: Array.from(detectedModelsSet),
          importedModels: Array.from(importedModelsSet),
          skippedModels: Array.from(skippedModelsSet),
          missingBomModels,
          grandTotalFileQty,
          validTotalQty,
          missingTotalQty
        });
      } catch (err: any) {
        reject(new Error(err?.message || 'Lỗi không xác định khi đọc file Excel'));
      }
    };

    reader.onerror = () => {
      reject(new Error('Không thể đọc file đã chọn!'));
    };

    reader.readAsArrayBuffer(file);
  });
}

function buildPOForSingleDay(
  modelId: string, 
  point: DailyPlanPoint, 
  usedIds: Set<string>
): ProductionOrder {
  const excelDate = point.date;
  const totalQty = point.qty;

  // 1. Mốc 08:00:00 của ngày trong file Excel (ví dụ: ngày 10/10/2026 08:00:00)
  const excelDate8AM = new Date(excelDate.getFullYear(), excelDate.getMonth(), excelDate.getDate(), 8, 0, 0, 0);
  
  // 2. Thời gian mong muốn hoàn thành: Lùi 48 tiếng (48h trước 08:00 của ngày trong file Excel)
  // Ví dụ: Trong file là ngày 10 lúc 8h sáng thì thời gian mong muốn hoàn thành là ngày 8 lúc 8h sáng
  const targetCompletion = new Date(excelDate8AM.getTime() - 48 * 60 * 60 * 1000);
  
  // 3. Thời gian bắt đầu: Khởi tạo cùng mốc deadline (sẽ được tự động lập lịch chi tiết lùi tiếp cho các công đoạn)
  const plannedStart = new Date(targetCompletion.getTime());

  const excelDateStr = formatDateToYYYYMMDD(excelDate);
  let baseId = `PO-${modelId}-${excelDateStr}`;
  let uniqueId = baseId;
  let counter = 1;
  while (usedIds.has(uniqueId)) {
    counter++;
    uniqueId = `${baseId}_${counter}`;
  }
  usedIds.add(uniqueId);

  const excelDateDisplay = formatDateToDDMMYYYY(excelDate);
  const note = `Kế hoạch ngày ${excelDateDisplay} (${totalQty.toLocaleString('vi-VN')} PCS)`;

  return {
    id: uniqueId,
    partId: modelId,
    targetQuantity: totalQty,
    producedQuantity: 0,
    exportedQuantity: 0,
    status: 'PENDING',
    createdAt: Date.now(),
    plannedStartTime: plannedStart.getTime(),
    expectedCompletionTime: targetCompletion.getTime(),
    note,
    excelRequiredDate: excelDate8AM.getTime()
  };
}

/**
 * Tạo và tải xuống file Excel mẫu kế hoạch sản xuất PO chuẩn
 */
export function downloadSamplePOExcelTemplate(availableModels?: string[]) {
  const wb = XLSX.utils.book_new();

  // Danh sách các ngày mẫu (từ 08/10/2026 đến 01/11/2026)
  const sampleDates = [
    { dateStr: '08/10/2026', dow: 'T5', dayNum: '8' },
    { dateStr: '09/10/2026', dow: 'T6', dayNum: '9' },
    { dateStr: '10/10/2026', dow: 'T7', dayNum: '10' },
    { dateStr: '11/10/2026', dow: 'CN', dayNum: '11' },
    { dateStr: '12/10/2026', dow: 'T2', dayNum: '12' },
    { dateStr: '13/10/2026', dow: 'T3', dayNum: '13' },
    { dateStr: '14/10/2026', dow: 'T4', dayNum: '14' },
    { dateStr: '15/10/2026', dow: 'T5', dayNum: '15' },
    { dateStr: '16/10/2026', dow: 'T6', dayNum: '16' },
    { dateStr: '17/10/2026', dow: 'T7', dayNum: '17' },
    { dateStr: '18/10/2026', dow: 'CN', dayNum: '18' },
    { dateStr: '19/10/2026', dow: 'T2', dayNum: '19' },
    { dateStr: '20/10/2026', dow: 'T3', dayNum: '20' },
    { dateStr: '21/10/2026', dow: 'T4', dayNum: '21' },
    { dateStr: '22/10/2026', dow: 'T5', dayNum: '22' },
    { dateStr: '23/10/2026', dow: 'T6', dayNum: '23' },
    { dateStr: '24/10/2026', dow: 'T7', dayNum: '24' },
    { dateStr: '25/10/2026', dow: 'CN', dayNum: '25' },
    { dateStr: '26/10/2026', dow: 'T2', dayNum: '26' },
    { dateStr: '27/10/2026', dow: 'T3', dayNum: '27' },
    { dateStr: '28/10/2026', dow: 'T4', dayNum: '28' },
    { dateStr: '29/10/2026', dow: 'T5', dayNum: '29' },
    { dateStr: '30/10/2026', dow: 'T6', dayNum: '30' },
    { dateStr: '31/10/2026', dow: 'T7', dayNum: '31' },
    { dateStr: '01/11/2026', dow: 'CN', dayNum: '1' }
  ];

  // Hàng 1: Ngày chi tiết thực tế
  const row1: any[] = ['', '', '', ...sampleDates.map(d => d.dateStr)];
  // Hàng 2: Thứ trong tuần
  const row2: any[] = ['', '', '', ...sampleDates.map(d => d.dow)];
  // Hàng 3: Tiêu đề cột
  const row3: any[] = ['Mã hàng', 'Model', 'Tên hàng', ...sampleDates.map(d => d.dayNum)];

  // Dữ liệu mẫu
  const sampleDataRows = [
    {
      code: 'RMVSHR76210CK',
      model: 'SHR76210CK',
      name: 'Máy lọc nước R.O nước nóng 10 lõi Sunhouse SHR76210CK',
      quantities: {} as Record<string, number>
    },
    {
      code: 'RMVSHA75102LA',
      model: 'SHA75102LA',
      name: 'Máy lọc nước ion kiềm Hydrogen UltraX Sunhouse SHA75102LA',
      quantities: { '24/10/2026': 200 }
    },
    {
      code: 'RMVSHA76222KL',
      model: 'SHA76222KL',
      name: 'Máy lọc nước RO UltraPURE Sunhouse 11 lõi SHA76222KL',
      quantities: { '08/10/2026': 300, '09/10/2026': 470, '10/10/2026': 290, '31/10/2026': 300 }
    },
    {
      code: 'RMVSHA76225L',
      model: 'SHA76225L',
      name: 'Máy lọc nước RO UltraPURE Sunhouse 11 lõi SHA76225L',
      quantities: { '11/10/2026': 100, '12/10/2026': 100 }
    },
    {
      code: 'RMVSHA76210KL',
      model: 'SHA76210KL',
      name: 'Máy lọc nước RO UltraPURE Sunhouse 10 lõi SHA76210KL',
      quantities: {}
    },
    {
      code: 'RMVSHA76211KL',
      model: 'SHA76211KL',
      name: 'Máy lọc nước RO UltraPURE Sunhouse 10 lõi SHA76211KL',
      quantities: { '11/10/2026': 150 }
    },
    {
      code: 'RMVSHA76212KL',
      model: 'SHA76212KL',
      name: 'Máy lọc nước RO UltraPURE Sunhouse 10 lõi SHA76212KL',
      quantities: {}
    }
  ];

  // Nếu có danh sách model từ hệ thống mà chưa có trong mẫu, bổ sung thêm
  if (availableModels && availableModels.length > 0) {
    const existingModels = new Set(sampleDataRows.map(r => r.model.toUpperCase()));
    for (const mId of availableModels) {
      if (!existingModels.has(mId.toUpperCase())) {
        sampleDataRows.push({
          code: `RMV${mId}`,
          model: mId,
          name: `Sản phẩm ${mId}`,
          quantities: {}
        });
      }
    }
  }

  const sheetData: any[][] = [row1, row2, row3];

  sampleDataRows.forEach(item => {
    const r: any[] = [item.code, item.model, item.name];
    sampleDates.forEach(d => {
      const q = item.quantities[d.dateStr];
      r.push(q !== undefined ? q : '');
    });
    sheetData.push(r);
  });

  const ws = XLSX.utils.aoa_to_sheet(sheetData);

  // Đặt độ rộng cột
  ws['!cols'] = [
    { wch: 20 }, // Mã hàng
    { wch: 18 }, // Model (Cột B)
    { wch: 55 }, // Tên hàng
    ...sampleDates.map(() => ({ wch: 12 })) // Các cột ngày
  ];

  // Đặt chiều cao dòng
  ws['!rows'] = [
    { hpt: 26 }, // Hàng 1
    { hpt: 22 }, // Hàng 2
    { hpt: 24 }, // Hàng 3
    ...sampleDataRows.map(() => ({ hpt: 24 }))
  ];

  // Format styles
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:AA30');
  
  for (let R = range.s.r; R <= range.e.r; ++R) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellRef = XLSX.utils.encode_cell({ c: C, r: R });
      const cell = ws[cellRef];
      if (!cell) continue;

      // Header rows (0, 1, 2)
      if (R <= 2) {
        cell.s = {
          fill: { fgColor: { rgb: "E5E7EB" } },
          font: { bold: true, sz: 11, color: { rgb: "111827" } },
          alignment: { horizontal: "center", vertical: "center" },
          border: {
            top: { style: "thin", color: { rgb: "D1D5DB" } },
            bottom: { style: "thin", color: { rgb: "D1D5DB" } },
            left: { style: "thin", color: { rgb: "D1D5DB" } },
            right: { style: "thin", color: { rgb: "D1D5DB" } }
          }
        };
        if (R === 0 && C >= 3) {
          cell.s.font = { bold: true, sz: 10, color: { rgb: "1F2937" } };
        }
      } else {
        // Data rows
        const isQtyCell = C >= 3 && typeof cell.v === 'number' && cell.v > 0;
        cell.s = {
          font: { sz: 11, bold: isQtyCell, color: { rgb: isQtyCell ? "15803D" : "1F2937" } },
          alignment: { 
            horizontal: C === 2 ? "left" : (C >= 3 ? "center" : "left"), 
            vertical: "center" 
          },
          border: {
            top: { style: "thin", color: { rgb: "E5E7EB" } },
            bottom: { style: "thin", color: { rgb: "E5E7EB" } },
            left: { style: "thin", color: { rgb: "E5E7EB" } },
            right: { style: "thin", color: { rgb: "E5E7EB" } }
          }
        };

        if (isQtyCell) {
          cell.s.fill = { fgColor: { rgb: "DCFCE7" } }; // Highlight green for sample quantities
        }
      }
    }
  }

  XLSX.utils.book_append_sheet(wb, ws, "KeHoach_SanXuat");
  XLSX.writeFile(wb, "Mau_KeHoach_SanXuat_PO.xlsx");
}

