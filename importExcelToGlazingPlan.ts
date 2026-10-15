import * as XLSX from 'xlsx-js-style';
import { GlazingPlan, GlazingPlanNorm } from './types';
import { storageService } from './storage';
import { cleanModelId, resolveColumnDate } from './importExcelToPO';

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

export interface MissingNormModelInfo {
  modelId: string;
  name?: string;
  totalQty: number;
  datesCount: number;
  hasGlazingNorm: boolean;
  issueReason: string;
}

export interface ImportGlazingPlanResult {
  plans: GlazingPlan[];
  detectedModels: string[];
  importedModels: string[];
  skippedModels: string[];
  missingNormModels: MissingNormModelInfo[];
  grandTotalFileQty: number;
  validTotalQty: number;
  missingTotalQty: number;
}

interface DailyPlanPoint {
  date: Date;
  dateKey: string;
  qty: number;
}

/**
 * Hàm phân tích file Excel Kế hoạch Sản xuất cho Công đoạn Dán Kính:
 * - Cấu trúc file y chang file bên PO tổng
 * - Cột Model: Ưu tiên cột B (index 1)
 * - Hàng 1: Ngày sản xuất trong file Excel
 * - Cơ chế nhận ngày mong muốn hoàn thành: Lấy ngày trong file Excel lúc 08:00 sáng lùi lại 48h
 * - Tạo kế hoạch theo từng model, nếu model chạy nhiều ngày thì tạo kế hoạch riêng theo từng ngày
 * - Lọc và kiểm tra theo Định mức Dán Kính (tab Cấu hình)
 */
export async function importExcelToGlazingPlan(
  file: File,
  glazingNorms: GlazingPlanNorm[]
): Promise<ImportGlazingPlanResult> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array', cellDates: false, raw: false });
        
        const firstSheetName = workbook.SheetNames[0];
        if (!firstSheetName) {
          throw new Error('File Excel không có sheet dữ liệu nào!');
        }

        const worksheet = workbook.Sheets[firstSheetName];
        const rawGrid: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

        if (rawGrid.length < 3) {
          throw new Error('File Excel không đúng định dạng kế hoạch sản xuất (quá ít hàng dữ liệu)!');
        }

        // Tạo tập hợp các model đã có định mức dán kính
        const validNormModelsSet = new Set<string>();
        glazingNorms.forEach(n => {
          if (n.appliedModel) {
            validNormModelsSet.add(cleanModelId(n.appliedModel).toUpperCase());
            validNormModelsSet.add(n.appliedModel.trim().toUpperCase());
          }
          if (n.modelName) {
            validNormModelsSet.add(cleanModelId(n.modelName).toUpperCase());
            validNormModelsSet.add(n.modelName.trim().toUpperCase());
          }
        });

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

        // 2. Xác định các cột ngày từ Hàng 1 (Row index 0)
        const dateColumns: { colIndex: number; date: Date }[] = [];
        const maxCols = Math.max(...rawGrid.map(r => (r ? r.length : 0)));

        for (let c = 0; c < maxCols; c++) {
          if (c === modelColIndex || c === 0 || c === modelNameColIndex) continue;
          const foundDate = resolveColumnDate(c, worksheet, rawGrid);
          if (foundDate) {
            dateColumns.push({ colIndex: c, date: foundDate });
          }
        }

        if (dateColumns.length === 0) {
          throw new Error('Không tìm thấy các cột ngày sản xuất ở hàng 1 của file Excel!');
        }

        // 3. Quét các dòng dữ liệu Model
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

          // Kiểm tra xem Model có Định mức dán kính không
          const upperModel = modelId.toUpperCase();
          const cleanUpper = cleanModelId(rawModel).toUpperCase();
          const isModelValidInGlazing = validNormModelsSet.size === 0 ||
            validNormModelsSet.has(upperModel) ||
            validNormModelsSet.has(cleanUpper) ||
            Array.from(validNormModelsSet).some(vm => vm === upperModel || vm.includes(upperModel) || upperModel.includes(vm));

          if (!isModelValidInGlazing) {
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

        // Tính toán tổng sản lượng toàn file và danh sách thiếu định mức dán kính
        let grandTotalFileQty = 0;
        for (const qty of modelTotalQtyMap.values()) {
          grandTotalFileQty += qty;
        }

        const missingNormModels: MissingNormModelInfo[] = [];
        let missingTotalQty = 0;

        for (const mId of detectedModelsSet) {
          const totalQty = modelTotalQtyMap.get(mId) || 0;
          if (totalQty <= 0) continue;

          const upperModel = mId.toUpperCase();
          const hasNorm = validNormModelsSet.size > 0 ? (
            validNormModelsSet.has(upperModel) ||
            Array.from(validNormModelsSet).some(vm => vm === upperModel || vm.includes(upperModel) || upperModel.includes(vm))
          ) : true;

          if (!hasNorm) {
            missingTotalQty += totalQty;
            missingNormModels.push({
              modelId: mId,
              name: modelNamesMap.get(mId) || '',
              totalQty,
              datesCount: 0,
              hasGlazingNorm: false,
              issueReason: 'Có lịch sản xuất trong file nhưng chưa cài đặt Định mức dán kính (tab Cấu hình)'
            });
          }
        }

        missingNormModels.sort((a, b) => b.totalQty - a.totalQty);

        // Nếu tất cả model bị lọc do chưa có định mức dán kính
        if (modelPointsMap.size === 0 && detectedModelsSet.size > 0 && validNormModelsSet.size > 0) {
          resolve({
            plans: [],
            detectedModels: Array.from(detectedModelsSet),
            importedModels: [],
            skippedModels: Array.from(skippedModelsSet),
            missingNormModels,
            grandTotalFileQty,
            validTotalQty: 0,
            missingTotalQty
          });
          return;
        }

        // 4. Phân tách Kế hoạch Dán Kính theo từng model và theo từng ngày
        const glazingPlans: GlazingPlan[] = [];
        const usedIds = new Set<string>();

        for (const [modelId, points] of modelPointsMap.entries()) {
          if (points.length === 0) continue;

          points.sort((a, b) => a.date.getTime() - b.date.getTime());

          for (const point of points) {
            const excelDate = point.date;
            const totalQty = point.qty;

            // Mốc 08:00 sáng của ngày trong file Excel
            const excelDate8AM = new Date(excelDate.getFullYear(), excelDate.getMonth(), excelDate.getDate(), 8, 0, 0, 0);

            // Thời gian mong muốn hoàn thành: Lùi 48 tiếng so với 08:00 sáng ngày trong file Excel
            const targetCompletion = new Date(excelDate8AM.getTime() - 48 * 60 * 60 * 1000);

            // Tính lịch tự động cho dán kính
            const schedule = storageService.getGlazingSchedule(modelId, totalQty, targetCompletion.getTime());

            const excelDateStr = formatDateToYYYYMMDD(excelDate);
            let baseId = `GLZ-PLN-${modelId}-${excelDateStr}`;
            let uniqueId = baseId;
            let counter = 1;
            while (usedIds.has(uniqueId)) {
              counter++;
              uniqueId = `${baseId}_${counter}`;
            }
            usedIds.add(uniqueId);

            const excelDateDisplay = formatDateToDDMMYYYY(excelDate);
            const note = `Kế hoạch ngày ${excelDateDisplay} (${totalQty.toLocaleString('vi-VN')} Máy)`;

            glazingPlans.push({
              id: uniqueId,
              modelId,
              targetQuantity: totalQty,
              targetCompletionTime: targetCompletion.getTime(),
              plannedStartTime: schedule.start,
              expectedCompletionTime: schedule.end,
              createdAt: Date.now(),
              status: 'PENDING',
              note,
              excelRequiredDate: excelDate8AM.getTime()
            });
          }
        }

        // Sắp xếp theo hạn hoàn thành
        glazingPlans.sort((a, b) => (a.targetCompletionTime || 0) - (b.targetCompletionTime || 0));

        const validTotalQty = glazingPlans.reduce((sum, p) => sum + p.targetQuantity, 0);

        resolve({
          plans: glazingPlans,
          detectedModels: Array.from(detectedModelsSet),
          importedModels: Array.from(importedModelsSet),
          skippedModels: Array.from(skippedModelsSet),
          missingNormModels,
          grandTotalFileQty,
          validTotalQty,
          missingTotalQty
        });
      } catch (err: any) {
        reject(new Error(err?.message || 'Lỗi không xác định khi đọc file Excel dán kính'));
      }
    };

    reader.onerror = () => {
      reject(new Error('Không thể đọc file đã chọn!'));
    };

    reader.readAsArrayBuffer(file);
  });
}
