import * as XLSX from 'xlsx-js-style';
import { format } from 'date-fns';
import { Part, ProductionOrder, GlazingPlan, GlazingPlanNorm, STAGES } from './types';
import { storageService } from './storage';

export const exportPreparationReport = (
  orders: ProductionOrder[],
  glazingPlans: GlazingPlan[],
  parts: Part[],
  glazingPlanNorms: GlazingPlanNorm[],
  selectedMasterPoIds: string[],
  weekName?: string
) => {
  const sheetData: any[][] = [];
  const cleanWeekName = weekName?.replace(/^tuần\s*/i, '') || '';
  const titleText = weekName ? `BÁO CÁO CHUẨN BỊ LINH KIỆN LẮP RÁP TUẦN ${cleanWeekName.toUpperCase()}` : `BÁO CÁO CHUẨN BỊ LINH KIỆN LẮP RÁP`;
  
  const headerRow: any[] = [
    'STT', 
    'MODEL', 
    'TÊN LK', 
    'KH DCCK DỰ KIẾN', 
    'KH DCCK THỰC TẾ', 
    'NGÀY CẦN', 
    'KẾ HOẠCH', 
    'SL THỰC TẾ', 
    'SL CÒN LẠI', 
    'KẾT QUẢ 36/48H', 
    'ĐỐI SÁCH', 
    'TÌNH TRẠNG'
  ];
  
  sheetData.push([titleText]); // Title row matches index 0
  sheetData.push(headerRow);   // Header row matches index 1

  let rowIndex = 1;
  const merges = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 11 } } // Merge title across all 12 columns
  ];

  const selectedMasterPOs = orders.filter(o => selectedMasterPoIds.includes(o.id));
  const selectedGlazingPlans = glazingPlans.filter(p => selectedMasterPoIds.includes(p.id));
  const allTransactions = storageService.getTransactions();
  
  selectedMasterPoIds.forEach(selectedId => {
    const masterPo = selectedMasterPOs.find(o => o.id === selectedId);
    const glazingPlan = selectedGlazingPlans.find(p => p.id === selectedId);

    let modelId = '';
    let displayModelName = '';
    let khDcckDuKien = '';
    let ngayCan = '';
    let ngayCanDate: Date | null = null;
    let partGroups: Record<string, { keHoach: number, slThucTe: number, completedAt: number }> = {};
    let subOrders: ProductionOrder[] = [];

    if (masterPo) {
      modelId = masterPo.partId;
      const baseModelName = parts.find(p => p.id === modelId)?.name || modelId;
      displayModelName = `${baseModelName} (${masterPo.targetQuantity})`;
      
      // 1. KH DCCK dự kiến: Lấy mốc thời gian là kế hoạch theo lệnh sản xuất đã được tạo ra
      khDcckDuKien = masterPo.expectedCompletionTime ? format(new Date(masterPo.expectedCompletionTime), 'dd/MM/yyyy') : '';

      // 2. NGÀY CẦN: Lấy ngày đúng với ngày trong file excel upload vào (bảng popup kế hoạch tự động đã lùi 48h)
      if (masterPo.excelRequiredDate) {
        ngayCanDate = new Date(masterPo.excelRequiredDate);
      } else if (masterPo.note && masterPo.note.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/)) {
        const m = masterPo.note.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        ngayCanDate = new Date(parseInt(m![3]), parseInt(m![2]) - 1, parseInt(m![1]), 8, 0, 0);
      } else if (masterPo.id && masterPo.id.match(/-(\d{4})(\d{2})(\d{2})/)) {
        const m = masterPo.id.match(/-(\d{4})(\d{2})(\d{2})/);
        ngayCanDate = new Date(parseInt(m![1]), parseInt(m![2]) - 1, parseInt(m![3]), 8, 0, 0);
      } else if (masterPo.expectedCompletionTime) {
        // Tự động bù lại 48h để ra đúng ngày trong file Excel
        ngayCanDate = new Date(masterPo.expectedCompletionTime + 48 * 60 * 60 * 1000);
      } else {
        ngayCanDate = new Date(masterPo.createdAt + 48 * 60 * 60 * 1000);
      }
      ngayCan = format(ngayCanDate, 'dd/MM/yyyy');

      subOrders = orders.filter(o => o.masterPoId === masterPo.id && (o.stageId === 'PAINTING' || o.stageId === 'GLAZING'));
      partGroups = subOrders.reduce((acc, sub) => {
         if (!acc[sub.partId]) {
           acc[sub.partId] = { keHoach: 0, slThucTe: 0, completedAt: 0 };
         }
         acc[sub.partId].keHoach += sub.targetQuantity;
         acc[sub.partId].slThucTe += sub.exportedQuantity;
         if (sub.completedAt && sub.completedAt > acc[sub.partId].completedAt) {
           acc[sub.partId].completedAt = sub.completedAt;
         }
         return acc;
      }, {} as Record<string, { keHoach: number, slThucTe: number, completedAt: number }>);
    } else if (glazingPlan) {
      modelId = glazingPlan.modelId;
      const baseModelName = parts.find(p => p.id === modelId)?.name || modelId;
      displayModelName = `${baseModelName} (${glazingPlan.targetQuantity})`;
      
      khDcckDuKien = glazingPlan.expectedCompletionTime 
        ? format(new Date(glazingPlan.expectedCompletionTime), 'dd/MM/yyyy') 
        : (glazingPlan.targetCompletionTime ? format(new Date(glazingPlan.targetCompletionTime), 'dd/MM/yyyy') : '');

      if ((glazingPlan as any).excelRequiredDate) {
        ngayCanDate = new Date((glazingPlan as any).excelRequiredDate);
      } else if (glazingPlan.expectedCompletionTime || glazingPlan.targetCompletionTime) {
        const baseTime = glazingPlan.expectedCompletionTime || glazingPlan.targetCompletionTime;
        ngayCanDate = new Date(baseTime + 48 * 60 * 60 * 1000);
      } else {
        ngayCanDate = new Date(glazingPlan.createdAt + 48 * 60 * 60 * 1000);
      }
      ngayCan = format(ngayCanDate, 'dd/MM/yyyy');

      const components = storageService.getNormsForModel(modelId);
      components.forEach(c => {
         partGroups[c.id] = { keHoach: 0, slThucTe: 0, completedAt: glazingPlan.expectedCompletionTime || 0 };
         partGroups[c.id].keHoach += glazingPlan.targetQuantity;
         partGroups[c.id].slThucTe += (glazingPlan.producedQuantities || {})[c.id] || 0;
      });
    }

    if (!modelId || Object.keys(partGroups).length === 0) return;

    const startRowIdx = sheetData.length;
    let modelPrinted = false;

    const inventory = storageService.getInventory();
    const transformations = storageService.getTransformations();
    const bomV2 = storageService.getBOMV2();

    // Helper to find all available inventory for a part, including its source parts
    const traceInventoryBackwards = (partId: string, neededQty: number, currentModelId: string, visited: Set<string> = new Set(), depth = 0): Record<string, number> => {
      if (visited.has(partId)) return {};
      visited.add(partId);

      const result: Record<string, number> = {};
      
      // Check direct inventory for this part
      const partInventory = inventory.filter(i => i.partId === partId && i.quantity > 0);
      partInventory.forEach(item => {
        const stageName = STAGES.find(s => s.id === item.stageId)?.name || item.stageId;
        const key = `${stageName} (${item.location})`;
        // Include part name if it's a source part
        const partName = depth > 0 ? (parts.find(p => p.id === partId)?.name || partId) : '';
        const displayKey = depth > 0 ? `${partName} ở ${key}` : key;
        result[displayKey] = (result[displayKey] || 0) + item.quantity;
      });

      // Check transformations
      const relatedTransforms = transformations.filter(t => t.targetPartId === partId && (!t.applicableModel || t.applicableModel === currentModelId));
      relatedTransforms.forEach(t => {
        const subResult = traceInventoryBackwards(t.sourcePartId, neededQty, currentModelId, visited, depth + 1);
        Object.entries(subResult).forEach(([k, v]) => {
          result[k] = (result[k] || 0) + v;
        });
      });

      // Check BOM V2
      const relatedBoms = bomV2.filter(b => b.resultPartId === partId && (!b.applicableModel || b.applicableModel === currentModelId));
      relatedBoms.forEach(b => {
        const subResult = traceInventoryBackwards(b.ingredientPartId, neededQty * b.quantity, currentModelId, visited, depth + 1);
        Object.entries(subResult).forEach(([k, v]) => {
          result[k] = (result[k] || 0) + v;
        });
      });

      return result;
    };

    // Mốc thời gian 8h sáng ngày cần để tính kết quả 36/48h
    const mox8hNgayCan = ngayCanDate 
      ? new Date(ngayCanDate.getFullYear(), ngayCanDate.getMonth(), ngayCanDate.getDate(), 8, 0, 0, 0).getTime()
      : 0;

    Object.entries(partGroups).forEach(([partId, data]) => {
       const part = parts.find(p => p.id === partId);
       const lkName = part?.name || partId;
       
       const { keHoach, slThucTe, completedAt } = data;
       const conLai = Math.max(keHoach - slThucTe, 0); 
       
       // 3. KH DCCK thực tế: Điền ngày bắt đầu nhận những số lượng vào các PO linh kiện sơn đó
       let khDcckThucTe = '';
       if (masterPo) {
         const partSubOrders = subOrders.filter(o => o.partId === partId);
         const partSubIds = new Set(partSubOrders.map(o => o.id));

         const txTimes: number[] = [];
         for (const tx of allTransactions) {
           if (tx.quantity <= 0) continue;
           if (tx.poId && partSubIds.has(tx.poId)) {
             txTimes.push(tx.timestamp);
           } else if ((tx.partId === partId || tx.originalPartId === partId) && (tx.stageId === 'PAINTING' || tx.stageId === 'GLAZING')) {
             if (tx.poId) {
               const txPo = orders.find(o => o.id === tx.poId);
               if (txPo && txPo.masterPoId === masterPo.id) {
                 txTimes.push(tx.timestamp);
               }
             }
           }
         }

         if (txTimes.length > 0) {
           khDcckThucTe = format(new Date(Math.min(...txTimes)), 'dd/MM/yyyy');
         } else {
           const subTotalExported = partSubOrders.reduce((sum, s) => sum + (s.exportedQuantity || 0), 0);
           const subTotalProduced = partSubOrders.reduce((sum, s) => sum + (s.producedQuantity || 0), 0);
           if (subTotalExported > 0 || subTotalProduced > 0 || slThucTe > 0) {
             const completedTimes = partSubOrders.filter(s => s.completedAt).map(s => s.completedAt!);
             if (completedTimes.length > 0) {
               khDcckThucTe = format(new Date(Math.min(...completedTimes)), 'dd/MM/yyyy');
             } else {
               khDcckThucTe = format(new Date(masterPo.createdAt), 'dd/MM/yyyy');
             }
           }
         }
       } else if (glazingPlan) {
         const txTimes: number[] = [];
         for (const tx of allTransactions) {
           if (tx.quantity <= 0) continue;
           if (tx.planId === glazingPlan.id || tx.poId === glazingPlan.id) {
             txTimes.push(tx.timestamp);
           } else if (tx.stageId === 'GLAZING' && (tx.partId === partId || tx.originalPartId === partId) && tx.timestamp >= glazingPlan.createdAt - 60 * 60 * 1000) {
             txTimes.push(tx.timestamp);
           }
         }

         if (txTimes.length > 0) {
           khDcckThucTe = format(new Date(Math.min(...txTimes)), 'dd/MM/yyyy');
         } else if (slThucTe > 0) {
           khDcckThucTe = format(new Date(glazingPlan.createdAt), 'dd/MM/yyyy');
         }
       }

       let ketQua = 'OK';
       let tinhTrang = '';
       
       if (conLai > 0) {
         // Không đáp ứng đủ số lượng ngày cần -> báo NG tô đỏ
         ketQua = 'NG';
         const stageCounts = traceInventoryBackwards(partId, conLai, modelId);
         
         tinhTrang = Object.entries(stageCounts).map(([key, qty]) => {
           return `${key}: ${qty}`;
         }).join('\n');
       } else {
         // Đáp ứng đủ số lượng -> xét thời gian hoàn thành so với mốc 8h ngày cần
         let actualFinishedTime = completedAt;
         if (!actualFinishedTime && masterPo) {
           const partSubOrders = subOrders.filter(o => o.partId === partId);
           const partSubIds = new Set(partSubOrders.map(o => o.id));
           const partTxs = allTransactions.filter(tx => 
             (tx.poId && partSubIds.has(tx.poId)) ||
             ((tx.partId === partId || tx.originalPartId === partId) && (tx.stageId === 'PAINTING' || tx.stageId === 'GLAZING'))
           );
           if (partTxs.length > 0) {
             actualFinishedTime = Math.max(...partTxs.map(t => t.timestamp));
           } else {
             actualFinishedTime = masterPo.completedAt || masterPo.createdAt;
           }
         }

         if (mox8hNgayCan > 0 && actualFinishedTime > 0) {
           const diffMs = mox8hNgayCan - actualFinishedTime;
           const diffHours = diffMs / (1000 * 60 * 60);

           if (diffHours < 0) {
             // Hoàn thành sau 8h sáng ngày cần -> trễ kế hoạch ngày cần -> báo NG
             ketQua = 'NG';
             const lateHours = Math.abs(Math.round(diffHours * 10) / 10);
             tinhTrang = `Hoàn thành trễ ${Number.isInteger(lateHours) ? lateHours : lateHours.toFixed(1)}h so với ngày cần`;
           } else if (diffHours >= 48) {
             // Hoàn thành trước 48h ngày cần sản xuất -> OK bình thường
             ketQua = 'OK';
           } else {
             // Hoàn thành trong khoảng thời gian 48h đổ lại tính từ 8h sáng ngày cần (0 <= diffHours < 48)
             // Ghi rõ chi tiết OK cách bao nhiêu tiếng lấy mốc 8h ngày cần
             const roundedHours = Math.round(diffHours * 10) / 10;
             const hStr = Number.isInteger(roundedHours) ? roundedHours.toString() : roundedHours.toFixed(1);
             ketQua = `OK cách ${hStr}h`;
           }
         } else {
           ketQua = 'OK';
         }
       }

       sheetData.push([
         rowIndex++,
         !modelPrinted ? displayModelName : '', // In tên model chỉ ở hàng đầu tiên của nhóm
         lkName,
         khDcckDuKien,
         khDcckThucTe,
         ngayCan,
         keHoach,
         slThucTe,
         conLai,
         ketQua,
         '', // Đối sách
         tinhTrang  // Tình trạng
       ]);
       modelPrinted = true;
    });

    const endRowIdx = sheetData.length - 1;
    if (endRowIdx > startRowIdx) {
      // Merge Model cell across rows for this model
      merges.push({ s: { r: startRowIdx, c: 1 }, e: { r: endRowIdx, c: 1 } });
    }
  });

  const ws = XLSX.utils.aoa_to_sheet(sheetData);

  // Styling
  const borderStyle = {
    top: { style: "thin", color: { rgb: "000000" } },
    bottom: { style: "thin", color: { rgb: "000000" } },
    left: { style: "thin", color: { rgb: "000000" } },
    right: { style: "thin", color: { rgb: "000000" } }
  };

  const headerStyle = {
    font: { bold: true, sz: 11 },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: borderStyle
  };

  const titleStyle = {
    font: { bold: true, sz: 16 },
    alignment: { horizontal: "center", vertical: "center" },
    border: borderStyle
  };

  ws['!merges'] = merges;

  // Set widths for 12 columns
  ws['!cols'] = [
    { wch: 5 },   // STT (0)
    { wch: 15 },  // MODEL (1)
    { wch: 45 },  // TÊN LK (2)
    { wch: 16 },  // KH DCCK DỰ KIẾN (3)
    { wch: 16 },  // KH DCCK THỰC TẾ (4)
    { wch: 14 },  // NGÀY CẦN (5)
    { wch: 12 },  // KẾ HOẠCH (6)
    { wch: 12 },  // SL THỰC TẾ (7)
    { wch: 12 },  // SL CÒN LẠI (8)
    { wch: 16 },  // KẾT QUẢ 36/48H (9)
    { wch: 25 },  // ĐỐI SÁCH (10)
    { wch: 28 }   // TÌNH TRẠNG (11)
  ];

  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  for (let R = 0; R <= range.e.r; ++R) {
    for (let C = 0; C <= range.e.c; ++C) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      if (!ws[cellAddress]) ws[cellAddress] = { t: 's', v: '' };

      if (R === 0) {
        ws[cellAddress].s = titleStyle;
      } else if (R === 1) {
        ws[cellAddress].s = headerStyle;
      } else {
        const val = ws[cellAddress].v;
        const sTemplate: any = { 
            border: borderStyle, 
            alignment: { vertical: "center", horizontal: "center", wrapText: true } 
        };
        
        // Cột KẾT QUẢ 36/48H ở cột index 9 (C === 9)
        if (C === 9 && String(val).startsWith('NG')) {
            sTemplate.fill = { fgColor: { rgb: "FF0000" } };
            sTemplate.font = { color: { rgb: "000000" }, bold: true };
        } else if (C === 9 && String(val).startsWith('OK')) {
            sTemplate.fill = { fgColor: { rgb: "00B050" } };
            sTemplate.font = { color: { rgb: "000000" }, bold: true };
        }
        
        ws[cellAddress].s = sTemplate;
      }
    }
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Lắp ráp 2');
  XLSX.writeFile(wb, `BaoCao_ChuanBi_LK_LapRap_${format(new Date(), 'yyyyMMdd')}.xlsx`);
};
