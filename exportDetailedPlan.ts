import * as XLSX from 'xlsx-js-style';
import { format, addDays, isSameDay, startOfDay, endOfDay } from 'date-fns';
import { ProductionOrder, Part, InventoryItem, ModelBOMDefinition, BOMDefinitionV2, StageId, GlazingPlan, GlazingPlanNorm } from './types';
import { storageService } from './storage';

export interface ExportPlanOptions {
  reportDateStr?: string; // YYYY-MM-DD
  selectedStages?: StageId[];
  glazingPlans?: GlazingPlan[];
  glazingPlanNorms?: GlazingPlanNorm[];
}

interface ModelStagePart {
  partId: string;
  partName: string;
  norm: number;
}

interface CellPlanMeta {
  isCoveredFull?: boolean;
  isPartialDeducted?: boolean;
  displayVal?: string;
  originalQty?: number;
  usedInv?: number;
  remainingNeeded?: number;
  isPastDate?: boolean;
  isTodayDate?: boolean;
}

export const exportDetailedProductionPlan = (
  orders: ProductionOrder[],
  parts: Part[],
  inventory: InventoryItem[],
  modelBom: ModelBOMDefinition[],
  bomV2: BOMDefinitionV2[],
  options: ExportPlanOptions = {}
) => {
  const reportDateStr = options.reportDateStr || format(new Date(), 'yyyy-MM-dd');
  const [rY, rM, rD] = reportDateStr.split('-').map(Number);
  const baseDate = new Date(rY, rM - 1, rD, 0, 0, 0, 0);

  // Time window: 3 days prior to reportDate and 30 days after (approx 1 month)
  const startDate = startOfDay(addDays(baseDate, -3));
  const endDate = endOfDay(addDays(baseDate, 30));
  const startMs = startDate.getTime();
  const endMs = endDate.getTime();

  const dayDates: Date[] = [];
  let curr = new Date(startDate);
  while (curr <= endDate) {
    dayDates.push(new Date(curr));
    curr = addDays(curr, 1);
  }

  const wb = XLSX.utils.book_new();

  const stagesToExport: { id: StageId; sheetName: string; titleName: string }[] = [
    { id: 'LASER', sheetName: 'KHSX Laser', titleName: 'KẾ HOẠCH SẢN XUẤT NHÓM LASER' },
    { id: 'BENDING', sheetName: 'KHSX Chấn Dập', titleName: 'KẾ HOẠCH SẢN XUẤT NHÓM CHẤN DẬP' },
    { id: 'WELDING', sheetName: 'KHSX Hàn', titleName: 'KẾ HOẠCH SẢN XUẤT NHÓM HÀN' },
    { id: 'PAINTING', sheetName: 'KHSX Sơn', titleName: 'KẾ HOẠCH SẢN XUẤT NHÓM SƠN' },
    { id: 'GLAZING', sheetName: 'KHSX Dán Kính', titleName: 'KẾ HOẠCH SẢN XUẤT NHÓM DÁN KÍNH' },
  ];

  const filteredStages = options.selectedStages && options.selectedStages.length > 0
    ? stagesToExport.filter(s => options.selectedStages?.includes(s.id))
    : stagesToExport;

  const borderStyle = {
    top: { style: 'thin', color: { rgb: '000000' } },
    bottom: { style: 'thin', color: { rgb: '000000' } },
    left: { style: 'thin', color: { rgb: '000000' } },
    right: { style: 'thin', color: { rgb: '000000' } }
  };

  const partsMap = new Map(parts.map(p => [p.id.trim().toUpperCase(), p]));
  const findPart = (id: string) => partsMap.get(id.trim().toUpperCase());

  // Helper to extract stage parts for a model
  const getStagePartsForModel = (modelId: string, stageId: StageId): ModelStagePart[] => {
    const requiredParts = new Map<string, { qty: number; minLevel: number }>();
    
    const traverse = (currentId: string, currentQty: number, level: number) => {
      if (level > 20) return;
      if (level > 0) {
        const existing = requiredParts.get(currentId);
        if (!existing) {
          requiredParts.set(currentId, { qty: currentQty, minLevel: level });
        } else {
          requiredParts.set(currentId, { qty: existing.qty + currentQty, minLevel: Math.min(existing.minLevel, level) });
        }
      }
      const v1 = modelBom.filter(b => b.modelId === currentId);
      for (const ing of v1) {
        traverse(ing.partId, currentQty * ing.quantity, level + 1);
      }
      const v2 = bomV2.filter(b => b.resultPartId === currentId && (!b.applicableModel || b.applicableModel === modelId));
      for (const ing of v2) {
        traverse(ing.ingredientPartId, currentQty * ing.quantity, level + 1);
      }
    };

    traverse(modelId, 1, 0);

    const partsList: ModelStagePart[] = [];
    const seenPartIds = new Set<string>();

    requiredParts.forEach((info, partId) => {
      const part = findPart(partId);
      const hasIngredients = bomV2.some(b => b.resultPartId === partId && (!b.applicableModel || b.applicableModel === modelId));
      
      let isForStage = false;
      if (stageId === 'LASER') {
        if (!hasIngredients && !part?.skipLaser) isForStage = true;
      } else if (stageId === 'BENDING') {
        if (!hasIngredients && !part?.skipBending) isForStage = true;
      } else if (stageId === 'WELDING') {
        if ((hasIngredients || info.minLevel <= 1) && !part?.skipWelding) isForStage = true;
      } else if (stageId === 'PAINTING') {
        if ((info.minLevel <= 1 || part?.hasPaintingPO) && !part?.skipPainting) isForStage = true;
      } else if (stageId === 'GLAZING') {
        if (info.minLevel <= 1) isForStage = true;
      }

      if (isForStage && !seenPartIds.has(partId)) {
        seenPartIds.add(partId);
        partsList.push({
          partId,
          partName: part?.name || partId,
          norm: info.qty
        });
      }
    });

    // Supplement from sub-POs if model has POs with sub-orders for this stage
    const masterPosForModel = orders.filter(o => !o.masterPoId && o.partId === modelId);
    masterPosForModel.forEach(master => {
      const subOrders = orders.filter(so => so.masterPoId === master.id && so.stageId === stageId);
      subOrders.forEach(sub => {
        if (!seenPartIds.has(sub.partId)) {
          seenPartIds.add(sub.partId);
          const computedNorm = master.targetQuantity > 0 ? Math.round((sub.targetQuantity / master.targetQuantity) * 100) / 100 : 1;
          partsList.push({
            partId: sub.partId,
            partName: findPart(sub.partId)?.name || sub.partId,
            norm: computedNorm || 1
          });
        }
      });
    });

    return partsList;
  };

  const getPoDateStr = (po: ProductionOrder): string => {
    const ts = po.plannedStartTime || po.createdAt;
    return format(new Date(ts), 'yyyy-MM-dd');
  };

  // 1. Tải hệ số quy đổi (HSQĐ) cho từng công đoạn
  // Sheet Cắt Laser: lấy từ tab HSQĐ Laser
  const laserHSQD = storageService.getLaserHSQD();
  const laserHSQDMap = new Map<string, number>();
  laserHSQD.forEach(n => {
    if (n.hsqd > 0) {
      if (n.partId) laserHSQDMap.set(n.partId.trim().toUpperCase(), n.hsqd);
      if (n.partName) laserHSQDMap.set(n.partName.trim().toUpperCase(), n.hsqd);
    }
  });

  // Sheet Chấn/Dập và Hàn: lấy từ tab HSQĐ Chấn/Hàn
  const bwHSQD = storageService.getBendingWeldingHSQD();
  const bwHSQDMap = new Map<string, number>();
  bwHSQD.forEach(n => {
    if (n.hsqd > 0) {
      if (n.partId) bwHSQDMap.set(n.partId.trim().toUpperCase(), n.hsqd);
      if (n.partName) bwHSQDMap.set(n.partName.trim().toUpperCase(), n.hsqd);
    }
  });

  // Sheet Sơn: lấy từ tab Lắp ráp (DCLR)
  const dclrNorms = storageService.getNorms().filter(n => n.stageId === 'DCLR');
  const dclrNormsMap = new Map<string, number>();
  dclrNorms.forEach(n => {
    if (n.partId && n.secondsPerUnit > 0) {
      dclrNormsMap.set(n.partId.trim().toUpperCase(), n.secondsPerUnit);
    }
  });

  // Sheet Dán Kính: lấy từ tab Dán Kính (GLAZING) trong Định mức năng suất
  const glazingNorms = storageService.getNorms().filter(n => n.stageId === 'GLAZING');
  const glazingNormsMap = new Map<string, number>();
  glazingNorms.forEach(n => {
    if (n.partId && n.secondsPerUnit > 0) {
      glazingNormsMap.set(n.partId.trim().toUpperCase(), n.secondsPerUnit);
    }
  });

  const getStageHSQD = (partId: string, partName: string, stageId: StageId, modelId?: string): number => {
    const pId = (partId || '').trim().toUpperCase();
    const pName = (partName || '').trim().toUpperCase();

    if (stageId === 'LASER') {
      if (laserHSQDMap.has(pId)) return laserHSQDMap.get(pId)!;
      if (laserHSQDMap.has(pName)) return laserHSQDMap.get(pName)!;
      const pt = parts.find(p => p.id.toUpperCase() === pId || p.name.toUpperCase() === pName);
      if (pt) {
        if (laserHSQDMap.has(pt.id.toUpperCase())) return laserHSQDMap.get(pt.id.toUpperCase())!;
        if (laserHSQDMap.has(pt.name.toUpperCase())) return laserHSQDMap.get(pt.name.toUpperCase())!;
      }
      return 1;
    }

    if (stageId === 'BENDING' || stageId === 'WELDING') {
      if (bwHSQDMap.has(pId)) return bwHSQDMap.get(pId)!;
      if (bwHSQDMap.has(pName)) return bwHSQDMap.get(pName)!;
      const pt = parts.find(p => p.id.toUpperCase() === pId || p.name.toUpperCase() === pName);
      if (pt) {
        if (bwHSQDMap.has(pt.id.toUpperCase())) return bwHSQDMap.get(pt.id.toUpperCase())!;
        if (bwHSQDMap.has(pt.name.toUpperCase())) return bwHSQDMap.get(pt.name.toUpperCase())!;
      }
      return 1;
    }

    if (stageId === 'GLAZING') {
      // 1. Tra cứu trực tiếp theo mã/tên trong tab Dán Kính của Định mức năng suất
      if (glazingNormsMap.has(pId)) return glazingNormsMap.get(pId)!;
      if (glazingNormsMap.has(pName)) return glazingNormsMap.get(pName)!;

      const cleanId = pId.replace(/^GLZ-OUT-/i, '').replace(/\s*-\s*DÁN KÍNH$/i, '').trim().toUpperCase();
      if (cleanId && glazingNormsMap.has(cleanId)) return glazingNormsMap.get(cleanId)!;

      const pt = parts.find(p => p.id.toUpperCase() === pId || p.name.toUpperCase() === pName || (cleanId && (p.id.toUpperCase() === cleanId || p.name.toUpperCase() === cleanId)));
      if (pt) {
        if (glazingNormsMap.has(pt.id.toUpperCase())) return glazingNormsMap.get(pt.id.toUpperCase())!;
        if (glazingNormsMap.has(pt.name.toUpperCase())) return glazingNormsMap.get(pt.name.toUpperCase())!;
      }

      for (const norm of glazingNorms) {
        if (norm.partId.toUpperCase().trim() === pId || (cleanId && norm.partId.toUpperCase().trim() === cleanId)) {
          return norm.secondsPerUnit;
        }
        const np = parts.find(p => p.id === norm.partId);
        if (np && (np.name.toUpperCase().trim() === pName || np.name.toUpperCase().trim() === cleanId || np.name.toUpperCase().trim() === pId)) {
          return norm.secondsPerUnit;
        }
      }

      if (modelId) {
        const mId = modelId.trim().toUpperCase();
        if (glazingNormsMap.has(mId)) return glazingNormsMap.get(mId)!;
        const mPart = parts.find(p => p.id.toUpperCase() === mId || p.name.toUpperCase() === mId);
        if (mPart) {
          if (glazingNormsMap.has(mPart.id.toUpperCase())) return glazingNormsMap.get(mPart.id.toUpperCase())!;
          if (glazingNormsMap.has(mPart.name.toUpperCase())) return glazingNormsMap.get(mPart.name.toUpperCase())!;
        }
      }

      // Dự phòng nếu người dùng đã lưu ở DCLR
      if (dclrNormsMap.has(pId)) return dclrNormsMap.get(pId)!;
      if (dclrNormsMap.has(pName)) return dclrNormsMap.get(pName)!;
      if (cleanId && dclrNormsMap.has(cleanId)) return dclrNormsMap.get(cleanId)!;
      if (pt) {
        if (dclrNormsMap.has(pt.id.toUpperCase())) return dclrNormsMap.get(pt.id.toUpperCase())!;
        if (dclrNormsMap.has(pt.name.toUpperCase())) return dclrNormsMap.get(pt.name.toUpperCase())!;
      }

      return 1;
    }

    // PAINTING -> Lấy HSQĐ tab Lắp ráp (DCLR)
    if (dclrNormsMap.has(pId)) return dclrNormsMap.get(pId)!;
    if (dclrNormsMap.has(pName)) return dclrNormsMap.get(pName)!;
    const cleanId = pId.replace(/^GLZ-OUT-/i, '').replace(/\s*-\s*DÁN KÍNH$/i, '').trim().toUpperCase();
    if (cleanId && dclrNormsMap.has(cleanId)) return dclrNormsMap.get(cleanId)!;
    const pt = parts.find(p => p.id.toUpperCase() === pId || p.name.toUpperCase() === pName || (cleanId && (p.id.toUpperCase() === cleanId || p.name.toUpperCase() === cleanId)));
    if (pt) {
      if (dclrNormsMap.has(pt.id.toUpperCase())) return dclrNormsMap.get(pt.id.toUpperCase())!;
      if (dclrNormsMap.has(pt.name.toUpperCase())) return dclrNormsMap.get(pt.name.toUpperCase())!;
    }
    for (const norm of dclrNorms) {
      const np = parts.find(p => p.id === norm.partId);
      if (np && (np.name.toUpperCase().trim() === pName || np.name.toUpperCase().trim() === cleanId)) {
        return norm.secondsPerUnit;
      }
    }
    return 1;
  };

  filteredStages.forEach(stageInfo => {
    interface StagePlanUnit {
      planId: string;
      modelId: string;
      modelDisplayName: string;
      planQty: number;
      modelDisplayLabel: string;
      orderDate: number;
      subPOs: ProductionOrder[];
      glazingDateStr?: string;
      glazingComponents?: { partId: string; partName: string; norm: number }[];
      glazingProducedQuantities?: Record<string, number>;
    }

    const stagePlanUnits: StagePlanUnit[] = [];

    // Xác định kho nhận tồn: Cắt Laser, Sơn, Dán Kính nhận kho OUT; Chấn/Dập và Hàn nhận kho IN
    const targetLoc: 'IN' | 'OUT' = (stageInfo.id === 'BENDING' || stageInfo.id === 'WELDING') ? 'IN' : 'OUT';

    const findInventoryForPart = (partId: string, partName: string) => {
      const cleanPartId = partId.trim().toUpperCase();
      const cleanPartName = partName.trim().toUpperCase();
      return inventory
        .filter(i => {
          if (i.stageId !== stageInfo.id || i.location !== targetLoc) return false;
          const iPartId = i.partId?.trim().toUpperCase() || '';
          return iPartId === cleanPartId || iPartId === cleanPartName || (cleanPartId.startsWith('GLZ-OUT-') && iPartId === cleanPartId.replace('GLZ-OUT-', ''));
        })
        .reduce((sum, item) => sum + item.quantity, 0);
    };

    // Global tracking of remaining inventory per part at this stage (consumed chronologically by KHSX priority)
    const remainingInventoryMap = new Map<string, number>();
    parts.forEach(p => {
      const lineInv = findInventoryForPart(p.id, p.name || p.id);
      remainingInventoryMap.set(p.id, lineInv);
    });

    if (stageInfo.id === 'GLAZING') {
      // 1. Lấy toàn bộ các kế hoạch dán kính (GlazingPlans) chưa hoàn thành trong hệ thống
      const allGlzPlans = (options.glazingPlans && options.glazingPlans.length > 0) 
        ? options.glazingPlans 
        : storageService.getGlazingPlans();
      const allGlzNorms = (options.glazingPlanNorms && options.glazingPlanNorms.length > 0)
        ? options.glazingPlanNorms
        : storageService.getGlazingPlanNorms();

      allGlzPlans.forEach(glz => {
        // Bỏ qua các kế hoạch dán kính đã hoàn thành (XONG)
        if (glz.status === 'COMPLETED') return;

        // Hạn của kế hoạch dán kính (targetCompletionTime)
        let dueTs = 0;
        if (typeof glz.targetCompletionTime === 'number' && glz.targetCompletionTime > 0) {
          dueTs = glz.targetCompletionTime;
        } else if (glz.targetCompletionTime) {
          dueTs = new Date(glz.targetCompletionTime).getTime();
        }
        if (!dueTs || isNaN(dueTs)) {
          dueTs = glz.expectedCompletionTime || glz.plannedStartTime || glz.createdAt;
        }

        if (dueTs >= startMs && dueTs <= endMs) {
          const modelName = findPart(glz.modelId)?.name || glz.modelId;
          const planQty = glz.targetQuantity;
          const modelDisplayLabel = `${modelName} (${planQty})`;
          const dueDateStr = format(new Date(dueTs), 'yyyy-MM-dd');

          // Lấy danh sách linh kiện dán kính của model này (định mức số lượng là 1 chi tiết/máy)
          const modelIdClean = glz.modelId?.trim().toLowerCase();
          const modelNameClean = modelName?.trim().toLowerCase();

          let components = allGlzNorms
            .filter(n => n.appliedModel?.trim().toLowerCase() === modelIdClean || n.modelName?.trim().toLowerCase() === modelNameClean || n.id?.trim().toLowerCase() === modelIdClean)
            .map(n => ({
              partId: n.id,
              partName: n.partName || findPart(n.id)?.name || n.id,
              norm: 1 // Định mức số lượng chi tiết/máy là 1
            }));

          if (components.length === 0) {
            const bomComponents = getStagePartsForModel(glz.modelId, 'GLAZING');
            if (bomComponents.length > 0) {
              components = bomComponents.map(b => ({ ...b, norm: 1 }));
            } else {
              components = [{
                partId: glz.modelId,
                partName: modelName,
                norm: 1
              }];
            }
          }

          // Kiểm tra xem kế hoạch này còn linh kiện nào cần in/làm tiếp không
          const producedMap = glz.producedQuantities || {};
          const hasRemainingWork = components.some(c => {
            const targetQty = Math.round((c.norm || 1) * planQty);
            const doneQty = producedMap[c.partId] || 0;
            return targetQty > doneQty;
          });

          if (!hasRemainingWork) return; // Đã in đủ 100% tất cả linh kiện -> Bỏ qua

          stagePlanUnits.push({
            planId: glz.id,
            modelId: glz.modelId,
            modelDisplayName: modelName,
            planQty,
            modelDisplayLabel,
            orderDate: dueTs,
            subPOs: [],
            glazingDateStr: dueDateStr,
            glazingComponents: components,
            glazingProducedQuantities: producedMap
          });
        }
      });
    } else {
      // 2. Master POs có công đoạn này (chỉ lấy kế hoạch chưa hoàn thành)
      const masterPOs = orders.filter(o => !o.masterPoId && !o.id.startsWith('REPAIR') && o.status !== 'COMPLETED');
      masterPOs.forEach(master => {
        const subOrders = orders.filter(so => so.masterPoId === master.id && so.stageId === stageInfo.id && so.status !== 'COMPLETED');
        const masterTs = master.plannedStartTime || master.createdAt;
        
        const hasSubInWindow = subOrders.some(so => {
          const ts = so.plannedStartTime || so.createdAt;
          const rem = Math.max(0, so.targetQuantity - (so.producedQuantity || 0));
          return rem > 0 && ts >= startMs && ts <= endMs;
        });
        const isMasterInWindow = masterTs >= startMs && masterTs <= endMs;

        if (hasSubInWindow || (subOrders.length > 0 && isMasterInWindow)) {
          // Kiểm tra nếu tất cả sub-orders đều đã làm xong 100% thì bỏ qua
          const hasUnfinishedSub = subOrders.some(so => Math.max(0, so.targetQuantity - (so.producedQuantity || 0)) > 0);
          if (!hasUnfinishedSub) return;

          const modelName = findPart(master.partId)?.name || master.partId;
          const planQty = master.targetQuantity;
          const modelDisplayLabel = `${modelName} (${planQty})`;
          const orderDate = masterTs;
          stagePlanUnits.push({
            planId: master.id,
            modelId: master.partId,
            modelDisplayName: modelName,
            planQty,
            modelDisplayLabel,
            orderDate,
            subPOs: subOrders
          });
        }
      });

      // 3. Standalone POs tạo trực tiếp cho công đoạn này (chưa hoàn thành)
      const standalonePOs = orders.filter(o => !o.masterPoId && o.stageId === stageInfo.id && o.status !== 'COMPLETED');
      standalonePOs.forEach(po => {
        const ts = po.plannedStartTime || po.createdAt;
        const rem = Math.max(0, po.targetQuantity - (po.producedQuantity || 0));
        if (rem > 0 && ts >= startMs && ts <= endMs) {
          if (!stagePlanUnits.some(u => u.planId === po.id)) {
            const modelName = findPart(po.partId)?.name || po.partId;
            const planQty = po.targetQuantity;
            stagePlanUnits.push({
              planId: po.id,
              modelId: po.partId,
              modelDisplayName: modelName,
              planQty,
              modelDisplayLabel: `${modelName} (${planQty})`,
              orderDate: ts,
              subPOs: [po]
            });
          }
        }
      });
    }

    // Sort KHSX units chronologically (earliest plan first)
    stagePlanUnits.sort((a, b) => a.orderDate - b.orderDate);

    // Chuẩn bị dữ liệu các dòng chi tiết linh kiện trước để tính toán các dòng tổng kết
    interface PlanRowItem {
      rowData: any[];
      metaList: { colIdx: number; meta: CellPlanMeta }[];
    }

    interface PlanUnitGroup {
      modelDisplayLabel: string;
      partCount: number;
      rows: PlanRowItem[];
    }

    const planUnitGroups: PlanUnitGroup[] = [];
    const colSumUnconverted = new Array(dayDates.length).fill(0);
    const colSumConverted = new Array(dayDates.length).fill(0);

    stagePlanUnits.forEach(planUnit => {
      let stageParts: { partId: string; partName: string; norm: number }[] = [];
      
      if (planUnit.glazingComponents && planUnit.glazingComponents.length > 0) {
        stageParts = [...planUnit.glazingComponents];
      } else {
        stageParts = getStagePartsForModel(planUnit.modelId, stageInfo.id);

        const existingPartIds = new Set(stageParts.map(p => p.partId));
        planUnit.subPOs.forEach(sub => {
          if (!existingPartIds.has(sub.partId)) {
            existingPartIds.add(sub.partId);
            const computedNorm = planUnit.planQty > 0 ? Math.round((sub.targetQuantity / planUnit.planQty) * 100) / 100 : 1;
            stageParts.push({
              partId: sub.partId,
              partName: findPart(sub.partId)?.name || sub.partId,
              norm: computedNorm || 1
            });
          }
        });

        if (stageParts.length === 0) {
          stageParts.push({
            partId: planUnit.modelId,
            partName: planUnit.modelDisplayName,
            norm: 1
          });
        }
      }

      const groupRows: PlanRowItem[] = [];

      stageParts.forEach((p, pIdx) => {
        const rowData: any[] = [];
        rowData.push(pIdx === 0 ? planUnit.modelDisplayLabel : '');
        rowData.push(p.partName);
        rowData.push(p.norm);

        // Lấy Hệ số quy đổi (HSQĐ) của linh kiện này
        const hsqdVal = getStageHSQD(p.partId, p.partName, stageInfo.id, planUnit.modelId);
        rowData.push(hsqdVal > 0 ? hsqdVal : 1);

        // Initial line inventory at this stage for this part
        const totalLineInv = findInventoryForPart(p.partId, p.partName);
        rowData.push(totalLineInv > 0 ? totalLineInv : '');

        const metaList: { colIdx: number; meta: CellPlanMeta }[] = [];

        // Process each date column in chronological order
        dayDates.forEach((d, dIdx) => {
          const dStr = format(d, 'yyyy-MM-dd');
          const colIdx = 5 + dIdx; // 5 columns before dates: Model, Tên LK, Định mức, HSQĐ, Tồn line
          
          const isPastDate = d < startOfDay(baseDate);
          const isTodayDate = isSameDay(d, baseDate);

          let dayKhsxQty = 0;

          if (planUnit.glazingDateStr) {
            // Kế hoạch dán kính: ngày chạy chính là Hạn (targetCompletionTime)
            if (planUnit.glazingDateStr === dStr) {
              const totalTarget = Math.round(p.norm * planUnit.planQty);
              const doneQty = planUnit.glazingProducedQuantities?.[p.partId] || 0;
              dayKhsxQty = Math.max(0, totalTarget - doneQty); // Tự động trừ số lượng đã in/làm
            }
          } else {
            // Find matching sub-POs for this plan unit and part on this date
            const matchingSubPOs = planUnit.subPOs.filter(sub => {
              if (sub.partId !== p.partId) return false;
              const poDate = getPoDateStr(sub);
              return poDate === dStr;
            });

            if (matchingSubPOs.length > 0) {
              dayKhsxQty = matchingSubPOs.reduce((sum, sub) => {
                if (sub.status === 'COMPLETED') return sum;
                const remQty = Math.max(0, sub.targetQuantity - (sub.producedQuantity || 0));
                return sum + remQty; // Tự động trừ số lượng đã hoàn thành của lệnh PO
              }, 0);
            } else if (planUnit.subPOs.length === 0 && format(new Date(planUnit.orderDate), 'yyyy-MM-dd') === dStr) {
              dayKhsxQty = Math.max(0, Math.round(p.norm * planUnit.planQty));
            }
          }

          // Tích lũy vào tổng sản lượng chưa quy đổi và quy đổi của cột ngày này
          if (dayKhsxQty > 0) {
            colSumUnconverted[dIdx] += dayKhsxQty;
            colSumConverted[dIdx] += (dayKhsxQty * (hsqdVal > 0 ? hsqdVal : 1));
          }

          if (dayKhsxQty === 0) {
            rowData.push('');
            metaList.push({ colIdx, meta: { isPastDate, isTodayDate } });
          } else {
            let remInventory = remainingInventoryMap.get(p.partId) || 0;

            if (remInventory >= dayKhsxQty) {
              // CASE 1: Inventory fully covers KHSX (Tồn >= KHSX)
              remInventory -= dayKhsxQty;
              remainingInventoryMap.set(p.partId, remInventory);
              const displayStr = `(${dayKhsxQty})`;
              rowData.push(displayStr);
              metaList.push({
                colIdx,
                meta: {
                  isCoveredFull: true,
                  displayVal: displayStr,
                  originalQty: dayKhsxQty,
                  isPastDate,
                  isTodayDate
                }
              });
            } else if (remInventory > 0 && remInventory < dayKhsxQty) {
              // CASE 2: Inventory partially covers KHSX (0 < Tồn < KHSX)
              const usedInv = remInventory;
              remInventory = 0; // inventory fully exhausted
              remainingInventoryMap.set(p.partId, 0);
              const remainingNeeded = dayKhsxQty - usedInv;
              const displayFormulaText = `${dayKhsxQty} - ${usedInv} = ${remainingNeeded}`;
              rowData.push(displayFormulaText);
              metaList.push({
                colIdx,
                meta: {
                  isPartialDeducted: true,
                  displayVal: displayFormulaText,
                  originalQty: dayKhsxQty,
                  usedInv,
                  remainingNeeded,
                  isPastDate,
                  isTodayDate
                }
              });
            } else {
              // CASE 3: No inventory left (Tồn = 0)
              rowData.push(dayKhsxQty);
              metaList.push({
                colIdx,
                meta: {
                  isPastDate,
                  isTodayDate
                }
              });
            }
          }
        });

        groupRows.push({ rowData, metaList });
      });

      planUnitGroups.push({
        modelDisplayLabel: planUnit.modelDisplayLabel,
        partCount: stageParts.length,
        rows: groupRows
      });
    });

    const wsData: any[][] = [];
    const merges: XLSX.Range[] = [];
    const cellMetaMatrix: Record<string, CellPlanMeta> = {};

    // ==========================================
    // 6 DÒNG TỔNG KẾT ĐẶT TRÊN ĐẦU TRANG (DÒNG 0 - 5):
    // (Đưa lên trên cùng, trước dòng KẾ HOẠCH SẢN XUẤT)
    // 1. Tổng sản phẩm chưa quy đổi
    // 2. Tổng sản phẩm quy đổi / ngày
    // 3. Tổng số người định biên
    // 4. Năng lực sản xuất dây chuyền 8h
    // 5. KHSX/NLSX
    // 6. Số giờ cần hoàn thành KHSX
    // ==========================================
    // Lấy số người định biên riêng của từng công đoạn từ popup nhân sự (ngày xuất báo cáo)
    const stagePlannedWorkers = storageService.getStagePlannedWorkers(reportDateStr);
    const defaultStageWorkers: Record<StageId, number> = {
      LASER: 4,
      BENDING: 8,
      WELDING: 8,
      PAINTING: 8,
      GLAZING: 8,
      DCLR: 8,
    };
    const plannedWorkersCount: number = (stageInfo.id in stagePlannedWorkers)
      ? stagePlannedWorkers[stageInfo.id as keyof typeof stagePlannedWorkers]
      : (defaultStageWorkers[stageInfo.id] || 8);

    const summaryUnconverted: any[] = ['Tổng sản phẩm chưa quy đổi', '', '', '', ''];
    const summaryConverted: any[] = ['Tổng sản phẩm quy đổi / ngày', '', '', '', ''];
    const summaryPeople: any[] = ['Tổng số người định biên', '', '', '', ''];
    const summaryCapacity: any[] = ['Năng lực sản xuất dây chuyền 8h', '', '', '', ''];
    const summaryRatio: any[] = ['KHSX/NLSX', '', '', '', ''];
    const summaryHours: any[] = ['Số giờ cần hoàn thành KHSX', '', '', '', ''];

    dayDates.forEach((d, dIdx) => {
      const dStr = format(d, 'yyyy-MM-dd');
      const unconverted = colSumUnconverted[dIdx] || 0;
      const converted = Math.round((colSumConverted[dIdx] || 0) * 10) / 10;

      // 1. Tổng số người định biên của mỗi sheet công đoạn:
      // Lấy số người định biên của từng công đoạn được nhập ở popup nhân sự (thay cho tổng số người cũ)
      const people = plannedWorkersCount;

      // 2. Năng lực sản xuất dây chuyền 8h thay đổi cách tính theo số người định biên của từng sheet:
      // - Sheet Cắt Laser: Lấy (số người định biên Laser) × 720
      // - Sheet Chấn/Dập: Lấy 172,7 × (số người định biên Chấn/Dập)
      // - Sheet Hàn: Lấy 77,2 × (số người định biên Hàn)
      // - Sheet Sơn: Lấy 127,5 × (số người định biên Sơn)
      // - Sheet Dán Kính: Lấy 242 × (số người định biên Dán Kính)
      let capacity = 0;
      if (stageInfo.id === 'LASER') {
        capacity = Math.round(people * 720);
      } else if (stageInfo.id === 'BENDING') {
        capacity = Math.round(172.7 * people);
      } else if (stageInfo.id === 'WELDING') {
        capacity = Math.round(77.2 * people);
      } else if (stageInfo.id === 'PAINTING') {
        capacity = Math.round(127.5 * people);
      } else if (stageInfo.id === 'GLAZING') {
        capacity = Math.round(242 * people);
      } else {
        capacity = Math.round(people * 23.8);
      }

      // KHSX / NLSX = Tổng sản phẩm quy đổi / Năng lực sản xuất dây chuyền 8h
      const ratioVal = capacity > 0 && converted > 0 ? (converted / capacity) : 0;
      const ratioStr = ratioVal > 0 ? `${Math.round(ratioVal * 100)}%` : '0%';

      // Số giờ cần hoàn thành KHSX = (KHSX / NLSX) * 8
      const hoursNeeded = ratioVal > 0 ? Math.round(ratioVal * 8 * 10) / 10 : 0;

      summaryUnconverted.push(unconverted > 0 ? unconverted : 0);
      summaryConverted.push(converted > 0 ? converted : 0);
      summaryPeople.push(people);
      summaryCapacity.push(capacity);
      summaryRatio.push(ratioStr);
      summaryHours.push(hoursNeeded);
    });

    // Push 6 dòng tổng kết (Hàng 0 đến 5)
    wsData.push(summaryUnconverted);
    merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 4 } });

    wsData.push(summaryConverted);
    merges.push({ s: { r: 1, c: 0 }, e: { r: 1, c: 4 } });

    wsData.push(summaryPeople);
    merges.push({ s: { r: 2, c: 0 }, e: { r: 2, c: 4 } });

    wsData.push(summaryCapacity);
    merges.push({ s: { r: 3, c: 0 }, e: { r: 3, c: 4 } });

    wsData.push(summaryRatio);
    merges.push({ s: { r: 4, c: 0 }, e: { r: 4, c: 4 } });

    wsData.push(summaryHours);
    merges.push({ s: { r: 5, c: 0 }, e: { r: 5, c: 4 } });

    // ==========================================
    // DÒNG 6: Banner Row 1: Tiêu đề nhóm (KẾ HOẠCH SẢN XUẤT...) + Ngày trong tuần
    // ==========================================
    const row6: any[] = [stageInfo.titleName, '', '', '', ''];
    merges.push({ s: { r: 6, c: 0 }, e: { r: 6, c: 4 } });

    dayDates.forEach(d => {
      const dayNum = d.getDay();
      const dayNames = ['Chủ nhật', 'THỨ 2', 'THỨ 3', 'THỨ 4', 'THỨ 5', 'THỨ 6', 'THỨ 7'];
      row6.push(dayNames[dayNum]);
    });
    wsData.push(row6);

    // ==========================================
    // DÒNG 7: Header Row 2: Model, Tên linh kiện, Định mức, Hệ số quy đổi, Tồn line + Ngày
    // ==========================================
    const row7: any[] = ['Model', 'Tên linh kiện', 'Định mức', 'Hệ số quy đổi', 'Tồn line'];
    dayDates.forEach(d => {
      row7.push(format(d, 'dd/MM/yyyy'));
    });
    wsData.push(row7);

    // ==========================================
    // DÒNG 8 TRỞ ĐI: CÁC DÒNG CHI TIẾT SẢN XUẤT
    // ==========================================
    let currentRowIdx = 8;

    if (stagePlanUnits.length === 0) {
      const emptyRow: any[] = ['Không có kế hoạch sản xuất trong khoảng thời gian này', '', '', '', ''];
      dayDates.forEach(() => emptyRow.push(''));
      merges.push({ s: { r: currentRowIdx, c: 0 }, e: { r: currentRowIdx, c: 4 + dayDates.length } });
      wsData.push(emptyRow);
      currentRowIdx++;
    } else {
      planUnitGroups.forEach(group => {
        const startModelRow = currentRowIdx;
        const endModelRow = currentRowIdx + group.partCount - 1;

        if (group.partCount > 1) {
          merges.push({ s: { r: startModelRow, c: 0 }, e: { r: endModelRow, c: 0 } });
        }

        group.rows.forEach(rItem => {
          wsData.push(rItem.rowData);
          rItem.metaList.forEach(m => {
            cellMetaMatrix[`${currentRowIdx}_${m.colIdx}`] = m.meta;
          });
          currentRowIdx++;
        });
      });
    }

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws['!merges'] = merges;

    // Set Column Widths (Generously sized so no component names or formulas are ever truncated!)
    ws['!cols'] = [
      { wch: 18 }, // Model
      { wch: 52 }, // Tên linh kiện (extra wide to guarantee full visibility)
      { wch: 12 }, // Định mức
      { wch: 14 }, // Hệ số quy đổi
      { wch: 14 }, // Tồn line
      ...dayDates.map(() => ({ wch: 24 })) // Date columns (width 24 for full formula display like "500 - 200 = 300")
    ];

    // Set Row Heights
    // Dòng 0-5 (Tổng kết): 24pt
    // Dòng 6 (Banner Tiêu đề): 30pt
    // Dòng 7 (Tiêu đề cột): 26pt
    // Dòng 8 trở đi (Chi tiết): 24pt
    ws['!rows'] = [
      { hpt: 24 }, // Row 0: Tổng sản phẩm chưa quy đổi
      { hpt: 24 }, // Row 1: Tổng sản phẩm quy đổi / ngày
      { hpt: 24 }, // Row 2: Tổng số người định biên
      { hpt: 24 }, // Row 3: NLSX dây chuyền 8h
      { hpt: 24 }, // Row 4: KHSX/NLSX
      { hpt: 24 }, // Row 5: Số giờ cần hoàn thành KHSX
      { hpt: 30 }, // Row 6: Title & Day of week
      { hpt: 26 }, // Row 7: Column Headers
      ...Array.from({ length: Math.max(0, wsData.length - 8) }).map(() => ({ hpt: 24 }))
    ];

    // Styling all cells
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');

    for (let R = 0; R <= range.e.r; ++R) {
      for (let C = 0; C <= range.e.c; ++C) {
        const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
        if (!ws[cellAddress]) ws[cellAddress] = { t: 's', v: '' };

        const isDateCol = C >= 5;
        let dateObj: Date | null = null;
        let isPastDate = false;
        let isTodayDate = false;
        let isSundayDate = false;

        if (isDateCol && dayDates[C - 5]) {
          dateObj = dayDates[C - 5];
          isPastDate = dateObj < startOfDay(baseDate);
          isTodayDate = isSameDay(dateObj, baseDate);
          isSundayDate = dateObj.getDay() === 0;
        }

        if (R >= 0 && R <= 5) {
          // ==========================================
          // 6 DÒNG TỔNG KẾT ĐẶT TRÊN CÙNG
          // ==========================================
          const isLabelCell = C < 5;
          let footerBg = 'F2F2F2';
          if (R === 0) footerBg = 'E2EFDA'; // Dòng 1: Tổng sản phẩm chưa quy đổi (Soft green)
          else if (R === 1) footerBg = 'C6E0B4'; // Dòng 2: Tổng sản phẩm quy đổi / ngày (Deeper green)
          else if (R === 2) footerBg = 'FFF2CC'; // Dòng 3: Tổng số người định biên (Soft amber)
          else if (R === 3) footerBg = 'FFE699'; // Dòng 4: NLSX dây chuyền 8h (Deeper amber)
          else if (R === 4) footerBg = 'E8E5F8'; // Dòng 5: KHSX/NLSX (Soft lavender)
          else if (R === 5) footerBg = 'FCE4D6'; // Dòng 6: Số giờ cần hoàn thành KHSX (Soft coral)

          ws[cellAddress].s = {
            font: { bold: true, sz: 10, color: { rgb: '000000' } },
            fill: { fgColor: { rgb: footerBg } },
            alignment: { horizontal: isLabelCell ? 'left' : 'center', vertical: 'center' },
            border: borderStyle
          };
        } else if (R === 6) {
          // Banner Row 1: Tiêu đề KẾ HOẠCH SẢN XUẤT... (r: 6)
          if (C < 5) {
            ws[cellAddress].s = {
              font: { bold: true, sz: 12, color: { rgb: '000000' } },
              fill: { fgColor: { rgb: 'B4C6E7' } }, // Light Steel Blue
              alignment: { horizontal: 'center', vertical: 'center' },
              border: borderStyle
            };
          } else {
            // Day of week header
            let headerBg = 'FFFF00'; // Bright yellow default
            let fontColor = isSundayDate ? 'FF0000' : '000000'; // Red for Sunday

            if (isPastDate) {
              headerBg = 'FCE4D6'; // Soft peach/orange for past 3 days
              fontColor = isSundayDate ? 'D32F2F' : '4A5568';
            } else if (isTodayDate) {
              headerBg = 'FEF08A'; // Distinctive highlight for today
              fontColor = '000000';
            }

            ws[cellAddress].s = {
              font: { bold: true, sz: 10, color: { rgb: fontColor } },
              fill: { fgColor: { rgb: headerBg } },
              alignment: { horizontal: 'center', vertical: 'center' },
              border: borderStyle
            };
          }
        } else if (R === 7) {
          // Header Row 2 (r: 7)
          if (C < 5) {
            const isTonLine = C === 4;
            const isHsqd = C === 3;
            ws[cellAddress].s = {
              font: { bold: true, sz: 11, color: isTonLine ? { rgb: 'D32F2F' } : isHsqd ? { rgb: '1E3A8A' } : { rgb: '000000' } },
              fill: { fgColor: { rgb: isHsqd ? 'E0E7FF' : 'FFFF00' } }, // Special tint for HSQD column header
              alignment: { horizontal: 'center', vertical: 'center' },
              border: borderStyle
            };
          } else {
            // Date string header (dd/MM/yyyy)
            let headerBg = 'FFFF00';
            let fontColor = isSundayDate ? 'FF0000' : '000000';

            if (isPastDate) {
              headerBg = 'FCE4D6';
              fontColor = isSundayDate ? 'D32F2F' : '4A5568';
            } else if (isTodayDate) {
              headerBg = 'FEF08A';
              fontColor = '000000';
            }

            ws[cellAddress].s = {
              font: { bold: true, sz: 10, color: { rgb: fontColor } },
              fill: { fgColor: { rgb: headerBg } },
              alignment: { horizontal: 'center', vertical: 'center' },
              border: borderStyle
            };
          }
        } else {
          // Data Rows (R >= 8)
          if (C === 0) {
            // Model column (merged vertically)
            ws[cellAddress].s = {
              font: { bold: true, sz: 10, color: { rgb: '000000' } },
              alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
              border: borderStyle
            };
          } else if (C === 1) {
            // Tên linh kiện (Left aligned, never truncated)
            ws[cellAddress].s = {
              font: { sz: 10, color: { rgb: '000000' } },
              alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
              border: borderStyle
            };
          } else if (C === 2) {
            // Định mức (Center)
            ws[cellAddress].s = {
              font: { sz: 10, color: { rgb: '000000' } },
              alignment: { horizontal: 'center', vertical: 'center' },
              border: borderStyle
            };
          } else if (C === 3) {
            // Hệ số quy đổi (Center, bold, distinctive font color)
            ws[cellAddress].s = {
              font: { bold: true, sz: 10, color: { rgb: '1E3A8A' } },
              alignment: { horizontal: 'center', vertical: 'center' },
              border: borderStyle
            };
          } else if (C === 4) {
            // Tồn line
            const cellVal = ws[cellAddress].v;
            const hasInv = typeof cellVal === 'number' && cellVal > 0;
            ws[cellAddress].s = {
              font: { sz: 10, bold: hasInv, color: hasInv ? { rgb: '1B5E20' } : { rgb: '000000' } },
              alignment: { horizontal: 'center', vertical: 'center' },
              border: borderStyle
            };
          } else {
            // Date data cells with inventory deduction recognition
            const cellKey = `${R}_${C}`;
            const meta = cellMetaMatrix[cellKey];

            if (meta?.isCoveredFull) {
              // CASE 1: Full inventory coverage -> Red font in parenthesis
              ws[cellAddress] = {
                t: 's',
                v: meta.displayVal,
                s: {
                  font: { bold: true, sz: 10, color: { rgb: 'D32F2F' } }, // Red font
                  fill: { fgColor: { rgb: 'FFEBEE' } }, // Soft light red background
                  alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                  border: borderStyle
                }
              };
            } else if (meta?.isPartialDeducted) {
              // CASE 2: Partial inventory deduction -> Explicit formula text e.g. "500 - 200 = 300" in Yellow highlight
              ws[cellAddress] = {
                t: 's',
                v: meta.displayVal,
                s: {
                  font: { bold: true, sz: 10, color: { rgb: '78350F' } }, // Deep amber text
                  fill: { fgColor: { rgb: 'FEF08A' } }, // Bright yellow highlight
                  alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                  border: borderStyle
                }
              };
            } else {
              // CASE 3: Normal cell
              const cellVal = ws[cellAddress].v;
              const hasVal = cellVal !== '' && cellVal !== undefined;

              let cellBg: string | undefined = undefined;
              if (isPastDate) {
                cellBg = 'FFF8F5'; // Light tint for past 3 days
              } else if (isTodayDate) {
                cellBg = 'FEFCE8'; // Light tint for today
              }

              ws[cellAddress].s = {
                font: { bold: hasVal, sz: 10, color: { rgb: '000000' } },
                fill: cellBg ? { fgColor: { rgb: cellBg } } : undefined,
                alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                border: borderStyle
              };
            }
          }
        }
      }
    }

    XLSX.utils.book_append_sheet(wb, ws, stageInfo.sheetName);
  });

  const fileDateStr = format(baseDate, 'yyyyMMdd');
  XLSX.writeFile(wb, `KeHoach_SanXuat_ChiTiet_${fileDateStr}.xlsx`);
};
