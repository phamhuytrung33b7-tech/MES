import * as XLSX from 'xlsx-js-style';
import { format } from 'date-fns';
import { Part, StageId, STAGES } from './types';
import { storageService } from './storage';

export type NormTabKey = StageId | 'NESTING' | 'BW_HSQD' | 'LASER_HSQD' | 'GLAZING_STAGE' | 'PAINTING_HSQD';
export type SettingsExportTabKey = 'parts' | 'bom' | 'bom_v2' | 'model_bom' | 'transformations';

const findPartInfo = (parts: Part[], idOrName: string): Part | undefined => {
  if (!idOrName) return undefined;
  const clean = idOrName.trim().toUpperCase();
  return parts.find(
    p => (p.id && p.id.trim().toUpperCase() === clean) || (p.name && p.name.trim().toUpperCase() === clean)
  );
};

/**
 * Tải file Excel mẫu cho từng tab trong phần Định mức Năng suất Sản xuất
 */
export function downloadNormTemplate(tabKey: NormTabKey) {
  const wb = XLSX.utils.book_new();

  if (tabKey === 'PAINTING_HSQD') {
    const headers = [
      'STT',
      'Mã linh kiện',
      'Tên linh kiện',
      'Hệ số quy đổi'
    ];

    const sampleRows = [
      [1, '04-29-08-SHA76210KL-0001', 'Thân tủ R.O SHA76210KL', 1],
      [2, '04-29-08-SHA76210KL-0002', 'Bộ linh kiện kèm theo SHA76210KL', 1.15],
    ];

    const ws = XLSX.utils.aoa_to_sheet([headers, ...sampleRows]);
    ws['!cols'] = [
      { wch: 8 },
      { wch: 32 },
      { wch: 42 },
      { wch: 20 },
    ];
    styleHeaderAndData(ws, headers.length, sampleRows.length);
    XLSX.utils.book_append_sheet(wb, ws, "HSQD_Son");
    XLSX.writeFile(wb, "Mau_HSQD_Son.xlsx");
    return;
  }

  if (tabKey === 'GLAZING_STAGE') {
    const headers = [
      'STT',
      'Mã linh kiện',
      'Tên linh kiện',
      'Hệ số quy đổi'
    ];

    const sampleRows = [
      [1, '04-29-08-SHA76210KL-0040', 'Gói cánh tủ SHA76210KL(BD)', 0.16],
      [2, '04-29-08-SHA76210KL-0041', 'Gói đỉnh tủ SHA76210KL(BD)', 0.13],
      [3, '04-29-08-SHA76211KL-0002', 'Linh kiện dán kính mẫu', 0.16],
      [4, '04-29-08-SHA76212KL-0002', 'Linh kiện dán kính mẫu 2', 0.16],
    ];

    const ws = XLSX.utils.aoa_to_sheet([headers, ...sampleRows]);
    ws['!cols'] = [
      { wch: 8 },
      { wch: 32 },
      { wch: 42 },
      { wch: 20 },
    ];
    styleHeaderAndData(ws, headers.length, sampleRows.length);
    XLSX.utils.book_append_sheet(wb, ws, "HSQD_Dan_Kinh_2");
    XLSX.writeFile(wb, "Mau_HSQD_Dan_Kinh_2.xlsx");
    return;
  }

  if (tabKey === 'NESTING') {
    const headers = [
      'Mã bàn (Nesting ID)',
      'Mã linh kiện',
      'Tên linh kiện',
      'Số lượng / Tấm',
      'Thời gian 1 bàn (Giây)',
      'Model áp dụng'
    ];

    const sampleRows = [
      ['NEST_TH_01', '04-29-08-SHA76210KL-0001', 'Thân tủ R.O SHA76210KL', 2, 120, 'SHA76210KL'],
      ['NEST_TH_01', '04-29-08-SHA76210KL-0002', 'Thanh giằng dưới thân tủ SHA76210KL', 4, 120, 'SHA76210KL'],
      ['NEST_TH_02', '04-29-08-SHA76222KL-0001', 'Khung sườn chính SHA76222KL', 1, 95, 'SHA76222KL'],
      ['NEST_TH_02', '04-29-08-SHA76222KL-0005', 'Tấm nắp che tủ phía sau SHA76222KL', 2, 95, 'SHA76222KL'],
    ];

    const ws = XLSX.utils.aoa_to_sheet([headers, ...sampleRows]);
    ws['!cols'] = [
      { wch: 22 },
      { wch: 30 },
      { wch: 42 },
      { wch: 18 },
      { wch: 24 },
      { wch: 18 },
    ];
    styleHeaderAndData(ws, headers.length, sampleRows.length);
    XLSX.utils.book_append_sheet(wb, ws, "ToHop_Laser");
    XLSX.writeFile(wb, "Mau_DinhMuc_ToHop_Laser.xlsx");
    return;
  }

  if (tabKey === 'LASER_HSQD') {
    const headers = [
      'STT',
      'Mã thành phẩm / Linh kiện',
      'Tên thành phẩm / Linh kiện',
      'HSQĐ'
    ];

    const sampleRows = [
      [1, '04-29-08-SHA76210KL-0001', 'Thân tủ R.O SHA76210KL', 1.0],
      [2, '04-29-08-SHA76210KL-0002', 'Thanh giằng dưới thân tủ SHA76210KL', 0.85],
      [3, '04-29-08-SHA76222KL-0001', 'Khung sườn chính SHA76222KL', 1.2],
    ];

    const ws = XLSX.utils.aoa_to_sheet([headers, ...sampleRows]);
    ws['!cols'] = [
      { wch: 8 },
      { wch: 32 },
      { wch: 40 },
      { wch: 14 },
    ];
    styleHeaderAndData(ws, headers.length, sampleRows.length);
    XLSX.utils.book_append_sheet(wb, ws, "HSQD_Laser");
    XLSX.writeFile(wb, "Mau_HSQD_Laser.xlsx");
    return;
  }

  if (tabKey === 'BW_HSQD') {
    const headers = [
      'STT',
      'Mã thành phẩm',
      'Tên thành phẩm',
      'HSQĐ'
    ];

    const sampleRows = [
      [1, '04-29-08-SHA76210KL-0001', 'Thân tủ R.O SHA76210KL', 1.15],
      [2, '04-29-08-SHA76210KL-0002', 'Thanh giằng dưới thân tủ SHA76210KL', 0.9],
      [3, '04-29-08-SHA76222KL-0001', 'Khung sườn chính SHA76222KL', 1.3],
    ];

    const ws = XLSX.utils.aoa_to_sheet([headers, ...sampleRows]);
    ws['!cols'] = [
      { wch: 8 },
      { wch: 32 },
      { wch: 40 },
      { wch: 14 },
    ];
    styleHeaderAndData(ws, headers.length, sampleRows.length);
    XLSX.utils.book_append_sheet(wb, ws, "HSQD_Chan_Han");
    XLSX.writeFile(wb, "Mau_HSQD_Chan_Han.xlsx");
    return;
  }

  const stageName = STAGES.find(s => s.id === tabKey)?.name || tabKey;
  const isCoefficient = (tabKey === 'DCLR' || tabKey === 'GLAZING');
  
  const headers = [
    'STT',
    'Mã linh kiện',
    'Tên linh kiện',
    isCoefficient ? 'Hệ số quy đổi' : 'Thời gian (Giây/Đơn vị)'
  ];

  let sampleRows: any[][] = [];
  if (tabKey === 'BENDING') {
    sampleRows = [
      [1, '04-29-08-SHA76210KL-0001', 'Thân tủ R.O SHA76210KL', 217.1],
      [2, '04-29-08-SHA76210KL-0002', 'Thanh giằng dưới thân tủ SHA76210KL', 18.2],
      [3, '04-29-08-SHA76210KL-0003', 'Thanh giằng trên thân tủ SHA76210KL', 18.2],
      [4, '04-29-08-SHA76210KL-0004', 'Nẹp đỡ khay trái/phải SHA76210KL', 10.3],
    ];
  } else if (tabKey === 'WELDING') {
    sampleRows = [
      [1, '04-29-08-SHA76210KL-0001', 'Thân tủ R.O SHA76210KL', 180.5],
      [2, '04-29-08-SHA76210KL-0002', 'Cụm khung đỡ hàn SHA76210KL', 95.0],
      [3, '04-29-08-SHA76222KL-0001', 'Khung thân máy hàn SHA76222KL', 145.0],
    ];
  } else if (tabKey === 'PAINTING') {
    sampleRows = [
      [1, '04-29-08-SHA76210KL-0001', 'Thân tủ R.O SHA76210KL', 120.0],
      [2, '04-29-08-SHA76210KL-0002', 'Vỏ ngoài sơn tĩnh điện SHA76210KL', 85.0],
    ];
  } else {
    sampleRows = [
      [1, '04-29-08-SHA76210KL-0001', 'Thân tủ R.O SHA76210KL', 1.0],
      [2, '04-29-08-SHA76210KL-0002', 'Bộ linh kiện kèm theo SHA76210KL', 1.15],
    ];
  }

  const ws = XLSX.utils.aoa_to_sheet([headers, ...sampleRows]);
  ws['!cols'] = [
    { wch: 8 },
    { wch: 32 },
    { wch: 42 },
    { wch: 26 },
  ];
  styleHeaderAndData(ws, headers.length, sampleRows.length);
  XLSX.utils.book_append_sheet(wb, ws, `DinhMuc_${tabKey}`);
  XLSX.writeFile(wb, `Mau_DinhMuc_${stageName.replace(/[\/\s]/g, '_')}.xlsx`);
}

/**
 * Xuất toàn bộ dữ liệu thực tế của từng tab Định mức Năng suất Sản xuất thành file Excel chuẩn mẫu
 * để lưu trữ và có thể upload khôi phục lại lên Supabase sau này.
 */
export function exportNormData(tabKey: NormTabKey, parts: Part[]) {
  const dateSuffix = format(new Date(), 'yyyyMMdd_HHmm');
  const wb = XLSX.utils.book_new();

  if (tabKey === 'NESTING') {
    const nesting = storageService.getLaserNesting();
    if (nesting.length === 0) {
      alert('Tab Tổ hợp Laser chưa có dữ liệu để xuất! Bạn có thể bấm "Tải file mẫu Excel" để lấy file mẫu.');
      return;
    }
    const headers = [
      'Mã bàn (Nesting ID)',
      'Mã linh kiện',
      'Tên linh kiện',
      'Số lượng / Tấm',
      'Thời gian 1 bàn (Giây)',
      'Model áp dụng'
    ];
    const rows = nesting.map(item => {
      const part = findPartInfo(parts, item.partId);
      return [
        item.nestingId,
        part?.id || item.partId,
        part?.name || '',
        item.qtyPerSheet,
        item.secondsPerSheet,
        item.applicableModel || ''
      ];
    });
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    ws['!cols'] = [
      { wch: 22 },
      { wch: 32 },
      { wch: 42 },
      { wch: 18 },
      { wch: 24 },
      { wch: 20 },
    ];
    styleHeaderAndData(ws, headers.length, rows.length);
    XLSX.utils.book_append_sheet(wb, ws, 'ToHop_Laser');
    XLSX.writeFile(wb, `DinhMuc_ToHop_Laser_${dateSuffix}.xlsx`);
    return;
  }

  if (tabKey === 'LASER_HSQD' || tabKey === 'BW_HSQD' || tabKey === 'GLAZING_STAGE' || tabKey === 'PAINTING_HSQD') {
    let list: { partId: string; partName?: string; hsqd: number }[] = [];
    let headers: string[] = [];
    let sheetName = '';
    let filePrefix = '';

    if (tabKey === 'LASER_HSQD') {
      list = storageService.getLaserHSQD();
      headers = ['STT', 'Mã thành phẩm / Linh kiện', 'Tên thành phẩm / Linh kiện', 'HSQĐ'];
      sheetName = 'HSQD_Laser';
      filePrefix = 'HSQD_Laser';
    } else if (tabKey === 'BW_HSQD') {
      list = storageService.getBendingWeldingHSQD();
      headers = ['STT', 'Mã thành phẩm', 'Tên thành phẩm', 'HSQĐ'];
      sheetName = 'HSQD_Chan_Han';
      filePrefix = 'HSQD_Chan_Han';
    } else if (tabKey === 'GLAZING_STAGE') {
      list = storageService.getGlazingStageHSQD();
      headers = ['STT', 'Mã linh kiện', 'Tên linh kiện', 'Hệ số quy đổi'];
      sheetName = 'HSQD_Dan_Kinh_2';
      filePrefix = 'HSQD_Dan_Kinh_2';
    } else {
      list = storageService.getPaintingHSQD();
      headers = ['STT', 'Mã linh kiện', 'Tên linh kiện', 'Hệ số quy đổi'];
      sheetName = 'HSQD_Son';
      filePrefix = 'HSQD_Son';
    }

    if (list.length === 0) {
      alert('Tab hiện tại chưa có dữ liệu để xuất! Bạn có thể bấm "Tải file mẫu Excel" để lấy file mẫu.');
      return;
    }

    const rows = list.map((item, idx) => {
      const part = findPartInfo(parts, item.partId);
      return [
        idx + 1,
        part?.id || item.partId,
        item.partName || part?.name || '',
        item.hsqd
      ];
    });

    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    ws['!cols'] = [
      { wch: 8 },
      { wch: 32 },
      { wch: 44 },
      { wch: 18 },
    ];
    styleHeaderAndData(ws, headers.length, rows.length);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, `${filePrefix}_${dateSuffix}.xlsx`);
    return;
  }

  // Standard stage tabs: BENDING, WELDING, PAINTING, DCLR, GLAZING
  const stageNorms = storageService.getNorms().filter(n => n.stageId === tabKey);
  if (stageNorms.length === 0) {
    alert('Tab hiện tại chưa có dữ liệu để xuất! Bạn có thể bấm "Tải file mẫu Excel" để lấy file mẫu.');
    return;
  }
  const stageName = STAGES.find(s => s.id === tabKey)?.name || tabKey;
  const isCoefficient = (tabKey === 'DCLR' || tabKey === 'GLAZING');
  const headers = [
    'STT',
    'Mã linh kiện',
    'Tên linh kiện',
    isCoefficient ? 'Hệ số quy đổi' : 'Thời gian (Giây/Đơn vị)'
  ];
  const rows = stageNorms.map((norm, idx) => {
    const part = findPartInfo(parts, norm.partId);
    return [
      idx + 1,
      part?.id || norm.partId,
      part?.name || '',
      norm.secondsPerUnit
    ];
  });

  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws['!cols'] = [
    { wch: 8 },
    { wch: 32 },
    { wch: 44 },
    { wch: 24 },
  ];
  styleHeaderAndData(ws, headers.length, rows.length);
  XLSX.utils.book_append_sheet(wb, ws, `DinhMuc_${tabKey}`);
  XLSX.writeFile(wb, `DinhMuc_${stageName.replace(/[\/\s()]/g, '_')}_${dateSuffix}.xlsx`);
}

/**
 * Xuất dữ liệu từng tab trong "Cài đặt linh kiện & Sao lưu" ra file Excel chuẩn mẫu nhập
 * với đầy đủ mã linh kiện, tên linh kiện để có thể upload khôi phục lại lên Supabase.
 */
export function exportSettingsTabExcel(tabKey: SettingsExportTabKey, parts: Part[]) {
  const dateSuffix = format(new Date(), 'yyyyMMdd_HHmm');
  const wb = XLSX.utils.book_new();

  if (tabKey === 'parts') {
    const headers = [
      'Cấp 1',
      'Cấp 2',
      'Cấp 3',
      'Mã linh kiện',
      'Tên linh kiện',
      'ĐVT',
      'Bỏ qua Laser',
      'Bỏ qua Chấn',
      'Bỏ qua Hàn',
      'Bỏ qua Sơn',
      'Có PO tại công đoạn sơn'
    ];
    const sorted = [...parts].sort((a, b) => ((a.level || 1) - (b.level || 1)) || a.id.localeCompare(b.id));
    const rows = sorted.length > 0
      ? sorted.map(p => {
          const lvl = p.level || 1;
          return [
            lvl === 1 ? 1 : '',
            lvl === 2 ? 2 : '',
            lvl === 3 ? 3 : '',
            p.id,
            p.name,
            p.unit || 'Cái',
            p.skipLaser ? 'X' : '',
            p.skipBending ? 'X' : '',
            p.skipWelding ? 'X' : '',
            p.skipPainting ? 'X' : '',
            p.hasPaintingPO ? 'X' : ''
          ];
        })
      : [
          [1, '', '', '04-29-08-SHA76623KL-0017', 'Gói khung tủ SHA76623KL-CT(BD)', 'Bộ', '', '', '', '', ''],
          ['', 2, '', '04-29-08-SHA76623KL-0018', 'Tấm khung sườn SHA76623KL', 'Cái', '', '', 'X', '', ''],
        ];

    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    ws['!cols'] = [
      { wch: 8 },
      { wch: 8 },
      { wch: 8 },
      { wch: 30 },
      { wch: 44 },
      { wch: 10 },
      { wch: 14 },
      { wch: 14 },
      { wch: 14 },
      { wch: 14 },
      { wch: 22 },
    ];
    styleHeaderAndData(ws, headers.length, rows.length);
    XLSX.utils.book_append_sheet(wb, ws, 'DanhMuc_LinhKien');
    XLSX.writeFile(wb, sorted.length > 0 ? `DanhMuc_LinhKien_${dateSuffix}.xlsx` : 'Mau_DanhMuc_LinhKien.xlsx');
    return;
  }

  if (tabKey === 'bom') {
    const bomList = storageService.getBOM();
    const headers = [
      'STT',
      'Mã linh kiện tôn',
      'Mã linh kiện',
      'Tên linh kiện',
      'KL Linh kiện (kg)',
      'KL Phế (kg)'
    ];
    const rows = bomList.length > 0
      ? bomList.map((b, idx) => {
          const child = findPartInfo(parts, b.childPartId);
          return [
            idx + 1,
            b.parentPartId || '',
            child?.id || b.childPartId,
            child?.name || '',
            Number(b.componentWeight || 0),
            Number(b.scrapWeight || 0)
          ];
        })
      : [
          [1, 'TON-0.6MM', '04-29-08-SHA76623KL-0018', 'Tấm khung sườn SHA76623KL', 1.25, 0.15]
        ];

    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    ws['!cols'] = [
      { wch: 8 },
      { wch: 28 },
      { wch: 30 },
      { wch: 42 },
      { wch: 18 },
      { wch: 16 },
    ];
    styleHeaderAndData(ws, headers.length, rows.length);
    XLSX.utils.book_append_sheet(wb, ws, 'DinhMuc_BOM');
    XLSX.writeFile(wb, bomList.length > 0 ? `DinhMuc_SanXuat_BOM_${dateSuffix}.xlsx` : 'Mau_DinhMuc_SanXuat_BOM.xlsx');
    return;
  }

  if (tabKey === 'bom_v2') {
    const bomV2List = storageService.getBOMV2();
    const headers = [
      'STT',
      'Model áp dụng',
      'Mã thành phẩm',
      'Tên thành phẩm',
      'Mã linh kiện',
      'Tên linh kiện',
      'Số lượng'
    ];
    const rows = bomV2List.length > 0
      ? bomV2List.map((b, idx) => {
          const resPart = findPartInfo(parts, b.resultPartId);
          const ingPart = findPartInfo(parts, b.ingredientPartId);
          return [
            idx + 1,
            b.applicableModel || '',
            resPart?.id || b.resultPartId,
            resPart?.name || '',
            ingPart?.id || b.ingredientPartId,
            ingPart?.name || '',
            b.quantity
          ];
        })
      : [
          [1, 'SHA76623KL', '04-29-08-SHA76623KL-0017', 'Gói khung tủ SHA76623KL-CT(BD)', '04-29-08-SHA76623KL-0018', 'Tấm khung sườn SHA76623KL', 1]
        ];

    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    ws['!cols'] = [
      { wch: 8 },
      { wch: 18 },
      { wch: 30 },
      { wch: 40 },
      { wch: 30 },
      { wch: 40 },
      { wch: 12 },
    ];
    styleHeaderAndData(ws, headers.length, rows.length);
    XLSX.utils.book_append_sheet(wb, ws, 'DinhMuc_Han_BOMv2');
    XLSX.writeFile(wb, bomV2List.length > 0 ? `DinhMuc_Han_BOMv2_${dateSuffix}.xlsx` : 'Mau_DinhMuc_Han_BOMv2.xlsx');
    return;
  }

  if (tabKey === 'model_bom') {
    const modelBomList = storageService.getModelBOM();
    const headers = [
      'STT',
      'Mã model',
      'Tên model',
      'Mã linh kiện',
      'Tên linh kiện',
      'Số lượng'
    ];
    const rows = modelBomList.length > 0
      ? modelBomList.map((b, idx) => {
          const mPart = findPartInfo(parts, b.modelId);
          const pPart = findPartInfo(parts, b.partId);
          return [
            idx + 1,
            b.modelId,
            mPart?.name || b.modelId,
            pPart?.id || b.partId,
            pPart?.name || '',
            b.quantity
          ];
        })
      : [
          [1, 'SHA76623KL', 'Máy lọc nước SHA76623KL', '04-29-08-SHA76623KL-0017', 'Gói khung tủ SHA76623KL-CT(BD)', 1]
        ];

    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    ws['!cols'] = [
      { wch: 8 },
      { wch: 22 },
      { wch: 32 },
      { wch: 30 },
      { wch: 42 },
      { wch: 12 },
    ];
    styleHeaderAndData(ws, headers.length, rows.length);
    XLSX.utils.book_append_sheet(wb, ws, 'DinhMuc_Model');
    XLSX.writeFile(wb, modelBomList.length > 0 ? `DinhMuc_Model_${dateSuffix}.xlsx` : 'Mau_DinhMuc_Model.xlsx');
    return;
  }

  if (tabKey === 'transformations') {
    const transList = storageService.getTransformations();
    const headers = [
      'STT',
      'Linh kiện nguồn (L2)',
      'Tên linh kiện nguồn',
      'Model áp dụng',
      'Linh kiện đích (L1)',
      'Tên linh kiện đích',
      'Công đoạn áp dụng'
    ];
    const rows = transList.length > 0
      ? transList.map((t, idx) => {
          const srcPart = findPartInfo(parts, t.sourcePartId);
          const tgtPart = findPartInfo(parts, t.targetPartId);
          const stageName = STAGES.find(s => s.id === t.targetStageId)?.name || t.targetStageId;
          return [
            idx + 1,
            srcPart?.id || t.sourcePartId,
            srcPart?.name || '',
            t.applicableModel || '',
            tgtPart?.id || t.targetPartId,
            tgtPart?.name || '',
            stageName
          ];
        })
      : [
          [1, '04-29-08-SHA76623KL-0018', 'Tấm vách trái SHA76623KL', 'SHA76623KL', '04-29-08-SHA76623KL-0017', 'Gói khung tủ SHA76623KL-CT(BD)', 'Sơn']
        ];

    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    ws['!cols'] = [
      { wch: 8 },
      { wch: 30 },
      { wch: 40 },
      { wch: 18 },
      { wch: 30 },
      { wch: 40 },
      { wch: 20 },
    ];
    styleHeaderAndData(ws, headers.length, rows.length);
    XLSX.utils.book_append_sheet(wb, ws, 'ChuyenDoi_LinhKien');
    XLSX.writeFile(wb, transList.length > 0 ? `ChuyenDoi_LinhKien_${dateSuffix}.xlsx` : 'Mau_ChuyenDoi_LinhKien.xlsx');
    return;
  }
}

function styleHeaderAndData(ws: XLSX.WorkSheet, colCount: number, rowCount: number) {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:Z50');
  
  for (let R = range.s.r; R <= range.e.r; ++R) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellRef = XLSX.utils.encode_cell({ c: C, r: R });
      const cell = ws[cellRef];
      if (!cell) continue;

      if (R === 0) {
        // Header style
        cell.s = {
          fill: { fgColor: { rgb: "1E293B" } }, // Slate 800
          font: { bold: true, sz: 11, color: { rgb: "FFFFFF" } },
          alignment: { horizontal: "center", vertical: "center", wrapText: true },
          border: {
            top: { style: "thin", color: { rgb: "94A3B8" } },
            bottom: { style: "thin", color: { rgb: "94A3B8" } },
            left: { style: "thin", color: { rgb: "94A3B8" } },
            right: { style: "thin", color: { rgb: "94A3B8" } }
          }
        };
      } else {
        // Data row style
        const isNumeric = typeof cell.v === 'number';
        cell.s = {
          font: { sz: 10, color: { rgb: "1E293B" }, bold: isNumeric && C >= 2 },
          alignment: {
            horizontal: C === 0 ? "center" : (isNumeric ? "center" : "left"),
            vertical: "center"
          },
          border: {
            top: { style: "thin", color: { rgb: "E2E8F0" } },
            bottom: { style: "thin", color: { rgb: "E2E8F0" } },
            left: { style: "thin", color: { rgb: "E2E8F0" } },
            right: { style: "thin", color: { rgb: "E2E8F0" } }
          }
        };
      }
    }
  }

  // Row heights
  ws['!rows'] = [
    { hpt: 28 }, // Header
    ...Array(rowCount).fill({ hpt: 22 })
  ];
}
