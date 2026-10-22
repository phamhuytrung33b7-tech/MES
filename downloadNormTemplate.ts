import * as XLSX from 'xlsx-js-style';
import { StageId, STAGES } from './types';

/**
 * Tải file Excel mẫu cho từng tab trong phần Định mức Năng suất Sản xuất
 */
export function downloadNormTemplate(tabKey: StageId | 'NESTING' | 'BW_HSQD' | 'LASER_HSQD' | 'GLAZING_STAGE') {
  const wb = XLSX.utils.book_new();

  if (tabKey === 'GLAZING_STAGE') {
    // Mẫu HSQĐ Dán Kính 2 (Tính riêng lẻ cho biểu đồ công đoạn)
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
      { wch: 8 },  // STT
      { wch: 32 }, // Mã linh kiện
      { wch: 42 }, // Tên linh kiện
      { wch: 20 }, // Hệ số quy đổi
    ];
    styleHeaderAndData(ws, headers.length, sampleRows.length);
    XLSX.utils.book_append_sheet(wb, ws, "HSQD_Dan_Kinh_2");
    XLSX.writeFile(wb, "Mau_HSQD_Dan_Kinh_2.xlsx");
    return;
  }

  if (tabKey === 'NESTING') {
    // 1. Mẫu Tổ hợp Laser (Nesting)
    const headers = [
      'Mã bàn (Nesting ID)',
      'Mã linh kiện',
      'Số lượng / Tấm',
      'Thời gian 1 bàn (Giây)',
      'Model áp dụng'
    ];

    const sampleRows = [
      ['NEST_TH_01', '04-29-08-SHA76210KL-0001', 2, 120, 'SHA76210KL'],
      ['NEST_TH_01', '04-29-08-SHA76210KL-0002', 4, 120, 'SHA76210KL'],
      ['NEST_TH_02', '04-29-08-SHA76222KL-0001', 1, 95, 'SHA76222KL'],
      ['NEST_TH_02', '04-29-08-SHA76222KL-0005', 2, 95, 'SHA76222KL'],
    ];

    const ws = XLSX.utils.aoa_to_sheet([headers, ...sampleRows]);
    ws['!cols'] = [
      { wch: 22 }, // Mã bàn
      { wch: 30 }, // Mã linh kiện
      { wch: 18 }, // Số lượng / Tấm
      { wch: 24 }, // Thời gian 1 bàn
      { wch: 18 }, // Model áp dụng
    ];
    styleHeaderAndData(ws, headers.length, sampleRows.length);
    XLSX.utils.book_append_sheet(wb, ws, "ToHop_Laser");
    XLSX.writeFile(wb, "Mau_DinhMuc_ToHop_Laser.xlsx");
    return;
  }

  if (tabKey === 'LASER_HSQD') {
    // 2. Mẫu HSQĐ Laser
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
      { wch: 8 },  // STT
      { wch: 32 }, // Mã
      { wch: 40 }, // Tên
      { wch: 14 }, // HSQĐ
    ];
    styleHeaderAndData(ws, headers.length, sampleRows.length);
    XLSX.utils.book_append_sheet(wb, ws, "HSQD_Laser");
    XLSX.writeFile(wb, "Mau_HSQD_Laser.xlsx");
    return;
  }

  if (tabKey === 'BW_HSQD') {
    // 3. Mẫu HSQĐ Chấn/Hàn
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
      { wch: 8 },  // STT
      { wch: 32 }, // Mã
      { wch: 40 }, // Tên
      { wch: 14 }, // HSQĐ
    ];
    styleHeaderAndData(ws, headers.length, sampleRows.length);
    XLSX.utils.book_append_sheet(wb, ws, "HSQD_Chan_Han");
    XLSX.writeFile(wb, "Mau_HSQD_Chan_Han.xlsx");
    return;
  }

  // 4. Các công đoạn tiêu chuẩn: BENDING, WELDING, PAINTING, DCLR, GLAZING
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
    { wch: 8 },  // STT
    { wch: 32 }, // Mã linh kiện (Mã hàng)
    { wch: 42 }, // Tên linh kiện
    { wch: 26 }, // Thời gian / Hệ số
  ];
  styleHeaderAndData(ws, headers.length, sampleRows.length);
  XLSX.utils.book_append_sheet(wb, ws, `DinhMuc_${tabKey}`);
  XLSX.writeFile(wb, `Mau_DinhMuc_${stageName.replace(/[\/\s]/g, '_')}.xlsx`);
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
