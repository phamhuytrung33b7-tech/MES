/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export type StageId = 'LASER' | 'BENDING' | 'WELDING' | 'PAINTING' | 'DCLR' | 'GLAZING';

export interface Part {
  id: string;
  name: string;
  unit: string;
  level?: number;
  skipLaser?: boolean;
  skipBending?: boolean;
  skipWelding?: boolean;
  skipPainting?: boolean;
  hasPaintingPO?: boolean;
}

export interface InventoryItem {
  partId: string;
  originalPartId?: string; // Tracks the source identity if transformed
  stageId: StageId;
  location: 'IN' | 'OUT' | 'DEFECT';
  quantity: number;
}

export interface Transaction {
  id: string;
  type: 'STAGE_OUT' | 'STAGE_IN' | 'DEFECT' | 'DISPOSAL';
  partId: string;
  partName?: string; // Cache the name at time of transaction
  originalPartId?: string; // If part was transformed upon entry
  quantity: number;
  stageId: StageId;
  location?: 'IN' | 'OUT' | 'DEFECT';
  timestamp: number;
  qrData?: string;
  sourceStageId?: StageId; // For STAGE_IN, where it came from
  targetStageId?: StageId; // For STAGE_OUT, where it is intended to go
  poId?: string; // Link to production order
  planId?: string; // Link to glazing plan
  printed?: boolean;
  defectReason?: string;
  defectCategory?: string;
  kpiRecorded?: boolean;
}

export interface BOMDefinition {
  parentPartId: string; // Level 3 part (Tôn tấm)
  childPartId: string;  // Level 2 part (Linh kiện)
  componentWeight: number; // KG of component per unit of child
  scrapWeight: number;     // KG of scrap per unit of child
}

export interface BOMDefinitionV2 {
  resultPartId: string;
  ingredientPartId: string;
  quantity: number; // Amount of ingredient per unit of result
  applicableModel?: string;
}

export interface ModelBOMDefinition {
  modelId: string;
  partId: string; // Level 1 part
  quantity: number;
}

export interface ProductionOrder {
  id: string;
  masterPoId?: string; // Links to the top-level Model PO
  partId: string;
  stageId?: StageId; // Which stage this PO is for (null for Model PO)
  targetQuantity: number;
  producedQuantity: number;
  exportedQuantity: number;
  status: 'PENDING' | 'IN_PROGRESS' | 'PAUSED' | 'COMPLETED' | 'PENDING_APPROVAL' | 'REJECTED';
  createdAt: number;
  completedAt?: number;
  plannedStartTime?: number;
  leadTime?: number;
  expectedCompletionTime?: number;
  note?: string;
  isOddPO?: boolean;
  excelRequiredDate?: number;
  createdBy?: string;
  createdByName?: string;
  approvedBy?: string;
  approvedByName?: string;
  approvedAt?: number;
}

export interface ProductivityNorm {
  partId: string;
  stageId: StageId;
  secondsPerUnit: number; // Time in seconds to process 1 unit
}

export interface LaserNesting {
  nestingId: string;
  partId: string;
  qtyPerSheet: number;
  secondsPerUnit: number; // Time for this individual part
  secondsPerSheet: number; // Total time for the full sheet
  applicableModel?: string;
}

export interface BreakTime {
  start: string; // HH:mm
  end: string;   // HH:mm
}

export interface WorkerOverride {
  modelId: string;
  workerCount: number;
}

export interface ShiftConfig {
  stageId: StageId;
  workerCount: number; // Number of people or parallel resources
  workerOverrides?: WorkerOverride[];
  workOnSunday?: boolean; // Toggles working on Sunday
  shifts: {
    start: string; // HH:mm
    end: string;   // HH:mm
  }[];
  breaks: BreakTime[];
}

export interface PartTransformation {
  sourcePartId: string;
  targetPartId: string;
  targetStageId: StageId;
  applicableModel?: string;
}

export interface GlazingConfig {
  partName: string;
  partId: string;
  sourceStageId: StageId;
}

export interface GlazingOutConfig {
  finalPartName: string;
  subParts: {
    partId: string;
    partName: string;
  }[];
}

export interface StagePlannedWorkers {
  LASER: number;
  BENDING: number;
  WELDING: number;
  PAINTING: number;
  GLAZING: number;
}

export interface GlazingPlanNorm {
  id: string;
  partName: string;
  norm: number;
  modelName: string;
  appliedModel?: string; // The model this component belongs to
}

export interface GlazingPlan {
  id: string;
  modelId: string;
  targetQuantity: number;
  targetCompletionTime: number;
  plannedStartTime?: number;
  expectedCompletionTime?: number;
  createdAt: number;
  status: 'PENDING' | 'IN_PROGRESS' | 'COMPLETED';
  producedQuantities?: Record<string, number>; // Track printed quantity per component ID
  note?: string;
  excelRequiredDate?: number;
}

export const STAGES: { id: StageId; name: string; nextStageId?: StageId }[] = [
  { id: 'LASER', name: 'Cắt Laser', nextStageId: 'BENDING' },
  { id: 'BENDING', name: 'Chấn/Dập', nextStageId: 'WELDING' },
  { id: 'WELDING', name: 'Hàn', nextStageId: 'PAINTING' },
  { id: 'PAINTING', name: 'Sơn', nextStageId: 'DCLR' },
  { id: 'DCLR', name: 'Lắp ráp (DCLR)', nextStageId: 'GLAZING' },
  { id: 'GLAZING', name: 'Dán Kính' },
];

export const DEFECT_REASONS = [
  { id: 'OPERATOR', name: 'Lỗi thao tác' },
  { id: 'SUPPLIER', name: 'Lỗi NCC' },
  { id: 'SCRATCH', name: 'Trầy xước / Móp méo' },
  { id: 'DIMENSION', name: 'Sai kích thước / Bản vẽ' },
  { id: 'WELD_FAIL', name: 'Mối hàn lỗi / Thủng' },
  { id: 'PAINT_FAIL', name: 'Bề mặt sơn không đạt / Chảy sơn' },
  { id: 'MATERIAL', name: 'Lỗi vật liệu (Phôi)' },
  { id: 'OTHER', name: 'Lỗi khác (Ghi chú thêm)' },
];

export const INITIAL_PARTS: Part[] = [];

export type View = 
  | 'dashboard' 
  | 'produce' 
  | 'inbound' 
  | 'laser_inbound' 
  | 'welding_inbound' 
  | 'glazing' 
  | 'labels' 
  | 'po' 
  | 'defects' 
  | 'working_hours' 
  | 'norms' 
  | 'odd_po' 
  | 'settings' 
  | 'history'
  | 'users';

export type UserRole = 'ADMIN' | 'USER';

export interface UserAccount {
  id: string;
  username: string;
  password?: string;
  fullName: string;
  role: UserRole;
  defaultStage?: StageId; // Công đoạn phụ trách / mặc định (Cắt Laser, Chấn/Dập, Hàn, Sơn, DCLR, Dán kính)
  allowedViews: View[];
  canApproveOddPO?: boolean;
  isActive: boolean;
  createdAt: number;
  lastLoginAt?: number;
}

/**
 * Lấy công đoạn mặc định của tài khoản theo phân quyền / cài đặt
 * Tự động nhận diện nếu user đã chọn defaultStage, hoặc thông minh suy đoán từ username/họ tên
 */
export function getUserDefaultStage(user: UserAccount | null | undefined): StageId {
  if (!user) return 'LASER';
  if (user.defaultStage && STAGES.some(s => s.id === user.defaultStage)) {
    return user.defaultStage;
  }
  
  // Tự động phân tích tên người dùng / họ tên nếu chưa cài đặt thủ công
  const text = `${user.username || ''} ${user.fullName || ''}`.toLowerCase();
  if (text.includes('laser')) return 'LASER';
  if (text.includes('chan') || text.includes('chấn') || text.includes('dap') || text.includes('dập') || text.includes('bending')) return 'BENDING';
  if (text.includes('han') || text.includes('hàn') || text.includes('weld')) return 'WELDING';
  if (text.includes('son') || text.includes('sơn') || text.includes('paint')) return 'PAINTING';
  if (text.includes('dclr') || text.includes('rap') || text.includes('ráp') || text.includes('assembly')) return 'DCLR';
  if (text.includes('kinh') || text.includes('kính') || text.includes('glazing')) return 'GLAZING';

  // Kiểm tra quyền allowedViews nếu user chỉ có 1 vài tab đặc thù
  if (user.allowedViews) {
    if (user.allowedViews.includes('laser_inbound') && !user.allowedViews.includes('welding_inbound')) return 'LASER';
    if (user.allowedViews.includes('welding_inbound') && !user.allowedViews.includes('laser_inbound')) return 'WELDING';
    if (user.allowedViews.includes('glazing') && user.allowedViews.length <= 4) return 'GLAZING';
  }

  return 'LASER';
}

export interface TabPermissionConfig {
  key: View;
  label: string;
  description: string;
  category: 'production' | 'warehouse' | 'planning' | 'system';
}

export const AVAILABLE_TABS: TabPermissionConfig[] = [
  { key: 'dashboard', label: 'Tổng quan (Dashboard)', description: 'Xem biểu đồ sản lượng, năng suất và tồn kho WIP', category: 'production' },
  { key: 'produce', label: 'Xuất kho (Finish)', description: 'Thực hiện xuất kho công đoạn và in nhãn QR', category: 'warehouse' },
  { key: 'inbound', label: 'Nhập kho (Quét mã)', description: 'Quét camera/máy quét QR nhập kho các công đoạn', category: 'warehouse' },
  { key: 'laser_inbound', label: 'Nhập kho Laser', description: 'Nhập kho thủ công phôi cắt tấm laser', category: 'warehouse' },
  { key: 'welding_inbound', label: 'Nhập kho Hàn', description: 'Nhập kho thủ công phôi công đoạn hàn', category: 'warehouse' },
  { key: 'glazing', label: 'Công đoạn Dán Kính', description: 'Theo dõi và thực hiện kế hoạch dán kính, in nhãn thành phẩm', category: 'production' },
  { key: 'labels', label: 'Danh sách nhãn QR', description: 'Xem và in lại danh sách nhãn mã QR đã tạo', category: 'warehouse' },
  { key: 'po', label: 'Lệnh sản xuất (PO)', description: 'Quản lý, import Excel và tạo các lệnh sản xuất PO', category: 'planning' },
  { key: 'defects', label: 'Báo cáo lỗi (Defect)', description: 'Ghi nhận và thống kê linh kiện lỗi, phế phẩm', category: 'production' },
  { key: 'working_hours', label: 'Thời gian Ca & Nghỉ', description: 'Cấu hình ca làm việc, giờ nghỉ giải lao xưởng', category: 'planning' },
  { key: 'norms', label: 'Định mức Năng suất', description: 'Cài đặt định mức thời gian sản xuất linh kiện', category: 'planning' },
  { key: 'odd_po', label: 'PO dư lẻ', description: 'Quản lý các lệnh sản xuất dư thừa, lẻ theo công đoạn', category: 'planning' },
  { key: 'settings', label: 'Cài đặt & Sao lưu dữ liệu', description: 'Quản lý BOM, danh mục vật tư, sao lưu và đồng bộ Cloud', category: 'system' },
];
