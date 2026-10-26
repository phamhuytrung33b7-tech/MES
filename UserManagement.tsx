/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo } from 'react';
import { 
  Users, 
  UserPlus, 
  ShieldCheck, 
  ShieldAlert, 
  Key, 
  Lock, 
  Unlock, 
  Trash2, 
  Edit3, 
  Check, 
  X, 
  Search, 
  UserCheck, 
  UserX, 
  AlertCircle,
  Eye,
  EyeOff,
  LogOut,
  Sparkles,
  CheckSquare,
  Square,
  Layers,
  LayoutDashboard,
  PackagePlus,
  QrCode,
  Flame,
  ClipboardList,
  Clock,
  Monitor,
  Settings
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { UserAccount, UserRole, View, AVAILABLE_TABS, TabPermissionConfig, StageId, STAGES, getUserDefaultStage } from './types';
import { storageService } from './storage';

interface UserManagementViewProps {
  key?: string;
  currentUser: UserAccount | null;
  onUserUpdated: () => void;
}

// Icon mapping for tabs in permissions list
const TAB_ICONS: Record<string, React.ReactNode> = {
  dashboard: <LayoutDashboard size={18} className="text-orange-500" />,
  produce: <PackagePlus size={18} className="text-indigo-500" />,
  inbound: <QrCode size={18} className="text-emerald-500" />,
  laser_inbound: <Layers size={18} className="text-blue-500" />,
  welding_inbound: <Flame size={18} className="text-amber-500" />,
  glazing: <Square size={18} className="text-cyan-500" />,
  labels: <QrCode size={18} className="text-purple-500" />,
  po: <ClipboardList size={18} className="text-teal-500" />,
  defects: <AlertCircle size={18} className="text-rose-500" />,
  working_hours: <Clock size={18} className="text-blue-600" />,
  norms: <Monitor size={18} className="text-violet-500" />,
  odd_po: <Layers size={18} className="text-amber-600" />,
  settings: <Settings size={18} className="text-gray-600" />,
};

export function UserManagementView({ currentUser, onUserUpdated }: UserManagementViewProps) {
  const [users, setUsers] = useState<UserAccount[]>(() => storageService.getUsers());
  const [searchTerm, setSearchTerm] = useState('');
  const [roleFilter, setRoleFilter] = useState<'ALL' | 'ADMIN' | 'USER'>('ALL');
  
  // Modal states
  const [showAddModal, setShowAddModal] = useState(false);
  const [editingUser, setEditingUser] = useState<UserAccount | null>(null);
  const [deletingUser, setDeletingUser] = useState<UserAccount | null>(null);
  const [toastMsg, setToastMsg] = useState<{ text: string; type: 'success' | 'error' } | null>(null);

  const showToast = (text: string, type: 'success' | 'error' = 'success') => {
    setToastMsg({ text, type });
    setTimeout(() => setToastMsg(null), 3500);
  };

  const reloadUsers = () => {
    const list = storageService.getUsers();
    setUsers(list);
    onUserUpdated();
  };

  // Filtered users
  const filteredUsers = useMemo(() => {
    return users.filter(u => {
      const matchSearch = u.username.toLowerCase().includes(searchTerm.toLowerCase()) || 
                          u.fullName.toLowerCase().includes(searchTerm.toLowerCase());
      const matchRole = roleFilter === 'ALL' || u.role === roleFilter;
      return matchSearch && matchRole;
    });
  }, [users, searchTerm, roleFilter]);

  // Statistics
  const stats = useMemo(() => {
    return {
      total: users.length,
      admin: users.filter(u => u.role === 'ADMIN').length,
      staff: users.filter(u => u.role === 'USER').length,
      active: users.filter(u => u.isActive).length,
    };
  }, [users]);

  const handleToggleActive = (user: UserAccount) => {
    if (user.id === currentUser?.id) {
      showToast('Không thể tự vô hiệu hóa tài khoản của chính bạn!', 'error');
      return;
    }
    const res = storageService.updateUser(user.id, { isActive: !user.isActive });
    if (res.success) {
      showToast(`Đã ${!user.isActive ? 'kích hoạt' : 'khóa'} tài khoản ${user.username}!`);
      reloadUsers();
    } else {
      showToast(res.error || 'Có lỗi xảy ra', 'error');
    }
  };

  const handleDeleteConfirm = () => {
    if (!deletingUser) return;
    const res = storageService.deleteUser(deletingUser.id);
    if (res.success) {
      showToast(`Đã xóa tài khoản ${deletingUser.username} thành công!`);
      setDeletingUser(null);
      reloadUsers();
    } else {
      showToast(res.error || 'Không thể xóa tài khoản', 'error');
    }
  };

  return (
    <div className="space-y-6">
      {/* Toast Notification */}
      <AnimatePresence>
        {toastMsg && (
          <motion.div
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -20 }}
            className={`fixed top-6 right-6 z-50 px-5 py-3 rounded-xl shadow-lg border text-sm font-bold flex items-center gap-2 ${
              toastMsg.type === 'success' 
                ? 'bg-emerald-50 text-emerald-800 border-emerald-200' 
                : 'bg-rose-50 text-rose-800 border-rose-200'
            }`}
          >
            {toastMsg.type === 'success' ? <Check size={18} className="text-emerald-600" /> : <AlertCircle size={18} className="text-rose-600" />}
            <span>{toastMsg.text}</span>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Header & Stats Banner */}
      <div className="bg-gradient-to-r from-orange-500 via-amber-500 to-orange-600 rounded-2xl p-6 text-white shadow-md">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2 text-orange-100 text-xs font-mono uppercase tracking-widest mb-1">
              <ShieldCheck size={16} />
              <span>Phân quyền & Kiểm soát truy cập</span>
            </div>
            <h1 className="text-2xl font-black tracking-tight">Quản trị Người dùng & Phân quyền chức năng</h1>
            <p className="text-sm text-orange-100 mt-1 max-w-2xl">
              Tạo tài khoản người dùng, thiết lập mật khẩu và phân quyền chi tiết các tab chức năng trong xưởng sản xuất. Dữ liệu được đồng bộ trực tuyến lên Supabase.
            </p>
          </div>
          
          <button
            onClick={() => setShowAddModal(true)}
            className="flex items-center gap-2 bg-white text-orange-600 hover:bg-orange-50 px-5 py-3 rounded-xl font-black text-sm transition-all shadow-md active:scale-95 whitespace-nowrap self-start md:self-auto"
          >
            <UserPlus size={18} />
            <span>Thêm tài khoản mới</span>
          </button>
        </div>

        {/* Quick Stats Grid */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mt-6 pt-5 border-t border-white/20">
          <div className="bg-white/10 backdrop-blur-sm px-4 py-3 rounded-xl">
            <div className="text-xs text-orange-100 uppercase tracking-wider font-bold">Tổng tài khoản</div>
            <div className="text-2xl font-black mt-1 font-mono">{stats.total}</div>
          </div>
          <div className="bg-white/10 backdrop-blur-sm px-4 py-3 rounded-xl">
            <div className="text-xs text-orange-100 uppercase tracking-wider font-bold">Quản trị viên (Admin)</div>
            <div className="text-2xl font-black mt-1 font-mono">{stats.admin}</div>
          </div>
          <div className="bg-white/10 backdrop-blur-sm px-4 py-3 rounded-xl">
            <div className="text-xs text-orange-100 uppercase tracking-wider font-bold">Nhân viên (User)</div>
            <div className="text-2xl font-black mt-1 font-mono">{stats.staff}</div>
          </div>
          <div className="bg-white/10 backdrop-blur-sm px-4 py-3 rounded-xl">
            <div className="text-xs text-orange-100 uppercase tracking-wider font-bold">Đang hoạt động</div>
            <div className="text-2xl font-black mt-1 font-mono text-emerald-300">{stats.active}</div>
          </div>
        </div>
      </div>

      {/* Filter and Search Bar */}
      <div className="bg-white p-4 rounded-2xl border border-gray-200 shadow-sm flex flex-col sm:flex-row items-center justify-between gap-4">
        <div className="relative w-full sm:w-80">
          <Search size={18} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-gray-400" />
          <input
            type="text"
            placeholder="Tìm theo username hoặc họ tên..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="w-full pl-10 pr-4 py-2 bg-gray-50 border border-gray-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-orange-500/20 focus:border-orange-500 font-medium"
          />
        </div>

        <div className="flex items-center gap-2 w-full sm:w-auto">
          <span className="text-xs font-bold text-gray-500 uppercase tracking-wider whitespace-nowrap">Vai trò:</span>
          <div className="inline-flex bg-gray-100 p-1 rounded-xl w-full sm:w-auto">
            <button
              onClick={() => setRoleFilter('ALL')}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                roleFilter === 'ALL' ? 'bg-white text-gray-900 shadow-sm' : 'text-gray-500 hover:text-gray-800'
              }`}
            >
              Tất cả ({users.length})
            </button>
            <button
              onClick={() => setRoleFilter('ADMIN')}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                roleFilter === 'ADMIN' ? 'bg-white text-orange-600 shadow-sm' : 'text-gray-500 hover:text-gray-800'
              }`}
            >
              Admin ({stats.admin})
            </button>
            <button
              onClick={() => setRoleFilter('USER')}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                roleFilter === 'USER' ? 'bg-white text-blue-600 shadow-sm' : 'text-gray-500 hover:text-gray-800'
              }`}
            >
              Nhân viên ({stats.staff})
            </button>
          </div>
        </div>
      </div>

      {/* Users Table */}
      <div className="bg-white rounded-2xl border border-gray-200 shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-gray-50/80 border-b border-gray-200 text-xs font-black uppercase text-gray-500 tracking-wider">
                <th className="py-3.5 px-4">Tài khoản & Họ tên</th>
                <th className="py-3.5 px-4">Vai trò</th>
                <th className="py-3.5 px-4">Quyền truy cập chức năng</th>
                <th className="py-3.5 px-4 text-center">Trạng thái</th>
                <th className="py-3.5 px-4 text-right">Thao tác</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100 text-sm">
              {filteredUsers.length === 0 ? (
                <tr>
                  <td colSpan={5} className="py-12 text-center text-gray-400">
                    <Users size={36} className="mx-auto mb-2 opacity-40" />
                    <p className="font-medium">Không tìm thấy tài khoản nào phù hợp điều kiện tìm kiếm.</p>
                  </td>
                </tr>
              ) : (
                filteredUsers.map((user) => {
                  const isCurrent = user.id === currentUser?.id;
                  const isAllViews = user.role === 'ADMIN' || user.allowedViews.length === AVAILABLE_TABS.length;

                  return (
                    <tr key={user.id} className="hover:bg-gray-50/70 transition-colors">
                      {/* User Info */}
                      <td className="py-4 px-4">
                        <div className="flex items-center gap-3">
                          <div className={`w-10 h-10 rounded-xl flex items-center justify-center font-black text-sm uppercase shadow-sm ${
                            user.role === 'ADMIN' 
                              ? 'bg-orange-100 text-orange-700 border border-orange-200' 
                              : 'bg-blue-100 text-blue-700 border border-blue-200'
                          }`}>
                            {user.fullName ? user.fullName.substring(0, 2) : user.username.substring(0, 2)}
                          </div>
                          <div>
                            <div className="flex items-center gap-2">
                              <span className="font-bold text-gray-900">{user.fullName || user.username}</span>
                              {isCurrent && (
                                <span className="px-2 py-0.5 bg-orange-100 text-orange-700 text-[10px] font-black rounded-full border border-orange-200">
                                  Bạn (Đang đăng nhập)
                                </span>
                              )}
                            </div>
                            <div className="text-xs text-gray-500 font-mono flex items-center gap-1.5 mt-0.5">
                              <span>@{user.username}</span>
                              <span className="opacity-30">•</span>
                              <span className="text-gray-400 text-[11px]">
                                {user.password ? '••••••••' : '(Không có mật khẩu)'}
                              </span>
                            </div>
                          </div>
                        </div>
                      </td>

                      {/* Role & Stage Badge */}
                      <td className="py-4 px-4 whitespace-nowrap">
                        <div className="flex flex-col gap-1.5">
                          {user.role === 'ADMIN' ? (
                            <div className="flex flex-col gap-1">
                              <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-black bg-orange-50 text-orange-700 border border-orange-200 w-fit">
                                <ShieldCheck size={14} className="text-orange-600" />
                                Quản trị viên
                              </span>
                              <span className="text-[10px] text-emerald-700 font-bold">
                                ✓ Toàn quyền duyệt PO lẻ
                              </span>
                            </div>
                          ) : (
                            <div className="flex flex-col gap-1">
                              <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-bold bg-blue-50 text-blue-700 border border-blue-200 w-fit">
                                <UserCheck size={14} className="text-blue-600" />
                                Nhân viên
                              </span>
                              {user.canApproveOddPO ? (
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-bold bg-amber-50 text-amber-800 border border-amber-200 w-fit">
                                  <ShieldCheck size={11} className="text-amber-600" />
                                  Có quyền duyệt PO lẻ
                                </span>
                              ) : (
                                <span className="text-[10px] text-gray-400">
                                  Chỉ được tạo (Chờ duyệt)
                                </span>
                              )}
                            </div>
                          )}

                          {/* Stage assignment badge */}
                          {user.defaultStage ? (
                            <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-md text-[11px] font-bold bg-indigo-50 text-indigo-700 border border-indigo-200 w-fit">
                              🎯 {STAGES.find(s => s.id === user.defaultStage)?.name || user.defaultStage}
                            </span>
                          ) : (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] text-gray-500 bg-gray-100 border border-gray-200 w-fit">
                              Tự do ({STAGES.find(s => s.id === getUserDefaultStage(user))?.name || 'Cắt Laser'})
                            </span>
                          )}
                        </div>
                      </td>

                      {/* Allowed Views */}
                      <td className="py-4 px-4">
                        {user.role === 'ADMIN' ? (
                          <div className="flex items-center gap-1.5 text-xs text-emerald-700 font-bold bg-emerald-50 px-2.5 py-1 rounded-lg border border-emerald-200 w-fit">
                            <Sparkles size={14} className="text-emerald-600" />
                            <span>Toàn quyền (Tất cả 13 chức năng + Quản trị tài khoản)</span>
                          </div>
                        ) : (
                          <div className="flex flex-wrap gap-1.5 max-w-md">
                            {isAllViews ? (
                              <span className="text-xs bg-gray-100 text-gray-700 px-2 py-0.5 rounded-md font-bold">
                                Tất cả 13 chức năng
                              </span>
                            ) : user.allowedViews.length === 0 ? (
                              <span className="text-xs text-rose-600 bg-rose-50 px-2 py-0.5 rounded-md font-bold border border-rose-200">
                                Chưa được cấp quyền nào
                              </span>
                            ) : (
                              <>
                                {user.allowedViews.slice(0, 4).map(vKey => {
                                  const tab = AVAILABLE_TABS.find(t => t.key === vKey);
                                  return (
                                    <span key={vKey} className="text-[11px] bg-gray-100 text-gray-700 px-2 py-0.5 rounded-md font-medium border border-gray-200">
                                      {tab ? tab.label.split('(')[0].trim() : vKey}
                                    </span>
                                  );
                                })}
                                {user.allowedViews.length > 4 && (
                                  <span className="text-[11px] bg-orange-50 text-orange-700 px-2 py-0.5 rounded-md font-bold border border-orange-200">
                                    +{user.allowedViews.length - 4} tab khác
                                  </span>
                                )}
                              </>
                            )}
                            {user.canViewAllStageLabels && (
                              <span className="text-[10px] bg-purple-50 text-purple-700 px-1.5 py-0.5 rounded-md font-bold border border-purple-200" title="Được quyền xem nhãn QR của toàn bộ công đoạn như Admin">
                                ✦ Xem QR mọi CĐ
                              </span>
                            )}
                            {user.canApproveOddPO && (
                              <span className="text-[10px] bg-amber-50 text-amber-700 px-1.5 py-0.5 rounded-md font-bold border border-amber-200" title="Được quyền phê duyệt PO lẻ">
                                ✦ Duyệt PO lẻ
                              </span>
                            )}
                          </div>
                        )}
                      </td>

                      {/* Status Toggle */}
                      <td className="py-4 px-4 text-center whitespace-nowrap">
                        <button
                          type="button"
                          onClick={() => handleToggleActive(user)}
                          disabled={isCurrent}
                          title={isCurrent ? 'Không thể khóa tài khoản đang sử dụng' : user.isActive ? 'Bấm để khóa' : 'Bấm để mở khóa'}
                          className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-bold transition-all ${
                            user.isActive 
                              ? 'bg-emerald-50 text-emerald-700 border border-emerald-200 hover:bg-emerald-100' 
                              : 'bg-rose-50 text-rose-700 border border-rose-200 hover:bg-rose-100'
                          } ${isCurrent ? 'opacity-60 cursor-not-allowed' : 'cursor-pointer'}`}
                        >
                          <span className={`w-2 h-2 rounded-full ${user.isActive ? 'bg-emerald-500' : 'bg-rose-500'}`} />
                          <span>{user.isActive ? 'Hoạt động' : 'Đã khóa'}</span>
                        </button>
                      </td>

                      {/* Actions */}
                      <td className="py-4 px-4 text-right whitespace-nowrap">
                        <div className="flex items-center justify-end gap-1.5">
                          <button
                            onClick={() => setEditingUser(user)}
                            title="Chỉnh sửa thông tin & phân quyền"
                            className="p-2 hover:bg-orange-50 text-gray-500 hover:text-orange-600 rounded-lg transition-colors"
                          >
                            <Edit3 size={16} />
                          </button>
                          
                          <button
                            onClick={() => setDeletingUser(user)}
                            disabled={isCurrent || (user.role === 'ADMIN' && stats.admin <= 1)}
                            title={isCurrent ? 'Không thể xóa chính bạn' : 'Xóa tài khoản'}
                            className={`p-2 rounded-lg transition-colors ${
                              isCurrent || (user.role === 'ADMIN' && stats.admin <= 1)
                                ? 'text-gray-300 cursor-not-allowed' 
                                : 'text-gray-500 hover:bg-rose-50 hover:text-rose-600'
                            }`}
                          >
                            <Trash2 size={16} />
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* MODAL: Thêm tài khoản mới */}
      <AnimatePresence>
        {showAddModal && (
          <UserFormModal
            title="Thêm tài khoản người dùng mới"
            onClose={() => setShowAddModal(false)}
            onSave={(userData) => {
              const res = storageService.createUser(userData);
              if (res.success) {
                showToast(`Đã thêm thành công tài khoản "${userData.username}"!`);
                setShowAddModal(false);
                reloadUsers();
              } else {
                showToast(res.error || 'Có lỗi xảy ra', 'error');
              }
            }}
          />
        )}
      </AnimatePresence>

      {/* MODAL: Sửa tài khoản */}
      <AnimatePresence>
        {editingUser && (
          <UserFormModal
            title={`Chỉnh sửa tài khoản: @${editingUser.username}`}
            initialData={editingUser}
            isEditing={true}
            onClose={() => setEditingUser(null)}
            onSave={(userData) => {
              const res = storageService.updateUser(editingUser.id, userData);
              if (res.success) {
                showToast(`Đã cập nhật thông tin tài khoản "${userData.username || editingUser.username}"!`);
                setEditingUser(null);
                reloadUsers();
              } else {
                showToast(res.error || 'Có lỗi xảy ra', 'error');
              }
            }}
          />
        )}
      </AnimatePresence>

      {/* MODAL: Xác nhận xóa tài khoản */}
      <AnimatePresence>
        {deletingUser && (
          <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-2xl max-w-md w-full p-6 shadow-2xl border border-gray-100"
            >
              <div className="w-12 h-12 rounded-2xl bg-rose-100 text-rose-600 flex items-center justify-center mx-auto mb-4">
                <Trash2 size={24} />
              </div>
              <h3 className="text-lg font-black text-center text-gray-900">Xác nhận xóa tài khoản</h3>
              <p className="text-sm text-gray-500 text-center mt-2">
                Bạn có chắc chắn muốn xóa tài khoản <span className="font-bold text-gray-900">@{deletingUser.username}</span> ({deletingUser.fullName}) không? Thao tác này sẽ đồng bộ trực tiếp lên cơ sở dữ liệu Supabase.
              </p>

              <div className="flex gap-3 mt-6">
                <button
                  type="button"
                  onClick={() => setDeletingUser(null)}
                  className="flex-1 py-2.5 px-4 bg-gray-100 hover:bg-gray-200 text-gray-700 font-bold rounded-xl text-sm transition-colors"
                >
                  Hủy bỏ
                </button>
                <button
                  type="button"
                  onClick={handleDeleteConfirm}
                  className="flex-1 py-2.5 px-4 bg-rose-600 hover:bg-rose-700 text-white font-bold rounded-xl text-sm transition-colors shadow-md shadow-rose-600/20"
                >
                  Xóa tài khoản
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

// ==============================================================================
// MODAL: FORM THÊM / SỬA TÀI KHOẢN VÀ CHỌN PHÂN QUYỀN
// ==============================================================================
interface UserFormModalProps {
  title: string;
  initialData?: UserAccount;
  isEditing?: boolean;
  onClose: () => void;
  onSave: (data: Omit<UserAccount, 'id' | 'createdAt'>) => void;
}

function UserFormModal({ title, initialData, isEditing = false, onClose, onSave }: UserFormModalProps) {
  const [username, setUsername] = useState(initialData?.username || '');
  const [fullName, setFullName] = useState(initialData?.fullName || '');
  const [password, setPassword] = useState(initialData?.password || '');
  const [showPassword, setShowPassword] = useState(false);
  const [role, setRole] = useState<UserRole>(initialData?.role || 'USER');
  const [defaultStage, setDefaultStage] = useState<StageId | 'ALL'>(initialData?.defaultStage || 'ALL');
  const [isActive, setIsActive] = useState<boolean>(initialData?.isActive ?? true);
  const [canApproveOddPO, setCanApproveOddPO] = useState<boolean>(initialData?.canApproveOddPO ?? false);
  const [canViewOwnStageExportedLabels, setCanViewOwnStageExportedLabels] = useState<boolean>(initialData?.canViewOwnStageExportedLabels ?? true);
  const [canViewAllStageLabels, setCanViewAllStageLabels] = useState<boolean>(initialData?.canViewAllStageLabels ?? (initialData?.role === 'ADMIN' ? true : false));
  const [allowedViews, setAllowedViews] = useState<View[]>(() => {
    if (initialData?.allowedViews) return [...initialData.allowedViews];
    return ['dashboard', 'produce', 'inbound'];
  });
  const [formError, setFormError] = useState<string | null>(null);

  // Toggle single tab permission
  const handleToggleTab = (tabKey: View) => {
    setAllowedViews(prev => {
      if (prev.includes(tabKey)) {
        return prev.filter(k => k !== tabKey);
      } else {
        return [...prev, tabKey];
      }
    });
  };

  // Presets
  const applyPreset = (preset: 'ALL' | 'WAREHOUSE' | 'PLANNING' | 'NONE') => {
    if (preset === 'ALL') {
      setAllowedViews(AVAILABLE_TABS.map(t => t.key));
    } else if (preset === 'WAREHOUSE') {
      setAllowedViews(['produce', 'inbound', 'laser_inbound', 'welding_inbound', 'labels']);
    } else if (preset === 'PLANNING') {
      setAllowedViews(['dashboard', 'po', 'odd_po', 'norms', 'working_hours', 'glazing']);
    } else if (preset === 'NONE') {
      setAllowedViews([]);
    }
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!username.trim()) {
      setFormError('Vui lòng nhập tên đăng nhập');
      return;
    }
    if (!isEditing && !password.trim()) {
      setFormError('Vui lòng đặt mật khẩu khởi tạo cho tài khoản');
      return;
    }

    onSave({
      username: username.trim(),
      fullName: fullName.trim() || username.trim(),
      password: password.trim(),
      role,
      defaultStage: defaultStage === 'ALL' ? undefined : (defaultStage as StageId),
      allowedViews: role === 'ADMIN' ? AVAILABLE_TABS.map(t => t.key) : allowedViews,
      canApproveOddPO: role === 'ADMIN' ? true : canApproveOddPO,
      canViewOwnStageExportedLabels: role === 'ADMIN' ? true : canViewOwnStageExportedLabels,
      canViewAllStageLabels: role === 'ADMIN' ? true : canViewAllStageLabels,
      isActive,
    });
  };

  return (
    <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4 overflow-y-auto">
      <motion.div
        initial={{ opacity: 0, scale: 0.95 }}
        animate={{ opacity: 1, scale: 1 }}
        exit={{ opacity: 0, scale: 0.95 }}
        className="bg-white rounded-3xl max-w-2xl w-full p-6 sm:p-8 shadow-2xl border border-gray-100 max-h-[90vh] flex flex-col my-auto"
      >
        {/* Header */}
        <div className="flex items-center justify-between pb-4 border-b border-gray-100">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-2xl bg-orange-100 text-orange-600 flex items-center justify-center font-bold">
              <Key size={20} />
            </div>
            <div>
              <h2 className="text-lg font-black text-gray-900">{title}</h2>
              <p className="text-xs text-gray-500">Phân quyền trực quan theo từng màn hình trong hệ thống</p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-2 text-gray-400 hover:text-gray-600 rounded-xl hover:bg-gray-100 transition-colors"
          >
            <X size={20} />
          </button>
        </div>

        {/* Form Body */}
        <form onSubmit={handleSubmit} className="flex-1 overflow-y-auto py-5 space-y-5 pr-1">
          {formError && (
            <div className="p-3.5 bg-rose-50 border border-rose-200 rounded-xl text-rose-700 text-xs font-bold flex items-center gap-2">
              <AlertCircle size={16} />
              <span>{formError}</span>
            </div>
          )}

          {/* Account credentials */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <div>
              <label className="block text-xs font-bold text-gray-700 uppercase tracking-wider mb-1.5">
                Tên đăng nhập (Username) *
              </label>
              <input
                type="text"
                required
                value={username}
                onChange={(e) => setUsername(e.target.value)}
                placeholder="VD: thukho01, son_laser..."
                className="w-full px-3.5 py-2.5 bg-gray-50 border border-gray-200 rounded-xl text-sm font-semibold focus:outline-none focus:ring-2 focus:ring-orange-500/20 focus:border-orange-500"
              />
            </div>

            <div>
              <label className="block text-xs font-bold text-gray-700 uppercase tracking-wider mb-1.5">
                Họ và tên nhân viên
              </label>
              <input
                type="text"
                value={fullName}
                onChange={(e) => setFullName(e.target.value)}
                placeholder="VD: Nguyễn Văn A"
                className="w-full px-3.5 py-2.5 bg-gray-50 border border-gray-200 rounded-xl text-sm font-semibold focus:outline-none focus:ring-2 focus:ring-orange-500/20 focus:border-orange-500"
              />
            </div>

            <div>
              <label className="block text-xs font-bold text-gray-700 uppercase tracking-wider mb-1.5">
                Mật khẩu đăng nhập {isEditing && '(Để trống nếu không đổi)'} *
              </label>
              <div className="relative">
                <input
                  type={showPassword ? 'text' : 'password'}
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="Nhập mật khẩu..."
                  className="w-full px-3.5 py-2.5 bg-gray-50 border border-gray-200 rounded-xl text-sm font-semibold focus:outline-none focus:ring-2 focus:ring-orange-500/20 focus:border-orange-500 pr-10"
                />
                <button
                  type="button"
                  onClick={() => setShowPassword(!showPassword)}
                  className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600"
                >
                  {showPassword ? <EyeOff size={16} /> : <Eye size={16} />}
                </button>
              </div>
            </div>

            <div>
              <label className="block text-xs font-bold text-gray-700 uppercase tracking-wider mb-1.5">
                Vai trò (Role)
              </label>
              <select
                value={role}
                onChange={(e) => {
                  const newRole = e.target.value as UserRole;
                  setRole(newRole);
                  if (newRole === 'ADMIN') {
                    setAllowedViews(AVAILABLE_TABS.map(t => t.key));
                  }
                }}
                className="w-full px-3.5 py-2.5 bg-gray-50 border border-gray-200 rounded-xl text-sm font-bold focus:outline-none focus:ring-2 focus:ring-orange-500/20 focus:border-orange-500"
              >
                <option value="USER">Nhân viên (Chỉ xem các tab được cấp quyền)</option>
                <option value="ADMIN">Quản trị viên (Toàn quyền hệ thống + Quản lý User)</option>
              </select>
            </div>

            {/* Công đoạn phụ trách (Mặc định tự chọn khi Xuất kho / Nhập kho) */}
            <div className="sm:col-span-2">
              <label className="block text-xs font-bold text-blue-900 uppercase tracking-wider mb-1.5 flex items-center gap-1.5">
                <span>🎯 Công đoạn phụ trách (Tự động chọn khi Xuất kho & Nhập kho)</span>
              </label>
              <select
                value={defaultStage}
                onChange={(e) => setDefaultStage(e.target.value as any)}
                className="w-full px-3.5 py-2.5 bg-blue-50/50 border border-blue-200 rounded-xl text-sm font-bold text-gray-900 focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500"
              >
                <option value="ALL">-- Tự do / Tất cả công đoạn (Mặc định: Cắt Laser) --</option>
                {STAGES.map(s => (
                  <option key={s.id} value={s.id}>
                    {s.name} ({s.id})
                  </option>
                ))}
              </select>
              <p className="text-[11px] text-gray-500 mt-1">
                Khi tài khoản này đăng nhập và mở màn hình <b>Xuất kho</b> (ảnh 1) hoặc <b>Nhập kho</b> (ảnh 2), hệ thống sẽ tự động chọn sẵn đúng công đoạn được phân quyền này.
              </p>
            </div>
          </div>

          {/* Active status */}
          <div className="flex items-center justify-between p-3.5 bg-gray-50 rounded-xl border border-gray-200">
            <div>
              <span className="text-xs font-bold text-gray-900">Cho phép tài khoản hoạt động</span>
              <p className="text-[11px] text-gray-500">Nếu tắt, tài khoản sẽ không thể đăng nhập vào hệ thống</p>
            </div>
            <label className="relative inline-flex items-center cursor-pointer">
              <input
                type="checkbox"
                checked={isActive}
                onChange={(e) => setIsActive(e.target.checked)}
                className="sr-only peer"
              />
              <div className="w-11 h-6 bg-gray-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-5 after:w-5 after:transition-all peer-checked:bg-orange-500"></div>
            </label>
          </div>

          {/* Quyền phê duyệt PO dư lẻ (Cắt Laser) */}
          <div className="flex items-center justify-between p-3.5 bg-amber-50/80 rounded-xl border border-amber-200">
            <div>
              <span className="text-xs font-bold text-amber-950 flex items-center gap-1.5">
                <ShieldCheck size={16} className="text-amber-600" />
                <span>Quyền Phê Duyệt PO Dư Lẻ (Cắt Laser)</span>
              </span>
              <p className="text-[11px] text-amber-700 mt-0.5">
                {role === 'ADMIN' 
                  ? 'Quản trị viên (Admin) luôn có toàn quyền phê duyệt trực tiếp mọi lệnh PO dư lẻ' 
                  : 'Cho phép tài khoản nhân viên này duyệt lệnh PO dư lẻ do người khác tạo'}
              </p>
            </div>
            <label className="relative inline-flex items-center cursor-pointer">
              <input
                type="checkbox"
                checked={role === 'ADMIN' || canApproveOddPO}
                disabled={role === 'ADMIN'}
                onChange={(e) => setCanApproveOddPO(e.target.checked)}
                className="sr-only peer"
              />
              <div className="w-11 h-6 bg-gray-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-5 after:w-5 after:transition-all peer-checked:bg-amber-600"></div>
            </label>
          </div>

          {/* Quyền xem nhãn QR do công đoạn mình xuất đi */}
          <div className="flex items-center justify-between p-3.5 bg-blue-50/80 rounded-xl border border-blue-200">
            <div>
              <span className="text-xs font-bold text-blue-950 flex items-center gap-1.5">
                <QrCode size={16} className="text-blue-600" />
                <span>Xem nhãn QR do công đoạn mình xuất ra</span>
              </span>
              <p className="text-[11px] text-blue-700 mt-0.5">
                {role === 'ADMIN' 
                  ? 'Quản trị viên (Admin) luôn có toàn quyền xem nhãn QR của tất cả công đoạn' 
                  : 'Cho phép xem tem do chính công đoạn mình xuất đi để kiểm tra đã in chưa hoặc xuất đi đâu (ngoài tem có đích đến là mình)'}
              </p>
            </div>
            <label className="relative inline-flex items-center cursor-pointer">
              <input
                type="checkbox"
                checked={role === 'ADMIN' || canViewOwnStageExportedLabels}
                disabled={role === 'ADMIN'}
                onChange={(e) => setCanViewOwnStageExportedLabels(e.target.checked)}
                className="sr-only peer"
              />
              <div className="w-11 h-6 bg-gray-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-5 after:w-5 after:transition-all peer-checked:bg-blue-600"></div>
            </label>
          </div>

          {/* Quyền xem nhãn QR của TẤT CẢ các công đoạn (như Admin) */}
          <div className="flex items-center justify-between p-3.5 bg-purple-50/80 rounded-xl border border-purple-200">
            <div>
              <span className="text-xs font-bold text-purple-950 flex items-center gap-1.5">
                <Layers size={16} className="text-purple-600" />
                <span>Xem nhãn QR tất cả các công đoạn (giống Admin)</span>
              </span>
              <p className="text-[11px] text-purple-700 mt-0.5">
                {role === 'ADMIN' 
                  ? 'Quản trị viên (Admin) luôn có toàn quyền xem nhãn QR của tất cả công đoạn' 
                  : 'Cho phép tài khoản này (Tổ trưởng, KCS, Quản lý...) xem và lọc toàn bộ danh sách tem nhãn QR của mọi công đoạn giống tài khoản Admin'}
              </p>
            </div>
            <label className="relative inline-flex items-center cursor-pointer">
              <input
                type="checkbox"
                checked={role === 'ADMIN' || canViewAllStageLabels}
                disabled={role === 'ADMIN'}
                onChange={(e) => setCanViewAllStageLabels(e.target.checked)}
                className="sr-only peer"
              />
              <div className="w-11 h-6 bg-gray-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-5 after:w-5 after:transition-all peer-checked:bg-purple-600"></div>
            </label>
          </div>

          {/* Permissions selection (13 tabs) */}
          <div className="space-y-3 pt-2">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
              <div>
                <label className="text-xs font-black text-gray-800 uppercase tracking-wider">
                  Phân quyền chức năng theo Tab ({allowedViews.length}/{AVAILABLE_TABS.length})
                </label>
                <p className="text-[11px] text-gray-500">Người dùng chỉ nhìn thấy và truy cập được những tab được chọn</p>
              </div>

              {/* Preset buttons */}
              {role !== 'ADMIN' && (
                <div className="flex flex-wrap gap-1.5">
                  <button
                    type="button"
                    onClick={() => applyPreset('ALL')}
                    className="text-[11px] font-bold px-2 py-1 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-lg transition-colors"
                  >
                    Chọn tất cả
                  </button>
                  <button
                    type="button"
                    onClick={() => applyPreset('WAREHOUSE')}
                    className="text-[11px] font-bold px-2 py-1 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 rounded-lg transition-colors"
                  >
                    Kho / Quét mã
                  </button>
                  <button
                    type="button"
                    onClick={() => applyPreset('PLANNING')}
                    className="text-[11px] font-bold px-2 py-1 bg-amber-50 hover:bg-amber-100 text-amber-700 rounded-lg transition-colors"
                  >
                    Kế hoạch & PO
                  </button>
                  <button
                    type="button"
                    onClick={() => applyPreset('NONE')}
                    className="text-[11px] font-bold px-2 py-1 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded-lg transition-colors"
                  >
                    Bỏ chọn
                  </button>
                </div>
              )}
            </div>

            {role === 'ADMIN' ? (
              <div className="p-4 bg-orange-50 border border-orange-200 rounded-2xl flex items-center gap-3">
                <Sparkles size={22} className="text-orange-600 shrink-0" />
                <div className="text-xs text-orange-900 font-medium">
                  <span className="font-bold">Tài khoản Quản trị viên (Admin)</span> tự động có toàn quyền truy cập tất cả 13 chức năng và menu Quản trị tài khoản, không cần cấu hình thủ công.
                </div>
              </div>
            ) : (
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5 max-h-64 overflow-y-auto p-1 border border-gray-100 rounded-2xl bg-gray-50/50">
                {AVAILABLE_TABS.map((tab) => {
                  const isChecked = allowedViews.includes(tab.key);
                  return (
                    <div
                      key={tab.key}
                      onClick={() => handleToggleTab(tab.key)}
                      className={`p-3 rounded-xl border transition-all cursor-pointer flex items-center justify-between ${
                        isChecked 
                          ? 'bg-white border-orange-300 shadow-xs' 
                          : 'bg-white/60 border-gray-200 opacity-60 hover:opacity-100'
                      }`}
                    >
                      <div className="flex items-center gap-2.5 min-w-0 pr-2">
                        <div className="shrink-0 p-1.5 bg-gray-50 rounded-lg border border-gray-100">
                          {TAB_ICONS[tab.key] || <Layers size={16} />}
                        </div>
                        <div className="truncate">
                          <div className="text-xs font-bold text-gray-900 truncate">{tab.label}</div>
                          <div className="text-[10px] text-gray-400 truncate">{tab.description}</div>
                        </div>
                      </div>

                      <div className={`w-5 h-5 rounded-md flex items-center justify-center shrink-0 transition-colors ${
                        isChecked ? 'bg-orange-500 text-white' : 'border border-gray-300 bg-white'
                      }`}>
                        {isChecked && <Check size={14} strokeWidth={3} />}
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>

          {/* Footer buttons */}
          <div className="flex items-center justify-end gap-3 pt-4 border-t border-gray-100">
            <button
              type="button"
              onClick={onClose}
              className="py-2.5 px-5 bg-gray-100 hover:bg-gray-200 text-gray-700 font-bold rounded-xl text-sm transition-colors"
            >
              Hủy bỏ
            </button>
            <button
              type="submit"
              className="py-2.5 px-6 bg-orange-500 hover:bg-orange-600 text-white font-black rounded-xl text-sm transition-all shadow-md shadow-orange-500/20 active:scale-95"
            >
              {isEditing ? 'Lưu thay đổi' : 'Tạo tài khoản'}
            </button>
          </div>
        </form>
      </motion.div>
    </div>
  );
}

// ==============================================================================
// MODAL: ĐĂNG NHẬP (LOGIN MODAL)
// ==============================================================================
interface LoginModalProps {
  onSuccess: (user: UserAccount) => void;
  onCancel?: () => void;
  canCancel?: boolean;
}

export function LoginModal({ onSuccess, onCancel, canCancel = false }: LoginModalProps) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [rememberMe, setRememberMe] = useState(true);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  const handleLogin = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    if (!username.trim()) {
      setErrorMsg('Vui lòng nhập tên đăng nhập');
      return;
    }

    setLoading(true);
    setErrorMsg(null);

    setTimeout(() => {
      const res = storageService.authenticate(username, password, rememberMe);
      setLoading(false);
      if (res.success && res.user) {
        onSuccess(res.user);
      } else {
        setErrorMsg(res.error || 'Đăng nhập không thành công');
      }
    }, 200);
  };

  const quickLogin = (u: string, p: string) => {
    setUsername(u);
    setPassword(p);
    const res = storageService.authenticate(u, p, true);
    if (res.success && res.user) {
      onSuccess(res.user);
    } else {
      setErrorMsg(res.error || 'Đăng nhập không thành công');
    }
  };

  return (
    <div className="fixed inset-0 z-[9999] bg-black/75 backdrop-blur-md flex items-center justify-center p-4">
      <motion.div
        initial={{ opacity: 0, scale: 0.9, y: 10 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.9 }}
        className="bg-white rounded-3xl max-w-md w-full p-8 shadow-2xl border border-gray-100 relative"
      >
        {canCancel && onCancel && (
          <button
            type="button"
            onClick={onCancel}
            className="absolute right-5 top-5 p-2 text-gray-400 hover:text-gray-600 rounded-full hover:bg-gray-100 transition-colors"
          >
            <X size={20} />
          </button>
        )}

        {/* Brand header */}
        <div className="text-center mb-6">
          <div className="w-14 h-14 bg-gradient-to-tr from-orange-500 to-amber-500 text-white rounded-2xl flex items-center justify-center mx-auto mb-3 shadow-lg shadow-orange-500/25">
            <Lock size={28} />
          </div>
          <h2 className="text-2xl font-black text-gray-900 tracking-tight">Đăng nhập Hệ thống WIP</h2>
          <p className="text-xs text-gray-500 mt-1">Xác thực tài khoản để truy cập các tính năng được phân quyền</p>
        </div>

        {errorMsg && (
          <div className="mb-4 p-3 bg-rose-50 border border-rose-200 rounded-xl text-rose-700 text-xs font-bold flex items-center gap-2">
            <AlertCircle size={16} className="shrink-0" />
            <span>{errorMsg}</span>
          </div>
        )}

        <form onSubmit={handleLogin} className="space-y-4">
          <div>
            <label className="block text-xs font-bold text-gray-700 uppercase tracking-wider mb-1.5">
              Tên đăng nhập
            </label>
            <input
              type="text"
              autoFocus
              required
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              placeholder="Nhập tên đăng nhập..."
              className="w-full px-4 py-3 bg-gray-50 border border-gray-200 rounded-xl text-sm font-semibold focus:outline-none focus:ring-2 focus:ring-orange-500/20 focus:border-orange-500"
            />
          </div>

          <div>
            <label className="block text-xs font-bold text-gray-700 uppercase tracking-wider mb-1.5">
              Mật khẩu
            </label>
            <div className="relative">
              <input
                type={showPassword ? 'text' : 'password'}
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder="Nhập mật khẩu..."
                className="w-full px-4 py-3 bg-gray-50 border border-gray-200 rounded-xl text-sm font-semibold focus:outline-none focus:ring-2 focus:ring-orange-500/20 focus:border-orange-500 pr-10"
              />
              <button
                type="button"
                onClick={() => setShowPassword(!showPassword)}
                className="absolute right-3.5 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600"
              >
                {showPassword ? <EyeOff size={18} /> : <Eye size={18} />}
              </button>
            </div>
          </div>

          {/* Remember Login Checkbox */}
          <div className="flex items-center justify-between pt-1">
            <label className="flex items-center gap-2.5 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={rememberMe}
                onChange={(e) => setRememberMe(e.target.checked)}
                className="w-4 h-4 text-orange-500 rounded border-gray-300 focus:ring-orange-500"
              />
              <span className="text-xs font-semibold text-gray-700">
                Ghi nhớ đăng nhập <span className="text-gray-400 font-normal">(Tự động vào lần sau)</span>
              </span>
            </label>
          </div>

          <button
            type="submit"
            disabled={loading}
            className="w-full py-3.5 bg-gradient-to-r from-orange-500 to-amber-500 hover:from-orange-600 hover:to-amber-600 text-white font-black rounded-xl text-sm transition-all shadow-md shadow-orange-500/20 active:scale-98 cursor-pointer mt-2"
          >
            {loading ? 'Đang xác thực...' : 'Đăng nhập ngay'}
          </button>
        </form>

      </motion.div>
    </div>
  );
}

// ==============================================================================
// MODAL: ĐỔI MẬT KHẨU CÁ NHÂN
// ==============================================================================
interface ChangePasswordModalProps {
  currentUser: UserAccount;
  onClose: () => void;
  onSuccess: () => void;
}

export function ChangePasswordModal({ currentUser, onClose, onSuccess }: ChangePasswordModalProps) {
  const [oldPassword, setOldPassword] = useState('');
  const [newPassword, setNewPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (currentUser.password && oldPassword !== currentUser.password) {
      setErrorMsg('Mật khẩu hiện tại không chính xác');
      return;
    }
    if (!newPassword.trim()) {
      setErrorMsg('Vui lòng nhập mật khẩu mới');
      return;
    }
    if (newPassword !== confirmPassword) {
      setErrorMsg('Mật khẩu mới và xác nhận mật khẩu không khớp');
      return;
    }

    const res = storageService.updateUser(currentUser.id, { password: newPassword });
    if (res.success) {
      onSuccess();
    } else {
      setErrorMsg(res.error || 'Không thể đổi mật khẩu');
    }
  };

  return (
    <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4">
      <motion.div
        initial={{ opacity: 0, scale: 0.95 }}
        animate={{ opacity: 1, scale: 1 }}
        exit={{ opacity: 0, scale: 0.95 }}
        className="bg-white rounded-3xl max-w-md w-full p-6 sm:p-8 shadow-2xl border border-gray-100"
      >
        <div className="flex items-center justify-between pb-4 border-b border-gray-100 mb-5">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-2xl bg-orange-100 text-orange-600 flex items-center justify-center font-bold">
              <Key size={20} />
            </div>
            <div>
              <h2 className="text-lg font-black text-gray-900">Đổi mật khẩu</h2>
              <p className="text-xs text-gray-500">Tài khoản: @{currentUser.username}</p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-2 text-gray-400 hover:text-gray-600 rounded-xl hover:bg-gray-100 transition-colors"
          >
            <X size={20} />
          </button>
        </div>

        {errorMsg && (
          <div className="mb-4 p-3 bg-rose-50 border border-rose-200 rounded-xl text-rose-700 text-xs font-bold flex items-center gap-2">
            <AlertCircle size={16} />
            <span>{errorMsg}</span>
          </div>
        )}

        <form onSubmit={handleSubmit} className="space-y-4">
          {currentUser.password && (
            <div>
              <label className="block text-xs font-bold text-gray-700 uppercase tracking-wider mb-1.5">
                Mật khẩu hiện tại
              </label>
              <input
                type="password"
                required
                value={oldPassword}
                onChange={(e) => setOldPassword(e.target.value)}
                placeholder="Nhập mật khẩu cũ..."
                className="w-full px-3.5 py-2.5 bg-gray-50 border border-gray-200 rounded-xl text-sm font-semibold focus:outline-none focus:ring-2 focus:ring-orange-500/20 focus:border-orange-500"
              />
            </div>
          )}

          <div>
            <label className="block text-xs font-bold text-gray-700 uppercase tracking-wider mb-1.5">
              Mật khẩu mới
            </label>
            <input
              type="password"
              required
              value={newPassword}
              onChange={(e) => setNewPassword(e.target.value)}
              placeholder="Nhập mật khẩu mới..."
              className="w-full px-3.5 py-2.5 bg-gray-50 border border-gray-200 rounded-xl text-sm font-semibold focus:outline-none focus:ring-2 focus:ring-orange-500/20 focus:border-orange-500"
            />
          </div>

          <div>
            <label className="block text-xs font-bold text-gray-700 uppercase tracking-wider mb-1.5">
              Xác nhận mật khẩu mới
            </label>
            <input
              type="password"
              required
              value={confirmPassword}
              onChange={(e) => setConfirmPassword(e.target.value)}
              placeholder="Nhập lại mật khẩu mới..."
              className="w-full px-3.5 py-2.5 bg-gray-50 border border-gray-200 rounded-xl text-sm font-semibold focus:outline-none focus:ring-2 focus:ring-orange-500/20 focus:border-orange-500"
            />
          </div>

          <div className="flex gap-3 pt-3">
            <button
              type="button"
              onClick={onClose}
              className="flex-1 py-2.5 px-4 bg-gray-100 hover:bg-gray-200 text-gray-700 font-bold rounded-xl text-sm transition-colors"
            >
              Hủy bỏ
            </button>
            <button
              type="submit"
              className="flex-1 py-2.5 px-4 bg-orange-500 hover:bg-orange-600 text-white font-bold rounded-xl text-sm transition-colors shadow-md shadow-orange-500/20"
            >
              Cập nhật
            </button>
          </div>
        </form>
      </motion.div>
    </div>
  );
}
