/**
 * Hybrid Sync Hook & Service: Supabase <-> LocalStorage
 * Designed for WMS / MES offline-first & cloud-resilient manufacturing applications.
 *
 * CHIẾN LƯỢC EVENT-DRIVEN (KHÔNG POLLING NGẦM):
 * 1. Không Polling định kỳ:
 *    - Tuyệt đối không dùng setInterval / setTimeout ngầm định kỳ kéo dữ liệu.
 *    - Giảm Supabase Egress & Log Ingestion về 0 cho các tác vụ ngầm.
 *
 * 2. Kiểm tra Cờ Reset khi Mở App / F5:
 *    - Khi mở ứng dụng hoặc F5, chỉ fetch đúng 1 dòng duy nhất từ bảng system_configs (key = 'data_reset_version').
 *    - Nếu version trên Cloud khác LocalStorage:
 *        + Gọi localStorage.clear(), gán version mới vào LocalStorage.
 *        + PULL toàn bộ dữ liệu sạch từ Supabase về LocalStorage.
 *    - Nếu version trên Cloud trùng khớp với LocalStorage:
 *        + Nếu LocalStorage đã có dữ liệu: KHÔNG fetch thêm bất kỳ bảng nào! Sử dụng trực tiếp LocalStorage (0ms lag, 0 request).
 *        + Nếu LocalStorage trống (lần đầu vào trình duyệt mới): PULL dữ liệu từ Supabase về 1 lần duy nhất.
 *
 * 3. Event-Driven Write-Through:
 *    - Mọi thao tác Ghi (Thêm/Sửa/Xóa đơn hàng, Scan Inbound/Outbound, Tạo nhãn QR, Kế hoạch dán kính)
 *      được kích hoạt trực tiếp từ hành động người dùng (Write-Through vào LocalStorage + Cloud).
 *    - Không tự động PUSH toàn bộ LocalStorage lên Cloud khi khởi động.
 */

import { useState, useEffect, useCallback, useRef } from 'react';
import { supabase, isSupabaseConfigured } from './supabaseClient';

export const RESET_VERSION_KEY = 'data_reset_version';
export const DEFAULT_RESET_VERSION = '1.0.0';

export interface SyncStatus {
  isOnline: boolean;
  isSyncing: boolean;
  lastSynced: number | null;
  error: string | null;
  resetDetected: boolean;
  currentResetVersion: string | null;
}

export interface UserActionPayload {
  table: 'production_orders' | 'wip_inventory' | 'wip_transactions' | 'glazing_plans' | 'qr_labels' | 'system_configs';
  action: 'INSERT' | 'UPDATE' | 'UPSERT' | 'DELETE';
  data: any;
  matchKey?: string;
  matchValue?: any;
}

/**
 * Kiểm tra xem LocalStorage hiện có dữ liệu cốt lõi hay không
 */
export function hasLocalCachedData(): boolean {
  try {
    const pos = localStorage.getItem('wip_production_orders');
    const inv = localStorage.getItem('wip_inventory');
    const parts = localStorage.getItem('wip_parts');
    if (pos && pos.trim() !== '[]' && pos.trim() !== '') return true;
    if (inv && inv.trim() !== '[]' && inv.trim() !== '') return true;
    if (parts && parts.trim() !== '[]' && parts.trim() !== '') return true;
    return false;
  } catch {
    return false;
  }
}

/**
 * 1. Service: Kiểm tra Cờ Reset Hệ Thống (CHỈ FETCH ĐÚNG 1 DÒNG TRONG system_configs)
 * Trả về { resetOccurred, remoteVersion, localVersion, isMatch }
 */
export async function checkAndHandleSystemReset(client = supabase): Promise<{
  resetOccurred: boolean;
  remoteVersion: string;
  localVersion: string | null;
  isMatch: boolean;
}> {
  if (!isSupabaseConfigured) {
    return {
      resetOccurred: false,
      remoteVersion: DEFAULT_RESET_VERSION,
      localVersion: null,
      isMatch: true
    };
  }

  try {
    // Chỉ truy vấn đúng 1 dòng key = 'data_reset_version'
    const { data, error } = await client
      .from('system_configs')
      .select('value')
      .eq('key', RESET_VERSION_KEY)
      .maybeSingle();

    if (error) {
      console.warn('[SyncService] Không thể đọc data_reset_version từ Supabase:', error.message);
      const localVersion = localStorage.getItem(RESET_VERSION_KEY);
      return {
        resetOccurred: false,
        remoteVersion: localVersion || DEFAULT_RESET_VERSION,
        localVersion,
        isMatch: true
      };
    }

    if (!data || data.value === undefined || data.value === null) {
      const localVersion = localStorage.getItem(RESET_VERSION_KEY);
      return {
        resetOccurred: false,
        remoteVersion: DEFAULT_RESET_VERSION,
        localVersion,
        isMatch: true
      };
    }

    // Chuẩn hóa version từ chuỗi hoặc JSON
    let remoteVersion = typeof data.value === 'string'
      ? data.value
      : (typeof data.value === 'object' && data.value.version
          ? String(data.value.version)
          : JSON.stringify(data.value));
    remoteVersion = remoteVersion.replace(/^"|"$/g, '').trim();

    const localVersion = localStorage.getItem(RESET_VERSION_KEY);

    // So sánh phiên bản Cloud và LocalStorage
    if (!localVersion || localVersion !== remoteVersion) {
      console.warn(`[SyncService] CỜ RESET THAY ĐỔI! Remote: "${remoteVersion}" !== Local: "${localVersion}". Tiến hành xóa sạch localStorage...`);

      // XÓA SẠCH LocalStorage của máy trạm khi có reset từ Admin
      localStorage.clear();

      // Lưu lại cờ version mới nhất vào LocalStorage
      localStorage.setItem(RESET_VERSION_KEY, remoteVersion);

      return {
        resetOccurred: true,
        remoteVersion,
        localVersion,
        isMatch: false
      };
    }

    return {
      resetOccurred: false,
      remoteVersion,
      localVersion,
      isMatch: true
    };
  } catch (err: any) {
    console.error('[SyncService] Lỗi khi kiểm tra cờ reset:', err);
    const localVersion = localStorage.getItem(RESET_VERSION_KEY);
    return {
      resetOccurred: false,
      remoteVersion: DEFAULT_RESET_VERSION,
      localVersion,
      isMatch: true
    };
  }
}

/**
 * 2. Service: PULL dữ liệu sạch từ Supabase về LocalStorage
 * Tuyệt đối KHÔNG đẩy dữ liệu từ client lên Supabase trong hàm này
 */
export async function pullLatestFromSupabase(client = supabase): Promise<{
  success: boolean;
  error?: string;
  counts?: Record<string, number>;
}> {
  if (!isSupabaseConfigured) {
    return { success: false, error: 'Chưa cấu hình Supabase Client' };
  }

  try {
    // 2.1 PULL system_configs
    const { data: configs, error: cfgErr } = await client
      .from('system_configs')
      .select('key, value');

    if (cfgErr) throw cfgErr;

    if (configs && configs.length > 0) {
      configs.forEach((item: any) => {
        if (item && item.key) {
          try {
            const valStr = typeof item.value === 'string' ? item.value : JSON.stringify(item.value);
            localStorage.setItem(item.key, valStr);
          } catch {}
        }
      });
    }

    // 2.2 PULL production_orders
    const { data: pos, error: poErr } = await client
      .from('production_orders')
      .select('*')
      .order('created_at', { ascending: false });

    if (poErr) throw poErr;
    if (pos) {
      localStorage.setItem('wip_production_orders', JSON.stringify(pos));
    }

    // 2.3 PULL wip_inventory
    const { data: inv, error: invErr } = await client
      .from('wip_inventory')
      .select('*');

    if (invErr) throw invErr;
    if (inv) {
      localStorage.setItem('wip_inventory', JSON.stringify(inv));
    }

    // 2.4 PULL wip_transactions (giới hạn 35 ngày gần nhất để tiết kiệm RAM & Egress)
    const cutoff35d = Date.now() - 35 * 24 * 60 * 60 * 1000;
    const { data: txs, error: txErr } = await client
      .from('wip_transactions')
      .select('*')
      .gte('timestamp', cutoff35d)
      .order('timestamp', { ascending: false });

    if (txErr) throw txErr;
    if (txs) {
      localStorage.setItem('wip_transactions', JSON.stringify(txs));
    }

    // 2.5 PULL glazing_plans
    const { data: plans, error: gpErr } = await client
      .from('glazing_plans')
      .select('*')
      .order('created_at', { ascending: false });

    if (gpErr) throw gpErr;
    if (plans) {
      localStorage.setItem('wip_glazing_plans', JSON.stringify(plans));
    }

    return {
      success: true,
      counts: {
        configs: configs?.length || 0,
        productionOrders: pos?.length || 0,
        inventory: inv?.length || 0,
        transactions: txs?.length || 0,
        glazingPlans: plans?.length || 0
      }
    };
  } catch (err: any) {
    console.error('[SyncService] Lỗi khi PULL dữ liệu từ Supabase:', err);
    return { success: false, error: err?.message || 'Lỗi mạng khi tải dữ liệu' };
  }
}

/**
 * 3. Service: PUSH thao tác của Người Dùng lên Supabase (Thêm/Sửa/Xóa/Scan)
 * Event-Driven Write-Through: Chỉ gọi trực tiếp khi người dùng bấm nút hoặc scan
 */
export async function pushUserActionToCloud(
  actionPayload: UserActionPayload,
  client = supabase
): Promise<{ success: boolean; error?: string }> {
  if (!isSupabaseConfigured) {
    return { success: false, error: 'Chưa cấu hình Supabase' };
  }

  try {
    const { table, action, data, matchKey, matchValue } = actionPayload;

    if (action === 'INSERT') {
      const { error } = await client.from(table).insert(data);
      if (error) throw error;
    } else if (action === 'UPSERT') {
      const { error } = await client.from(table).upsert(data, {
        onConflict: matchKey || 'id'
      });
      if (error) throw error;
    } else if (action === 'UPDATE') {
      if (!matchKey || matchValue === undefined) {
        throw new Error('Thao tác UPDATE yêu cầu matchKey và matchValue');
      }
      const { error } = await client.from(table).update(data).eq(matchKey, matchValue);
      if (error) throw error;
    } else if (action === 'DELETE') {
      if (!matchKey || matchValue === undefined) {
        throw new Error('Thao tác DELETE yêu cầu matchKey và matchValue');
      }
      const { error } = await client.from(table).delete().eq(matchKey, matchValue);
      if (error) throw error;
    }

    return { success: true };
  } catch (err: any) {
    console.error(`[SyncService] Lỗi khi PUSH action (${actionPayload.action} trên ${actionPayload.table}):`, err);
    return { success: false, error: err?.message || 'Không thể đồng bộ thao tác lên Cloud' };
  }
}

/**
 * 4. Service: Admin kích hoạt làm sạch dữ liệu và cập nhật version reset trên Supabase
 */
export async function triggerAdminCloudReset(
  newVersionTag?: string,
  client = supabase
): Promise<{ success: boolean; newVersion: string; error?: string }> {
  if (!isSupabaseConfigured) {
    return { success: false, newVersion: '', error: 'Chưa cấu hình Supabase' };
  }

  const newVersion = newVersionTag || `reset_${Date.now()}`;

  try {
    // 1. Truncate/Xóa dữ liệu giao dịch trên Cloud
    await Promise.all([
      client.from('production_orders').delete().neq('id', ''),
      client.from('wip_inventory').delete().neq('id', ''),
      client.from('wip_transactions').delete().neq('id', ''),
      client.from('glazing_plans').delete().neq('id', ''),
      client.from('qr_labels').delete().neq('id', '')
    ]);

    // 2. Cập nhật cờ reset phiên bản mới lên system_configs
    const { error: resetErr } = await client.from('system_configs').upsert({
      key: RESET_VERSION_KEY,
      value: newVersion,
      updated_at: new Date().toISOString()
    }, { onConflict: 'key' });

    if (resetErr) throw resetErr;

    // 3. Xóa cache của máy Admin hiện tại và gán version mới
    localStorage.clear();
    localStorage.setItem(RESET_VERSION_KEY, newVersion);

    return { success: true, newVersion };
  } catch (err: any) {
    return { success: false, newVersion, error: err?.message || 'Lỗi khi kích hoạt Cloud Reset' };
  }
}

/**
 * React Hook: useSupabaseSync
 * Chạy 1 lần duy nhất khi mount (App Open / F5).
 * - Fetch đúng 1 dòng key='data_reset_version' từ system_configs.
 * - Nếu version mismatch: Xóa localStorage, PULL dữ liệu mới.
 * - Nếu version khớp và đã có dữ liệu: Không kéo thêm gì cả (0 Cloud request).
 * - Cung cấp hàm refreshFromCloud để người dùng bấm nút làm mới thủ công khi cần.
 * - HOÀN TOÀN KHÔNG CHẠY BẤT KỲ POLLING ĐỊNH KỲ NÀO.
 */
export function useSupabaseSync(onDataReloaded?: () => void) {
  const [status, setStatus] = useState<SyncStatus>({
    isOnline: typeof navigator !== 'undefined' ? navigator.onLine : true,
    isSyncing: false,
    lastSynced: null,
    error: null,
    resetDetected: false,
    currentResetVersion: typeof window !== 'undefined' ? localStorage.getItem(RESET_VERSION_KEY) : null
  });

  const isMountedRef = useRef(true);

  // Khởi động đồng bộ an toàn: Reset Check (1 row) -> Pull nếu cần
  const executeSyncCheck = useCallback(async (forcePull: boolean = false) => {
    if (!isSupabaseConfigured) return;

    setStatus(prev => ({ ...prev, isSyncing: true, error: null }));

    try {
      // BƯỚC 1: Kiểm tra Cờ Reset Hệ Thống (CHỈ 1 DÒNG DUY NHẤT)
      const resetResult = await checkAndHandleSystemReset();

      if (!isMountedRef.current) return;

      if (resetResult.resetOccurred) {
        setStatus(prev => ({
          ...prev,
          resetDetected: true,
          currentResetVersion: resetResult.remoteVersion
        }));
      }

      // BƯỚC 2: Quyết định có cần PULL toàn bộ bảng hay không
      const hasLocal = hasLocalCachedData();
      const needPull = forcePull || resetResult.resetOccurred || !hasLocal;

      if (needPull) {
        const pullResult = await pullLatestFromSupabase();

        if (!isMountedRef.current) return;

        if (pullResult.success) {
          setStatus(prev => ({
            ...prev,
            isSyncing: false,
            lastSynced: Date.now(),
            error: null,
            currentResetVersion: resetResult.remoteVersion
          }));

          if (onDataReloaded) {
            onDataReloaded();
          }
        } else {
          setStatus(prev => ({
            ...prev,
            isSyncing: false,
            error: pullResult.error || 'Lỗi khi đồng bộ dữ liệu'
          }));
        }
      } else {
        // Cờ reset khớp & đã có sẵn LocalStorage: KHÔNG gọi query Supabase nào nữa!
        console.log(`[useSupabaseSync] Cờ data_reset_version khớp ("${resetResult.remoteVersion}"). Sử dụng dữ liệu LocalStorage có sẵn (0 Cloud query ngầm).`);
        setStatus(prev => ({
          ...prev,
          isSyncing: false,
          lastSynced: Date.now(),
          error: null,
          currentResetVersion: resetResult.remoteVersion
        }));
      }
    } catch (err: any) {
      if (!isMountedRef.current) return;
      setStatus(prev => ({
        ...prev,
        isSyncing: false,
        error: err?.message || 'Không thể đồng bộ với Supabase'
      }));
    }
  }, [onDataReloaded]);

  // Hook mount: chạy 1 lần duy nhất khi mở app hoặc F5. Tuyệt đối KHÔNG có setInterval polling!
  useEffect(() => {
    isMountedRef.current = true;

    const handleOnline = () => setStatus(prev => ({ ...prev, isOnline: true }));
    const handleOffline = () => setStatus(prev => ({ ...prev, isOnline: false }));

    window.addEventListener('online', handleOnline);
    window.addEventListener('offline', handleOffline);

    // Chạy kiểm tra reset version khi mở app
    executeSyncCheck(false);

    return () => {
      isMountedRef.current = false;
      window.removeEventListener('online', handleOnline);
      window.removeEventListener('offline', handleOffline);
    };
  }, [executeSyncCheck]);

  // Dispatch thao tác của người dùng lên Supabase (Write-Through)
  const dispatchAction = useCallback(async (action: UserActionPayload) => {
    return await pushUserActionToCloud(action);
  }, []);

  // Nút bấm thủ công cho người dùng: PULL dữ liệu mới nhất
  const refreshFromCloud = useCallback(async () => {
    return await executeSyncCheck(true);
  }, [executeSyncCheck]);

  return {
    ...status,
    refreshFromCloud,
    dispatchAction,
    triggerAdminReset: triggerAdminCloudReset
  };
}
