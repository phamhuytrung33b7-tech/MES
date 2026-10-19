/**
 * Hybrid Sync Hook & Service: Supabase <-> LocalStorage
 * Designed for WMS / MES offline-first & cloud-resilient manufacturing applications.
 *
 * Core Principles:
 * 1. System Reset Flag (data_reset_version):
 *    - Cloud is authoritative on schema & system resets.
 *    - When Admin resets data on Supabase (e.g. TRUNCATE), `data_reset_version` in `system_configs` is bumped.
 *    - Any client opening the app checks this key first. If local != remote, client executes `localStorage.clear()`,
 *      updates local flag, and PREVENTS any push/upsert of stale data back to Supabase.
 *
 * 2. Hybrid Read-Heavy / Action-Push Pattern (Save Free Tier Egress & API Quotas):
 *    - Startup & Background Polling: PULL-ONLY (Supabase -> LocalStorage).
 *    - Client never automatically pushes local state unconditionally on startup.
 *    - Cloud PUSH occurs exclusively on deliberate user actions (Create, Edit, Delete, Inbound/Outbound Scan).
 *
 * 3. Fallback Offline Resilience:
 *    - If network is offline or Supabase is unreachable, local cached data remains available.
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
 * SQL Scripts for Supabase SQL Editor:
 * ------------------------------------
 * -- 1. Khởi tạo bảng system_configs và giá trị cờ ban đầu:
 * CREATE TABLE IF NOT EXISTS public.system_configs (
 *     key VARCHAR(100) PRIMARY KEY,
 *     value JSONB NOT NULL,
 *     updated_at TIMESTAMPTZ DEFAULT NOW()
 * );
 * 
 * INSERT INTO public.system_configs (key, value)
 * VALUES ('data_reset_version', '"1.0.0"'::jsonb)
 * ON CONFLICT (key) DO NOTHING;
 * 
 * -- 2. Lệnh Admin làm sạch dữ liệu (TRUNCATE) & Kích hoạt xóa cache trên toàn bộ máy trạm:
 * TRUNCATE TABLE public.production_orders, public.wip_inventory, public.wip_transactions, public.glazing_plans, public.qr_labels;
 * 
 * -- Cập nhật cờ reset sang phiên bản mới (ví dụ: dùng timestamp):
 * INSERT INTO public.system_configs (key, value, updated_at)
 * VALUES ('data_reset_version', to_jsonb(to_char(NOW(), 'YYYYMMDD_HH24MISS')), NOW())
 * ON CONFLICT (key) DO UPDATE 
 * SET value = EXCLUDED.value, updated_at = NOW();
 */

/**
 * 1. Service: Kiểm tra Cờ Reset Hệ Thống (System Reset Flag)
 * Trả về true nếu vừa phát hiện reset và đã xóa sạch LocalStorage máy trạm
 */
export async function checkAndHandleSystemReset(client = supabase): Promise<{
  resetOccurred: boolean;
  remoteVersion: string;
  localVersion: string | null;
}> {
  if (!isSupabaseConfigured) {
    return { resetOccurred: false, remoteVersion: DEFAULT_RESET_VERSION, localVersion: null };
  }

  try {
    const { data, error } = await client
      .from('system_configs')
      .select('value')
      .eq('key', RESET_VERSION_KEY)
      .maybeSingle();

    if (error) {
      console.warn('[SyncService] Không thể đọc data_reset_version từ Supabase:', error.message);
      return { resetOccurred: false, remoteVersion: DEFAULT_RESET_VERSION, localVersion: localStorage.getItem(RESET_VERSION_KEY) };
    }

    if (!data || data.value === undefined || data.value === null) {
      return { resetOccurred: false, remoteVersion: DEFAULT_RESET_VERSION, localVersion: localStorage.getItem(RESET_VERSION_KEY) };
    }

    // Chuẩn hóa version từ JSONB
    let remoteVersion = typeof data.value === 'string' ? data.value : JSON.stringify(data.value);
    remoteVersion = remoteVersion.replace(/^"|"$/g, '').trim();

    const localVersion = localStorage.getItem(RESET_VERSION_KEY);

    // Nếu máy trạm chưa có cờ hoặc cờ khác phiên bản trên Cloud
    if (!localVersion || localVersion !== remoteVersion) {
      console.warn(`[SyncService] CỜ RESET THAY ĐỔI! Remote: "${remoteVersion}" !== Local: "${localVersion}". Tiến hành xóa sạch localStorage...`);

      // XÓA SẠCH LocalStorage của máy trạm
      localStorage.clear();

      // Cập nhật lại cờ reset mới nhất
      localStorage.setItem(RESET_VERSION_KEY, remoteVersion);

      return {
        resetOccurred: true,
        remoteVersion,
        localVersion
      };
    }

    return {
      resetOccurred: false,
      remoteVersion,
      localVersion
    };
  } catch (err: any) {
    console.error('[SyncService] Lỗi ngoại lệ khi kiểm tra cờ reset:', err);
    return { resetOccurred: false, remoteVersion: DEFAULT_RESET_VERSION, localVersion: localStorage.getItem(RESET_VERSION_KEY) };
  }
}

/**
 * 2. Service: PULL dữ liệu từ Supabase về ghi đè LocalStorage
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
 * Chỉ gọi khi người dùng thực hiện một hành động cụ thể trên UI
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
 * Tự động kiểm tra cờ reset khi mount, sau đó PULL dữ liệu mới nhất về LocalStorage.
 * Cung cấp hàm dispatchUserAction để PUSH có chủ đích khi người dùng thao tác.
 */
export function useSupabaseSync(onDataReloaded?: () => void) {
  const [status, setStatus] = useState<SyncStatus>({
    isOnline: navigator.onLine,
    isSyncing: false,
    lastSynced: null,
    error: null,
    resetDetected: false,
    currentResetVersion: typeof window !== 'undefined' ? localStorage.getItem(RESET_VERSION_KEY) : null
  });

  const isMountedRef = useRef(true);

  // Khởi động đồng bộ an toàn: Reset Check -> Pull
  const initializeSync = useCallback(async () => {
    if (!isSupabaseConfigured) return;

    setStatus(prev => ({ ...prev, isSyncing: true, error: null }));

    try {
      // BƯỚC 1: Kiểm tra Cờ Reset Hệ Thống trước tiên
      const resetResult = await checkAndHandleSystemReset();

      if (!isMountedRef.current) return;

      if (resetResult.resetOccurred) {
        setStatus(prev => ({
          ...prev,
          resetDetected: true,
          currentResetVersion: resetResult.remoteVersion
        }));
      }

      // BƯỚC 2: PULL dữ liệu mới nhất từ Supabase về LocalStorage
      // (Tuyệt đối KHÔNG push dữ liệu từ localStorage lên)
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
    } catch (err: any) {
      if (!isMountedRef.current) return;
      setStatus(prev => ({
        ...prev,
        isSyncing: false,
        error: err?.message || 'Không thể đồng bộ với Supabase'
      }));
    }
  }, [onDataReloaded]);

  // Hook mount
  useEffect(() => {
    isMountedRef.current = true;

    const handleOnline = () => setStatus(prev => ({ ...prev, isOnline: true }));
    const handleOffline = () => setStatus(prev => ({ ...prev, isOnline: false }));

    window.addEventListener('online', handleOnline);
    window.addEventListener('offline', handleOffline);

    initializeSync();

    return () => {
      isMountedRef.current = false;
      window.removeEventListener('online', handleOnline);
      window.removeEventListener('offline', handleOffline);
    };
  }, [initializeSync]);

  // Dispatch thao tác của người dùng lên Supabase
  const dispatchAction = useCallback(async (action: UserActionPayload) => {
    return await pushUserActionToCloud(action);
  }, []);

  return {
    ...status,
    refreshFromCloud: initializeSync,
    dispatchAction,
    triggerAdminReset: triggerAdminCloudReset
  };
}
