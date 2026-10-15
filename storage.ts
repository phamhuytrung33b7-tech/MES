/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { InventoryItem, Transaction, StageId, STAGES, INITIAL_PARTS, Part, BOMDefinition, BOMDefinitionV2, ProductionOrder, ModelBOMDefinition, ProductivityNorm, LaserNesting, ShiftConfig, PartTransformation } from './types';
import { format, addMilliseconds, setHours, setMinutes, setSeconds, getHours, getMinutes, isBefore, isAfter, startOfDay, addDays } from 'date-fns';
import { supabase, isSupabaseConfigured } from './supabaseClient';

const STORAGE_KEYS = {
  INVENTORY: 'wip_inventory',
  TRANSACTIONS: 'wip_transactions',
  PARTS: 'wip_parts',
  BOM: 'wip_bom',
  BOM_V2: 'wip_bom_v2',
  LABEL_SETTINGS: 'wip_label_settings',
  PRODUCTION_ORDERS: 'wip_production_orders',
  MODEL_BOM: 'wip_model_bom',
  NORMS: 'wip_productivity_norms',
  LASER_NESTING: 'wip_laser_nesting',
  SHIFT_CONFIGS: 'wip_shift_configs',
  TRANSFORMATIONS: 'wip_transformations',
  GLAZING_CONFIGS: 'wip_glazing_configs',
  GLAZING_OUT_CONFIGS: 'wip_glazing_out_configs',
  QUICK_PRINT_PARTS: 'wip_quick_print_parts',
  GLAZING_PLAN_NORMS: 'wip_glazing_plan_norms',
  GLAZING_PLANS: 'wip_glazing_plans',
  PEOPLE_PER_DAY: 'wip_people_per_day',
  MANDAYS_PER_DAY: 'wip_mandays_per_day',
  EXPORT_WEEK_NAME: 'wip_export_week_name',
  HOURLY_PEOPLE_LASER: 'wip_hourly_people_laser',
  HOURLY_PEOPLE_PAINTING: 'wip_hourly_people_painting',
  HOURLY_PEOPLE_GLAZING: 'wip_hourly_people_glazing',
  HOURLY_PEOPLE_BENDING: 'wip_hourly_people_bending',
  HOURLY_PEOPLE_WELDING: 'wip_hourly_people_welding',
  BENDING_WELDING_HSQD: 'wip_bending_welding_hsqd',
  LASER_HSQD: 'wip_laser_hsqd',
  STAGE_PLANNED_WORKERS: 'wip_stage_planned_workers',
};

// In-memory cache to eliminate redundant reads & preserve Free tier egress
const cache: Record<string, any> = {};
const effectivePartIdCache = new Map<string, string>();

// Supabase sync state & listeners
let isSyncing = false;
let lastSyncTime: number | null = null;
let syncError: string | null = null;
const syncListeners = new Set<(status: { connected: boolean; syncing: boolean; lastSynced: number | null; error: string | null }) => void>();

function notifySyncChange() {
  const status = {
    connected: isSupabaseConfigured && !syncError,
    syncing: isSyncing,
    lastSynced: lastSyncTime,
    error: syncError
  };
  syncListeners.forEach(listener => {
    try { listener(status); } catch {}
  });
}

// ---------------------------------------------------------
// Background Async Persistence Queues (Batch & Debounce)
// ---------------------------------------------------------
const pendingConfigSaves = new Map<string, any>();
let configSaveTimer: any = null;

function queueConfigPersist(key: string, value: any) {
  pendingConfigSaves.set(key, value);
  if (!configSaveTimer) {
    configSaveTimer = setTimeout(flushPendingConfigSaves, 400);
  }
}

async function flushPendingConfigSaves() {
  configSaveTimer = null;
  if (!isSupabaseConfigured || pendingConfigSaves.size === 0) return;
  const entries = Array.from(pendingConfigSaves.entries());
  pendingConfigSaves.clear();
  try {
    const rows = entries.map(([key, value]) => ({ key, value }));
    const { error } = await supabase.from('system_configs').upsert(rows, { onConflict: 'key' });
    if (error) throw error;
  } catch (err: any) {
    console.warn('[Supabase] Failed to persist system_configs:', err?.message || err);
  }
}

let pendingPoSave: ProductionOrder[] | null = null;
let poSaveTimer: any = null;

function queueProductionOrdersPersist(orders: ProductionOrder[]) {
  pendingPoSave = orders;
  if (!poSaveTimer) {
    poSaveTimer = setTimeout(flushPendingPoSaves, 500);
  }
}

async function flushPendingPoSaves() {
  poSaveTimer = null;
  if (!isSupabaseConfigured || !pendingPoSave) return;
  const orders = pendingPoSave;
  pendingPoSave = null;
  await persistProductionOrdersToSupabase(orders);
}

async function persistProductionOrdersToSupabase(orders: ProductionOrder[]) {
  if (!isSupabaseConfigured || !orders) return;
  try {
    const dbOrders = orders.map(po => ({
      id: po.id,
      master_po_id: po.masterPoId || null,
      part_id: po.partId,
      stage_id: po.stageId || null,
      target_quantity: po.targetQuantity,
      produced_quantity: po.producedQuantity || 0,
      exported_quantity: po.exportedQuantity || 0,
      status: po.status || 'PENDING',
      created_at: po.createdAt || Date.now(),
      completed_at: po.completedAt || null,
      planned_start_time: po.plannedStartTime || null,
      lead_time: po.leadTime || null,
      expected_completion_time: po.expectedCompletionTime || null,
      note: po.note || null,
      is_odd_po: Boolean(po.isOddPO),
      excel_required_date: po.excelRequiredDate || null,
    }));
    // Batch upsert in chunks of 150 to keep request size optimal
    for (let i = 0; i < dbOrders.length; i += 150) {
      const chunk = dbOrders.slice(i, i + 150);
      const { error } = await supabase.from('production_orders').upsert(chunk, { onConflict: 'id' });
      if (error) throw error;
    }
  } catch (err: any) {
    console.warn('[Supabase] persistProductionOrders error:', err?.message || err);
  }
}

async function deletePOFromSupabase(id: string) {
  if (!isSupabaseConfigured) return;
  try {
    await supabase.from('production_orders').delete().or(`id.eq.${id},master_po_id.eq.${id}`);
  } catch (err: any) {
    console.warn('[Supabase] deletePO error:', err?.message || err);
  }
}

let pendingInvSave: InventoryItem[] | null = null;
let invSaveTimer: any = null;

function queueInventoryPersist(inventory: InventoryItem[]) {
  pendingInvSave = inventory;
  if (!invSaveTimer) {
    invSaveTimer = setTimeout(flushPendingInvSaves, 500);
  }
}

async function flushPendingInvSaves() {
  invSaveTimer = null;
  if (!isSupabaseConfigured || !pendingInvSave) return;
  const inv = pendingInvSave;
  pendingInvSave = null;
  await persistInventoryToSupabase(inv);
}

async function persistInventoryToSupabase(inventory: InventoryItem[]) {
  if (!isSupabaseConfigured || !inventory) return;
  try {
    const dbRows = inventory.map(item => ({
      id: `${item.partId}_${item.stageId}_${item.location}_${item.originalPartId || ''}`,
      part_id: item.partId,
      original_part_id: item.originalPartId || '',
      stage_id: item.stageId,
      location: item.location,
      quantity: item.quantity
    }));
    for (let i = 0; i < dbRows.length; i += 150) {
      const chunk = dbRows.slice(i, i + 150);
      const { error } = await supabase.from('wip_inventory').upsert(chunk, { onConflict: 'id' });
      if (error) throw error;
    }
  } catch (err: any) {
    console.warn('[Supabase] persistInventory error:', err?.message || err);
  }
}

async function deleteInventoryFromSupabase(partId: string, stageId: StageId, location: string, originalPartId?: string) {
  if (!isSupabaseConfigured) return;
  try {
    const rowId = `${partId}_${stageId}_${location}_${originalPartId || ''}`;
    await supabase.from('wip_inventory').delete().eq('id', rowId);
  } catch (err: any) {
    console.warn('[Supabase] deleteInventory error:', err?.message || err);
  }
}

async function clearStageInventoryFromSupabase(stageId: StageId, location: string) {
  if (!isSupabaseConfigured) return;
  try {
    await supabase.from('wip_inventory').delete().eq('stage_id', stageId).eq('location', location);
  } catch (err: any) {
    console.warn('[Supabase] clearStageInventory error:', err?.message || err);
  }
}

let pendingTxs: Transaction[] = [];
let txSaveTimer: any = null;

function queueTransactionInsert(tx: Transaction) {
  pendingTxs.push(tx);
  if (!txSaveTimer) {
    txSaveTimer = setTimeout(flushPendingTransactions, 300);
  }
}

async function flushPendingTransactions() {
  txSaveTimer = null;
  if (!isSupabaseConfigured || pendingTxs.length === 0) return;
  const toSave = [...pendingTxs];
  pendingTxs = [];
  await persistTransactionsToSupabase(toSave);
}

async function persistTransactionsToSupabase(transactions: Transaction[]) {
  if (!isSupabaseConfigured || !transactions || transactions.length === 0) return;
  try {
    const cutoff = Date.now() - 35 * 24 * 60 * 60 * 1000;
    const recent = transactions.filter(t => t && t.timestamp && t.timestamp >= cutoff);
    const dbRows = recent.map(t => ({
      id: t.id,
      type: t.type,
      part_id: t.partId,
      part_name: t.partName || null,
      original_part_id: t.originalPartId || null,
      quantity: t.quantity,
      stage_id: t.stageId,
      location: t.location || null,
      timestamp: t.timestamp,
      qr_data: t.qrData || null,
      source_stage_id: t.sourceStageId || null,
      target_stage_id: t.targetStageId || null,
      po_id: t.poId || null,
      plan_id: t.planId || null,
      printed: Boolean(t.printed),
      defect_reason: t.defectReason || null,
      defect_category: t.defectCategory || null,
      kpi_recorded: Boolean(t.kpiRecorded)
    }));
    for (let i = 0; i < dbRows.length; i += 150) {
      const chunk = dbRows.slice(i, i + 150);
      const { error } = await supabase.from('wip_transactions').upsert(chunk, { onConflict: 'id' });
      if (error) throw error;
    }
  } catch (err: any) {
    console.warn('[Supabase] persistTransactions error:', err?.message || err);
  }
}

let pendingGpSave: import('./types').GlazingPlan[] | null = null;
let gpSaveTimer: any = null;

function queueGlazingPlansPersist(plans: import('./types').GlazingPlan[]) {
  pendingGpSave = plans;
  if (!gpSaveTimer) {
    gpSaveTimer = setTimeout(flushPendingGpSaves, 500);
  }
}

async function flushPendingGpSaves() {
  gpSaveTimer = null;
  if (!isSupabaseConfigured || !pendingGpSave) return;
  const plans = pendingGpSave;
  pendingGpSave = null;
  await persistGlazingPlansToSupabase(plans);
}

async function persistGlazingPlansToSupabase(plans: import('./types').GlazingPlan[]) {
  if (!isSupabaseConfigured || !plans) return;
  try {
    const dbRows = plans.map(p => ({
      id: p.id,
      model_id: p.modelId,
      target_quantity: p.targetQuantity,
      target_completion_time: p.targetCompletionTime,
      planned_start_time: p.plannedStartTime || null,
      expected_completion_time: p.expectedCompletionTime || null,
      created_at: p.createdAt,
      status: p.status || 'PENDING',
      produced_quantities: p.producedQuantities || {},
      note: p.note || null,
      excel_required_date: p.excelRequiredDate || null
    }));
    for (let i = 0; i < dbRows.length; i += 150) {
      const chunk = dbRows.slice(i, i + 150);
      const { error } = await supabase.from('glazing_plans').upsert(chunk, { onConflict: 'id' });
      if (error) throw error;
    }
  } catch (err: any) {
    console.warn('[Supabase] persistGlazingPlans error:', err?.message || err);
  }
}

async function deleteGlazingPlanFromSupabase(id: string) {
  if (!isSupabaseConfigured) return;
  try {
    await supabase.from('glazing_plans').delete().eq('id', id);
  } catch (err: any) {
    console.warn('[Supabase] deleteGlazingPlan error:', err?.message || err);
  }
}

let pendingLabels: Transaction[] = [];
let labelSaveTimer: any = null;

function queueLabelInsert(label: Transaction) {
  pendingLabels.push(label);
  if (!labelSaveTimer) {
    labelSaveTimer = setTimeout(flushPendingLabels, 300);
  }
}

async function flushPendingLabels() {
  labelSaveTimer = null;
  if (!isSupabaseConfigured || pendingLabels.length === 0) return;
  const toSave = [...pendingLabels];
  pendingLabels = [];
  await persistLabelsToSupabase(toSave);
}

async function persistLabelsToSupabase(labels: Transaction[]) {
  if (!isSupabaseConfigured || !labels || labels.length === 0) return;
  try {
    const cutoff = Date.now() - 24 * 60 * 60 * 1000;
    const recent = labels.filter(l => l && l.timestamp && l.timestamp >= cutoff);
    const dbRows = recent.map(l => ({
      id: l.id,
      part_id: l.partId || null,
      stage_id: l.stageId || null,
      quantity: l.quantity || 0,
      timestamp: l.timestamp,
      qr_data: l.qrData || null,
      metadata: l
    }));
    for (let i = 0; i < dbRows.length; i += 150) {
      const chunk = dbRows.slice(i, i + 150);
      const { error } = await supabase.from('qr_labels').upsert(chunk, { onConflict: 'id' });
      if (error) throw error;
    }
  } catch (err: any) {
    console.warn('[Supabase] persistLabels error:', err?.message || err);
  }
}

async function deleteLabelFromSupabase(id: string) {
  if (!isSupabaseConfigured) return;
  try {
    await supabase.from('qr_labels').delete().eq('id', id);
  } catch (err: any) {
    console.warn('[Supabase] deleteLabel error:', err?.message || err);
  }
}

function getCached<T>(key: string, fetchFn: () => T): T {
  if (cache[key] !== undefined) return cache[key];
  const data = fetchFn();
  cache[key] = data;
  return data;
}

function clearCache(key?: string) {
  if (key) {
    delete cache[key];
    if (key === STORAGE_KEYS.TRANSFORMATIONS || key === STORAGE_KEYS.PARTS) {
      effectivePartIdCache.clear();
    }
  } else {
    Object.keys(cache).forEach(k => delete cache[k]);
    effectivePartIdCache.clear();
  }
}

export const storageService = {
  normalize(s: string): string {
    if (!s) return '';
    let res = s.toUpperCase().normalize('NFC');
    // Strip common prefixes
    res = res.replace(/^(TẤM|CHI TIẾT|PHỤ TÙNG|BẢN|KHO|THÀNH PHẨM|L-)\s+/g, '');
    // Strip common suffixes
    res = res.replace(/\s*-\s*(CD|H|C|P|G|W|B|L|CT|BD)$/g, ''); 
    return res.split('(')[0].trim();
  },

  getLabelSettings() {
    return getCached(STORAGE_KEYS.LABEL_SETTINGS, () => {
      const data = localStorage.getItem(STORAGE_KEYS.LABEL_SETTINGS);
      // Ensure settings align with standard A7 (75x100mm portrait)
      if (data) {
        try {
          let parsed = JSON.parse(data);
          if (
            !parsed.width || !parsed.height ||
            parsed.width > 150 || parsed.height > 180 ||
            parsed.width < 50 || parsed.height < 50 ||
            parsed.qrSize > 150 || parsed.qrSize < 50 ||
            parsed.fontSize > 16 || parsed.fontSize < 6
          ) {
            parsed = { width: 75, height: 100, fontSize: 10, qrSize: 110 };
            localStorage.setItem(STORAGE_KEYS.LABEL_SETTINGS, JSON.stringify(parsed));
          }
          return parsed;
        } catch {
          // invalid json fallback
        }
      }
      const defaultSettings = { width: 75, height: 100, fontSize: 10, qrSize: 110 };
      localStorage.setItem(STORAGE_KEYS.LABEL_SETTINGS, JSON.stringify(defaultSettings));
      return defaultSettings;
    });
  },

  saveLabelSettings(settings: any) {
    localStorage.setItem(STORAGE_KEYS.LABEL_SETTINGS, JSON.stringify(settings));
    cache[STORAGE_KEYS.LABEL_SETTINGS] = settings;
    queueConfigPersist(STORAGE_KEYS.LABEL_SETTINGS, settings);
  },

  getParts(): Part[] {
    const data = getCached(STORAGE_KEYS.PARTS, () => {
      const data = localStorage.getItem(STORAGE_KEYS.PARTS);
      return data ? JSON.parse(data) : INITIAL_PARTS;
    });
    return [...data];
  },

  saveParts(parts: Part[]) {
    localStorage.setItem(STORAGE_KEYS.PARTS, JSON.stringify(parts));
    cache[STORAGE_KEYS.PARTS] = parts;
    effectivePartIdCache.clear();
    queueConfigPersist(STORAGE_KEYS.PARTS, parts);
  },

  getBOM(): BOMDefinition[] {
    return getCached(STORAGE_KEYS.BOM, () => {
      const data = localStorage.getItem(STORAGE_KEYS.BOM);
      return data ? JSON.parse(data) : [];
    });
  },

  saveBOM(bom: BOMDefinition[]) {
    localStorage.setItem(STORAGE_KEYS.BOM, JSON.stringify(bom));
    cache[STORAGE_KEYS.BOM] = bom;
    queueConfigPersist(STORAGE_KEYS.BOM, bom);
  },

  getBOMV2(): BOMDefinitionV2[] {
    return getCached(STORAGE_KEYS.BOM_V2, () => {
      const data = localStorage.getItem(STORAGE_KEYS.BOM_V2);
      return data ? JSON.parse(data) : [];
    });
  },

  saveBOMV2(bom: BOMDefinitionV2[]) {
    localStorage.setItem(STORAGE_KEYS.BOM_V2, JSON.stringify(bom));
    cache[STORAGE_KEYS.BOM_V2] = bom;
    queueConfigPersist(STORAGE_KEYS.BOM_V2, bom);
  },

  getModelBOM(): ModelBOMDefinition[] {
    return getCached(STORAGE_KEYS.MODEL_BOM, () => {
      const data = localStorage.getItem(STORAGE_KEYS.MODEL_BOM);
      return data ? JSON.parse(data) : [];
    });
  },

  saveModelBOM(bom: ModelBOMDefinition[]) {
    localStorage.setItem(STORAGE_KEYS.MODEL_BOM, JSON.stringify(bom));
    cache[STORAGE_KEYS.MODEL_BOM] = bom;
    queueConfigPersist(STORAGE_KEYS.MODEL_BOM, bom);
  },
  
  getNorms(): ProductivityNorm[] {
    return getCached(STORAGE_KEYS.NORMS, () => {
      const data = localStorage.getItem(STORAGE_KEYS.NORMS);
      return data ? JSON.parse(data) : [];
    });
  },

  saveNorms(norms: ProductivityNorm[]) {
    localStorage.setItem(STORAGE_KEYS.NORMS, JSON.stringify(norms));
    cache[STORAGE_KEYS.NORMS] = norms;
    queueConfigPersist(STORAGE_KEYS.NORMS, norms);
  },

  getLaserNesting(): LaserNesting[] {
    return getCached(STORAGE_KEYS.LASER_NESTING, () => {
      const data = localStorage.getItem(STORAGE_KEYS.LASER_NESTING);
      return data ? JSON.parse(data) : [];
    });
  },

  saveLaserNesting(nesting: LaserNesting[]) {
    localStorage.setItem(STORAGE_KEYS.LASER_NESTING, JSON.stringify(nesting));
    cache[STORAGE_KEYS.LASER_NESTING] = nesting;
    queueConfigPersist(STORAGE_KEYS.LASER_NESTING, nesting);
  },
  
  getShiftConfigs(): ShiftConfig[] {
    return getCached(STORAGE_KEYS.SHIFT_CONFIGS, () => {
      const data = localStorage.getItem(STORAGE_KEYS.SHIFT_CONFIGS);
      let configs: ShiftConfig[] = data ? JSON.parse(data) : [];
      
      const defaults: ShiftConfig[] = [
        {
          stageId: 'LASER',
          workerCount: 1,
          shifts: [{ start: '06:00', end: '18:00' }, { start: '18:00', end: '06:00' }],
          breaks: [{ start: '10:00', end: '10:10' }, { start: '11:50', end: '13:00' }, { start: '15:00', end: '15:10' }, { start: '18:00', end: '18:30' }]
        },
        {
          stageId: 'BENDING',
          workerCount: 2,
          shifts: [{ start: '08:00', end: '20:00' }],
          breaks: [{ start: '10:00', end: '10:10' }, { start: '11:50', end: '13:00' }, { start: '15:00', end: '15:10' }]
        },
        {
          stageId: 'WELDING',
          workerCount: 5,
          shifts: [{ start: '08:00', end: '20:00' }],
          breaks: [{ start: '10:00', end: '10:10' }, { start: '11:50', end: '13:00' }, { start: '15:00', end: '15:10' }, { start: '17:00', end: '17:10' }]
        },
        {
          stageId: 'PAINTING',
          workerCount: 1,
          shifts: [{ start: '08:00', end: '20:00' }],
          breaks: [{ start: '10:00', end: '10:10' }, { start: '11:50', end: '13:00' }, { start: '15:00', end: '15:10' }, { start: '17:00', end: '17:10' }]
        },
        {
          stageId: 'GLAZING',
          workerCount: 2,
          shifts: [{ start: '08:00', end: '20:00' }],
          breaks: [{ start: '10:00', end: '10:10' }, { start: '11:50', end: '13:00' }, { start: '15:00', end: '15:10' }, { start: '17:00', end: '17:10' }]
        }
      ];

      // Add missing defaults to existing configs
      defaults.forEach(def => {
        if (!configs.find(c => c.stageId === def.stageId)) {
          configs.push(def);
        }
      });
      
      return configs;
    });
  },

  saveShiftConfigs(configs: ShiftConfig[]) {
    localStorage.setItem(STORAGE_KEYS.SHIFT_CONFIGS, JSON.stringify(configs));
    cache[STORAGE_KEYS.SHIFT_CONFIGS] = configs;
    queueConfigPersist(STORAGE_KEYS.SHIFT_CONFIGS, configs);
  },

  getTransformations(): PartTransformation[] {
    return getCached(STORAGE_KEYS.TRANSFORMATIONS, () => {
      const data = localStorage.getItem(STORAGE_KEYS.TRANSFORMATIONS);
      return data ? JSON.parse(data) : [];
    });
  },

  saveTransformations(transformations: PartTransformation[]) {
    localStorage.setItem(STORAGE_KEYS.TRANSFORMATIONS, JSON.stringify(transformations));
    cache[STORAGE_KEYS.TRANSFORMATIONS] = transformations;
    effectivePartIdCache.clear();
    queueConfigPersist(STORAGE_KEYS.TRANSFORMATIONS, transformations);
  },

  getGlazingConfigs(): import('./types').GlazingConfig[] {
    return getCached(STORAGE_KEYS.GLAZING_CONFIGS, () => {
      const data = localStorage.getItem(STORAGE_KEYS.GLAZING_CONFIGS);
      return data ? JSON.parse(data) : [];
    });
  },

  saveGlazingConfigs(configs: import('./types').GlazingConfig[]) {
    localStorage.setItem(STORAGE_KEYS.GLAZING_CONFIGS, JSON.stringify(configs));
    cache[STORAGE_KEYS.GLAZING_CONFIGS] = configs;
    queueConfigPersist(STORAGE_KEYS.GLAZING_CONFIGS, configs);
  },

  getGlazingOutConfigs(): import('./types').GlazingOutConfig[] {
    return getCached(STORAGE_KEYS.GLAZING_OUT_CONFIGS, () => {
      const data = localStorage.getItem(STORAGE_KEYS.GLAZING_OUT_CONFIGS);
      return data ? JSON.parse(data) : [];
    });
  },

  saveGlazingOutConfigs(configs: import('./types').GlazingOutConfig[]) {
    localStorage.setItem(STORAGE_KEYS.GLAZING_OUT_CONFIGS, JSON.stringify(configs));
    cache[STORAGE_KEYS.GLAZING_OUT_CONFIGS] = configs;
    queueConfigPersist(STORAGE_KEYS.GLAZING_OUT_CONFIGS, configs);
  },

  getQuickPrintParts(): {id: string, name: string, quantity: number}[] {
    return getCached(STORAGE_KEYS.QUICK_PRINT_PARTS, () => {
      const data = localStorage.getItem(STORAGE_KEYS.QUICK_PRINT_PARTS);
      return data ? JSON.parse(data) : [];
    });
  },

  saveQuickPrintParts(parts: {id: string, name: string, quantity: number}[]) {
    localStorage.setItem(STORAGE_KEYS.QUICK_PRINT_PARTS, JSON.stringify(parts));
    cache[STORAGE_KEYS.QUICK_PRINT_PARTS] = parts;
    queueConfigPersist(STORAGE_KEYS.QUICK_PRINT_PARTS, parts);
  },

  getGlazingPlanNorms(): import('./types').GlazingPlanNorm[] {
    return getCached(STORAGE_KEYS.GLAZING_PLAN_NORMS, () => {
      const data = localStorage.getItem(STORAGE_KEYS.GLAZING_PLAN_NORMS);
      return data ? JSON.parse(data) : [];
    });
  },

  getNormsForModel(modelId: string): import('./types').GlazingPlanNorm[] {
    const norms = this.getGlazingPlanNorms();
    return norms.filter(n => n.appliedModel === modelId || n.modelName === modelId);
  },

  saveGlazingPlanNorms(norms: import('./types').GlazingPlanNorm[]) {
    localStorage.setItem(STORAGE_KEYS.GLAZING_PLAN_NORMS, JSON.stringify(norms));
    cache[STORAGE_KEYS.GLAZING_PLAN_NORMS] = norms;
    queueConfigPersist(STORAGE_KEYS.GLAZING_PLAN_NORMS, norms);
  },

  getGlazingPlans(): import('./types').GlazingPlan[] {
    return getCached(STORAGE_KEYS.GLAZING_PLANS, () => {
      const data = localStorage.getItem(STORAGE_KEYS.GLAZING_PLANS);
      return data ? JSON.parse(data) : [];
    });
  },

  saveGlazingPlans(plans: import('./types').GlazingPlan[]) {
    localStorage.setItem(STORAGE_KEYS.GLAZING_PLANS, JSON.stringify(plans));
    cache[STORAGE_KEYS.GLAZING_PLANS] = plans;
    queueGlazingPlansPersist(plans);
  },

  getPeoplePerDay(): Record<string, number> {
    return getCached(STORAGE_KEYS.PEOPLE_PER_DAY, () => {
      const data = localStorage.getItem(STORAGE_KEYS.PEOPLE_PER_DAY);
      return data ? JSON.parse(data) : {};
    });
  },

  savePeoplePerDay(record: Record<string, number>) {
    localStorage.setItem(STORAGE_KEYS.PEOPLE_PER_DAY, JSON.stringify(record));
    cache[STORAGE_KEYS.PEOPLE_PER_DAY] = record;
    queueConfigPersist(STORAGE_KEYS.PEOPLE_PER_DAY, record);
  },

  getMandaysPerDay(): Record<string, number> {
    return getCached(STORAGE_KEYS.MANDAYS_PER_DAY, () => {
      const data = localStorage.getItem(STORAGE_KEYS.MANDAYS_PER_DAY);
      return data ? JSON.parse(data) : {};
    });
  },

  saveMandaysPerDay(record: Record<string, number>) {
    localStorage.setItem(STORAGE_KEYS.MANDAYS_PER_DAY, JSON.stringify(record));
    cache[STORAGE_KEYS.MANDAYS_PER_DAY] = record;
    queueConfigPersist(STORAGE_KEYS.MANDAYS_PER_DAY, record);
  },

  getExportWeekName(): string {
    return getCached(STORAGE_KEYS.EXPORT_WEEK_NAME, () => {
      return localStorage.getItem(STORAGE_KEYS.EXPORT_WEEK_NAME) || "";
    });
  },

  saveExportWeekName(weekName: string) {
    localStorage.setItem(STORAGE_KEYS.EXPORT_WEEK_NAME, weekName);
    cache[STORAGE_KEYS.EXPORT_WEEK_NAME] = weekName;
    queueConfigPersist(STORAGE_KEYS.EXPORT_WEEK_NAME, weekName);
  },

  _resolveHourlyData(rawData: Record<string, number>, date?: string): Record<string, number> {
    const timeSlots = ['6h-8h', '8h-10h', '10h-12h', '13h-15h', '15h-17h', '17h-19h', '19h-20h', '20h-22h', '22h-0h', '0h-2h', '2h-4h', '4h-6h'];
    if (!date) {
      return {};
    }

    // Mỗi ngày bắt điền mới, không tự động lưu/kế thừa người của ngày trước
    const explicit: Record<string, number> = {};
    timeSlots.forEach(slot => {
      const key = `${date}_${slot}`;
      if (rawData[key] !== undefined) {
        explicit[slot] = rawData[key];
      }
    });
    return explicit;
  },

  getHourlyPeoplePainting(date?: string): Record<string, number> {
    const rawData = getCached(STORAGE_KEYS.HOURLY_PEOPLE_PAINTING, () => {
      const data = localStorage.getItem(STORAGE_KEYS.HOURLY_PEOPLE_PAINTING);
      return data ? JSON.parse(data) : {};
    });
    return this._resolveHourlyData(rawData, date);
  },

  saveHourlyPeoplePainting(record: Record<string, number>, date?: string) {
    const rawData = getCached(STORAGE_KEYS.HOURLY_PEOPLE_PAINTING, () => {
      const data = localStorage.getItem(STORAGE_KEYS.HOURLY_PEOPLE_PAINTING);
      return data ? JSON.parse(data) : {};
    });
    const updated = { ...rawData };
    if (date) {
      Object.entries(record).forEach(([slot, val]) => {
        updated[`${date}_${slot}`] = val;
      });
    } else {
      Object.assign(updated, record);
    }
    localStorage.setItem(STORAGE_KEYS.HOURLY_PEOPLE_PAINTING, JSON.stringify(updated));
    cache[STORAGE_KEYS.HOURLY_PEOPLE_PAINTING] = updated;
    queueConfigPersist(STORAGE_KEYS.HOURLY_PEOPLE_PAINTING, updated);
  },

  getHourlyPeopleGlazing(date?: string): Record<string, number> {
    const rawData = getCached(STORAGE_KEYS.HOURLY_PEOPLE_GLAZING, () => {
      const data = localStorage.getItem(STORAGE_KEYS.HOURLY_PEOPLE_GLAZING);
      return data ? JSON.parse(data) : {};
    });
    return this._resolveHourlyData(rawData, date);
  },

  saveHourlyPeopleGlazing(record: Record<string, number>, date?: string) {
    const rawData = getCached(STORAGE_KEYS.HOURLY_PEOPLE_GLAZING, () => {
      const data = localStorage.getItem(STORAGE_KEYS.HOURLY_PEOPLE_GLAZING);
      return data ? JSON.parse(data) : {};
    });
    const updated = { ...rawData };
    if (date) {
      Object.entries(record).forEach(([slot, val]) => {
        updated[`${date}_${slot}`] = val;
      });
    } else {
      Object.assign(updated, record);
    }
    localStorage.setItem(STORAGE_KEYS.HOURLY_PEOPLE_GLAZING, JSON.stringify(updated));
    cache[STORAGE_KEYS.HOURLY_PEOPLE_GLAZING] = updated;
    queueConfigPersist(STORAGE_KEYS.HOURLY_PEOPLE_GLAZING, updated);
  },

  getHourlyPeopleBending(date?: string): Record<string, number> {
    const rawData = getCached(STORAGE_KEYS.HOURLY_PEOPLE_BENDING, () => {
      const data = localStorage.getItem(STORAGE_KEYS.HOURLY_PEOPLE_BENDING);
      return data ? JSON.parse(data) : {};
    });
    return this._resolveHourlyData(rawData, date);
  },

  saveHourlyPeopleBending(record: Record<string, number>, date?: string) {
    const rawData = getCached(STORAGE_KEYS.HOURLY_PEOPLE_BENDING, () => {
      const data = localStorage.getItem(STORAGE_KEYS.HOURLY_PEOPLE_BENDING);
      return data ? JSON.parse(data) : {};
    });
    const updated = { ...rawData };
    if (date) {
      Object.entries(record).forEach(([slot, val]) => {
        updated[`${date}_${slot}`] = val;
      });
    } else {
      Object.assign(updated, record);
    }
    localStorage.setItem(STORAGE_KEYS.HOURLY_PEOPLE_BENDING, JSON.stringify(updated));
    cache[STORAGE_KEYS.HOURLY_PEOPLE_BENDING] = updated;
    queueConfigPersist(STORAGE_KEYS.HOURLY_PEOPLE_BENDING, updated);
  },

  getHourlyPeopleWelding(date?: string): Record<string, number> {
    const rawData = getCached(STORAGE_KEYS.HOURLY_PEOPLE_WELDING, () => {
      const data = localStorage.getItem(STORAGE_KEYS.HOURLY_PEOPLE_WELDING);
      return data ? JSON.parse(data) : {};
    });
    return this._resolveHourlyData(rawData, date);
  },

  saveHourlyPeopleWelding(record: Record<string, number>, date?: string) {
    const rawData = getCached(STORAGE_KEYS.HOURLY_PEOPLE_WELDING, () => {
      const data = localStorage.getItem(STORAGE_KEYS.HOURLY_PEOPLE_WELDING);
      return data ? JSON.parse(data) : {};
    });
    const updated = { ...rawData };
    if (date) {
      Object.entries(record).forEach(([slot, val]) => {
        updated[`${date}_${slot}`] = val;
      });
    } else {
      Object.assign(updated, record);
    }
    localStorage.setItem(STORAGE_KEYS.HOURLY_PEOPLE_WELDING, JSON.stringify(updated));
    cache[STORAGE_KEYS.HOURLY_PEOPLE_WELDING] = updated;
    queueConfigPersist(STORAGE_KEYS.HOURLY_PEOPLE_WELDING, updated);
  },

  getHourlyPeopleLaser(date?: string): Record<string, number> {
    const rawData = getCached(STORAGE_KEYS.HOURLY_PEOPLE_LASER, () => {
      const data = localStorage.getItem(STORAGE_KEYS.HOURLY_PEOPLE_LASER);
      return data ? JSON.parse(data) : {};
    });
    return this._resolveHourlyData(rawData, date);
  },

  saveHourlyPeopleLaser(record: Record<string, number>, date?: string) {
    const rawData = getCached(STORAGE_KEYS.HOURLY_PEOPLE_LASER, () => {
      const data = localStorage.getItem(STORAGE_KEYS.HOURLY_PEOPLE_LASER);
      return data ? JSON.parse(data) : {};
    });
    const updated = { ...rawData };
    if (date) {
      Object.entries(record).forEach(([slot, val]) => {
        updated[`${date}_${slot}`] = val;
      });
    } else {
      Object.assign(updated, record);
    }
    localStorage.setItem(STORAGE_KEYS.HOURLY_PEOPLE_LASER, JSON.stringify(updated));
    cache[STORAGE_KEYS.HOURLY_PEOPLE_LASER] = updated;
    queueConfigPersist(STORAGE_KEYS.HOURLY_PEOPLE_LASER, updated);
  },

  getDailyTotalPeopleAverage(date: string): number | null {
    const laser = this.getHourlyPeopleLaser(date);
    const bending = this.getHourlyPeopleBending(date);
    const welding = this.getHourlyPeopleWelding(date);
    const painting = this.getHourlyPeoplePainting(date);
    const glazing = this.getHourlyPeopleGlazing(date);
    const timeSlots = ['6h-8h', '8h-10h', '10h-12h', '13h-15h', '15h-17h', '17h-19h', '19h-20h', '20h-22h', '22h-0h', '0h-2h', '2h-4h', '4h-6h'];
    let totalPeople = 0;
    let validSlots = 0;
    timeSlots.forEach(slot => {
      const l = laser[slot] !== undefined ? laser[slot] : 0;
      const b = bending[slot] !== undefined ? bending[slot] : 0;
      const w = welding[slot] !== undefined ? welding[slot] : 0;
      const p = painting[slot] !== undefined ? painting[slot] : 0;
      const g = glazing[slot] !== undefined ? glazing[slot] : 0;
      const slotTotal = l + b + w + p + g;
      if (slotTotal > 0) {
        totalPeople += slotTotal;
        validSlots++;
      }
    });
    return validSlots > 0 ? Math.round(totalPeople / validSlots) : null;
  },

  getDailyStagePeopleAverage(stageId: StageId, date: string): number | null {
    // 1. Cắt laser: Số lượng cố định là 4, không cần lấy dữ liệu gì
    if (stageId === 'LASER') {
      return 4;
    }

    let hourly: Record<string, number> = {};
    if (stageId === 'BENDING') hourly = this.getHourlyPeopleBending(date);
    else if (stageId === 'WELDING') hourly = this.getHourlyPeopleWelding(date);
    else if (stageId === 'PAINTING') hourly = this.getHourlyPeoplePainting(date);
    else if (stageId === 'GLAZING') hourly = this.getHourlyPeopleGlazing(date);

    // Khung giờ 6h sáng đến 19h tối (ca ngày)
    const daySlots = ['6h-8h', '8h-10h', '10h-12h', '13h-15h', '15h-17h', '17h-19h'];
    // Khung giờ 19h tối đến 6h sáng (ca đêm)
    const nightSlots = ['19h-20h', '20h-22h', '22h-0h', '0h-2h', '2h-4h', '4h-6h'];

    let dayTotal = 0;
    let dayCount = 0;
    daySlots.forEach(slot => {
      const p = hourly[slot];
      if (p !== undefined && p > 0) {
        dayTotal += p;
        dayCount++;
      }
    });

    let nightTotal = 0;
    let nightCount = 0;
    nightSlots.forEach(slot => {
      const p = hourly[slot];
      if (p !== undefined && p > 0) {
        nightTotal += p;
        nightCount++;
      }
    });

    // Nếu cả ca ngày và ca đêm đều không có dữ liệu
    if (dayCount === 0 && nightCount === 0) {
      return null;
    }

    const dayAvg = dayCount > 0 ? (dayTotal / dayCount) : 0;
    const nightAvg = nightCount > 0 ? (nightTotal / nightCount) : 0;

    // Tổng số người = trung bình ca ngày + trung bình ca đêm (nếu ca đêm không điền thì chỉ tính trung bình ca ngày)
    const totalPeople = nightCount > 0 ? (dayAvg + nightAvg) : dayAvg;
    return Math.round(totalPeople * 10) / 10;
  },

  getStagePlannedWorkers(date?: string): import('./types').StagePlannedWorkers {
    const rawData = getCached(STORAGE_KEYS.STAGE_PLANNED_WORKERS, () => {
      const data = localStorage.getItem(STORAGE_KEYS.STAGE_PLANNED_WORKERS);
      return data ? JSON.parse(data) : {};
    });
    const defaults: import('./types').StagePlannedWorkers = {
      LASER: 4,
      BENDING: 8,
      WELDING: 8,
      PAINTING: 8,
      GLAZING: 8,
    };
    if (date && rawData[date]) {
      return { ...defaults, ...rawData[date] };
    }
    if (rawData['default']) {
      return { ...defaults, ...rawData['default'] };
    }
    return defaults;
  },

  saveStagePlannedWorkers(workers: Partial<import('./types').StagePlannedWorkers>, date?: string) {
    const rawData = getCached(STORAGE_KEYS.STAGE_PLANNED_WORKERS, () => {
      const data = localStorage.getItem(STORAGE_KEYS.STAGE_PLANNED_WORKERS);
      return data ? JSON.parse(data) : {};
    });
    const updated = { ...rawData };
    const defaults: import('./types').StagePlannedWorkers = {
      LASER: 4,
      BENDING: 8,
      WELDING: 8,
      PAINTING: 8,
      GLAZING: 8,
    };
    const current = (date && updated[date]) ? updated[date] : (updated['default'] || defaults);
    const merged = { ...current, ...workers };
    if (date) {
      updated[date] = merged;
      updated['default'] = merged;
    } else {
      updated['default'] = merged;
    }
    localStorage.setItem(STORAGE_KEYS.STAGE_PLANNED_WORKERS, JSON.stringify(updated));
    cache[STORAGE_KEYS.STAGE_PLANNED_WORKERS] = updated;
    queueConfigPersist(STORAGE_KEYS.STAGE_PLANNED_WORKERS, updated);
  },

  getDailyTotalMandays(date: string): number {
    const laser = this.getHourlyPeopleLaser(date);
    const bending = this.getHourlyPeopleBending(date);
    const welding = this.getHourlyPeopleWelding(date);
    const painting = this.getHourlyPeoplePainting(date);
    const glazing = this.getHourlyPeopleGlazing(date);

    const slotConfigs = [
      { slot: '6h-8h', hours: 2 },
      { slot: '8h-10h', hours: 2 },
      { slot: '10h-12h', hours: 2 },
      { slot: '13h-15h', hours: 2 },
      { slot: '15h-17h', hours: 2 },
      { slot: '17h-19h', hours: 2 },
      { slot: '19h-20h', hours: 1 },
      { slot: '20h-22h', hours: 2 },
      { slot: '22h-0h', hours: 2 },
      { slot: '0h-2h', hours: 2 },
      { slot: '2h-4h', hours: 2 },
      { slot: '4h-6h', hours: 2 }
    ];

    let totalMandays = 0;
    slotConfigs.forEach(({ slot, hours }) => {
      const l = laser[slot] !== undefined ? laser[slot] : 0;
      const b = bending[slot] !== undefined ? bending[slot] : 0;
      const w = welding[slot] !== undefined ? welding[slot] : 0;
      const p = painting[slot] !== undefined ? painting[slot] : 0;
      const g = glazing[slot] !== undefined ? glazing[slot] : 0;
      const slotTotalPeople = l + b + w + p + g;
      // Khung giờ nào có nhân sự ở bất kỳ công đoạn nào thì tính công của khung giờ đó
      // Khung giờ nào không có nhân sự nào hết ở tất cả công đoạn sẽ không tính công khung giờ đó
      if (slotTotalPeople > 0) {
        // Công của khung giờ = (Tổng số người của khung giờ đó) * (hours / 8)
        // (Khung giờ 2h = slotTotalPeople * 0.25; Khung giờ 1h = slotTotalPeople * 0.125)
        const slotMandays = slotTotalPeople * (hours / 8);
        totalMandays += slotMandays;
      }
    });

    return Math.round(totalMandays * 10) / 10;
  },

  getDailyPaintingGlazingAverage(date: string): number | null {
    // Tự động chuyển tiếp sang trung bình cộng tất cả các công đoạn của ngày đó
    return this.getDailyTotalPeopleAverage(date);
  },

  getBendingWeldingHSQD(): { partId: string; partName?: string; hsqd: number }[] {
    return getCached(STORAGE_KEYS.BENDING_WELDING_HSQD, () => {
      const data = localStorage.getItem(STORAGE_KEYS.BENDING_WELDING_HSQD);
      return data ? JSON.parse(data) : [];
    });
  },

  saveBendingWeldingHSQD(data: { partId: string; partName?: string; hsqd: number }[]) {
    localStorage.setItem(STORAGE_KEYS.BENDING_WELDING_HSQD, JSON.stringify(data));
    cache[STORAGE_KEYS.BENDING_WELDING_HSQD] = data;
    queueConfigPersist(STORAGE_KEYS.BENDING_WELDING_HSQD, data);
  },

  getLaserHSQD(): { partId: string; partName?: string; hsqd: number }[] {
    return getCached(STORAGE_KEYS.LASER_HSQD, () => {
      const data = localStorage.getItem(STORAGE_KEYS.LASER_HSQD);
      return data ? JSON.parse(data) : [];
    });
  },

  saveLaserHSQD(data: { partId: string; partName?: string; hsqd: number }[]) {
    localStorage.setItem(STORAGE_KEYS.LASER_HSQD, JSON.stringify(data));
    cache[STORAGE_KEYS.LASER_HSQD] = data;
    queueConfigPersist(STORAGE_KEYS.LASER_HSQD, data);
  },

  createGlazingPlan(modelId: string, quantity: number, targetCompletion: number, note?: string, excelRequiredDate?: number, customId?: string) {
    const plans = this.getGlazingPlans();
    const schedule = this.getGlazingSchedule(modelId, quantity, targetCompletion);
    
    const newPlan: import('./types').GlazingPlan = {
      id: customId || `GLZ-PLN-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
      modelId,
      targetQuantity: quantity,
      targetCompletionTime: targetCompletion,
      plannedStartTime: schedule.start,
      expectedCompletionTime: schedule.end,
      createdAt: Date.now(),
      status: 'PENDING',
      note,
      excelRequiredDate
    };
    plans.push(newPlan);
    this.saveGlazingPlans(plans);
    return newPlan;
  },

  createBatchGlazingPlans(planItems: {
    modelId: string;
    quantity: number;
    targetCompletion: number;
    note?: string;
    excelRequiredDate?: number;
    id?: string;
  }[]) {
    if (!planItems || planItems.length === 0) return [];
    const plans = this.getGlazingPlans();
    const newPlans: import('./types').GlazingPlan[] = [];
    const allNorms = this.getGlazingPlanNorms();
    const shiftConfigs = this.getShiftConfigs();
    const context = { allNorms, shiftConfigs };

    for (const item of planItems) {
      const schedule = this.getGlazingSchedule(item.modelId, item.quantity, item.targetCompletion, context);
      newPlans.push({
        id: item.id || `GLZ-PLN-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
        modelId: item.modelId,
        targetQuantity: item.quantity,
        targetCompletionTime: item.targetCompletion,
        plannedStartTime: schedule.start,
        expectedCompletionTime: schedule.end,
        createdAt: Date.now(),
        status: 'PENDING',
        note: item.note,
        excelRequiredDate: item.excelRequiredDate
      });
    }

    const updated = [...newPlans, ...plans];
    this.saveGlazingPlans(updated);
    return newPlans;
  },

  getGlazingSchedule(modelId: string, quantity: number, targetCompletion: number, context?: { allNorms?: import('./types').GlazingPlanNorm[], shiftConfigs?: ShiftConfig[] }) {
    const allNorms = context?.allNorms || this.getGlazingPlanNorms();
    const norms = allNorms.filter(n => n.appliedModel === modelId);
    const shiftConfigs = context?.shiftConfigs || this.getShiftConfigs();
    const glazingConfig = shiftConfigs.find(c => c.stageId === 'GLAZING');
    const workerCount = glazingConfig?.workerCount || 2;

    const totalDurationMs = norms.length > 0 
      ? norms.reduce((sum, n) => sum + (n.norm * quantity * 1000) / workerCount, 0)
      : (quantity * 300 * 1000) / workerCount; // Default 5 min per unit if no norms

    const runForward = (baseStartTime: number) => {
      const actualStart = this.getNextWorkingTime(baseStartTime, 'GLAZING', shiftConfigs);
      const end = this.calculateEndTime(actualStart, totalDurationMs, 'GLAZING', shiftConfigs);
      return { start: actualStart, end };
    };

    let low = targetCompletion - 60 * 24 * 60 * 60 * 1000;
    let high = targetCompletion;
    let best = runForward(low);

    for (let i = 0; i < 20; i++) {
      if (high - low < 60000) break;
      const mid = low + Math.floor((high - low) / 2);
      const res = runForward(mid);
      if (res.end <= targetCompletion) {
        best = res;
        low = mid + 1;
      } else {
        high = mid - 1;
      }
    }
    return best;
  },

  calculateGlazingStartTime(modelId: string, quantity: number, targetCompletion: number) {
    const schedule = this.getGlazingSchedule(modelId, quantity, targetCompletion);
    return schedule.start;
  },

  deleteGlazingPlan(id: string) {
    const plans = this.getGlazingPlans().filter(p => p.id !== id);
    this.saveGlazingPlans(plans);
    deleteGlazingPlanFromSupabase(id);
  },

  completeGlazingPlan(id: string) {
    const plans = this.getGlazingPlans();
    const planIndex = plans.findIndex(p => p.id === id);
    if (planIndex !== -1) {
      plans[planIndex].status = 'COMPLETED';
      this.saveGlazingPlans(plans);
    }
  },

  updateGlazingPlanProgress(planId: string, partId: string, quantity: number) {
    const plans = this.getGlazingPlans();
    const planIndex = plans.findIndex(p => p.id === planId);
    if (planIndex !== -1) {
      const plan = plans[planIndex];
      const producedQuantities = { ...(plan.producedQuantities || {}) };
      producedQuantities[partId] = (producedQuantities[partId] || 0) + quantity;
      
      plans[planIndex] = { ...plan, producedQuantities, status: 'IN_PROGRESS' };
      this.saveGlazingPlans(plans);
    }
  },

  updateGlazingPlanPrinted(planId: string, partId: string, quantity: number) {
    const plans = this.getGlazingPlans();
    const planIndex = plans.findIndex(p => p.id === planId);
    if (planIndex !== -1) {
      const plan = plans[planIndex];
      const printedQuantities = { ...(plan.printedQuantities || {}) };
      printedQuantities[partId] = (printedQuantities[partId] || 0) + quantity;
      plans[planIndex] = { ...plan, printedQuantities };
      this.saveGlazingPlans(plans);
    }
  },

  getInventory(): InventoryItem[] {
    const data = getCached(STORAGE_KEYS.INVENTORY, () => {
      const data = localStorage.getItem(STORAGE_KEYS.INVENTORY);
      return data ? JSON.parse(data) : [];
    });
    return [...data];
  },

  getTransactions(): Transaction[] {
    const data = getCached(STORAGE_KEYS.TRANSACTIONS, () => {
      const data = localStorage.getItem(STORAGE_KEYS.TRANSACTIONS);
      let parsed = data ? JSON.parse(data) : [];
      // Auto purge scan log data older than 35 days to free local storage
      const cutoff = Date.now() - 35 * 24 * 60 * 60 * 1000;
      const initialLength = parsed.length;
      parsed = parsed.filter((t: any) => t && t.timestamp && t.timestamp >= cutoff);
      if (parsed.length !== initialLength) {
        localStorage.setItem(STORAGE_KEYS.TRANSACTIONS, JSON.stringify(parsed));
      }
      return parsed;
    });
    return [...data];
  },

  saveInventory(inventory: InventoryItem[]) {
    localStorage.setItem(STORAGE_KEYS.INVENTORY, JSON.stringify(inventory));
    cache[STORAGE_KEYS.INVENTORY] = inventory;
    queueInventoryPersist(inventory);
  },

  saveTransactions(transactions: Transaction[]) {
    const cutoff = Date.now() - 35 * 24 * 60 * 60 * 1000;
    const filtered = transactions.filter((t: any) => t && t.timestamp && t.timestamp >= cutoff);
    localStorage.setItem(STORAGE_KEYS.TRANSACTIONS, JSON.stringify(filtered));
    cache[STORAGE_KEYS.TRANSACTIONS] = filtered;
    queueTransactionInsert(filtered[0] || ({} as any));
    if (filtered.length > 1) {
      persistTransactionsToSupabase(filtered);
    }
  },

  saveLabels(labels: Transaction[]) {
    // Only store QR labels within the last 24 hours
    const cutoff = Date.now() - 24 * 60 * 60 * 1000;
    const filtered = labels.filter(l => l && l.timestamp && l.timestamp >= cutoff);
    localStorage.setItem('wip_labels', JSON.stringify(filtered));
    cache['wip_labels'] = filtered;
    persistLabelsToSupabase(filtered);
  },

  getLabels(): Transaction[] {
    const data = getCached('wip_labels', () => {
      const data = localStorage.getItem('wip_labels');
      let parsed = data ? JSON.parse(data) : [];
      // Only keep QR labels created within the last 24 hours (24h cutoff)
      const cutoff = Date.now() - 24 * 60 * 60 * 1000;
      const initialLength = parsed.length;
      parsed = parsed.filter((t: any) => t && t.timestamp && t.timestamp >= cutoff);
      if (parsed.length !== initialLength) {
        localStorage.setItem('wip_labels', JSON.stringify(parsed));
      }
      return parsed;
    });
    return [...data];
  },

  saveLabel(label: Transaction) {
    const cutoff = Date.now() - 24 * 60 * 60 * 1000;
    const existing = this.getLabels().filter(l => l && l.timestamp && l.timestamp >= cutoff && l.id !== label.id);
    const labels = [label, ...existing];
    localStorage.setItem('wip_labels', JSON.stringify(labels));
    cache['wip_labels'] = labels;
    queueLabelInsert(label);
  },

  deleteLabel(id: string) {
    const cutoff = Date.now() - 24 * 60 * 60 * 1000;
    const labels = this.getLabels().filter(l => l.id !== id && l.timestamp >= cutoff);
    localStorage.setItem('wip_labels', JSON.stringify(labels));
    cache['wip_labels'] = labels;
    deleteLabelFromSupabase(id);
  },

  markLabelAsPrinted(id: string) {
    const cutoff = Date.now() - 24 * 60 * 60 * 1000;
    const labels = this.getLabels()
      .filter(l => l.timestamp >= cutoff)
      .map(l => l.id === id ? { ...l, printed: true } : l);
    localStorage.setItem('wip_labels', JSON.stringify(labels));
    cache['wip_labels'] = labels;
    const target = labels.find(l => l.id === id);
    if (target) queueLabelInsert(target);
  },

  rollbackTransaction(txId: string) {
    const transactions = this.getTransactions();
    const txIndex = transactions.findIndex(t => t.id === txId);
    
    if (txIndex === -1) {
      // Check labels if not in transactions
      const labels = this.getLabels();
      const label = labels.find(l => l.id === txId);
      if (!label) return;
      
      // Rollback logic for label (Export)
      // 1. Add back to inventory (OUT location is "OK" stock for the stage) (Only if it's not a Quick Print label which doesn't deduct inventory)
      if (!label.planId) {
        this.updateInventory(label.partId, label.stageId, 'OUT', label.quantity);
      }
      
      // 2. Update PO exportedQuantity if label has poId
      if (label.poId) {
        const pos = this.getProductionOrders();
        const po = pos.find(p => p.id === label.poId);
        if (po) {
          po.exportedQuantity = Math.max(0, (po.exportedQuantity || 0) - label.quantity);
          po.status = 'IN_PROGRESS';
          po.completedAt = undefined;
          
          if (po.masterPoId) {
            const masterPo = pos.find(p => p.id === po.masterPoId);
            if (masterPo) {
              masterPo.status = 'IN_PROGRESS';
              masterPo.completedAt = undefined;
            }
          }
          this.saveProductionOrders(pos);
        }
      }
      
      // Rollback Glazing Plan progress if planId exists
      if (label.planId) {
        const plans = this.getGlazingPlans();
        const planIndex = plans.findIndex(p => p.id === label.planId);
        if (planIndex > -1) {
          const plan = plans[planIndex];
          if (plan.producedQuantities && plan.producedQuantities[label.partId]) {
            plan.producedQuantities[label.partId] = Math.max(0, plan.producedQuantities[label.partId] - label.quantity);
          }
          plan.status = 'IN_PROGRESS';
          this.saveGlazingPlans(plans);
        }
      }

      this.deleteLabel(txId);
      return;
    }

    const tx = transactions[txIndex];
    
    // Rollback for standard transactions (STAGE_IN / STAGE_OUT)
    // 1. Revert inventory
    if (tx.type === 'STAGE_IN') {
      // If it was an inbound, deduct from inventory
      this.updateInventory(tx.partId, tx.stageId, tx.location || 'IN', -tx.quantity, tx.originalPartId);
      
      // If it updated producedQuantity (usually for OUT location inbound)
      if (tx.location === 'OUT' && tx.poId) {
        const pos = this.getProductionOrders();
        const po = pos.find(p => p.id === tx.poId);
        if (po) {
          po.producedQuantity = Math.max(0, po.producedQuantity - tx.quantity);
          po.status = 'IN_PROGRESS';
          po.completedAt = undefined;

          if (po.masterPoId) {
            const masterPo = pos.find(p => p.id === po.masterPoId);
            if (masterPo) {
              masterPo.status = 'IN_PROGRESS';
              masterPo.completedAt = undefined;
            }
          }
          this.saveProductionOrders(pos);
        }
      }
    } else if (tx.type === 'STAGE_OUT' || tx.type === 'DISPOSAL') {
      // If it was an outbound/disposal, add back to inventory
      this.updateInventory(tx.partId, tx.stageId, tx.location || 'OUT', tx.quantity, tx.originalPartId);

      // Rollback PO progress for STAGE_OUT
      if (tx.type === 'STAGE_OUT' && tx.poId) {
        const pos = this.getProductionOrders();
        const po = pos.find(p => p.id === tx.poId);
        if (po) {
          if (tx.location === 'IN') {
            // Rollback Move In -> Out (increased producedQty)
            po.producedQuantity = Math.max(0, po.producedQuantity - tx.quantity);
          } else if (tx.location === 'OUT') {
            // Rollback Export (increased exportedQty)
            po.exportedQuantity = Math.max(0, (po.exportedQuantity || 0) - tx.quantity);
            // If producedQuantity was automatically bumped to match exportedQty, pull it back too
            if (po.producedQuantity > po.exportedQuantity) {
               // We only pull back if it's likely it was bumped
               // A better way is to check the transaction history, but staying simple:
               // If producedQuantity == exportedQuantity (after subtracting), then it was likely synced
            }
            // Simple heuristic: if we were at 100/100 and rollback 100 export, go to 0/0
            if (po.producedQuantity > po.exportedQuantity + tx.quantity * 0.1) { 
               // If there was significantly more produced than exported, keep produced as is?
               // Actually, let's just match them if produced was only driven by export
               po.producedQuantity = Math.max(po.exportedQuantity, po.producedQuantity - tx.quantity);
            } else {
               po.producedQuantity = Math.max(0, po.producedQuantity - tx.quantity);
            }
          }
          po.status = 'IN_PROGRESS';
          po.completedAt = undefined;
          
          if (po.masterPoId) {
            const masterPo = pos.find(p => p.id === po.masterPoId);
            if (masterPo) {
              masterPo.status = 'IN_PROGRESS';
              masterPo.completedAt = undefined;
            }
          }
          this.saveProductionOrders(pos);
        }
      }
    }

    // Remove from transactions
    transactions.splice(txIndex, 1);
    this.saveTransactions(transactions);
    
    // Also remove any related labels if exist
    this.deleteLabel(txId);
  },

  getEffectivePartId(partId: string, stageId: StageId, poId?: string): string {
    if (!partId) return '';
    
    const cacheKey = `${partId}|${stageId}|${poId || ''}`;
    if (effectivePartIdCache.has(cacheKey)) {
      return effectivePartIdCache.get(cacheKey)!;
    }

    const transformationsList = this.getTransformations();
    const stageTransformations = transformationsList.filter(t => t.targetStageId === stageId);
    if (stageTransformations.length === 0) {
      effectivePartIdCache.set(cacheKey, partId);
      return partId;
    }

    const std = (s: string) => s ? s.normalize('NFC').trim() : '';
    const inputStd = std(partId);
    const inputUpper = inputStd.toUpperCase();

    // Lấy thông tin từ danh mục để có Tên đầy đủ (vì Quy tắc thường lưu theo Tên)
    const parts = this.getParts();
    const partInCatalog = parts.find(p => p.id === partId || p.name === partId);
    const partIdInCatalog = partInCatalog ? std(partInCatalog.id) : '';
    const partNameInCatalog = partInCatalog ? std(partInCatalog.name) : '';

    let candidates = stageTransformations.filter(t => {
      const sId = std(t.sourcePartId);
      const sIdUpper = sId.toUpperCase();
      
      // 1. So khớp trực tiếp với chuỗi đầu vào
      if (sId === inputStd || sIdUpper === inputUpper) return true;
      
      // 2. So khớp với ID hoặc Tên từ danh mục linh kiện
      if (partIdInCatalog && (sId === partIdInCatalog || sIdUpper === partIdInCatalog.toUpperCase())) return true;
      if (partNameInCatalog && (sId === partNameInCatalog || sIdUpper === partNameInCatalog.toUpperCase())) return true;

      // 3. Chuẩn hóa nâng cao (bỏ tiền tố, hậu tố nhiễu)
      const normalize = (str: string) => {
        let res = str.toUpperCase().normalize('NFC');
        res = res.replace(/^(TẤM|CHI TIẾT|PHỤ TÙNG|BẢN|KHO|THÀNH PHẨM|L-)\s+/g, '');
        res = res.replace(/\s*-\s*(CD|H|C|P|G|W|B|L)$/g, ''); 
        return res.split(' (')[0].split('(')[0].trim();
      };
      
      const normRule = normalize(sId);
      if (normRule === normalize(inputStd)) return true;
      if (partNameInCatalog && normRule === normalize(partNameInCatalog)) return true;

      return false;
    });

    if (candidates.length === 0) {
      effectivePartIdCache.set(cacheKey, partId);
      return partId;
    }

    let bestMatch = candidates[0];
    if (poId) {
      const orders = this.getProductionOrders();
      const po = orders.find(p => p.id === poId);
      const parentPoId = po?.masterPoId || poId;
      const modelPo = orders.find(p => p.id === parentPoId);
      
      if (modelPo) {
        const modelId = modelPo.partId.toUpperCase();
        const modelMatch = candidates.find(t => 
          t.applicableModel && 
          (modelId === t.applicableModel.trim().toUpperCase() || modelId.includes(t.applicableModel.trim().toUpperCase()))
        );
        if (modelMatch) bestMatch = modelMatch;
      }
    } else {
      const generic = candidates.find(t => !t.applicableModel || t.applicableModel.trim() === '');
      if (generic) bestMatch = generic;
    }

    const targetPartIdStr = bestMatch.targetPartId;
    // Map targetPartIdStr to catalog ID if catalog has it
    const upperTargetStr = targetPartIdStr.toUpperCase().normalize('NFC').replace(/\s+/g, '');
    const targetCatalogPart = parts.find(p => {
      const pId = p.id.toUpperCase();
      const pName = p.name.toUpperCase();
      if (p.id === targetPartIdStr || p.name === targetPartIdStr || pId === targetPartIdStr.toUpperCase() || pName === targetPartIdStr.toUpperCase()) return true;
      
      const combined1 = (pName + pId).normalize('NFC').replace(/\s+/g, '');
      const combined2 = (pId + pName).normalize('NFC').replace(/\s+/g, '');
      if (upperTargetStr === combined1 || upperTargetStr === combined2) return true;
      
      // If the target string contains the ID (which is usually highly specific) and it's longer than 5 chars
      if (pId.length > 5 && upperTargetStr.includes(pId.normalize('NFC').replace(/\s+/g, ''))) return true;
      
      return false;
    });
    const finalResult = targetCatalogPart ? targetCatalogPart.id : targetPartIdStr;
    effectivePartIdCache.set(cacheKey, finalResult);
    return finalResult;
  },

  setAbsoluteInventory(partId: string, stageId: StageId, location: 'IN' | 'OUT' | 'DEFECT', newQuantity: number) {
    const inventory = this.getInventory();
    const cleanIdUpper = partId.trim().toUpperCase();
    const parts = this.getParts();
    const partInCatalog = parts.find(p => p.id.toUpperCase() === cleanIdUpper || p.name.toUpperCase().trim() === cleanIdUpper);
    const targetNameUpper = partInCatalog ? partInCatalog.name.toUpperCase().trim() : '';

    // Remove all existing matching entries
    const newInventory = inventory.filter(item => {
      const itemPartId = item.partId.toUpperCase().trim();
      const matchesPart = itemPartId === cleanIdUpper || (targetNameUpper && itemPartId === targetNameUpper);
      return !(matchesPart && item.stageId === stageId && item.location === location);
    });

    // Add back the new quantity if > 0
    if (newQuantity > 0) {
      newInventory.push({
        partId: partInCatalog ? partInCatalog.id : cleanIdUpper,
        stageId,
        location,
        quantity: newQuantity
      });
    }

    this.saveInventory(newInventory);
  },

  updateInventory(partId: string, stageId: StageId, location: 'IN' | 'OUT' | 'DEFECT', delta: number, originalPartId?: string) {
    const inventory = this.getInventory();
    const parts = this.getParts();
    
    // Find absolute correct case in catalog if exists
    const cleanIdUpper = partId.trim().toUpperCase();
    const cleanOrigIdUpper = originalPartId?.trim().toUpperCase() || '';
    
    const partInCatalog = parts.find(p => p.id.toUpperCase() === cleanIdUpper || p.name.toUpperCase().trim() === cleanIdUpper);
    const targetId = partInCatalog ? partInCatalog.id : partId.trim();
    const targetNameUpper = partInCatalog ? partInCatalog.name.toUpperCase().trim() : '';
    
    const origPartInCatalog = originalPartId ? parts.find(p => p.id.toUpperCase() === cleanOrigIdUpper || p.name.toUpperCase().trim() === cleanOrigIdUpper) : undefined;
    const targetOrigId = origPartInCatalog ? origPartInCatalog.id : (originalPartId?.trim() || '');

    const index = inventory.findIndex(
      (item) => {
        const itemPartId = item.partId.toUpperCase().trim();
        const itemOrigId = (item.originalPartId || '').toUpperCase().trim();
        const matchesPart = itemPartId === cleanIdUpper || (targetNameUpper && itemPartId === targetNameUpper);
        return matchesPart && item.stageId === stageId && item.location === location && itemOrigId === cleanOrigIdUpper;
      }
    );

    if (index >= 0) {
      inventory[index].quantity += delta;
      inventory[index].quantity = Math.round(inventory[index].quantity * 10000) / 10000;
      if (inventory[index].quantity < 0) inventory[index].quantity = 0;
      // Option: update it to the proper case if it was wrong
      inventory[index].partId = targetId; 
      if (targetOrigId) inventory[index].originalPartId = targetOrigId;
    } else {
      if (delta > 0) {
        inventory.push({ 
          partId: targetId, 
          originalPartId: targetOrigId || undefined,
          stageId, 
          location, 
          quantity: Math.max(0, delta) 
        });
      }
    }

    this.saveInventory(inventory);
  },

  setInventoryQuantity(partId: string, stageId: StageId, location: 'IN' | 'OUT' | 'DEFECT', quantity: number, originalPartId?: string) {
    const inventory = this.getInventory();
    const parts = this.getParts();
    
    // Find absolute correct case in catalog if exists
    const cleanIdUpper = partId.trim().toUpperCase();
    const cleanOrigIdUpper = originalPartId?.trim().toUpperCase() || '';
    
    const partInCatalog = parts.find(p => p.id.toUpperCase() === cleanIdUpper || p.name.toUpperCase().trim() === cleanIdUpper);
    const targetId = partInCatalog ? partInCatalog.id : partId.trim();
    const targetNameUpper = partInCatalog ? partInCatalog.name.toUpperCase().trim() : '';

    const origPartInCatalog = originalPartId ? parts.find(p => p.id.toUpperCase() === cleanOrigIdUpper || p.name.toUpperCase().trim() === cleanOrigIdUpper) : undefined;
    const targetOrigId = origPartInCatalog ? origPartInCatalog.id : (originalPartId?.trim() || '');

    const index = inventory.findIndex(
      (item) => {
        const itemPartId = item.partId.toUpperCase().trim();
        const itemOrigId = (item.originalPartId || '').toUpperCase().trim();
        const matchesPart = itemPartId === cleanIdUpper || (targetNameUpper && itemPartId === targetNameUpper);
        return matchesPart && item.stageId === stageId && item.location === location && itemOrigId === cleanOrigIdUpper;
      }
    );

    if (index >= 0) {
      inventory[index].quantity = Math.max(0, quantity);
      inventory[index].partId = targetId;
      if (targetOrigId) inventory[index].originalPartId = targetOrigId;
    } else {
      inventory.push({ 
        partId: targetId, 
        originalPartId: targetOrigId || undefined,
        stageId, 
        location, 
        quantity: Math.max(0, quantity) 
      });
    }

    this.saveInventory(inventory);
  },

  deleteInventoryItem(partId: string, stageId: StageId, location: 'IN' | 'OUT' | 'DEFECT', originalPartId?: string) {
    const inventory = this.getInventory();
    const parts = this.getParts();
    const cleanIdUpper = partId.trim().toUpperCase();
    const cleanOrigIdUpper = originalPartId?.trim().toUpperCase() || '';

    const partInCatalog = parts.find(p => p.id.toUpperCase() === cleanIdUpper || p.name.toUpperCase().trim() === cleanIdUpper);
    const targetNameUpper = partInCatalog ? partInCatalog.name.toUpperCase().trim() : '';

    const filtered = inventory.filter(
      (item) => {
        const itemPartId = item.partId.toUpperCase().trim();
        const itemOrigId = (item.originalPartId || '').toUpperCase().trim();
        const matchesPart = itemPartId === cleanIdUpper || (targetNameUpper && itemPartId === targetNameUpper);
        return !(matchesPart && item.stageId === stageId && item.location === location && itemOrigId === cleanOrigIdUpper);
      }
    );
    this.saveInventory(filtered);
    deleteInventoryFromSupabase(partId, stageId, location, originalPartId);
  },

  applyBOMDeduction(partId: string, stageId: StageId, quantity: number, poId?: string) {
    const parts = this.getParts();
    // Strip suffixes added by display logic (e.g., " - CD", " - H") to ensure BOM lookups match the original part ID
    const cleanId = partId.startsWith('GLZ-OUT-') ? partId : partId.split(' - ')[0];

    let currentModelId: string | undefined;
    if (poId) {
      const parentPoId = this.getProductionOrders().find(p => p.id === poId)?.masterPoId || poId;
      currentModelId = this.getProductionOrders().find(p => p.id === parentPoId)?.partId;
    }
    
    // Laser stage specific logic (BOM V1):
    // Deduct Level 3 parts from Laser IN based on BOM when Level 2 is produced
    if (stageId === 'LASER') {
      const bom = this.getBOM();
      const bomDef = bom.find(b => b.childPartId === cleanId);
      
      if (bomDef) {
        const totalConsumption = quantity * (bomDef.componentWeight + bomDef.scrapWeight);
        const totalScrap = quantity * bomDef.scrapWeight;
        
        const inventory = this.getInventory();
        const matchingStocks = inventory.filter(i => {
          const itPartId = i.partId.toUpperCase();
          const targetId = bomDef.parentPartId.toUpperCase();
          return itPartId === targetId && i.stageId === 'LASER' && i.location === 'IN';
        });
        const totalStock = matchingStocks.reduce((sum, i) => sum + i.quantity, 0);
        
        if (totalStock < totalConsumption) {
          const parentPart = parts.find(p => p.id === bomDef.parentPartId);
          throw new Error(`Lỗi: Không đủ tồn kho ${parentPart?.name || bomDef.parentPartId} tại LASER_IN. Cần ${totalConsumption.toFixed(4)} kg, hiện có ${totalStock} kg`);
        }
        
        let remainingToDeduct = totalConsumption;
        for (const stock of matchingStocks) {
          if (remainingToDeduct <= 0) break;
          const toTake = Math.min(stock.quantity, remainingToDeduct);
          this.updateInventory(stock.partId, 'LASER', 'IN', -toTake, stock.originalPartId);
          remainingToDeduct -= toTake;
        }
        const scrapPart = parts.find(p => p.id === 'PL-TON-SX' || p.name.toLowerCase().includes('phế liệu'));
        if (scrapPart) {
          this.updateInventory(scrapPart.id, 'LASER', 'OUT', totalScrap);
        }
      }
    }

    // Welding stage specific logic (BOM V2):
    if (stageId === 'WELDING') {
      const bomV2 = this.getBOMV2();
      // Only deduct ingredients that DON'T skip welding
      const allIngredients = bomV2.filter(b => b.resultPartId === cleanId && (!b.applicableModel || b.applicableModel === currentModelId));
      const ingredients = allIngredients.filter(ing => {
        const p = parts.find(part => part.id === ing.ingredientPartId);
        return !p?.skipWelding;
      });
      
      if (ingredients.length > 0) {
        const inventory = this.getInventory();
        for (const ing of ingredients) {
          const needed = quantity * ing.quantity;
          const effectiveIngId = this.getEffectivePartId(ing.ingredientPartId, 'WELDING', poId);
          
          const partInCatalog = parts.find(p => p.id.toUpperCase() === effectiveIngId.toUpperCase());
          const targetName = partInCatalog ? partInCatalog.name.toUpperCase().trim() : '';

          const matchingStocks = inventory.filter(i => {
            const itemPartId = i.partId.toUpperCase().trim();
            const matchesPart = itemPartId === effectiveIngId.toUpperCase() || (targetName && itemPartId === targetName);
            return matchesPart && i.stageId === 'WELDING' && i.location === 'IN';
          });
          const totalStock = matchingStocks.reduce((sum, i) => sum + i.quantity, 0);

          if (totalStock < needed) {
            const ingPart = parts.find(p => p.id === effectiveIngId);
            throw new Error(`Lỗi: Không đủ tồn kho ${ingPart?.name || effectiveIngId} tại WELDING_IN. Cần ${needed} ${ingPart?.unit || ''}, hiện có ${totalStock}`);
          }
        }
        for (const ing of ingredients) {
          const effectiveIngId = this.getEffectivePartId(ing.ingredientPartId, 'WELDING', poId);
          const partInCatalog = parts.find(p => p.id.toUpperCase() === effectiveIngId.toUpperCase());
          const targetName = partInCatalog ? partInCatalog.name.toUpperCase().trim() : '';
          const matchingStocks = inventory.filter(i => {
            const itemPartId = i.partId.toUpperCase().trim();
            const matchesPart = itemPartId === effectiveIngId.toUpperCase() || (targetName && itemPartId === targetName);
            return matchesPart && i.stageId === 'WELDING' && i.location === 'IN';
          });
          let remainingToDeduct = quantity * ing.quantity;
          for (const stock of matchingStocks) {
            if (remainingToDeduct <= 0) break;
            const toTake = Math.min(stock.quantity, remainingToDeduct);
            this.updateInventory(stock.partId, 'WELDING', 'IN', -toTake, stock.originalPartId);
            remainingToDeduct -= toTake;
          }
        }
      }
    }

    // Painting stage deduction for ingredients that skipped welding
    if (stageId === 'PAINTING') {
      // User request: Skip BOM V2 requirement when producing (IN->OUT) for Painting
      return;
    }
  },

  getNextValidStageId(partId: string, currentStageId: StageId): StageId | null {
    const parts = this.getParts();
    const part = parts.find(p => p.id === partId);
    if (!part) return null;

    const currentStageIndex = STAGES.findIndex(s => s.id === currentStageId);
    if (currentStageIndex === -1) return null;

    for (let i = currentStageIndex + 1; i < STAGES.length; i++) {
      const nextStage = STAGES[i];
      if (nextStage.id === 'LASER' && part.skipLaser) continue;
      if (nextStage.id === 'BENDING' && part.skipBending) continue;
      if (nextStage.id === 'WELDING' && part.skipWelding) continue;
      if (nextStage.id === 'PAINTING' && part.skipPainting) continue;
      return nextStage.id;
    }
    return null;
  },

  recordStageOut(partId: string, stageId: StageId, quantity: number, sourceLocation: 'IN' | 'OUT' = 'IN', targetStageId?: StageId, poId?: string, force?: boolean, customTimestamp?: number) {
    const cleanId = partId.startsWith('GLZ-OUT-') ? partId.trim().toUpperCase() : partId.split(' - ')[0].trim().toUpperCase();
    const pos = this.getProductionOrders();
    const poIndex = poId 
      ? pos.findIndex(p => p.id === poId)
      : pos.findIndex(p => p.partId === cleanId && p.stageId === stageId && p.status !== 'COMPLETED');
      
    let linkedPoId = poId;
    if (poIndex !== -1) {
      linkedPoId = pos[poIndex].id;
    }

    // Validation: Check if source location has enough quantity
    const inventory = this.getInventory();
    const effectiveId = this.getEffectivePartId(cleanId, stageId, linkedPoId);
    
    // 0. Validate PO limit before modifying state (STRICT: Never allow exceeding PO limit at any stage)
    if (poIndex !== -1) {
      const po = pos[poIndex];
      const stageName = STAGES.find(s => s.id === stageId)?.name || stageId;
      const partObj = this.getParts().find(p => p.id.toUpperCase() === cleanId.toUpperCase() || p.name.toUpperCase() === cleanId.toUpperCase());
      const partDisplayName = partObj?.name || cleanId;

      if (sourceLocation === 'IN') {
        const nextQty = po.producedQuantity + quantity;
        if (nextQty > po.targetQuantity) {
          const maxRemaining = Math.max(0, po.targetQuantity - po.producedQuantity);
          throw new Error(`OVER_PO:Số lượng xuất/hoàn thành (${nextQty}) sẽ vượt quá số lượng mục tiêu của PO [${po.id}] (${po.targetQuantity}) cho linh kiện "${partDisplayName}" tại công đoạn ${stageName} (Hiện đã làm ${po.producedQuantity}, còn lại ${maxRemaining}).\n\nHệ thống KHÔNG cho phép xuất vượt quá PO!`);
        }
      } else if (sourceLocation === 'OUT') {
        const currentExp = po.exportedQuantity || 0;
        const nextExp = currentExp + quantity;
        if (nextExp > po.targetQuantity) {
          const maxRemaining = Math.max(0, po.targetQuantity - currentExp);
          throw new Error(`OVER_PO:Số lượng xuất kho (${nextExp}) sẽ vượt quá số lượng mục tiêu của PO [${po.id}] (${po.targetQuantity}) cho linh kiện "${partDisplayName}" tại kho OUT ${stageName} (Hiện đã xuất ${currentExp}, còn lại ${maxRemaining}).\n\nHệ thống KHÔNG cho phép xuất vượt quá PO!`);
        }
      }
    }
    
    const partsInCatalog = this.getParts();
    const selectedPartInCatalog = partsInCatalog.find(p => p.id === cleanId || p.name === cleanId);
    const selectedPartName = selectedPartInCatalog?.name.toUpperCase();
    const selectedPartId = selectedPartInCatalog?.id.toUpperCase();

    const partInCatalogForEffectiveId = partsInCatalog.find(p => p.id.toUpperCase() === effectiveId.toUpperCase());
    const effectivePartName = partInCatalogForEffectiveId ? partInCatalogForEffectiveId.name.toUpperCase().trim() : '';

    const matchingStocks = inventory.filter((item) => {
      if (item.stageId === stageId && item.location === sourceLocation) {
        const itemPartId = item.partId.toUpperCase().trim();
        const targetId = effectiveId.toUpperCase();
        return itemPartId === targetId || (effectivePartName && itemPartId === effectivePartName);
      }
      return false;
    });
    const totalStock = matchingStocks.reduce((sum, item) => sum + item.quantity, 0);

    if (totalStock < quantity) {
      if (!force) {
        const part = this.getParts().find(p => p.id.toUpperCase() === effectiveId.toUpperCase());
        throw new Error(`Lỗi: Số lượng xuất (${quantity}) lớn hơn tổng tồn kho ${part?.name || effectiveId} tại ${STAGES.find(s => s.id === stageId)?.name}_${sourceLocation} (Hiện có ${totalStock})`);
      }
    }

    // 0. Update Production Order progress
    if (sourceLocation === 'IN' && poIndex !== -1) {
      const po = pos[poIndex];
      po.producedQuantity += quantity;
      const isProduced = po.producedQuantity >= po.targetQuantity;
      const isExported = (po.exportedQuantity || 0) >= po.targetQuantity;
      po.status = (isProduced && isExported) ? 'COMPLETED' : 'IN_PROGRESS';
      linkedPoId = po.id;

      if (po.masterPoId) {
        const masterPo = pos.find(p => p.id === po.masterPoId);
        if (masterPo) {
          const allSubsCompleted = po.status === 'COMPLETED' && pos.filter(p => p.masterPoId === po.masterPoId && p.id !== po.id).every(s => s.status === 'COMPLETED');
          masterPo.status = allSubsCompleted ? 'COMPLETED' : 'IN_PROGRESS';
        }
      }
      this.saveProductionOrders(pos);
    } else if (sourceLocation === 'OUT' && poIndex !== -1) {
      const po = pos[poIndex];
      if ((po.exportedQuantity || 0) + quantity > po.producedQuantity) {
        po.producedQuantity = (po.exportedQuantity || 0) + quantity;
      }
      po.exportedQuantity = (po.exportedQuantity || 0) + quantity;
      const isProduced = po.producedQuantity >= po.targetQuantity;
      const isExported = po.exportedQuantity >= po.targetQuantity;
      if (isProduced && isExported && po.status !== 'COMPLETED') po.completedAt = Date.now();
      po.status = (isProduced && isExported) ? 'COMPLETED' : 'IN_PROGRESS';

      if (po.masterPoId) {
        const masterPo = pos.find(p => p.id === po.masterPoId);
        if (masterPo && pos.filter(p => p.masterPoId === po.masterPoId).every(p => p.status === 'COMPLETED')) {
          if (masterPo.status !== 'COMPLETED') masterPo.completedAt = Date.now();
          masterPo.status = 'COMPLETED';
        }
      }
      this.saveProductionOrders(pos);
    }

    // 1. Inventory movement
    let remainingToDeduct = quantity;
    let lastOriginalId: string | undefined;
    matchingStocks.sort((a, b) => (a.originalPartId || '').localeCompare(b.originalPartId || ''));

    if (sourceLocation === 'IN') {
      this.applyBOMDeduction(cleanId, stageId, quantity, linkedPoId);
      for (const stock of matchingStocks) {
        if (remainingToDeduct <= 0) break;
        const toTake = Math.min(stock.quantity, remainingToDeduct);
        this.updateInventory(effectiveId, stageId, 'IN', -toTake, stock.originalPartId);
        this.updateInventory(effectiveId, stageId, 'OUT', toTake, stock.originalPartId);
        lastOriginalId = stock.originalPartId;
        remainingToDeduct -= toTake;
      }
    } else {
      for (const stock of matchingStocks) {
        if (remainingToDeduct <= 0) break;
        const toTake = Math.min(stock.quantity, remainingToDeduct);
        this.updateInventory(effectiveId, stageId, 'OUT', -toTake, stock.originalPartId);
        lastOriginalId = stock.originalPartId;
        remainingToDeduct -= toTake;
      }
    }

    // 2. Record transaction
    const transactions = this.getTransactions();
    const po = pos.find(p => p.id === linkedPoId);
    const masterPoId = po?.masterPoId || '';
    const subPoTargetQty = po?.targetQuantity || 0;
    const masterPo = masterPoId ? pos.find(p => p.id === masterPoId) : undefined;
    const masterPoTargetQty = masterPo?.targetQuantity || 0;
    
    const txId = Math.random().toString(36).substring(2, 12).toUpperCase();
    const timestamp = Date.now();
    const qrData = sourceLocation === 'OUT' 
      ? `${linkedPoId || effectiveId}|${quantity}|${stageId}|${timestamp}|${txId}|${targetStageId || ''}|${effectiveId}|${masterPoId}|${subPoTargetQty}|${masterPoTargetQty}`
      : undefined;

    const newTransaction: Transaction = {
      id: txId,
      type: 'STAGE_OUT',
      partId: effectiveId,
      partName: this.getParts().find(p => p.id === effectiveId)?.name || effectiveId,
      originalPartId: lastOriginalId,
      quantity,
      stageId,
      location: sourceLocation,
      targetStageId,
      timestamp: customTimestamp || Date.now(),
      qrData,
      poId: linkedPoId,
      printed: stageId === 'GLAZING' && sourceLocation === 'OUT' ? false : undefined,
      kpiRecorded: stageId === 'GLAZING' && sourceLocation === 'OUT' ? true : undefined
    };
    transactions.unshift(newTransaction);
    this.saveTransactions(transactions);

    if (qrData) {
      this.saveLabel(newTransaction);
    }

    return newTransaction;
  },

  setTransactionPrinted(txId: string, printed: boolean) {
    const transactions: Transaction[] = this.getTransactions();
    const idx = transactions.findIndex(t => t.id === txId);
    if (idx !== -1) {
      transactions[idx].printed = printed;
      // Also update in label history
      const labels: Transaction[] = this.getLabels();
      const lIdx = labels.findIndex(l => l.id === txId);
      if (lIdx !== -1) {
        labels[lIdx].printed = printed;
        this.saveLabels(labels);
      }
      this.saveTransactions(transactions);
    }
  },

  recordStageIn(qrData: string, currentStageId: StageId, targetLocation: 'IN' | 'OUT' = 'IN') {
    if (!qrData || typeof qrData !== 'string') {
      throw new Error('Mã QR không hợp lệ');
    }
    const parts = qrData.split('|');
    if (parts.length < 5) {
      throw new Error('Định dạng mã QR không hợp lệ hoặc không phải mã xuất kho OUT.');
    }
    
    // Format: poIdOrPartId|quantity|sourceStageId|timestamp|txId|targetStageId
    const [idOrPo, quantityStr, sourceStageId, , sourceTxId, targetStageId] = parts;

    if (idOrPo === 'DISPOSAL') {
      throw new Error('Lỗi: Đây là mã QR XUẤT HỦY. Hàng này không thể nhập lại vào kho sản xuất!');
    }

    const quantity = parseFloat(quantityStr);

    let partId = idOrPo.startsWith('GLZ-OUT-') ? idOrPo : idOrPo.split(' - ')[0];
    let linkedPoId: string | undefined;

    // Recognize PO IDs (PO- or REPAIR-)
    if (idOrPo.startsWith('PO-') || idOrPo.startsWith('REPAIR-')) {
      linkedPoId = idOrPo;
      const pos = this.getProductionOrders();
      const po = pos.find(p => p.id === idOrPo);
      if (po) {
        partId = po.partId;
      }
    }

    // 1. Check if this QR (Transaction ID) has already been scanned
    const transactions = this.getTransactions();
    
    // Interlock: Check if this label was marked as Defect (if we ever support that)
    const isDefect = transactions.some(tx => tx.id === sourceTxId && tx.type === 'DEFECT');
    if (isDefect) {
      throw new Error('CẢNH BÁO: Nhãn này đã bị đánh dấu là HÀNG LỖI (DEFECT). Không thể nhập kho công đoạn tiếp theo!');
    }

    const alreadyScanned = transactions.some(tx => tx.type === 'STAGE_IN' && tx.qrData?.includes(sourceTxId));
    if (alreadyScanned) {
      throw new Error('Lỗi: Mã QR này đã được sử dụng để nhập kho trước đó. Không thể nhập lại.');
    }

    // 2. Prevent scanning into the same stage it was exported from
    if (sourceStageId === currentStageId) {
      throw new Error('Lỗi: Không được phép nhập lại tại chính công đoạn đã xuất kho.');
    }

    // 3. If a target stage was specified, ensure it matches current stage
    if (targetStageId && targetStageId !== currentStageId) {
      const targetStageName = STAGES.find(s => s.id === targetStageId)?.name || targetStageId;
      throw new Error(`Lỗi: Nhãn này được chỉ định cho công đoạn ${targetStageName}. Bạn đang ở công đoạn ${STAGES.find(s => s.id === currentStageId)?.name}.`);
    }

    if (!linkedPoId && !partId.startsWith('GLZ-OUT-') && sourceStageId !== 'GLAZING') {
      throw new Error('Lỗi: Nhãn QR này không chứa thông tin Lệnh sản xuất (PO). Không thể nhập kho.');
    }

    // 3. Add to currentStage target location
    // FORCE targetLocation to 'IN' when scanning QR code as per user request
    const finalTargetLocation = 'IN';
    
    // Part Transformation Logic
    const finalPartId = this.getEffectivePartId(partId, currentStageId, linkedPoId);

    // Strict PO limit check for currentStageId (Never allow exceeding PO limit at any stage)
    if (linkedPoId) {
      const pos = this.getProductionOrders();
      let targetPo = pos.find(p => p.id === linkedPoId && p.stageId === currentStageId);
      if (!targetPo) {
        const sourcePo = pos.find(p => p.id === linkedPoId);
        if (sourcePo && sourcePo.masterPoId) {
          targetPo = pos.find(p => p.masterPoId === sourcePo.masterPoId && p.stageId === currentStageId && (p.partId.toUpperCase() === finalPartId.toUpperCase() || p.partId.toUpperCase() === partId.toUpperCase()));
        }
      }
      if (!targetPo) {
        targetPo = pos.find(p => (p.partId.toUpperCase() === finalPartId.toUpperCase() || p.partId.toUpperCase() === partId.toUpperCase()) && p.stageId === currentStageId && p.status !== 'COMPLETED');
      }

      if (targetPo) {
        const transactions = this.getTransactions();
        const alreadyInbound = transactions
          .filter(tx => tx.type === 'STAGE_IN' && tx.stageId === currentStageId && (tx.poId === targetPo.id || (targetPo.masterPoId && tx.poId && pos.find(p => p.id === tx.poId)?.masterPoId === targetPo.masterPoId && (tx.partId.toUpperCase() === finalPartId.toUpperCase() || tx.partId.toUpperCase() === partId.toUpperCase()))))
          .reduce((sum, tx) => sum + tx.quantity, 0);

        const nextInbound = alreadyInbound + quantity;
        if (nextInbound > targetPo.targetQuantity) {
          const maxRemaining = Math.max(0, targetPo.targetQuantity - alreadyInbound);
          const stageName = STAGES.find(s => s.id === currentStageId)?.name || currentStageId;
          const partObj = this.getParts().find(p => p.id.toUpperCase() === finalPartId.toUpperCase() || p.id.toUpperCase() === partId.toUpperCase());
          const partDisplayName = partObj?.name || finalPartId;
          throw new Error(`OVER_PO:Số lượng quét nhập kho (${nextInbound}) sẽ vượt quá số lượng mục tiêu của PO [${targetPo.id}] (${targetPo.targetQuantity}) cho linh kiện "${partDisplayName}" tại công đoạn ${stageName} (Hiện đã nhập ${alreadyInbound}, còn lại ${maxRemaining}).\n\nHệ thống KHÔNG cho phép nhập vượt quá PO!`);
        }
      }
    }
    
    // So khớp chuẩn hóa để xác định xem có sự thay đổi thực sự không
    const std = (s: string) => s ? s.normalize('NFC').trim().toUpperCase() : '';
    const originalPartId = (std(finalPartId) !== std(partId)) ? partId : undefined;

    this.updateInventory(finalPartId, currentStageId, finalTargetLocation, quantity, originalPartId);

    // 4. Record transaction
    const newTransaction: Transaction = {
      id: Math.random().toString(36).substring(2, 12).toUpperCase(),
      type: 'STAGE_IN',
      partId: finalPartId,
      originalPartId: originalPartId, // Track original if transformed
      quantity,
      stageId: currentStageId,
      location: finalTargetLocation,
      sourceStageId: sourceStageId as StageId,
      timestamp: Date.now(),
      qrData,
      poId: linkedPoId
    };
    transactions.unshift(newTransaction);
    this.saveTransactions(transactions);

    // 5. Update inventory for Level 2 parts if needed (BOM logic for Laser stage)
    // Actually, recordManualInbound has this logic, but recordStageIn should probably have it too if we scan a label.
    // Let's keep it consistent.
    
    return newTransaction;
  },

  recordManualInbound(partId: string, stageId: StageId, location: 'IN' | 'OUT', quantity: number, poId?: string, force?: boolean, customTimestamp?: number) {
    const cleanId = partId.startsWith('GLZ-OUT-') ? partId.trim().toUpperCase() : partId.split(' - ')[0].trim().toUpperCase();
    let linkedPoId = poId;

    // Check for required PO
    const pos = this.getProductionOrders();
    const poIndex = poId 
      ? pos.findIndex(p => p.id === poId)
      : pos.findIndex(p => p.partId === cleanId && p.stageId === stageId && p.status !== 'COMPLETED');
      
    const currentPoId = poIndex !== -1 ? pos[poIndex].id : poId;

    // Apply BOM logic if entering into OUT (Production result)
    if (location === 'OUT') {
      // Validate PO limit BEFORE deducting BOM and saving anything (STRICT: Never allow exceeding PO limit)
      if (poIndex !== -1) {
        const po = pos[poIndex];
        const nextProd = po.producedQuantity + quantity;
        if (nextProd > po.targetQuantity) {
          const maxRemaining = Math.max(0, po.targetQuantity - po.producedQuantity);
          const stageName = STAGES.find(s => s.id === stageId)?.name || stageId;
          const partObj = this.getParts().find(p => p.id.toUpperCase() === cleanId.toUpperCase() || p.name.toUpperCase() === cleanId.toUpperCase());
          const partDisplayName = partObj?.name || cleanId;
          throw new Error(`OVER_PO:Số lượng thêm vào (${nextProd}) sẽ vượt quá mục tiêu PO (${po.targetQuantity}) cho ${partDisplayName} tại ${stageName} (Hiện đã làm ${po.producedQuantity}, còn lại ${maxRemaining}).\n\nHệ thống KHÔNG cho phép nhập vượt quá PO!`);
        }
      }

      this.applyBOMDeduction(cleanId, stageId, quantity, currentPoId);

      // Update PO progress
      if (poIndex !== -1) {
        const po = pos[poIndex];
        po.producedQuantity += quantity;
        const isProduced = po.producedQuantity >= po.targetQuantity;
        const isExported = (po.exportedQuantity || 0) >= po.targetQuantity;
        po.status = (isProduced && isExported) ? 'COMPLETED' : 'IN_PROGRESS';
        linkedPoId = po.id;

        // Check if all sub-POs for this master are completed
        if (po.masterPoId) {
          const masterPo = pos.find(p => p.id === po.masterPoId);
          if (masterPo) {
            const otherSubs = pos.filter(p => p.masterPoId === po.masterPoId && p.id !== po.id);
            const allSubsCompleted = po.status === 'COMPLETED' && otherSubs.every(s => s.status === 'COMPLETED');
            if (allSubsCompleted) {
              masterPo.status = 'COMPLETED';
            } else {
              masterPo.status = 'IN_PROGRESS';
            }
          }
        }

        this.saveProductionOrders(pos);
      }
    } else if (location === 'IN') {
      // Validate PO limit for IN (Never allow exceeding PO limit at any stage)
      if (poIndex !== -1) {
        const po = pos[poIndex];
        const transactions = this.getTransactions();
        const alreadyIn = transactions
          .filter(tx => tx.type === 'STAGE_IN' && tx.stageId === stageId && tx.poId === po.id && tx.location === 'IN')
          .reduce((sum, tx) => sum + tx.quantity, 0);
        const nextIn = alreadyIn + quantity;
        if (nextIn > po.targetQuantity) {
          const maxRemaining = Math.max(0, po.targetQuantity - alreadyIn);
          const stageName = STAGES.find(s => s.id === stageId)?.name || stageId;
          const partObj = this.getParts().find(p => p.id.toUpperCase() === cleanId.toUpperCase() || p.name.toUpperCase() === cleanId.toUpperCase());
          const partDisplayName = partObj?.name || cleanId;
          throw new Error(`OVER_PO:Số lượng nhập kho IN (${nextIn}) sẽ vượt quá số lượng mục tiêu của PO [${po.id}] (${po.targetQuantity}) cho ${partDisplayName} tại ${stageName} (Hiện đã nhập ${alreadyIn}, còn lại ${maxRemaining}).\n\nHệ thống KHÔNG cho phép nhập vượt quá PO!`);
        }
      }
    }

    // Part Transformation Logic
    const finalPartId = this.getEffectivePartId(cleanId, stageId, currentPoId);
    
    const std = (s: string) => s ? s.normalize('NFC').trim().toUpperCase() : '';
    const originalPartId = (std(finalPartId) !== std(cleanId)) ? cleanId : undefined;

    this.updateInventory(finalPartId, stageId, location, quantity, originalPartId);
    
    const transactions = this.getTransactions();
    const newTransaction: Transaction = {
      id: Math.random().toString(36).substring(2, 12).toUpperCase(),
      type: 'STAGE_IN',
      partId: finalPartId,
      originalPartId: originalPartId,
      quantity,
      stageId,
      location, // Store the target location (IN/OUT)
      timestamp: customTimestamp || Date.now(),
      qrData: 'MANUAL_ENTRY',
      poId: linkedPoId
    };
    transactions.unshift(newTransaction);
    this.saveTransactions(transactions);

    return newTransaction;
  },

  getInventoryAtStage(stageId: StageId) {
    const inventory = this.getInventory();
    return {
      in: inventory.filter((item) => item.stageId === stageId && item.location === 'IN'),
      out: inventory.filter((item) => item.stageId === stageId && item.location === 'OUT'),
      defect: inventory.filter((item) => item.stageId === stageId && item.location === 'DEFECT'),
    };
  },

  recordDefect(partId: string, stageId: StageId, location: 'IN' | 'OUT', quantity: number, reason: string, category: string, poId?: string) {
    const cleanId = partId.startsWith('GLZ-OUT-') ? partId.trim().toUpperCase() : partId.split(' - ')[0].trim().toUpperCase();
    
    // 1. Validation: Ensure we have enough stock in IN to mark as defect
    const inventory = this.getInventory();
    const effectiveId = this.getEffectivePartId(cleanId, stageId, poId);
    
    const partsInCatalog = this.getParts();
    const partInCatalogForEffectiveId = partsInCatalog.find(p => p.id.toUpperCase() === effectiveId.toUpperCase());
    const effectivePartName = partInCatalogForEffectiveId ? partInCatalogForEffectiveId.name.toUpperCase().trim() : '';

    const matchingStocks = inventory.filter(i => {
      const itPartId = i.partId.toUpperCase().trim();
      const targetId = effectiveId.toUpperCase();
      const matchesPart = itPartId === targetId || (effectivePartName && itPartId === effectivePartName);
      return matchesPart && i.stageId === stageId && i.location === location;  // Fix: Check specific location
    });
    const totalStock = matchingStocks.reduce((sum, item) => sum + item.quantity, 0);

    if (stageId !== 'GLAZING' && totalStock < quantity) {
      const part = this.getParts().find(p => p.id === effectiveId);
      throw new Error(`Lỗi: Số lượng báo lỗi (${quantity}) lớn hơn tổng tồn kho IN của ${part?.name || effectiveId} tại ${STAGES.find(s => s.id === stageId)?.name} (Hiện có ${totalStock})`);
    }

    // 2. Inventory movement: Deduct from IN, Add to DEFECT
    let remainingToDeduct = quantity;
    let deductedFromOriginalId: string | undefined;

    matchingStocks.sort((a, b) => (a.originalPartId || '').localeCompare(b.originalPartId || ''));

    for (const stock of matchingStocks) {
      if (remainingToDeduct <= 0) break;
      const toTake = Math.min(stock.quantity, remainingToDeduct);
      this.updateInventory(effectiveId, stageId, location, -toTake, stock.originalPartId);
      this.updateInventory(effectiveId, stageId, 'DEFECT', toTake, stock.originalPartId);
      deductedFromOriginalId = stock.originalPartId;
      remainingToDeduct -= toTake;
    }

    if (remainingToDeduct > 0) {
      // Bypassed check cases (e.g. GLAZING)
      this.updateInventory(effectiveId, stageId, 'DEFECT', remainingToDeduct, deductedFromOriginalId);
    }

    // 3. Record transaction
    const transactions = this.getTransactions();
    const txId = `DF-${Math.random().toString(36).substring(2, 8).toUpperCase()}`;
    
    const timestamp = Date.now();
    const qrData = stageId === 'GLAZING' ? `DISPOSAL|${quantity}|${stageId}|${timestamp}|${txId}|DEFECT_ONLY|${reason || ''}|${category || ''}` : undefined;
    
    const newTransaction: Transaction = {
      id: txId,
      type: 'DEFECT',
      partId: effectiveId,
      originalPartId: deductedFromOriginalId,
      quantity,
      stageId,
      timestamp,
      defectReason: reason,
      defectCategory: category,
      poId: poId,
      qrData
    };
    transactions.unshift(newTransaction);
    this.saveTransactions(transactions);

    // Auto-create supplementary POs for compensation for any stage
    if (stageId !== 'LASER') {
      try {
        const repairId = deductedFromOriginalId || cleanId;
        this.createMasterPO(repairId, quantity, Date.now(), undefined, "REPAIR", stageId);
      } catch (err) {
        console.error("Failed to create supplementary PO:", err);
      }
    }

    return newTransaction;
  },

  recordDisposal(partId: string, stageId: StageId, quantity: number, reasonName?: string, categoryName?: string) {
    const cleanId = partId.startsWith('GLZ-OUT-') ? partId.trim().toUpperCase() : partId.split(' - ')[0].trim().toUpperCase();
    
    // 1. Validation
    const inventory = this.getInventory();
    const effectiveId = this.getEffectivePartId(cleanId, stageId);
    const partsInCatalog = this.getParts();
    const partInCatalogForEffectiveId = partsInCatalog.find(p => p.id.toUpperCase() === effectiveId.toUpperCase());
    const effectivePartName = partInCatalogForEffectiveId ? partInCatalogForEffectiveId.name.toUpperCase().trim() : '';
    const matchingStocks = inventory.filter(i => {
      const itPartId = i.partId.toUpperCase().trim();
      const targetId = effectiveId.toUpperCase();
      const matchesPart = itPartId === targetId || (effectivePartName && itPartId === effectivePartName);
      return matchesPart && i.stageId === stageId && i.location === 'DEFECT';
    });
    const totalStock = matchingStocks.reduce((sum, i) => sum + i.quantity, 0);

    if (totalStock < quantity) {
      const part = this.getParts().find(p => p.id === effectiveId);
      throw new Error(`Lỗi: Số lượng xuất hủy (${quantity}) lớn hơn tồn kho DEFECT của ${part?.name || effectiveId} tại ${STAGES.find(s => s.id === stageId)?.name} (Hiện có ${totalStock})`);
    }

    // 2. Inventory move: Deduct from DEFECT
    let remainingToDeduct = quantity;
    for (const stock of matchingStocks) {
      if (remainingToDeduct <= 0) break;
      const toTake = Math.min(stock.quantity, remainingToDeduct);
      this.updateInventory(stock.partId, stageId, 'DEFECT', -toTake, stock.originalPartId);
      remainingToDeduct -= toTake;
    }

    // 3. Record transaction
    const transactions = this.getTransactions();
    const txId = `DS-${Math.random().toString(36).substring(2, 8).toUpperCase()}`;
    const timestamp = Date.now();
    
    // QR data for disposal
    const qrData = `DISPOSAL|${quantity}|${stageId}|${timestamp}|${txId}|DISPOSAL_ONLY|${reasonName || ''}|${categoryName || ''}`;

    const newTransaction: Transaction = {
      id: txId,
      type: 'DISPOSAL',
      partId: effectiveId,
      originalPartId: (effectiveId !== cleanId) ? cleanId : undefined,
      quantity,
      stageId,
      timestamp,
      qrData,
      defectReason: reasonName,
      defectCategory: categoryName,
    };
    transactions.unshift(newTransaction);
    this.saveTransactions(transactions);
    
    // Save to label history for reprint
    this.saveLabel(newTransaction);

    return newTransaction;
  },

  filterExpiredCompletedPOs(orders: ProductionOrder[]): ProductionOrder[] {
    if (!Array.isArray(orders) || orders.length === 0) return [];
    const cutoff = Date.now() - 35 * 24 * 60 * 60 * 1000; // 35 days cutoff

    // Find master POs that have at least one uncompleted sub-PO
    const masterIdsWithIncompleteChildren = new Set<string>();
    orders.forEach(po => {
      const isDone = po.status === 'COMPLETED' || 
        (po.targetQuantity > 0 && po.producedQuantity >= po.targetQuantity && (po.exportedQuantity || 0) >= po.targetQuantity);
      if (!isDone && po.masterPoId) {
        masterIdsWithIncompleteChildren.add(po.masterPoId);
      }
    });

    return orders.filter(po => {
      const isDone = po.status === 'COMPLETED' || 
        (po.targetQuantity > 0 && po.producedQuantity >= po.targetQuantity && (po.exportedQuantity || 0) >= po.targetQuantity);

      // "chưa hoàn thành để lại" - All unfinished POs are preserved
      if (!isDone) return true;

      // If a sub-PO has an unfinished sibling under the same master PO, preserve it
      if (po.masterPoId && masterIdsWithIncompleteChildren.has(po.masterPoId)) return true;

      // If a master PO has any unfinished sub-PO, preserve it
      if (!po.masterPoId && masterIdsWithIncompleteChildren.has(po.id)) return true;

      // Completed PO: check if completed over 35 days ago
      const poTime = po.completedAt || po.expectedCompletionTime || po.plannedStartTime || po.createdAt || 0;
      if (poTime > 0 && poTime < cutoff) {
        return false; // Remove completed PO older than 35 days
      }

      return true;
    });
  },

  getProductionOrders(): ProductionOrder[] {
    const data = getCached(STORAGE_KEYS.PRODUCTION_ORDERS, () => {
      const data = localStorage.getItem(STORAGE_KEYS.PRODUCTION_ORDERS);
      let parsed: ProductionOrder[] = data ? JSON.parse(data) : [];
      const initialLength = parsed.length;
      parsed = this.filterExpiredCompletedPOs(parsed);
      if (parsed.length !== initialLength) {
        localStorage.setItem(STORAGE_KEYS.PRODUCTION_ORDERS, JSON.stringify(parsed));
      }
      return parsed;
    });
    return [...data];
  },

  saveProductionOrders(orders: ProductionOrder[]) {
    const filtered = this.filterExpiredCompletedPOs(orders);
    localStorage.setItem(STORAGE_KEYS.PRODUCTION_ORDERS, JSON.stringify(filtered));
    cache[STORAGE_KEYS.PRODUCTION_ORDERS] = filtered;
    queueProductionOrdersPersist(filtered);
  },

  resetShiftConfigs() {
    localStorage.removeItem(STORAGE_KEYS.SHIFT_CONFIGS);
    clearCache(STORAGE_KEYS.SHIFT_CONFIGS);
  },

  getNextWorkingTime(timestamp: number, stageId: StageId, shiftConfigs: ShiftConfig[]): number {
    const config = shiftConfigs.find(c => c.stageId === stageId);
    if (!config) return timestamp;

    const timeToDate = (timeStr: string, baseDate: Date) => {
      const [h, m] = timeStr.split(':').map(Number);
      return setSeconds(setMinutes(setHours(baseDate, h), m), 0);
    };

    let checkTime = new Date(timestamp);
    for (let day = 0; day < 100; day++) { // Increase day limit
      const baseDay = startOfDay(checkTime);
      const workingIntervals: { start: Date, end: Date }[] = [];
      
      if (!(baseDay.getDay() === 0 && !config.workOnSunday)) {
        config.shifts.forEach(shift => {
          const s = timeToDate(shift.start, baseDay);
          let e = timeToDate(shift.end, baseDay);
          if (isBefore(e, s)) e = addDays(e, 1);
          
          let intervals = [{ start: s, end: e }];
          config.breaks.forEach(brk => {
            const bs = timeToDate(brk.start, baseDay);
            const be = timeToDate(brk.end, baseDay);
            const newIntervals: typeof intervals = [];
            intervals.forEach(inv => {
              if (isAfter(be, inv.start) && isBefore(bs, inv.end)) {
                if (isAfter(bs, inv.start)) newIntervals.push({ start: inv.start, end: bs });
                if (isBefore(be, inv.end)) newIntervals.push({ start: be, end: inv.end });
              } else {
                newIntervals.push(inv);
              }
            });
            intervals = newIntervals;
          });
          workingIntervals.push(...intervals);
        });
        workingIntervals.sort((a, b) => a.start.getTime() - b.start.getTime());

        for (const inv of workingIntervals) {
          if (isBefore(checkTime, inv.end)) {
            if (isBefore(checkTime, inv.start)) return inv.start.getTime();
            else return checkTime.getTime();
          }
        }
      }
      checkTime = startOfDay(addDays(baseDay, 1));
    }
    return timestamp;
  },

  getPreviousWorkingTime(timestamp: number, stageId: StageId, shiftConfigs: ShiftConfig[]): number {
    const config = shiftConfigs.find(c => c.stageId === stageId);
    if (!config) return timestamp;

    const timeToDate = (timeStr: string, baseDate: Date) => {
      const [h, m] = timeStr.split(':').map(Number);
      return setSeconds(setMinutes(setHours(baseDate, h), m), 0);
    };

    let checkTime = new Date(timestamp);
    for (let day = 0; day < 100; day++) { // Increase day limit
      const baseDay = startOfDay(checkTime);
      const workingIntervals: { start: Date, end: Date }[] = [];
      
      if (!(baseDay.getDay() === 0 && !config.workOnSunday)) {
        config.shifts.forEach(shift => {
          const s = timeToDate(shift.start, baseDay);
          let e = timeToDate(shift.end, baseDay);
          if (isBefore(e, s)) e = addDays(e, 1);
          
          let intervals = [{ start: s, end: e }];
          config.breaks.forEach(brk => {
            const bs = timeToDate(brk.start, baseDay);
            const be = timeToDate(brk.end, baseDay);
            const newIntervals: typeof intervals = [];
            intervals.forEach(inv => {
              if (isAfter(be, inv.start) && isBefore(bs, inv.end)) {
                if (isAfter(bs, inv.start)) newIntervals.push({ start: inv.start, end: bs });
                if (isBefore(be, inv.end)) newIntervals.push({ start: be, end: inv.end });
              } else {
                newIntervals.push(inv);
              }
            });
            intervals = newIntervals;
          });
          workingIntervals.push(...intervals);
        });
        // Sort intervals descending to easily find previous time
        workingIntervals.sort((a, b) => b.start.getTime() - a.start.getTime());

        for (const inv of workingIntervals) {
          if (isAfter(checkTime, inv.start)) {
            if (isAfter(checkTime, inv.end)) return inv.end.getTime();
            else return checkTime.getTime();
          }
        }
      }
      checkTime = new Date(baseDay.getTime() - 1); // 23:59:59.999 of previous day
    }
    return timestamp;
  },

  calculateStartTime(endTime: number, durationMs: number, stageId: StageId, shiftConfigs: ShiftConfig[]): number {
    if (durationMs <= 0) return endTime;
    let remaining = durationMs;
    let currentTime = this.getPreviousWorkingTime(endTime, stageId, shiftConfigs);
    const config = shiftConfigs.find(c => c.stageId === stageId);
    if (!config) return endTime - durationMs;

    const timeToDate = (timeStr: string, baseDate: Date) => {
      const [h, m] = timeStr.split(':').map(Number);
      return setSeconds(setMinutes(setHours(baseDate, h), m), 0);
    };

    while (remaining > 0) {
      const baseDay = startOfDay(new Date(currentTime));
      const workingIntervals: { start: Date, end: Date }[] = [];
      
      if (!(baseDay.getDay() === 0 && !config.workOnSunday)) {
        config.shifts.forEach(shift => {
          const s = timeToDate(shift.start, baseDay);
          let e = timeToDate(shift.end, baseDay);
          if (isBefore(e, s)) e = addDays(e, 1);
          let intervals = [{ start: s, end: e }];
          config.breaks.forEach(brk => {
            const bs = timeToDate(brk.start, baseDay);
            const be = timeToDate(brk.end, baseDay);
            const newIntervals: typeof intervals = [];
            intervals.forEach(inv => {
              if (isAfter(be, inv.start) && isBefore(bs, inv.end)) {
                if (isAfter(bs, inv.start)) newIntervals.push({ start: inv.start, end: bs });
                if (isBefore(be, inv.end)) newIntervals.push({ start: be, end: inv.end });
              } else {
                newIntervals.push(inv);
              }
            });
            intervals = newIntervals;
          });
          workingIntervals.push(...intervals);
        });
        // Sort descending for backward search
        workingIntervals.sort((a, b) => b.end.getTime() - a.end.getTime());
      }

      let moved = false;
      for (const inv of workingIntervals) {
        if (isAfter(currentTime, inv.start)) {
          const endInInv = isBefore(currentTime, inv.end) ? currentTime : inv.end.getTime();
          const available = endInInv - inv.start.getTime();
          const consume = Math.min(remaining, available);
          remaining -= consume;
          currentTime = endInInv - consume;
          if (remaining <= 0) return currentTime;
          moved = true;
        }
      }
      if (!moved || remaining > 0) {
        currentTime = this.getPreviousWorkingTime(currentTime, stageId, shiftConfigs);
      }
    }
    return currentTime;
  },

  calculateEndTime(startTime: number, durationMs: number, stageId: StageId, shiftConfigs: ShiftConfig[]): number {
    if (durationMs <= 0) return startTime;
    let remaining = durationMs;
    let currentTime = this.getNextWorkingTime(startTime, stageId, shiftConfigs);
    const config = shiftConfigs.find(c => c.stageId === stageId);
    if (!config) return startTime + durationMs;

    const timeToDate = (timeStr: string, baseDate: Date) => {
      const [h, m] = timeStr.split(':').map(Number);
      return setSeconds(setMinutes(setHours(baseDate, h), m), 0);
    };

    while (remaining > 0) {
      const baseDay = startOfDay(new Date(currentTime));
      const workingIntervals: { start: Date, end: Date }[] = [];
      
      if (!(baseDay.getDay() === 0 && !config.workOnSunday)) {
        config.shifts.forEach(shift => {
          const s = timeToDate(shift.start, baseDay);
          let e = timeToDate(shift.end, baseDay);
          if (isBefore(e, s)) e = addDays(e, 1);
          let intervals = [{ start: s, end: e }];
          config.breaks.forEach(brk => {
            const bs = timeToDate(brk.start, baseDay);
            const be = timeToDate(brk.end, baseDay);
            const newIntervals: typeof intervals = [];
            intervals.forEach(inv => {
              if (isAfter(be, inv.start) && isBefore(bs, inv.end)) {
                if (isAfter(bs, inv.start)) newIntervals.push({ start: inv.start, end: bs });
                if (isBefore(be, inv.end)) newIntervals.push({ start: be, end: inv.end });
              } else {
                newIntervals.push(inv);
              }
            });
            intervals = newIntervals;
          });
          workingIntervals.push(...intervals);
        });
        workingIntervals.sort((a, b) => a.start.getTime() - b.start.getTime());
      }

      let moved = false;
      for (const inv of workingIntervals) {
        if (isBefore(currentTime, inv.end)) {
          const startInInv = isAfter(currentTime, inv.start) ? currentTime : inv.start.getTime();
          const available = inv.end.getTime() - startInInv;
          const consume = Math.min(remaining, available);
          remaining -= consume;
          currentTime = startInInv + consume;
          if (remaining <= 0) return currentTime;
          moved = true;
        }
      }
      if (!moved || remaining > 0) {
        currentTime = this.getNextWorkingTime(currentTime, stageId, shiftConfigs);
      }
    }
    return currentTime;
  },

  forceCompleteMasterPO(masterPoId: string) {
    const pos = this.getProductionOrders();
    const masterPo = pos.find(p => p.id === masterPoId);
    if (!masterPo) return;
    
    // Complete master PO
    masterPo.status = 'COMPLETED';
    if (!masterPo.completedAt) masterPo.completedAt = Date.now();
    
    // Complete sub POs that are not already completed
    const subPos = pos.filter(p => p.masterPoId === masterPoId);
    subPos.forEach(sub => {
      sub.status = 'COMPLETED';
      if (!sub.completedAt) sub.completedAt = Date.now();
    });
    
    this.saveProductionOrders(pos);
  },

  createMasterPO(modelId: string, quantity: number, plannedStartTime?: number, customLeadTime?: number, idPrefix: string = "PO", targetStageId?: string, selectedLaserPartIds?: string[], customNote?: string, excelRequiredDate?: number) {
    const { masterPo, allChildPOs } = this.calculateMasterPOSchedule(modelId, quantity, plannedStartTime, customLeadTime, idPrefix, targetStageId, selectedLaserPartIds, customNote, excelRequiredDate);
    const pos = this.getProductionOrders();
    const updatedPOs = [masterPo, ...allChildPOs, ...pos];
    this.saveProductionOrders(updatedPOs);
    return masterPo;
  },

  createBatchMasterPOs(orders: {
    modelId: string;
    quantity: number;
    expectedCompletionTime?: number;
    customLeadTime?: number;
    idPrefix?: string;
    customNote?: string;
    excelRequiredDate?: number;
  }[]) {
    if (!orders || orders.length === 0) return [];
    const pos = this.getProductionOrders();
    const newMasterPOs: ProductionOrder[] = [];
    const newChildPOs: ProductionOrder[] = [];
    const existingIdSet = new Set(pos.map(p => p.id));
    const context = {
      shiftConfigs: this.getShiftConfigs(),
      norms: this.getNorms(),
      partsList: this.getParts(),
      modelBom: this.getModelBOM(),
      bomV2: this.getBOMV2(),
      laserNesting: this.getLaserNesting()
    };

    for (const item of orders) {
      const { masterPo, allChildPOs } = this.calculateMasterPOSchedule(
        item.modelId,
        item.quantity,
        item.expectedCompletionTime,
        item.customLeadTime,
        item.idPrefix || "PO",
        undefined,
        undefined,
        item.customNote,
        item.excelRequiredDate,
        existingIdSet,
        context
      );
      existingIdSet.add(masterPo.id);
      allChildPOs.forEach(cp => existingIdSet.add(cp.id));
      newMasterPOs.push(masterPo);
      newChildPOs.push(...allChildPOs);
    }

    const updatedPOs = [...newMasterPOs, ...newChildPOs, ...pos];
    this.saveProductionOrders(updatedPOs);
    return newMasterPOs;
  },

  previewMasterPOStart(modelId: string, quantity: number, targetCompletionTime?: number, customLeadTime?: number): number {
    const { masterPo } = this.calculateMasterPOSchedule(modelId, quantity, targetCompletionTime, customLeadTime, "PO");
    return masterPo.plannedStartTime || Date.now();
  },

  calculateMasterPOSchedule(
    modelId: string, 
    quantity: number, 
    targetCompletionTime?: number, 
    customLeadTime?: number, 
    idPrefix: string = "PO", 
    targetStageId?: string, 
    selectedLaserPartIds?: string[], 
    customNote?: string, 
    excelRequiredDate?: number,
    existingIdSet?: Set<string>,
    context?: {
      shiftConfigs?: ShiftConfig[];
      norms?: ProductivityNorm[];
      partsList?: Part[];
      modelBom?: ModelBOMDefinition[];
      bomV2?: BOMDefinitionV2[];
      laserNesting?: LaserNesting[];
    }
  ) {
    const isOddPO = Array.isArray(selectedLaserPartIds) && selectedLaserPartIds.length > 0;
    if (isOddPO && idPrefix === "PO") {
      idPrefix = "PO-LE";
    }
    const pos = existingIdSet ? [] : this.getProductionOrders(); // Still needed for unique ID check if no existingIdSet provided
    const timestamp = Date.now();
    const shiftConfigs = context?.shiftConfigs || this.getShiftConfigs();
    const baseEndTime = targetCompletionTime || timestamp;
    const dateStr = format(baseEndTime, 'ddMM');
    const modelPrefix = modelId.length > 8 ? modelId.substring(0, 8).toUpperCase() : modelId.toUpperCase();
    const generateUniqueId = (prefix: string, suffix: string = "") => {
      let newId = "";
      do {
        newId = `${prefix}-${Math.random().toString(36).substring(2, 7).toUpperCase()}${suffix}`;
      } while (existingIdSet ? existingIdSet.has(newId) : pos.some(p => p.id === newId));
      if (existingIdSet) existingIdSet.add(newId);
      return newId;
    };

    const masterPoId = generateUniqueId(`${idPrefix}-${modelPrefix}-${dateStr}`);
    const norms = context?.norms || this.getNorms();
    const partsList: Part[] = context?.partsList || this.getParts();
    const partsMap = new Map<string, Part>(partsList.map(p => [p.id.trim().toUpperCase(), p]));
    const findPart = (id: string): Part | undefined => {
      return partsMap.get(id.trim().toUpperCase());
    };
    const modelBom = context?.modelBom || this.getModelBOM();
    const bomV2 = context?.bomV2 || this.getBOMV2();

    const requiredParts = new Map<string, { qty: number; minLevel: number; }>();
    const level1Children = new Map<string, string[]>();
    const parentsMap = new Map<string, string[]>();

    const traverseBOM = (currentId: string, currentQty: number, level: number, parentId: string | null) => {
      if (level > 20) return; // infinite loop guard
      if (level > 0 || (idPrefix.startsWith("REPAIR") && level === 0)) {
        const existing = requiredParts.get(currentId);
        if (!existing) {
          requiredParts.set(currentId, { qty: currentQty, minLevel: level });
        } else {
          requiredParts.set(currentId, { qty: existing.qty + currentQty, minLevel: Math.min(existing.minLevel, level) });
        }
        if (level >= 2 && parentId) {
          const children = level1Children.get(parentId) || [];
          if (!children.includes(currentId)) children.push(currentId);
          level1Children.set(parentId, children);
        }
      }
      if (parentId) {
        const parents = parentsMap.get(currentId) || [];
        if (!parents.includes(parentId)) parents.push(parentId);
        parentsMap.set(currentId, parents);
      }
      const v1Idx = modelBom.filter(b => b.modelId === currentId);
      for (const ing of v1Idx) {
        traverseBOM(ing.partId, currentQty * ing.quantity, level + 1, (level === 0) ? null : (level === 1 ? currentId : parentId));
      }
      const v2Idx = bomV2.filter(b => b.resultPartId === currentId && (!b.applicableModel || b.applicableModel === modelId));
      for (const ing of v2Idx) {
        const nextParentId = (level === 1) ? currentId : (level >= 2 ? parentId : null);
        traverseBOM(ing.ingredientPartId, currentQty * ing.quantity, level + 1, nextParentId);
      }
    };
    traverseBOM(modelId, quantity, 0, null);

    const laserPOs: ProductionOrder[] = [];
    const bendingPOs: ProductionOrder[] = [];
    const weldingPOs: ProductionOrder[] = [];
    const paintingPOs: ProductionOrder[] = [];

    const STAGE_ORDER = ['LASER', 'BENDING', 'WELDING', 'PAINTING'];
    const targetStageIdx = targetStageId ? STAGE_ORDER.indexOf(targetStageId) : 999;

    const addPoToList = (partId: string, info: any, stageId: StageId, list: ProductionOrder[]) => {
      if (STAGE_ORDER.indexOf(stageId) >= targetStageIdx) return;
      const { qty, minLevel: level } = info;
      const part = findPart(partId);
      if (stageId === 'LASER' && part?.skipLaser) return;
      if (stageId === 'BENDING' && part?.skipBending) return;
      if (stageId === 'WELDING' && part?.skipWelding) return;
      if (!isOddPO && stageId === 'PAINTING' && (part?.skipPainting || (level > 1 && !part?.hasPaintingPO))) return;
      if (isOddPO && stageId === 'PAINTING' && part?.skipPainting) return;
      let idSuffix = "";
      if (stageId === 'WELDING') idSuffix = "- H";
      else if (stageId === 'BENDING') idSuffix = "- CD";
      list.push({
        id: generateUniqueId(`${idPrefix}-${modelPrefix}-${dateStr}-${stageId}`, idSuffix),
        masterPoId: masterPoId,
        partId: partId,
        stageId: stageId,
        targetQuantity: qty,
        producedQuantity: 0,
        exportedQuantity: 0,
        status: 'PENDING',
        createdAt: timestamp
      });
    };

    requiredParts.forEach((info, partId) => {
      if (isOddPO) {
        if (!selectedLaserPartIds!.includes(partId)) return;
        // PO Tồn Lẻ: Khi tick chọn linh kiện ở Laser, tự động sinh ra Laser, Chấn/dập, Hàn, Sơn
        addPoToList(partId, info, 'LASER', laserPOs);
        addPoToList(partId, info, 'BENDING', bendingPOs);
        addPoToList(partId, info, 'WELDING', weldingPOs);
        addPoToList(partId, info, 'PAINTING', paintingPOs);
        return;
      }

      const { minLevel: level } = info;
      const hasIngredients = bomV2.some(b => b.resultPartId === partId && (!b.applicableModel || b.applicableModel === modelId));
      if (!hasIngredients) {
        addPoToList(partId, info, 'LASER', laserPOs);
        addPoToList(partId, info, 'BENDING', bendingPOs);
      }
      if (hasIngredients || level <= 1) {
        addPoToList(partId, info, 'WELDING', weldingPOs);
      }
      const part = findPart(partId);
      if (level <= 1 || part?.hasPaintingPO) {
        addPoToList(partId, info, 'PAINTING', paintingPOs);
      }
    });

    const sortByLevelDesc = (a: ProductionOrder, b: ProductionOrder) => {
      return (requiredParts.get(b.partId)?.minLevel || 0) - (requiredParts.get(a.partId)?.minLevel || 0);
    };

    laserPOs.sort(sortByLevelDesc);
    bendingPOs.sort(sortByLevelDesc);
    weldingPOs.sort(sortByLevelDesc);
    paintingPOs.sort(sortByLevelDesc);

    const laserNesting = context?.laserNesting || this.getLaserNesting();

    const runForwardPass = (globalStart: number) => {
      const outChildPOs: ProductionOrder[] = [];
      const partStageFinishTime = new Map<string, Map<StageId, number>>(); 
      const recordFinish = (po: ProductionOrder, time: number) => {
        let stages = partStageFinishTime.get(po.partId);
        if (!stages) {
          stages = new Map<StageId, number>();
          partStageFinishTime.set(po.partId, stages);
        }
        stages.set(po.stageId as StageId, time);
      };
      const getFinishTime = (partId: string, stageId: StageId) => {
        return partStageFinishTime.get(partId)?.get(stageId) || globalStart;
      };

      // 1. LASER
      let fFreeLaser = globalStart;
      const laserConfig = shiftConfigs.find(c => c.stageId === 'LASER');
      
      if (laserNesting.length > 0) {
        const nestedGroupMap = new Map<string, ProductionOrder[]>();
        const individualLaserPOs: ProductionOrder[] = [];
        laserPOs.forEach(po => {
          const p = {...po};
          const nest = laserNesting.find(ln => ln.partId === po.partId && (!ln.applicableModel || ln.applicableModel === modelId));
          if (nest) {
            const group = nestedGroupMap.get(nest.nestingId) || [];
            group.push(p);
            nestedGroupMap.set(nest.nestingId, group);
          } else {
            individualLaserPOs.push(p);
          }
        });
        individualLaserPOs.forEach(p => {
          const start = this.getNextWorkingTime(fFreeLaser, 'LASER', shiftConfigs);
          p.plannedStartTime = start;
          const norm = norms.find(n => n.partId === p.partId && n.stageId === 'LASER');
          
          // Worker override for individual laser PO
          const partName = partsList.find(pl => pl.id === p.partId)?.name;
          const override = laserConfig?.workerOverrides?.find(o => o.modelId === p.partId || o.modelId === partName || o.modelId === modelId);
          const currentLaserWorkers = override ? override.workerCount : (laserConfig?.workerCount || 1);

          const duration = norm ? (p.targetQuantity * norm.secondsPerUnit * 1000) / currentLaserWorkers : 0;
          const end = this.calculateEndTime(start, duration, 'LASER', shiftConfigs);
          p.expectedCompletionTime = end;
          fFreeLaser = end;
          recordFinish(p, end);
          outChildPOs.push(p);
        });
        nestedGroupMap.forEach((groupPOs, nestingId) => {
          let maxPlates = 0;
          let secondsPerSheet = 0;
          
          // Find the best worker count for the nesting group
          let groupMaxWorkers = laserConfig?.workerCount || 1;
          const batchModelOverride = laserConfig?.workerOverrides?.find(o => o.modelId === modelId);
          if (batchModelOverride) groupMaxWorkers = batchModelOverride.workerCount;

          // Calculate max required plates for this nesting group
          groupPOs.forEach(p => {
            const nest = laserNesting.find(ln => ln.partId === p.partId && ln.nestingId === nestingId && (!ln.applicableModel || ln.applicableModel === modelId));
            if (nest && nest.qtyPerSheet > 0) {
              const plates = Math.ceil(p.targetQuantity / nest.qtyPerSheet);
              if (plates > maxPlates) maxPlates = plates;
              if (nest.secondsPerSheet) secondsPerSheet = nest.secondsPerSheet;
            }

            // Check if this specific part in the group has an override
            const partName = partsList.find(pl => pl.id === p.partId)?.name;
            const pOverride = laserConfig?.workerOverrides?.find(o => o.modelId === p.partId || o.modelId === partName);
            if (pOverride && pOverride.workerCount > groupMaxWorkers) {
              groupMaxWorkers = pOverride.workerCount;
            }
          });
          
          const totalDur = maxPlates * secondsPerSheet * 1000;
          const adjustedDur = totalDur / groupMaxWorkers;
          const start = this.getNextWorkingTime(fFreeLaser, 'LASER', shiftConfigs);
          const end = this.calculateEndTime(start, adjustedDur, 'LASER', shiftConfigs);
          groupPOs.forEach(p => {
            p.plannedStartTime = start;
            p.expectedCompletionTime = end;
            recordFinish(p, end);
            outChildPOs.push(p);
          });
          fFreeLaser = end;
        });
      } else {
        laserPOs.forEach(po => {
          const p = {...po};
          const start = this.getNextWorkingTime(fFreeLaser, 'LASER', shiftConfigs);
          p.plannedStartTime = start;
          const norm = norms.find(n => n.partId === p.partId && n.stageId === 'LASER');

          const partName = partsList.find(pl => pl.id === p.partId)?.name;
          const override = laserConfig?.workerOverrides?.find(o => o.modelId === p.partId || o.modelId === partName || o.modelId === modelId);
          const currentLaserWorkers = override ? override.workerCount : (laserConfig?.workerCount || 1);

          const duration = norm ? (p.targetQuantity * norm.secondsPerUnit * 1000) / currentLaserWorkers : 0;
          const end = this.calculateEndTime(start, duration, 'LASER', shiftConfigs);
          p.expectedCompletionTime = end;
          fFreeLaser = end;
          recordFinish(p, end);
          outChildPOs.push(p);
        });
      }

      // 2. BENDING
      let fFreeBending = globalStart;
      const bendConfig = shiftConfigs.find(c => c.stageId === 'BENDING');
      bendingPOs.forEach(po => {
        const p = {...po};
        const partName = partsList.find(pl => pl.id === p.partId)?.name;
        const override = bendConfig?.workerOverrides?.find(o => o.modelId === p.partId || o.modelId === partName || o.modelId === modelId);
        const bendWorkers = override ? override.workerCount : (bendConfig?.workerCount || 1);

        const laserEnd = getFinishTime(p.partId, 'LASER');
        const start = this.getNextWorkingTime(Math.max(fFreeBending, laserEnd), 'BENDING', shiftConfigs);
        p.plannedStartTime = start;
        const norm = norms.find(n => n.partId === p.partId && n.stageId === 'BENDING');
        const duration = norm ? (p.targetQuantity * norm.secondsPerUnit * 1000) / bendWorkers : 0;
        const end = this.calculateEndTime(start, duration, 'BENDING', shiftConfigs);
        p.expectedCompletionTime = end;
        fFreeBending = end;
        recordFinish(p, end);
        outChildPOs.push(p);
      });

      // 3. WELDING
      let fFreeWelding = globalStart;
      const weldConfig = shiftConfigs.find(c => c.stageId === 'WELDING');
      weldingPOs.forEach(po => {
        const p = {...po};
        const partName = partsList.find(pl => pl.id === p.partId)?.name;
        const override = weldConfig?.workerOverrides?.find(o => o.modelId === p.partId || o.modelId === partName || o.modelId === modelId);
        const weldWorkers = override ? override.workerCount : (weldConfig?.workerCount || 1);

        const children = level1Children.get(p.partId) || [];
        const componentsReadyTime = children.length === 0 ? globalStart : Math.max(...children.map(cid => {
          const wEnd = partStageFinishTime.get(cid)?.get('WELDING');
          if (wEnd !== undefined) return wEnd;
          const bEnd = partStageFinishTime.get(cid)?.get('BENDING');
          if (bEnd !== undefined) return bEnd;
          return getFinishTime(cid, 'LASER');
        }));
        
        let myBendingEnd = partStageFinishTime.get(p.partId)?.get('BENDING');
        let myLaserEnd = partStageFinishTime.get(p.partId)?.get('LASER');
        let myReadyTime = myBendingEnd !== undefined ? myBendingEnd : (myLaserEnd !== undefined ? myLaserEnd : globalStart);
        
        const start = this.getNextWorkingTime(Math.max(fFreeWelding, componentsReadyTime, myReadyTime), 'WELDING', shiftConfigs);
        p.plannedStartTime = start;
        const norm = norms.find(n => n.partId === p.partId && n.stageId === 'WELDING');
        const duration = norm ? (p.targetQuantity * norm.secondsPerUnit * 1000) / weldWorkers : 0;
        const end = this.calculateEndTime(start, duration, 'WELDING', shiftConfigs);
        p.expectedCompletionTime = end;
        fFreeWelding = end;
        recordFinish(p, end);
        outChildPOs.push(p);
      });

      // 4. PAINTING
      let fFreePainting = globalStart;
      const paintConfig = shiftConfigs.find(c => c.stageId === 'PAINTING');
      paintingPOs.forEach(po => {
        const p = {...po};
        const partName = partsList.find(pl => pl.id === p.partId)?.name;
        const override = paintConfig?.workerOverrides?.find(o => o.modelId === p.partId || o.modelId === partName || o.modelId === modelId);
        const paintWorkers = override ? override.workerCount : (paintConfig?.workerCount || 1);

        let weldEnd = partStageFinishTime.get(p.partId)?.get('WELDING');
        let bendEnd = partStageFinishTime.get(p.partId)?.get('BENDING');
        let laserEnd = partStageFinishTime.get(p.partId)?.get('LASER');
        let readyTime = weldEnd !== undefined ? weldEnd : (bendEnd !== undefined ? bendEnd : (laserEnd !== undefined ? laserEnd : globalStart));

        const start = this.getNextWorkingTime(Math.max(fFreePainting, readyTime), 'PAINTING', shiftConfigs);
        p.plannedStartTime = start;
        const norm = norms.find(n => n.partId === p.partId && n.stageId === 'PAINTING');
        const duration = norm ? (p.targetQuantity * norm.secondsPerUnit * 1000) / paintWorkers : 0;
        const end = this.calculateEndTime(start, duration, 'PAINTING', shiftConfigs);
        p.expectedCompletionTime = end;
        fFreePainting = end;
        recordFinish(p, end);
        outChildPOs.push(p);
      });

      const maxEnd = outChildPOs.length > 0 
        ? Math.max(...outChildPOs.filter(p => p.expectedCompletionTime).map(p => p.expectedCompletionTime!))
        : globalStart;
        
      const actualMinStart = outChildPOs.length > 0
        ? Math.min(...outChildPOs.filter(p => p.plannedStartTime).map(p => p.plannedStartTime!))
        : globalStart;

      return { outChildPOs, maxEnd, minStart: actualMinStart };
    };

    // Backward scheduling logic: find the latest globalStart that meets the target date
    // while respecting the machine queuing (maxExisting...End)
    let low = targetCompletionTime ? targetCompletionTime - 60 * 24 * 60 * 60 * 1000 : timestamp;
    let high = targetCompletionTime ? targetCompletionTime : timestamp;
    let bestChildPOs: ProductionOrder[] = [];
    let bestStart = low;
    let bestEnd = low;

    if (targetCompletionTime) {
        // Binary search for the latest start time that finishes by the deadline
        for (let i = 0; i < 20; i++) {
            if (high - low < 60000) break;
            const mid = low + Math.floor((high - low) / 2);
            const { outChildPOs, maxEnd, minStart } = runForwardPass(mid);
            if (maxEnd <= targetCompletionTime) {
                bestChildPOs = outChildPOs;
                bestStart = minStart;
                bestEnd = maxEnd;
                low = mid + 1; // Try starting later
            } else {
                high = mid - 1; // Too late, must start earlier
            }
        }
    }

    // Fallback or if no target date: Use Forward Scheduling starting from now (queuing)
    if (bestChildPOs.length === 0) {
        const { outChildPOs, maxEnd, minStart } = runForwardPass(timestamp);
        bestChildPOs = outChildPOs;
        bestStart = minStart;
        bestEnd = maxEnd;
    }

    const masterPo: ProductionOrder = {
      id: masterPoId,
      partId: modelId,
      targetQuantity: quantity,
      producedQuantity: 0,
      exportedQuantity: 0,
      status: 'PENDING',
      createdAt: timestamp,
      plannedStartTime: bestStart,
      leadTime: customLeadTime,
      expectedCompletionTime: bestEnd,
      isOddPO: isOddPO || undefined,
      note: customNote || (isOddPO ? `PO Dư Lẻ (${selectedLaserPartIds!.length} LK Cắt Laser)` : undefined),
      excelRequiredDate: excelRequiredDate
    };
    return { masterPo, allChildPOs: bestChildPOs };
  },

  getModelLaserComponents(modelId: string, quantity: number = 1) {
    const partsList = this.getParts();
    const findPart = (id: string) => {
      const cleanId = id.trim().toUpperCase();
      return partsList.find(p => p.id.trim().toUpperCase() === cleanId);
    };
    const modelBom = this.getModelBOM();
    const bomV2 = this.getBOMV2();

    const requiredParts = new Map<string, { qty: number; minLevel: number; }>();

    const traverseBOM = (currentId: string, currentQty: number, level: number) => {
      if (level > 20) return;
      if (level > 0) {
        const existing = requiredParts.get(currentId);
        if (!existing) {
          requiredParts.set(currentId, { qty: currentQty, minLevel: level });
        } else {
          requiredParts.set(currentId, { qty: existing.qty + currentQty, minLevel: Math.min(existing.minLevel, level) });
        }
      }
      const v1Idx = modelBom.filter(b => b.modelId === currentId);
      for (const ing of v1Idx) {
        traverseBOM(ing.partId, currentQty * ing.quantity, level + 1);
      }
      const v2Idx = bomV2.filter(b => b.resultPartId === currentId && (!b.applicableModel || b.applicableModel === modelId));
      for (const ing of v2Idx) {
        traverseBOM(ing.ingredientPartId, currentQty * ing.quantity, level + 1);
      }
    };
    traverseBOM(modelId, 1, 0);

    const laserComponents: {
      partId: string;
      partName: string;
      ratio: number;
      totalQty: number;
      unit: string;
    }[] = [];

    requiredParts.forEach((info, partId) => {
      const hasIngredients = bomV2.some(b => b.resultPartId === partId && (!b.applicableModel || b.applicableModel === modelId));
      if (!hasIngredients) {
        const part = findPart(partId);
        if (!part?.skipLaser) {
          laserComponents.push({
            partId,
            partName: part?.name || partId,
            ratio: info.qty,
            totalQty: Math.round(info.qty * quantity * 100) / 100,
            unit: part?.unit || 'cái'
          });
        }
      }
    });

    return laserComponents;
  },

  createOddPO(modelId: string, quantity: number, selectedLaserPartIds: string[], targetCompletionTime?: number) {
    const { masterPo, allChildPOs } = this.calculateMasterPOSchedule(
      modelId,
      quantity,
      targetCompletionTime,
      undefined,
      "PO-LE",
      undefined,
      selectedLaserPartIds
    );
    const pos = this.getProductionOrders();
    const updatedPOs = [masterPo, ...allChildPOs, ...pos];
    this.saveProductionOrders(updatedPOs);
    return masterPo;
  },

  createSingleOddPO(partId: string, stageId: StageId, quantity: number): ProductionOrder {
    const pos = this.getProductionOrders();
    const dateStr = format(Date.now(), 'ddMM');
    const cleanPart = partId.length > 8 ? partId.substring(0, 8).toUpperCase() : partId.toUpperCase();
    const generateUniqueId = () => {
      let newId = "";
      do {
        newId = `PO-LE-${stageId}-${cleanPart}-${dateStr}-${Math.random().toString(36).substring(2, 6).toUpperCase()}`;
      } while (pos.some(p => p.id === newId));
      return newId;
    };
    const stageName = STAGES.find(s => s.id === stageId)?.name || stageId;
    const newPo: ProductionOrder = {
      id: generateUniqueId(),
      partId: partId,
      stageId: stageId,
      targetQuantity: quantity,
      producedQuantity: 0,
      exportedQuantity: 0,
      status: 'PENDING',
      createdAt: Date.now(),
      isOddPO: true,
      note: `PO Dư Lẻ (${stageName})`
    };
    this.saveProductionOrders([newPo, ...pos]);
    return newPo;
  },

  deletePO(id: string) {
    const pos = this.getProductionOrders();
    // If it's a master PO, delete all its children too
    const filtered = pos.filter(p => p.id !== id && p.masterPoId !== id);
    this.saveProductionOrders(filtered);
    deletePOFromSupabase(id);
  },

  clearStageInventory(stageId: StageId, location: 'IN' | 'OUT') {
    const inventory = this.getInventory();
    const newInventory = inventory.filter(item => !(item.stageId === stageId && item.location === location));
    this.saveInventory(newInventory);
    clearStageInventoryFromSupabase(stageId, location);
  },

  updateSubPoQty(id: string, qty: number) {
    const pos = this.getProductionOrders();
    const targetPo = pos.find(p => p.id === id);
    if (!targetPo) return;

    // Only update the specific target PO, not all related POs
    targetPo.targetQuantity = qty;
    
    // Recalculate status
    const isProduced = targetPo.producedQuantity >= targetPo.targetQuantity;
    const isExported = (targetPo.exportedQuantity || 0) >= targetPo.targetQuantity;
    
    // If it was already completed but now target is higher, put back to IN_PROGRESS. 
    // If it wasn't completed but now meets target, complete it.
    if (isProduced && isExported) {
      targetPo.status = 'COMPLETED';
      if (!targetPo.completedAt) targetPo.completedAt = Date.now();
    } else {
      targetPo.status = (targetPo.producedQuantity > 0 || (targetPo.exportedQuantity && targetPo.exportedQuantity > 0)) ? 'IN_PROGRESS' : 'PENDING';
      targetPo.completedAt = undefined;
    }

    this.saveProductionOrders(pos);
  },

  updateSubPoTime(id: string, newExpectedTime: number) {
    const pos = this.getProductionOrders();
    const targetPo = pos.find(p => p.id === id);
    if (!targetPo || !targetPo.expectedCompletionTime) return;

    const timeDiff = newExpectedTime - targetPo.expectedCompletionTime;
    targetPo.expectedCompletionTime = newExpectedTime;
    if (targetPo.plannedStartTime) {
      targetPo.plannedStartTime += timeDiff;
    }

    this.saveProductionOrders(pos);
  },

  getStorageStats() {
    const parts = this.getParts();
    const pos = this.getProductionOrders();
    const inventory = this.getInventory();
    const transactions = this.getTransactions();
    const labels = this.getLabels();
    const bom = this.getBOM();
    const bomV2 = this.getBOMV2();
    const modelBom = this.getModelBOM();
    const norms = this.getNorms();
    const laserNesting = this.getLaserNesting();
    const shiftConfigs = this.getShiftConfigs();
    const transformations = this.getTransformations();
    const glazingPlans = this.getGlazingPlans();

    let totalBytes = 0;
    let totalKeys = 0;
    if (typeof window !== 'undefined' && window.localStorage) {
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && (key.startsWith('wip_') || Object.values(STORAGE_KEYS).includes(key) || key === 'wip_labels')) {
          totalKeys++;
          const val = localStorage.getItem(key) || '';
          totalBytes += (key.length + val.length) * 2;
        }
      }
    }

    return {
      partsCount: parts.length,
      posCount: pos.length,
      inventoryCount: inventory.length,
      transactionsCount: transactions.length,
      labelsCount: labels.length,
      bomCount: bom.length,
      bomV2Count: bomV2.length,
      modelBomCount: modelBom.length,
      normsCount: norms.length,
      laserNestingCount: laserNesting.length,
      shiftConfigsCount: shiftConfigs.length,
      transformationsCount: transformations.length,
      glazingPlansCount: glazingPlans.length,
      totalKeys,
      totalBytes,
      sizeFormatted: totalBytes > 1024 * 1024 
        ? `${(totalBytes / (1024 * 1024)).toFixed(2)} MB` 
        : `${(totalBytes / 1024).toFixed(1)} KB`
    };
  },

  exportBackupData() {
    const stats = this.getStorageStats();
    const storageDump: Record<string, any> = {};

    if (typeof window !== 'undefined' && window.localStorage) {
      // 1. Gather all known STORAGE_KEYS
      Object.values(STORAGE_KEYS).forEach(k => {
        const val = localStorage.getItem(k);
        if (val !== null) {
          try {
            storageDump[k] = JSON.parse(val);
          } catch {
            storageDump[k] = val;
          }
        }
      });

      // 2. Gather wip_labels and any other dynamic wip_ keys
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (k && (k.startsWith('wip_') || k === 'wip_labels')) {
          if (!(k in storageDump)) {
            const val = localStorage.getItem(k);
            if (val !== null) {
              try {
                storageDump[k] = JSON.parse(val);
              } catch {
                storageDump[k] = val;
              }
            }
          }
        }
      }
    }

    return {
      app: 'WIP_TRACKING_SYSTEM',
      version: '1.0',
      exportedAt: new Date().toISOString(),
      exportDateFormatted: format(new Date(), 'dd/MM/yyyy HH:mm:ss'),
      summary: stats,
      storage: storageDump
    };
  },

  downloadBackupFile(customFilename?: string): string {
    const backup = this.exportBackupData();
    const dateStr = format(new Date(), 'yyyyMMdd_HHmmss');
    const filename = customFilename || `WIP_LocalStorage_Backup_${dateStr}.json`;
    const jsonStr = JSON.stringify(backup, null, 2);
    const blob = new Blob([jsonStr], { type: 'application/json;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', filename);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    return filename;
  },

  validateBackupJson(rawStringOrObject: any): { 
    valid: boolean; 
    error?: string; 
    storageData?: Record<string, any>; 
    metadata?: any;
    stats?: any;
  } {
    let parsed: any;
    if (typeof rawStringOrObject === 'string') {
      try {
        parsed = JSON.parse(rawStringOrObject);
      } catch (err: any) {
        return { valid: false, error: 'Tệp không phải định dạng JSON hợp lệ: ' + (err?.message || '') };
      }
    } else {
      parsed = rawStringOrObject;
    }

    if (!parsed || typeof parsed !== 'object') {
      return { valid: false, error: 'Dữ liệu JSON rỗng hoặc không đúng cấu trúc đối tượng' };
    }

    let storageData: Record<string, any> = {};
    let metadata: any = null;

    if (parsed.storage && typeof parsed.storage === 'object') {
      storageData = parsed.storage;
      metadata = {
        app: parsed.app,
        version: parsed.version,
        exportedAt: parsed.exportedAt,
        exportDateFormatted: parsed.exportDateFormatted,
        summary: parsed.summary
      };
    } else if (parsed.data && typeof parsed.data === 'object') {
      storageData = parsed.data;
      metadata = { exportedAt: parsed.exportedAt };
    } else {
      storageData = parsed;
    }

    const keys = Object.keys(storageData);
    const hasWipKeys = keys.some(k => k.startsWith('wip_') || Object.values(STORAGE_KEYS).includes(k) || k === 'wip_labels');

    if (!hasWipKeys) {
      return { 
        valid: false, 
        error: 'Không tìm thấy dữ liệu hợp lệ của hệ thống WIP (không có khóa dữ liệu wip_* nào trong file).' 
      };
    }

    const partsCount = Array.isArray(storageData[STORAGE_KEYS.PARTS]) ? storageData[STORAGE_KEYS.PARTS].length : 0;
    const posCount = Array.isArray(storageData[STORAGE_KEYS.PRODUCTION_ORDERS]) ? storageData[STORAGE_KEYS.PRODUCTION_ORDERS].length : 0;
    const inventoryCount = Array.isArray(storageData[STORAGE_KEYS.INVENTORY]) ? storageData[STORAGE_KEYS.INVENTORY].length : 0;
    const transactionsCount = Array.isArray(storageData[STORAGE_KEYS.TRANSACTIONS]) ? storageData[STORAGE_KEYS.TRANSACTIONS].length : 0;
    const labelsCount = Array.isArray(storageData['wip_labels']) ? storageData['wip_labels'].length : 0;

    return {
      valid: true,
      storageData,
      metadata,
      stats: {
        totalKeys: keys.length,
        partsCount,
        posCount,
        inventoryCount,
        transactionsCount,
        labelsCount
      }
    };
  },

  importBackupData(storageData: Record<string, any>, mode: 'overwrite' | 'merge' = 'overwrite'): {
    success: boolean;
    keysRestored: number;
    message: string;
  } {
    if (typeof window === 'undefined' || !window.localStorage) {
      return { success: false, keysRestored: 0, message: 'Trình duyệt không hỗ trợ LocalStorage.' };
    }

    try {
      const priorityOrder = [
        STORAGE_KEYS.PARTS,
        STORAGE_KEYS.INVENTORY,
        STORAGE_KEYS.PRODUCTION_ORDERS,
        STORAGE_KEYS.BOM,
        STORAGE_KEYS.BOM_V2,
        STORAGE_KEYS.MODEL_BOM,
        STORAGE_KEYS.NORMS,
        STORAGE_KEYS.SHIFT_CONFIGS,
        STORAGE_KEYS.LABEL_SETTINGS,
        STORAGE_KEYS.TRANSFORMATIONS,
        STORAGE_KEYS.LASER_NESTING,
        STORAGE_KEYS.GLAZING_PLANS,
        STORAGE_KEYS.GLAZING_PLAN_NORMS,
        STORAGE_KEYS.TRANSACTIONS,
        'wip_labels'
      ];

      const sortKeyPriority = (a: string, b: string) => {
        const idxA = priorityOrder.indexOf(a as any);
        const idxB = priorityOrder.indexOf(b as any);
        const pA = idxA === -1 ? 99 : idxA;
        const pB = idxB === -1 ? 99 : idxB;
        return pA - pB;
      };

      if (mode === 'overwrite') {
        const keysToRemove: string[] = [];
        for (let i = 0; i < localStorage.length; i++) {
          const k = localStorage.key(i);
          if (k && (k.startsWith('wip_') || Object.values(STORAGE_KEYS).includes(k) || k === 'wip_labels')) {
            keysToRemove.push(k);
          }
        }
        keysToRemove.forEach(k => localStorage.removeItem(k));

        let restoredCount = 0;
        const sortedEntries = Object.entries(storageData).sort(([a], [b]) => sortKeyPriority(a, b));

        for (const [key, value] of sortedEntries) {
          if (key.startsWith('wip_') || Object.values(STORAGE_KEYS).includes(key) || key === 'wip_labels') {
            try {
              const strVal = typeof value === 'string' ? value : JSON.stringify(value);
              localStorage.setItem(key, strVal);
              restoredCount++;
            } catch (err: any) {
              // If quota exceeded on large log/history arrays, trim to recent items
              if (Array.isArray(value) && value.length > 500) {
                try {
                  const trimmed = value.slice(-2000);
                  localStorage.setItem(key, JSON.stringify(trimmed));
                  restoredCount++;
                } catch {
                  console.warn(`Could not save key ${key} due to storage limits.`);
                }
              } else {
                console.warn(`Could not save key ${key}:`, err);
              }
            }
          }
        }

        clearCache();
        return {
          success: true,
          keysRestored: restoredCount,
          message: `Đã khôi phục thành công ${restoredCount} danh mục dữ liệu (Ghi đè hoàn toàn).`
        };
      } else {
        // Merge mode
        let restoredCount = 0;
        const sortedEntries = Object.entries(storageData).sort(([a], [b]) => sortKeyPriority(a, b));

        sortedEntries.forEach(([key, incomingVal]) => {
          if (!key.startsWith('wip_') && !Object.values(STORAGE_KEYS).includes(key) && key !== 'wip_labels') {
            return;
          }

          const existingRaw = localStorage.getItem(key);
          if (!existingRaw) {
            try {
              const strVal = typeof incomingVal === 'string' ? incomingVal : JSON.stringify(incomingVal);
              localStorage.setItem(key, strVal);
              restoredCount++;
            } catch {}
            return;
          }

          try {
            const existingVal = JSON.parse(existingRaw);
            let incomingParsed = incomingVal;
            if (typeof incomingVal === 'string') {
              try { incomingParsed = JSON.parse(incomingVal); } catch {}
            }

            if (Array.isArray(existingVal) && Array.isArray(incomingParsed)) {
              if (key === STORAGE_KEYS.PARTS) {
                const map = new Map<string, any>();
                existingVal.forEach(p => p && p.id && map.set(p.id, p));
                incomingParsed.forEach(p => p && p.id && map.set(p.id, p));
                localStorage.setItem(key, JSON.stringify(Array.from(map.values())));
              } else if (key === STORAGE_KEYS.PRODUCTION_ORDERS) {
                const map = new Map<string, any>();
                existingVal.forEach(p => p && p.id && map.set(p.id, p));
                incomingParsed.forEach(p => p && p.id && map.set(p.id, p));
                localStorage.setItem(key, JSON.stringify(Array.from(map.values())));
              } else if (key === STORAGE_KEYS.INVENTORY) {
                const map = new Map<string, any>();
                existingVal.forEach(item => {
                  const k = `${item.partId}_${item.stageId}_${item.location || 'IN'}`;
                  map.set(k, item);
                });
                incomingParsed.forEach(item => {
                  const k = `${item.partId}_${item.stageId}_${item.location || 'IN'}`;
                  map.set(k, item);
                });
                localStorage.setItem(key, JSON.stringify(Array.from(map.values())));
              } else if (key === STORAGE_KEYS.TRANSACTIONS || key === 'wip_labels') {
                const map = new Map<string, any>();
                existingVal.forEach(item => item && item.id && map.set(item.id, item));
                incomingParsed.forEach(item => item && item.id && map.set(item.id, item));
                localStorage.setItem(key, JSON.stringify(Array.from(map.values())));
              } else if (key === STORAGE_KEYS.BOM) {
                const map = new Map<string, any>();
                existingVal.forEach(item => item && item.childPartId && map.set(item.childPartId, item));
                incomingParsed.forEach(item => item && item.childPartId && map.set(item.childPartId, item));
                localStorage.setItem(key, JSON.stringify(Array.from(map.values())));
              } else if (key === STORAGE_KEYS.BOM_V2) {
                const map = new Map<string, any>();
                existingVal.forEach(item => item && map.set(`${item.parentPartId}_${item.childPartId}`, item));
                incomingParsed.forEach(item => item && map.set(`${item.parentPartId}_${item.childPartId}`, item));
                localStorage.setItem(key, JSON.stringify(Array.from(map.values())));
              } else if (key === STORAGE_KEYS.MODEL_BOM) {
                const map = new Map<string, any>();
                existingVal.forEach(item => item && map.set(`${item.modelId}_${item.childPartId}`, item));
                incomingParsed.forEach(item => item && map.set(`${item.modelId}_${item.childPartId}`, item));
                localStorage.setItem(key, JSON.stringify(Array.from(map.values())));
              } else if (key === STORAGE_KEYS.NORMS) {
                const map = new Map<string, any>();
                existingVal.forEach(item => item && map.set(`${item.stageId}_${item.partId}`, item));
                incomingParsed.forEach(item => item && map.set(`${item.stageId}_${item.partId}`, item));
                localStorage.setItem(key, JSON.stringify(Array.from(map.values())));
              } else {
                const combined = [...existingVal, ...incomingParsed];
                localStorage.setItem(key, JSON.stringify(combined));
              }
              restoredCount++;
            } else if (typeof existingVal === 'object' && typeof incomingParsed === 'object' && existingVal && incomingParsed) {
              const merged = { ...existingVal, ...incomingParsed };
              localStorage.setItem(key, JSON.stringify(merged));
              restoredCount++;
            } else {
              const strVal = typeof incomingVal === 'string' ? incomingVal : JSON.stringify(incomingVal);
              localStorage.setItem(key, strVal);
              restoredCount++;
            }
          } catch {
            try {
              const strVal = typeof incomingVal === 'string' ? incomingVal : JSON.stringify(incomingVal);
              localStorage.setItem(key, strVal);
              restoredCount++;
            } catch {}
          }
        });

        clearCache();
        return {
          success: true,
          keysRestored: restoredCount,
          message: `Đã hợp nhất thành công ${restoredCount} danh mục dữ liệu.`
        };
      }
    } catch (err: any) {
      return {
        success: false,
        keysRestored: 0,
        message: 'Lỗi khi nạp dữ liệu vào LocalStorage: ' + (err?.message || '')
      };
    }
  },

  async autoCleanupOldData(): Promise<{ labelsRemoved: number; posRemoved: number; txsRemoved: number }> {
    const now = Date.now();
    const labelCutoff = now - 24 * 60 * 60 * 1000; // 24 hours
    const thirtyFiveDaysCutoff = now - 35 * 24 * 60 * 60 * 1000; // 35 days

    let labelsRemoved = 0;
    let posRemoved = 0;
    let txsRemoved = 0;

    // 1. Cleanup labels (keep only within 24h)
    try {
      const rawLabels = localStorage.getItem('wip_labels');
      if (rawLabels) {
        const parsedLabels = JSON.parse(rawLabels);
        if (Array.isArray(parsedLabels)) {
          const validLabels = parsedLabels.filter((l: any) => l && l.timestamp && l.timestamp >= labelCutoff);
          labelsRemoved = parsedLabels.length - validLabels.length;
          if (labelsRemoved > 0) {
            localStorage.setItem('wip_labels', JSON.stringify(validLabels));
            cache['wip_labels'] = validLabels;
          }
        }
      }
    } catch {}

    // 2. Cleanup transactions / scan logs (keep only within 35 days)
    try {
      const rawTxs = localStorage.getItem(STORAGE_KEYS.TRANSACTIONS);
      if (rawTxs) {
        const parsedTxs = JSON.parse(rawTxs);
        if (Array.isArray(parsedTxs)) {
          const validTxs = parsedTxs.filter((t: any) => t && t.timestamp && t.timestamp >= thirtyFiveDaysCutoff);
          txsRemoved = parsedTxs.length - validTxs.length;
          if (txsRemoved > 0) {
            localStorage.setItem(STORAGE_KEYS.TRANSACTIONS, JSON.stringify(validTxs));
            cache[STORAGE_KEYS.TRANSACTIONS] = validTxs;
          }
        }
      }
    } catch {}

    // 3. Cleanup Production Orders (delete completed POs older than 35 days, preserve incomplete POs)
    try {
      const rawPos = localStorage.getItem(STORAGE_KEYS.PRODUCTION_ORDERS);
      if (rawPos) {
        const parsedPos: ProductionOrder[] = JSON.parse(rawPos);
        if (Array.isArray(parsedPos)) {
          const validPos = this.filterExpiredCompletedPOs(parsedPos);
          posRemoved = parsedPos.length - validPos.length;
          if (posRemoved > 0) {
            localStorage.setItem(STORAGE_KEYS.PRODUCTION_ORDERS, JSON.stringify(validPos));
            cache[STORAGE_KEYS.PRODUCTION_ORDERS] = validPos;
          }
        }
      }
    } catch {}

    // 4. Directly cleanup old data on Supabase database (Free tier storage & egress optimization)
    if (isSupabaseConfigured) {
      try {
        // Try calling stored procedure if user ran schema
        const { error: rpcError } = await supabase.rpc('cleanup_old_wip_data');
        if (rpcError) {
          // Fallback to direct DELETE operations with DB-side filtering
          await supabase.from('wip_transactions').delete().lt('timestamp', thirtyFiveDaysCutoff);
          await supabase.from('qr_labels').delete().lt('timestamp', labelCutoff);
          await supabase.from('production_orders')
            .delete()
            .eq('status', 'COMPLETED')
            .lt('completed_at', thirtyFiveDaysCutoff);
        }
      } catch (err: any) {
        console.warn('[Supabase] DB cleanup error:', err?.message || err);
      }
    }

    return { labelsRemoved, posRemoved, txsRemoved };
  },

  getSyncStatus() {
    return {
      connected: isSupabaseConfigured && !syncError,
      syncing: isSyncing,
      lastSynced: lastSyncTime,
      error: syncError
    };
  },

  onSyncChange(listener: (status: { connected: boolean; syncing: boolean; lastSynced: number | null; error: string | null }) => void): () => void {
    syncListeners.add(listener);
    // Call immediately with current status
    listener(this.getSyncStatus());
    return () => {
      syncListeners.delete(listener);
    };
  },

  async flushAllPending(): Promise<void> {
    await Promise.allSettled([
      flushPendingConfigSaves(),
      flushPendingPoSaves(),
      flushPendingInvSaves(),
      flushPendingTransactions(),
      flushPendingGpSaves(),
      flushPendingLabels()
    ]);
  },

  /**
   * Initializes data from Supabase online into memory cache
   * Uses selective columns to optimize egress on Free tier
   */
  async init(): Promise<boolean> {
    if (!isSupabaseConfigured) return false;
    isSyncing = true;
    syncError = null;
    notifySyncChange();

    try {
      // 1. Fetch system_configs (only key, value columns)
      const { data: configRows, error: configErr } = await supabase
        .from('system_configs')
        .select('key, value');

      if (configErr) {
        if (configErr.code === 'PGRST205' || configErr.message?.includes('not found') || configErr.message?.includes('does not exist')) {
          syncError = 'Chưa tìm thấy bảng dữ liệu trên Supabase. Vui lòng chạy mã lệnh trong file supabase_schema.sql trên Supabase SQL Editor.';
          isSyncing = false;
          notifySyncChange();
          return false;
        }
        throw configErr;
      }

      if (configRows && configRows.length > 0) {
        configRows.forEach((row: any) => {
          if (row && row.key) {
            cache[row.key] = row.value;
            try {
              localStorage.setItem(row.key, typeof row.value === 'string' ? row.value : JSON.stringify(row.value));
            } catch {}
          }
        });
      }

      // 2. Fetch active production_orders (selective columns only)
      const { data: poRows, error: poErr } = await supabase
        .from('production_orders')
        .select('id, master_po_id, part_id, stage_id, target_quantity, produced_quantity, exported_quantity, status, created_at, completed_at, planned_start_time, lead_time, expected_completion_time, note, is_odd_po, excel_required_date');

      if (!poErr && poRows && poRows.length > 0) {
        const pos: ProductionOrder[] = poRows.map((r: any) => ({
          id: r.id,
          masterPoId: r.master_po_id || undefined,
          partId: r.part_id,
          stageId: r.stage_id || undefined,
          targetQuantity: Number(r.target_quantity) || 0,
          producedQuantity: Number(r.produced_quantity) || 0,
          exportedQuantity: Number(r.exported_quantity) || 0,
          status: r.status,
          createdAt: Number(r.created_at) || Date.now(),
          completedAt: r.completed_at ? Number(r.completed_at) : undefined,
          plannedStartTime: r.planned_start_time ? Number(r.planned_start_time) : undefined,
          leadTime: r.lead_time ? Number(r.lead_time) : undefined,
          expectedCompletionTime: r.expected_completion_time ? Number(r.expected_completion_time) : undefined,
          note: r.note || undefined,
          isOddPO: Boolean(r.is_odd_po),
          excelRequiredDate: r.excel_required_date ? Number(r.excel_required_date) : undefined
        }));
        const validPos = this.filterExpiredCompletedPOs(pos);
        cache[STORAGE_KEYS.PRODUCTION_ORDERS] = validPos;
        try {
          localStorage.setItem(STORAGE_KEYS.PRODUCTION_ORDERS, JSON.stringify(validPos));
        } catch {}
      }

      // 3. Fetch wip_inventory (selective columns)
      const { data: invRows, error: invErr } = await supabase
        .from('wip_inventory')
        .select('part_id, original_part_id, stage_id, location, quantity');

      if (!invErr && invRows && invRows.length > 0) {
        const inv: InventoryItem[] = invRows.map((r: any) => ({
          partId: r.part_id,
          originalPartId: r.original_part_id || undefined,
          stageId: r.stage_id,
          location: r.location,
          quantity: Number(r.quantity) || 0
        }));
        cache[STORAGE_KEYS.INVENTORY] = inv;
        try {
          localStorage.setItem(STORAGE_KEYS.INVENTORY, JSON.stringify(inv));
        } catch {}
      }

      // 4. Fetch wip_transactions (only within 35 days, selective columns)
      const cutoff35d = Date.now() - 35 * 24 * 60 * 60 * 1000;
      const { data: txRows, error: txErr } = await supabase
        .from('wip_transactions')
        .select('id, type, part_id, part_name, original_part_id, quantity, stage_id, location, timestamp, qr_data, source_stage_id, target_stage_id, po_id, plan_id, printed, defect_reason, defect_category, kpi_recorded')
        .gte('timestamp', cutoff35d)
        .order('timestamp', { ascending: false });

      if (!txErr && txRows && txRows.length > 0) {
        const txs: Transaction[] = txRows.map((r: any) => ({
          id: r.id,
          type: r.type,
          partId: r.part_id,
          partName: r.part_name || undefined,
          originalPartId: r.original_part_id || undefined,
          quantity: Number(r.quantity) || 0,
          stageId: r.stage_id,
          location: r.location || undefined,
          timestamp: Number(r.timestamp),
          qrData: r.qr_data || undefined,
          sourceStageId: r.source_stage_id || undefined,
          targetStageId: r.target_stage_id || undefined,
          poId: r.po_id || undefined,
          planId: r.plan_id || undefined,
          printed: Boolean(r.printed),
          defectReason: r.defect_reason || undefined,
          defectCategory: r.defect_category || undefined,
          kpiRecorded: Boolean(r.kpi_recorded)
        }));
        cache[STORAGE_KEYS.TRANSACTIONS] = txs;
        try {
          localStorage.setItem(STORAGE_KEYS.TRANSACTIONS, JSON.stringify(txs));
        } catch {}
      }

      // 5. Fetch glazing_plans
      const { data: gpRows, error: gpErr } = await supabase
        .from('glazing_plans')
        .select('id, model_id, target_quantity, target_completion_time, planned_start_time, expected_completion_time, created_at, status, produced_quantities, note, excel_required_date')
        .order('created_at', { ascending: false });

      if (!gpErr && gpRows && gpRows.length > 0) {
        const plans: import('./types').GlazingPlan[] = gpRows.map((r: any) => ({
          id: r.id,
          modelId: r.model_id,
          targetQuantity: Number(r.target_quantity) || 0,
          targetCompletionTime: Number(r.target_completion_time),
          plannedStartTime: r.planned_start_time ? Number(r.planned_start_time) : undefined,
          expectedCompletionTime: r.expected_completion_time ? Number(r.expected_completion_time) : undefined,
          createdAt: Number(r.created_at),
          status: r.status,
          producedQuantities: r.produced_quantities || {},
          note: r.note || undefined,
          excelRequiredDate: r.excel_required_date ? Number(r.excel_required_date) : undefined
        }));
        cache[STORAGE_KEYS.GLAZING_PLANS] = plans;
        try {
          localStorage.setItem(STORAGE_KEYS.GLAZING_PLANS, JSON.stringify(plans));
        } catch {}
      }

      // 6. Fetch qr_labels (last 24 hours)
      const cutoff24h = Date.now() - 24 * 60 * 60 * 1000;
      const { data: qrRows, error: qrErr } = await supabase
        .from('qr_labels')
        .select('id, part_id, stage_id, quantity, timestamp, qr_data, metadata')
        .gte('timestamp', cutoff24h)
        .order('timestamp', { ascending: false });

      if (!qrErr && qrRows && qrRows.length > 0) {
        const labels: Transaction[] = qrRows.map((r: any) => {
          if (r.metadata && typeof r.metadata === 'object' && r.metadata.id) {
            return r.metadata;
          }
          return {
            id: r.id,
            type: 'STAGE_OUT',
            partId: r.part_id || '',
            quantity: Number(r.quantity) || 0,
            stageId: r.stage_id || 'LASER',
            timestamp: Number(r.timestamp),
            qrData: r.qr_data || undefined
          };
        });
        cache['wip_labels'] = labels;
        try {
          localStorage.setItem('wip_labels', JSON.stringify(labels));
        } catch {}
      }

      // If Supabase tables are freshly created and empty, but localStorage has data, auto-migrate to Supabase!
      const isOnlineEmpty = (!configRows || configRows.length === 0) && (!poRows || poRows.length === 0);
      const hasLocalData = (localStorage.getItem(STORAGE_KEYS.PARTS) || localStorage.getItem(STORAGE_KEYS.PRODUCTION_ORDERS));
      if (isOnlineEmpty && hasLocalData) {
        this.syncLocalToSupabase().catch(e => console.warn('[Supabase] Initial seed error:', e));
      }

      lastSyncTime = Date.now();
      isSyncing = false;
      syncError = null;
      notifySyncChange();
      return true;
    } catch (err: any) {
      isSyncing = false;
      syncError = err?.message || 'Lỗi kết nối Supabase';
      notifySyncChange();
      console.warn('[Supabase] Init error:', err);
      return false;
    }
  },

  async refreshFromSupabase(): Promise<boolean> {
    clearCache();
    return await this.init();
  },

  /**
   * Upload all local cache data to Supabase (Batch Upsert in chunks)
   */
  async syncLocalToSupabase(): Promise<{ success: boolean; message: string; details?: any }> {
    if (!isSupabaseConfigured) {
      return { success: false, message: 'Chưa cấu hình SUPABASE_URL hoặc SUPABASE_ANON_KEY' };
    }

    isSyncing = true;
    notifySyncChange();

    try {
      // 1. Sync system_configs
      const configKeys = [
        STORAGE_KEYS.PARTS,
        STORAGE_KEYS.BOM,
        STORAGE_KEYS.BOM_V2,
        STORAGE_KEYS.MODEL_BOM,
        STORAGE_KEYS.NORMS,
        STORAGE_KEYS.LASER_NESTING,
        STORAGE_KEYS.SHIFT_CONFIGS,
        STORAGE_KEYS.LABEL_SETTINGS,
        STORAGE_KEYS.TRANSFORMATIONS,
        STORAGE_KEYS.GLAZING_CONFIGS,
        STORAGE_KEYS.GLAZING_OUT_CONFIGS,
        STORAGE_KEYS.QUICK_PRINT_PARTS,
        STORAGE_KEYS.GLAZING_PLAN_NORMS,
        STORAGE_KEYS.STAGE_PLANNED_WORKERS,
        STORAGE_KEYS.BENDING_WELDING_HSQD,
        STORAGE_KEYS.LASER_HSQD,
        STORAGE_KEYS.PEOPLE_PER_DAY,
        STORAGE_KEYS.MANDAYS_PER_DAY,
        STORAGE_KEYS.EXPORT_WEEK_NAME,
        STORAGE_KEYS.HOURLY_PEOPLE_LASER,
        STORAGE_KEYS.HOURLY_PEOPLE_PAINTING,
        STORAGE_KEYS.HOURLY_PEOPLE_GLAZING,
        STORAGE_KEYS.HOURLY_PEOPLE_BENDING,
        STORAGE_KEYS.HOURLY_PEOPLE_WELDING,
      ];

      const configRows: { key: string; value: any }[] = [];
      configKeys.forEach(k => {
        const val = cache[k] !== undefined ? cache[k] : (() => {
          const raw = localStorage.getItem(k);
          if (!raw) return null;
          try { return JSON.parse(raw); } catch { return raw; }
        })();
        if (val !== null && val !== undefined) {
          configRows.push({ key: k, value: val });
        }
      });

      if (configRows.length > 0) {
        const { error: cfgErr } = await supabase.from('system_configs').upsert(configRows, { onConflict: 'key' });
        if (cfgErr) throw cfgErr;
      }

      // 2. Sync Production Orders in batch chunks
      const pos = this.getProductionOrders();
      if (pos.length > 0) {
        await persistProductionOrdersToSupabase(pos);
      }

      // 3. Sync Inventory
      const inventory = this.getInventory();
      if (inventory.length > 0) {
        await persistInventoryToSupabase(inventory);
      }

      // 4. Sync Transactions (last 35 days)
      const txs = this.getTransactions();
      if (txs.length > 0) {
        await persistTransactionsToSupabase(txs);
      }

      // 5. Sync Glazing Plans
      const plans = this.getGlazingPlans();
      if (plans.length > 0) {
        await persistGlazingPlansToSupabase(plans);
      }

      // 6. Sync QR Labels
      const labels = this.getLabels();
      if (labels.length > 0) {
        await persistLabelsToSupabase(labels);
      }

      lastSyncTime = Date.now();
      isSyncing = false;
      syncError = null;
      notifySyncChange();

      return {
        success: true,
        message: `Đã đồng bộ thành công lên Supabase: ${pos.length} lệnh sản xuất, ${inventory.length} mục tồn kho, ${txs.length} giao dịch, ${plans.length} kế hoạch dán kính, ${configRows.length} cấu hình hệ thống.`
      };
    } catch (err: any) {
      isSyncing = false;
      syncError = err?.message || 'Lỗi khi đẩy dữ liệu lên Supabase';
      notifySyncChange();
      return {
        success: false,
        message: `Không thể đồng bộ lên Supabase: ${err?.message || err}`,
        details: err
      };
    }
  },

  clearCache() {
    clearCache();
  },

  async resetAllData() {
    clearCache();
    if (typeof window !== 'undefined' && window.localStorage) {
      const keysToRemove: string[] = [];
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (k && (k.startsWith('wip_') || Object.values(STORAGE_KEYS).includes(k) || k === 'wip_labels')) {
          keysToRemove.push(k);
        }
      }
      keysToRemove.forEach(k => localStorage.removeItem(k));
    }

    if (isSupabaseConfigured) {
      try {
        await Promise.allSettled([
          supabase.from('system_configs').delete().neq('key', ''),
          supabase.from('production_orders').delete().neq('id', ''),
          supabase.from('wip_inventory').delete().neq('id', ''),
          supabase.from('wip_transactions').delete().neq('id', ''),
          supabase.from('glazing_plans').delete().neq('id', ''),
          supabase.from('qr_labels').delete().neq('id', '')
        ]);
      } catch (e) {
        console.warn('[Supabase] resetAllData error:', e);
      }
    }
  },
};

