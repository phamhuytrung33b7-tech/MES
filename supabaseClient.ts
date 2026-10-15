/**
 * Supabase Client Connection
 * Configured with client credentials and network resilience options
 */

import { createClient } from '@supabase/supabase-js';

export const SUPABASE_URL = (import.meta as any).env?.VITE_SUPABASE_URL || 'https://kopsoonfrmtmvpbjyhdu.supabase.co';
export const SUPABASE_ANON_KEY = (import.meta as any).env?.VITE_SUPABASE_ANON_KEY || 'sb_publishable_nHl6OWIGVSmKs7oCj3qFrQ_oK90ZDuM';

export const isSupabaseConfigured = Boolean(SUPABASE_URL && SUPABASE_ANON_KEY);

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  auth: {
    persistSession: false,
    autoRefreshToken: false,
    detectSessionInUrl: false
  },
  global: {
    headers: {
      'x-application-name': 'wip-inventory-manager'
    }
  }
});

/**
 * Quick ping to verify Supabase table connectivity
 */
export async function testSupabaseConnection(): Promise<{ ok: boolean; message: string; details?: any }> {
  try {
    const { data, error } = await supabase.from('system_configs').select('key').limit(1);
    if (error) {
      if (error.code === 'PGRST205' || error.message?.includes('not found') || error.message?.includes('does not exist')) {
        return {
          ok: false,
          message: 'Kết nối thành công nhưng chưa tìm thấy các bảng trong cơ sở dữ liệu Supabase. Vui lòng chạy mã lệnh SQL trong file supabase_schema.sql trên Supabase SQL Editor.',
          details: error
        };
      }
      return { ok: false, message: `Lỗi kết nối Supabase: ${error.message}`, details: error };
    }
    return { ok: true, message: 'Kết nối cơ sở dữ liệu Supabase online thành công!', details: data };
  } catch (err: any) {
    return { ok: false, message: `Không thể kết nối đến Supabase: ${err?.message || err}`, details: err };
  }
}

export const SUPABASE_SQL_SCHEMA = `-- 1. Table: system_configs
CREATE TABLE IF NOT EXISTS public.system_configs (
    key VARCHAR(100) PRIMARY KEY,
    value JSONB NOT NULL,
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- 2. Table: production_orders
CREATE TABLE IF NOT EXISTS public.production_orders (
    id VARCHAR(100) PRIMARY KEY,
    master_po_id VARCHAR(100),
    part_id VARCHAR(100) NOT NULL,
    stage_id VARCHAR(50),
    target_quantity NUMERIC NOT NULL,
    produced_quantity NUMERIC DEFAULT 0,
    exported_quantity NUMERIC DEFAULT 0,
    status VARCHAR(50) DEFAULT 'PENDING',
    created_at BIGINT,
    completed_at BIGINT,
    planned_start_time BIGINT,
    lead_time NUMERIC,
    expected_completion_time BIGINT,
    note TEXT,
    is_odd_po BOOLEAN DEFAULT FALSE,
    excel_required_date BIGINT,
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_po_master_id ON public.production_orders(master_po_id);
CREATE INDEX IF NOT EXISTS idx_po_stage_id ON public.production_orders(stage_id);
CREATE INDEX IF NOT EXISTS idx_po_status ON public.production_orders(status);
CREATE INDEX IF NOT EXISTS idx_po_created_at ON public.production_orders(created_at DESC);

-- 3. Table: wip_inventory
CREATE TABLE IF NOT EXISTS public.wip_inventory (
    id VARCHAR(255) PRIMARY KEY,
    part_id VARCHAR(100) NOT NULL,
    original_part_id VARCHAR(100) DEFAULT '',
    stage_id VARCHAR(50) NOT NULL,
    location VARCHAR(50) NOT NULL,
    quantity NUMERIC NOT NULL DEFAULT 0,
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_inv_stage_loc ON public.wip_inventory(stage_id, location);
CREATE INDEX IF NOT EXISTS idx_inv_part ON public.wip_inventory(part_id);

-- 4. Table: wip_transactions
CREATE TABLE IF NOT EXISTS public.wip_transactions (
    id VARCHAR(100) PRIMARY KEY,
    type VARCHAR(50) NOT NULL,
    part_id VARCHAR(100) NOT NULL,
    part_name TEXT,
    original_part_id VARCHAR(100),
    quantity NUMERIC NOT NULL,
    stage_id VARCHAR(50) NOT NULL,
    location VARCHAR(50),
    timestamp BIGINT NOT NULL,
    qr_data TEXT,
    source_stage_id VARCHAR(50),
    target_stage_id VARCHAR(50),
    po_id VARCHAR(100),
    plan_id VARCHAR(100),
    printed BOOLEAN DEFAULT FALSE,
    defect_reason TEXT,
    defect_category TEXT,
    kpi_recorded BOOLEAN DEFAULT FALSE,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_tx_timestamp ON public.wip_transactions(timestamp DESC);
CREATE INDEX IF NOT EXISTS idx_tx_part ON public.wip_transactions(part_id);
CREATE INDEX IF NOT EXISTS idx_tx_stage ON public.wip_transactions(stage_id);
CREATE INDEX IF NOT EXISTS idx_tx_po ON public.wip_transactions(po_id);

-- 5. Table: glazing_plans
CREATE TABLE IF NOT EXISTS public.glazing_plans (
    id VARCHAR(100) PRIMARY KEY,
    model_id VARCHAR(100) NOT NULL,
    target_quantity NUMERIC NOT NULL,
    target_completion_time BIGINT NOT NULL,
    planned_start_time BIGINT,
    expected_completion_time BIGINT,
    created_at BIGINT NOT NULL,
    status VARCHAR(50) DEFAULT 'PENDING',
    produced_quantities JSONB DEFAULT '{}'::jsonb,
    note TEXT,
    excel_required_date BIGINT,
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_gp_created_at ON public.glazing_plans(created_at DESC);
CREATE INDEX IF NOT EXISTS idx_gp_status ON public.glazing_plans(status);

-- 6. Table: qr_labels
CREATE TABLE IF NOT EXISTS public.qr_labels (
    id VARCHAR(100) PRIMARY KEY,
    part_id VARCHAR(100),
    stage_id VARCHAR(50),
    quantity NUMERIC,
    timestamp BIGINT NOT NULL,
    qr_data TEXT,
    metadata JSONB DEFAULT '{}'::jsonb,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_qr_timestamp ON public.qr_labels(timestamp DESC);

-- Enable RLS and grant permissions
ALTER TABLE public.system_configs ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.production_orders ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.wip_inventory ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.wip_transactions ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.glazing_plans ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.qr_labels ENABLE ROW LEVEL SECURITY;

DO $$ 
BEGIN
    DROP POLICY IF EXISTS "Allow anon all on system_configs" ON public.system_configs;
    CREATE POLICY "Allow anon all on system_configs" ON public.system_configs FOR ALL TO anon, authenticated USING (true) WITH CHECK (true);

    DROP POLICY IF EXISTS "Allow anon all on production_orders" ON public.production_orders;
    CREATE POLICY "Allow anon all on production_orders" ON public.production_orders FOR ALL TO anon, authenticated USING (true) WITH CHECK (true);

    DROP POLICY IF EXISTS "Allow anon all on wip_inventory" ON public.wip_inventory;
    CREATE POLICY "Allow anon all on wip_inventory" ON public.wip_inventory FOR ALL TO anon, authenticated USING (true) WITH CHECK (true);

    DROP POLICY IF EXISTS "Allow anon all on wip_transactions" ON public.wip_transactions;
    CREATE POLICY "Allow anon all on wip_transactions" ON public.wip_transactions FOR ALL TO anon, authenticated USING (true) WITH CHECK (true);

    DROP POLICY IF EXISTS "Allow anon all on glazing_plans" ON public.glazing_plans;
    CREATE POLICY "Allow anon all on glazing_plans" ON public.glazing_plans FOR ALL TO anon, authenticated USING (true) WITH CHECK (true);

    DROP POLICY IF EXISTS "Allow anon all on qr_labels" ON public.qr_labels;
    CREATE POLICY "Allow anon all on qr_labels" ON public.qr_labels FOR ALL TO anon, authenticated USING (true) WITH CHECK (true);
END $$;

GRANT ALL ON ALL TABLES IN SCHEMA public TO anon, authenticated;
GRANT ALL ON ALL SEQUENCES IN SCHEMA public TO anon, authenticated;

-- Function: Auto cleanup old data (>35 days and QR labels >24 hours)
CREATE OR REPLACE FUNCTION public.cleanup_old_wip_data()
RETURNS json AS $$
DECLARE
    cutoff_35d BIGINT;
    cutoff_24h BIGINT;
    deleted_txs INT;
    deleted_labels INT;
    deleted_pos INT;
BEGIN
    cutoff_35d := (EXTRACT(EPOCH FROM (NOW() - INTERVAL '35 days')) * 1000)::BIGINT;
    cutoff_24h := (EXTRACT(EPOCH FROM (NOW() - INTERVAL '24 hours')) * 1000)::BIGINT;

    DELETE FROM public.wip_transactions WHERE timestamp < cutoff_35d;
    GET DIAGNOSTICS deleted_txs = ROW_COUNT;

    DELETE FROM public.qr_labels WHERE timestamp < cutoff_24h;
    GET DIAGNOSTICS deleted_labels = ROW_COUNT;

    DELETE FROM public.production_orders 
    WHERE status = 'COMPLETED' 
      AND completed_at IS NOT NULL 
      AND completed_at < cutoff_35d
      AND (master_po_id IS NULL OR master_po_id NOT IN (
          SELECT DISTINCT master_po_id FROM public.production_orders WHERE status != 'COMPLETED' AND master_po_id IS NOT NULL
      ));
    GET DIAGNOSTICS deleted_pos = ROW_COUNT;

    RETURN json_build_object(
        'deleted_transactions', deleted_txs,
        'deleted_labels', deleted_labels,
        'deleted_production_orders', deleted_pos
    );
END;
$$ LANGUAGE plpgsql;`;
