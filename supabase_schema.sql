-- ==============================================================================
-- WIP Inventory Management System - Supabase Database Schema
-- Run this script in the Supabase SQL Editor: https://supabase.com/dashboard/project/kopsoonfrmtmvpbjyhdu/sql
-- ==============================================================================

-- 1. Table: system_configs
-- Stores system configurations, master parts, BOMs, productivity norms, shift configs, transformations, and data_reset_version
CREATE TABLE IF NOT EXISTS public.system_configs (
    key VARCHAR(100) PRIMARY KEY,
    value JSONB NOT NULL,
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Khởi tạo cờ System Reset Flag: data_reset_version (mặc định '1.0.0')
INSERT INTO public.system_configs (key, value)
VALUES ('data_reset_version', '"1.0.0"'::jsonb)
ON CONFLICT (key) DO NOTHING;

-- ==============================================================================
-- QUY TRÌNH ADMIN LÀM SẠCH DỮ LIỆU (TRUNCATE) & KÍCH HOẠT XÓA CACHE CLIENT:
-- 1) Xóa trắng dữ liệu giao dịch:
--    TRUNCATE TABLE public.production_orders, public.wip_inventory, public.wip_transactions, public.glazing_plans, public.qr_labels;
-- 2) Cập nhật phiên bản cờ reset sang version mới (ví dụ: dùng timestamp):
--    INSERT INTO public.system_configs (key, value, updated_at)
--    VALUES ('data_reset_version', to_jsonb(to_char(NOW(), 'YYYYMMDD_HH24MISS')), NOW())
--    ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_at = NOW();
-- -> Tất cả máy trạm mở app lên sẽ phát hiện cờ thay đổi, tự động xóa sạch LocalStorage
--    và KHÔNG BAO GIỜ đẩy dữ liệu cũ lên Supabase!
-- ==============================================================================

-- 2. Table: production_orders
-- Stores all Master POs and Sub-POs across all production stages
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

-- Indexes for production_orders
CREATE INDEX IF NOT EXISTS idx_po_master_id ON public.production_orders(master_po_id);
CREATE INDEX IF NOT EXISTS idx_po_stage_id ON public.production_orders(stage_id);
CREATE INDEX IF NOT EXISTS idx_po_status ON public.production_orders(status);
CREATE INDEX IF NOT EXISTS idx_po_created_at ON public.production_orders(created_at DESC);

-- 3. Table: wip_inventory
-- Stores real-time inventory quantity across parts, stages, and locations (IN, OUT, DEFECT)
CREATE TABLE IF NOT EXISTS public.wip_inventory (
    id VARCHAR(255) PRIMARY KEY, -- synthesized as: part_id || '_' || stage_id || '_' || location || '_' || COALESCE(original_part_id, '')
    part_id VARCHAR(100) NOT NULL,
    original_part_id VARCHAR(100) DEFAULT '',
    stage_id VARCHAR(50) NOT NULL,
    location VARCHAR(50) NOT NULL,
    quantity NUMERIC NOT NULL DEFAULT 0,
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Indexes for wip_inventory
CREATE INDEX IF NOT EXISTS idx_inv_stage_loc ON public.wip_inventory(stage_id, location);
CREATE INDEX IF NOT EXISTS idx_inv_part ON public.wip_inventory(part_id);

-- 4. Table: wip_transactions
-- Stores transaction history, stage movements, inbound/outbound scans, defect and disposal logs
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

-- Indexes for wip_transactions
CREATE INDEX IF NOT EXISTS idx_tx_timestamp ON public.wip_transactions(timestamp DESC);
CREATE INDEX IF NOT EXISTS idx_tx_part ON public.wip_transactions(part_id);
CREATE INDEX IF NOT EXISTS idx_tx_stage ON public.wip_transactions(stage_id);
CREATE INDEX IF NOT EXISTS idx_tx_po ON public.wip_transactions(po_id);

-- 5. Table: glazing_plans
-- Stores Glazing assembly production schedules and progress
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

-- Indexes for glazing_plans
CREATE INDEX IF NOT EXISTS idx_gp_created_at ON public.glazing_plans(created_at DESC);
CREATE INDEX IF NOT EXISTS idx_gp_status ON public.glazing_plans(status);

-- 6. Table: qr_labels
-- Stores generated and printed QR labels with 24-hour retention
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

-- Indexes for qr_labels
CREATE INDEX IF NOT EXISTS idx_qr_timestamp ON public.qr_labels(timestamp DESC);

-- ==============================================================================
-- Row Level Security (RLS) & Public / Anon Access Policies
-- ==============================================================================
ALTER TABLE public.system_configs ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.production_orders ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.wip_inventory ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.wip_transactions ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.glazing_plans ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.qr_labels ENABLE ROW LEVEL SECURITY;

DO $$ 
BEGIN
    -- system_configs
    DROP POLICY IF EXISTS "Allow anon all on system_configs" ON public.system_configs;
    CREATE POLICY "Allow anon all on system_configs" ON public.system_configs FOR ALL TO anon, authenticated USING (true) WITH CHECK (true);

    -- production_orders
    DROP POLICY IF EXISTS "Allow anon all on production_orders" ON public.production_orders;
    CREATE POLICY "Allow anon all on production_orders" ON public.production_orders FOR ALL TO anon, authenticated USING (true) WITH CHECK (true);

    -- wip_inventory
    DROP POLICY IF EXISTS "Allow anon all on wip_inventory" ON public.wip_inventory;
    CREATE POLICY "Allow anon all on wip_inventory" ON public.wip_inventory FOR ALL TO anon, authenticated USING (true) WITH CHECK (true);

    -- wip_transactions
    DROP POLICY IF EXISTS "Allow anon all on wip_transactions" ON public.wip_transactions;
    CREATE POLICY "Allow anon all on wip_transactions" ON public.wip_transactions FOR ALL TO anon, authenticated USING (true) WITH CHECK (true);

    -- glazing_plans
    DROP POLICY IF EXISTS "Allow anon all on glazing_plans" ON public.glazing_plans;
    CREATE POLICY "Allow anon all on glazing_plans" ON public.glazing_plans FOR ALL TO anon, authenticated USING (true) WITH CHECK (true);

    -- qr_labels
    DROP POLICY IF EXISTS "Allow anon all on qr_labels" ON public.qr_labels;
    CREATE POLICY "Allow anon all on qr_labels" ON public.qr_labels FOR ALL TO anon, authenticated USING (true) WITH CHECK (true);
END $$;

-- Grant permissions to public anon and authenticated roles
GRANT ALL ON ALL TABLES IN SCHEMA public TO anon, authenticated;
GRANT ALL ON ALL SEQUENCES IN SCHEMA public TO anon, authenticated;

-- ==============================================================================
-- Function & Trigger: Automatic updated_at timestamp
-- ==============================================================================
CREATE OR REPLACE FUNCTION public.handle_updated_at()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = NOW();
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS tr_system_configs_updated_at ON public.system_configs;
CREATE TRIGGER tr_system_configs_updated_at BEFORE UPDATE ON public.system_configs FOR EACH ROW EXECUTE FUNCTION public.handle_updated_at();

DROP TRIGGER IF EXISTS tr_production_orders_updated_at ON public.production_orders;
CREATE TRIGGER tr_production_orders_updated_at BEFORE UPDATE ON public.production_orders FOR EACH ROW EXECUTE FUNCTION public.handle_updated_at();

DROP TRIGGER IF EXISTS tr_wip_inventory_updated_at ON public.wip_inventory;
CREATE TRIGGER tr_wip_inventory_updated_at BEFORE UPDATE ON public.wip_inventory FOR EACH ROW EXECUTE FUNCTION public.handle_updated_at();

DROP TRIGGER IF EXISTS tr_glazing_plans_updated_at ON public.glazing_plans;
CREATE TRIGGER tr_glazing_plans_updated_at BEFORE UPDATE ON public.glazing_plans FOR EACH ROW EXECUTE FUNCTION public.handle_updated_at();

-- ==============================================================================
-- Function: Auto cleanup old data directly on Supabase database
-- Deletes transactions older than 35 days, labels older than 24 hours, and expired completed POs
-- ==============================================================================
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

    -- 1. Delete transactions older than 35 days
    DELETE FROM public.wip_transactions WHERE timestamp < cutoff_35d;
    GET DIAGNOSTICS deleted_txs = ROW_COUNT;

    -- 2. Delete QR labels older than 24 hours
    DELETE FROM public.qr_labels WHERE timestamp < cutoff_24h;
    GET DIAGNOSTICS deleted_labels = ROW_COUNT;

    -- 3. Delete completed POs older than 35 days without active child POs
    DELETE FROM public.production_orders 
    WHERE status = 'COMPLETED' 
      AND completed_at IS NOT NULL 
      AND completed_at < cutoff_35d
      AND (master_po_id IS NULL OR master_po_id NOT IN (
          SELECT DISTINCT master_po_id 
          FROM public.production_orders 
          WHERE status != 'COMPLETED' AND master_po_id IS NOT NULL
      ));
    GET DIAGNOSTICS deleted_pos = ROW_COUNT;

    RETURN json_build_object(
        'deleted_transactions', deleted_txs,
        'deleted_labels', deleted_labels,
        'deleted_production_orders', deleted_pos
    );
END;
$$ LANGUAGE plpgsql;
