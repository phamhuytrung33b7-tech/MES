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
    detectSessionInUrl: false,
    debug: false,
  },
  global: {
    headers: {
      'x-application-name': 'wip-inventory-manager'
    }
  },
  db: {
    schema: 'public'
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

export const SUPABASE_SQL_SCHEMA = `-- 1. Table: system_configs (Lưu cấu hình hệ thống và cờ reset data_reset_version)
CREATE TABLE IF NOT EXISTS public.system_configs (
    key VARCHAR(100) PRIMARY KEY,
    value JSONB NOT NULL,
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Khởi tạo cờ data_reset_version ban đầu ('1.0.0')
INSERT INTO public.system_configs (key, value)
VALUES ('data_reset_version', '"1.0.0"'::jsonb)
ON CONFLICT (key) DO NOTHING;

-- Khởi tạo / Cập nhật Hệ số quy đổi (HSQĐ) Dán Kính 2 (Dùng cho Biểu đồ công đoạn Dán Kính riêng lẻ)
INSERT INTO public.system_configs (key, value, updated_at)
VALUES (
    'wip_glazing_stage_hsqd',
    '[
        {"partId": "04-29-08-SHA76210KL-0040", "partName": "Gói cánh tủ SHA76210KL(BD)", "hsqd": 0.16},
        {"partId": "04-29-08-SHA76210KL-0041", "partName": "Gói đỉnh tủ SHA76210KL(BD)", "hsqd": 0.13},
        {"partId": "04-29-08-SHA76211KL-0002", "partName": "Linh kiện dán kính mẫu 1", "hsqd": 0.16},
        {"partId": "04-29-08-SHA76212KL-0002", "partName": "Linh kiện dán kính mẫu 2", "hsqd": 0.16}
    ]'::jsonb,
    NOW()
)
ON CONFLICT (key) DO UPDATE 
SET value = EXCLUDED.value, updated_at = NOW();

-- ==============================================================================
-- QUY TRÌNH ADMIN LÀM SẠCH DỮ LIỆU (TRUNCATE) & KÍCH HOẠT XÓA CACHE CLIENT:
-- 1) Xóa trắng dữ liệu giao dịch:
--    TRUNCATE TABLE public.production_orders, public.wip_inventory, public.wip_transactions, public.glazing_plans, public.qr_labels;
-- 2) Cập nhật phiên bản cờ reset sang version mới:
--    INSERT INTO public.system_configs (key, value, updated_at)
--    VALUES ('data_reset_version', to_jsonb(to_char(NOW(), 'YYYYMMDD_HH24MISS')), NOW())
--    ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_at = NOW();
-- -> Tất cả máy trạm mở app lên sẽ phát hiện cờ thay đổi, tự động xóa sạch LocalStorage
--    và KHÔNG BAO GIỜ đẩy dữ liệu cũ lên Supabase!
-- ==============================================================================

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
$$ LANGUAGE plpgsql;

-- 7. Table: wip_balance (Cân đối tồn kho thời gian thực)
CREATE TABLE IF NOT EXISTS public.wip_balance (
    id VARCHAR(255) PRIMARY KEY,
    part_code VARCHAR(100) NOT NULL,
    stage_id VARCHAR(50) NOT NULL,
    location VARCHAR(50) NOT NULL DEFAULT 'OUT',
    quantity NUMERIC NOT NULL DEFAULT 0,
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_wip_bal_lookup ON public.wip_balance(part_code, stage_id, location);
ALTER TABLE public.wip_balance ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "Allow anon all on wip_balance" ON public.wip_balance;
CREATE POLICY "Allow anon all on wip_balance" ON public.wip_balance FOR ALL TO anon, authenticated USING (true) WITH CHECK (true);
GRANT ALL ON TABLE public.wip_balance TO anon, authenticated, service_role;

-- Helper Function: normalize_stage_code
CREATE OR REPLACE FUNCTION public.normalize_stage_code(p_stage TEXT)
RETURNS TEXT AS $$
DECLARE
    v_clean TEXT;
BEGIN
    IF p_stage IS NULL THEN RETURN 'UNKNOWN'; END IF;
    v_clean := UPPER(TRIM(p_stage));
    IF v_clean LIKE '%SƠN%' OR v_clean LIKE '%SON%' OR v_clean = 'PAINTING' OR v_clean = 'KHO SƠN' OR v_clean = 'KHO_SON' THEN RETURN 'PAINTING'; END IF;
    IF v_clean LIKE '%DCLR%' OR v_clean LIKE '%LẮP RÁP%' OR v_clean LIKE '%LAP RAP%' THEN RETURN 'DCLR'; END IF;
    IF v_clean LIKE '%LASER%' OR v_clean LIKE '%CẮT%' OR v_clean LIKE '%CAT%' THEN RETURN 'LASER'; END IF;
    IF v_clean LIKE '%CHẤN%' OR v_clean LIKE '%CHAN%' OR v_clean LIKE '%DẬP%' OR v_clean LIKE '%DAP%' OR v_clean = 'BENDING' THEN RETURN 'BENDING'; END IF;
    IF v_clean LIKE '%HÀN%' OR v_clean LIKE '%HAN%' OR v_clean = 'WELDING' THEN RETURN 'WELDING'; END IF;
    IF v_clean LIKE '%DÁN KÍNH%' OR v_clean LIKE '%DAN KINH%' OR v_clean LIKE '%KÍNH%' OR v_clean = 'GLAZING' THEN RETURN 'GLAZING'; END IF;
    RETURN v_clean;
END;
$$ LANGUAGE plpgsql IMMUTABLE;
GRANT EXECUTE ON FUNCTION public.normalize_stage_code(TEXT) TO anon, authenticated, service_role;

-- 8. Stored Procedure (RPC): transfer_wip_stage_atomic
CREATE OR REPLACE FUNCTION public.transfer_wip_stage_atomic(
    p_qr_id TEXT,
    p_part_code TEXT,
    p_from_stage TEXT,
    p_to_stage TEXT,
    p_qty NUMERIC,
    p_from_location TEXT DEFAULT 'OUT',
    p_to_location TEXT DEFAULT 'IN',
    p_po_id TEXT DEFAULT NULL,
    p_operator TEXT DEFAULT 'SYSTEM'
)
RETURNS JSONB AS $$
DECLARE
    v_from_stage TEXT;
    v_to_stage TEXT;
    v_clean_part TEXT;
    v_current_stock NUMERIC := 0;
    v_from_bal_id TEXT;
    v_to_bal_id TEXT;
    v_to_inv_id TEXT;
    v_new_source_stock NUMERIC := 0;
    v_new_target_stock NUMERIC := 0;
    v_tx_id TEXT;
    v_timestamp_ms BIGINT;
    v_label_row RECORD;
BEGIN
    IF p_qty IS NULL OR p_qty <= 0 THEN
        RAISE EXCEPTION 'INVALID_QUANTITY: Số lượng chuyển phải lớn hơn 0 (Hiện nhận: %)', p_qty;
    END IF;

    IF p_part_code IS NULL OR TRIM(p_part_code) = '' THEN
        RAISE EXCEPTION 'INVALID_PART_CODE: Mã linh kiện không được để trống';
    END IF;

    v_from_stage := public.normalize_stage_code(p_from_stage);
    v_to_stage := public.normalize_stage_code(p_to_stage);
    v_clean_part := TRIM(p_part_code);
    v_timestamp_ms := (EXTRACT(EPOCH FROM NOW()) * 1000)::BIGINT;
    v_tx_id := 'TR-' || UPPER(SUBSTRING(MD5(RANDOM()::TEXT || NOW()::TEXT) FROM 1 FOR 8));

    IF v_from_stage = v_to_stage THEN
        RAISE EXCEPTION 'SAME_STAGE: Không thể chuyển kho trong cùng một công đoạn (%)', v_from_stage;
    END IF;

    -- Kiểm tra nhãn QR
    IF p_qr_id IS NOT NULL AND TRIM(p_qr_id) <> '' THEN
        SELECT * INTO v_label_row 
        FROM public.qr_labels 
        WHERE id = p_qr_id OR qr_data LIKE ('%' || p_qr_id || '%')
        FOR UPDATE;

        IF FOUND THEN
            IF (v_label_row.metadata->>'status') = 'TRANSFERRED' OR (v_label_row.metadata->>'status') = 'COMPLETED' THEN
                RAISE EXCEPTION 'QR_ALREADY_USED: Mã QR [%] đã được quét chuyển kho trước đó vào lúc %', 
                    p_qr_id, COALESCE(v_label_row.metadata->>'transferred_at', 'trước đó');
            END IF;
        END IF;
    END IF;

    -- Khóa và kiểm tra tồn kho nguồn
    SELECT COALESCE(SUM(quantity), 0) INTO v_current_stock
    FROM public.wip_inventory
    WHERE (part_id = v_clean_part OR part_id ILIKE (v_clean_part || '%'))
      AND stage_id = v_from_stage 
      AND location = p_from_location
    FOR UPDATE;

    IF v_current_stock = 0 THEN
        SELECT COALESCE(SUM(quantity), 0) INTO v_current_stock
        FROM public.wip_balance
        WHERE (part_code = v_clean_part OR part_code ILIKE (v_clean_part || '%'))
          AND stage_id = v_from_stage 
          AND location = p_from_location
        FOR UPDATE;
    END IF;

    IF v_current_stock < p_qty THEN
        RAISE EXCEPTION 'INSUFFICIENT_STOCK: Kho nguồn [%] (vị trí %) của linh kiện [%] không đủ số lượng để chuyển! Cần %, nhưng tồn kho chỉ còn %.',
            v_from_stage, p_from_location, v_clean_part, p_qty, v_current_stock;
    END IF;

    -- Trừ tồn kho nguồn
    UPDATE public.wip_inventory
    SET quantity = GREATEST(0, quantity - p_qty), updated_at = NOW()
    WHERE (part_id = v_clean_part OR part_id ILIKE (v_clean_part || '%'))
      AND stage_id = v_from_stage 
      AND location = p_from_location;

    v_from_bal_id := v_clean_part || '_' || v_from_stage || '_' || p_from_location;
    INSERT INTO public.wip_balance (id, part_code, stage_id, location, quantity, updated_at)
    VALUES (v_from_bal_id, v_clean_part, v_from_stage, p_from_location, 0, NOW())
    ON CONFLICT (id) DO UPDATE
    SET quantity = GREATEST(0, public.wip_balance.quantity - p_qty), updated_at = NOW();

    SELECT COALESCE(SUM(quantity), 0) INTO v_new_source_stock
    FROM public.wip_inventory
    WHERE (part_id = v_clean_part OR part_id ILIKE (v_clean_part || '%'))
      AND stage_id = v_from_stage 
      AND location = p_from_location;

    -- Cộng tồn kho đích
    v_to_inv_id := v_clean_part || '_' || v_to_stage || '_' || p_to_location || '_';
    INSERT INTO public.wip_inventory (id, part_id, original_part_id, stage_id, location, quantity, updated_at)
    VALUES (v_to_inv_id, v_clean_part, '', v_to_stage, p_to_location, p_qty, NOW())
    ON CONFLICT (id) DO UPDATE
    SET quantity = public.wip_inventory.quantity + p_qty, updated_at = NOW();

    v_to_bal_id := v_clean_part || '_' || v_to_stage || '_' || p_to_location;
    INSERT INTO public.wip_balance (id, part_code, stage_id, location, quantity, updated_at)
    VALUES (v_to_bal_id, v_clean_part, v_to_stage, p_to_location, p_qty, NOW())
    ON CONFLICT (id) DO UPDATE
    SET quantity = public.wip_balance.quantity + p_qty, updated_at = NOW();

    SELECT COALESCE(SUM(quantity), 0) INTO v_new_target_stock
    FROM public.wip_inventory
    WHERE (part_id = v_clean_part OR part_id ILIKE (v_clean_part || '%'))
      AND stage_id = v_to_stage 
      AND location = p_to_location;

    -- Cập nhật nhãn QR
    IF p_qr_id IS NOT NULL AND TRIM(p_qr_id) <> '' THEN
        INSERT INTO public.qr_labels (id, part_id, stage_id, quantity, timestamp, qr_data, metadata, created_at)
        VALUES (
            p_qr_id, v_clean_part, v_to_stage, p_qty, v_timestamp_ms, p_qr_id,
            jsonb_build_object(
                'status', 'TRANSFERRED',
                'from_stage', v_from_stage,
                'to_stage', v_to_stage,
                'transferred_at', to_char(NOW(), 'YYYY-MM-DD"T"HH24:MI:SS.MS"Z"'),
                'operator', p_operator,
                'po_id', p_po_id
            ),
            NOW()
        )
        ON CONFLICT (id) DO UPDATE
        SET stage_id = v_to_stage,
            metadata = jsonb_set(
                jsonb_set(
                    jsonb_set(
                        COALESCE(public.qr_labels.metadata, '{}'::jsonb),
                        '{status}', '"TRANSFERRED"'
                    ),
                    '{transferred_to_stage}', to_jsonb(v_to_stage)
                ),
                '{transferred_at}', to_jsonb(to_char(NOW(), 'YYYY-MM-DD"T"HH24:MI:SS.MS"Z"'))
            );
    END IF;

    -- Ghi nhật ký giao dịch
    INSERT INTO public.wip_transactions (
        id, type, part_id, quantity, stage_id, location,
        timestamp, qr_data, source_stage_id, target_stage_id, po_id, created_at
    ) VALUES (
        v_tx_id, 'STAGE_IN', v_clean_part, p_qty, v_to_stage, p_to_location,
        v_timestamp_ms, p_qr_id, v_from_stage, v_to_stage, p_po_id, NOW()
    );

    RETURN jsonb_build_object(
        'success', TRUE,
        'message', 'Chuyển kho nguyên tử thành công: đã trừ kho nguồn và cộng kho đích.',
        'transaction_id', v_tx_id,
        'qr_id', p_qr_id,
        'part_code', v_clean_part,
        'from_stage', v_from_stage,
        'to_stage', v_to_stage,
        'from_location', p_from_location,
        'to_location', p_to_location,
        'transferred_quantity', p_qty,
        'source_remaining_stock', v_new_source_stock,
        'target_new_stock', v_new_target_stock,
        'timestamp', v_timestamp_ms
    );
EXCEPTION WHEN OTHERS THEN
    RAISE;
END;
$$ LANGUAGE plpgsql;
GRANT EXECUTE ON FUNCTION public.transfer_wip_stage_atomic(TEXT, TEXT, TEXT, TEXT, NUMERIC, TEXT, TEXT, TEXT, TEXT) TO anon, authenticated, service_role;

-- 9. Stored Procedure (RPC): export_wip_stage_atomic
CREATE OR REPLACE FUNCTION public.export_wip_stage_atomic(
    p_part_code TEXT,
    p_stage TEXT,
    p_qty NUMERIC,
    p_source_location TEXT DEFAULT 'OUT',
    p_target_stage TEXT DEFAULT NULL,
    p_po_id TEXT DEFAULT NULL,
    p_qr_id TEXT DEFAULT NULL,
    p_qr_data TEXT DEFAULT NULL,
    p_operator TEXT DEFAULT 'SYSTEM'
)
RETURNS JSONB AS $$
DECLARE
    v_stage TEXT;
    v_clean_part TEXT;
    v_current_stock NUMERIC := 0;
    v_new_stock NUMERIC := 0;
    v_tx_id TEXT;
    v_timestamp_ms BIGINT;
BEGIN
    IF p_qty IS NULL OR p_qty <= 0 THEN RAISE EXCEPTION 'INVALID_QUANTITY: Số lượng xuất phải lớn hơn 0'; END IF;
    v_stage := public.normalize_stage_code(p_stage);
    v_clean_part := TRIM(p_part_code);
    v_timestamp_ms := (EXTRACT(EPOCH FROM NOW()) * 1000)::BIGINT;
    v_tx_id := COALESCE(p_qr_id, 'EX-' || UPPER(SUBSTRING(MD5(RANDOM()::TEXT || NOW()::TEXT) FROM 1 FOR 8)));

    SELECT COALESCE(SUM(quantity), 0) INTO v_current_stock
    FROM public.wip_inventory
    WHERE (part_id = v_clean_part OR part_id ILIKE (v_clean_part || '%')) AND stage_id = v_stage AND location = p_source_location
    FOR UPDATE;

    IF v_current_stock = 0 THEN
        SELECT COALESCE(SUM(quantity), 0) INTO v_current_stock
        FROM public.wip_balance
        WHERE (part_code = v_clean_part OR part_code ILIKE (v_clean_part || '%')) AND stage_id = v_stage AND location = p_source_location
        FOR UPDATE;
    END IF;

    IF v_current_stock < p_qty THEN
        RAISE EXCEPTION 'INSUFFICIENT_STOCK: Tồn kho tại công đoạn [%] (vị trí %) của [%] không đủ! Hiện có %, yêu cầu xuất %',
            v_stage, p_source_location, v_clean_part, v_current_stock, p_qty;
    END IF;

    UPDATE public.wip_inventory
    SET quantity = GREATEST(0, quantity - p_qty), updated_at = NOW()
    WHERE (part_id = v_clean_part OR part_id ILIKE (v_clean_part || '%')) AND stage_id = v_stage AND location = p_source_location;

    UPDATE public.wip_balance
    SET quantity = GREATEST(0, quantity - p_qty), updated_at = NOW()
    WHERE (part_code = v_clean_part OR part_code ILIKE (v_clean_part || '%')) AND stage_id = v_stage AND location = p_source_location;

    SELECT COALESCE(SUM(quantity), 0) INTO v_new_stock
    FROM public.wip_inventory
    WHERE (part_id = v_clean_part OR part_id ILIKE (v_clean_part || '%')) AND stage_id = v_stage AND location = p_source_location;

    INSERT INTO public.qr_labels (id, part_id, stage_id, quantity, timestamp, qr_data, metadata, created_at)
    VALUES (
        v_tx_id, v_clean_part, v_stage, p_qty, v_timestamp_ms, COALESCE(p_qr_data, v_tx_id),
        jsonb_build_object(
            'status', 'ACTIVE', 'exported_from_stage', v_stage, 'target_stage', p_target_stage,
            'source_location', p_source_location, 'po_id', p_po_id, 'operator', p_operator
        ),
        NOW()
    )
    ON CONFLICT (id) DO UPDATE SET stage_id = v_stage, quantity = p_qty, qr_data = COALESCE(p_qr_data, EXCLUDED.qr_data);

    INSERT INTO public.wip_transactions (
        id, type, part_id, quantity, stage_id, location, timestamp, qr_data, target_stage_id, po_id, created_at
    ) VALUES (
        v_tx_id, 'STAGE_OUT', v_clean_part, p_qty, v_stage, p_source_location, v_timestamp_ms, COALESCE(p_qr_data, v_tx_id), p_target_stage, p_po_id, NOW()
    );

    IF p_po_id IS NOT NULL AND TRIM(p_po_id) <> '' THEN
        UPDATE public.production_orders
        SET exported_quantity = COALESCE(exported_quantity, 0) + p_qty, updated_at = NOW()
        WHERE id = p_po_id;
    END IF;

    RETURN jsonb_build_object(
        'success', TRUE,
        'message', 'Xuất kho và tạo mã QR nguyên tử thành công',
        'transaction_id', v_tx_id,
        'part_code', v_clean_part,
        'stage', v_stage,
        'location', p_source_location,
        'exported_quantity', p_qty,
        'remaining_stock', v_new_stock,
        'timestamp', v_timestamp_ms
    );
EXCEPTION WHEN OTHERS THEN
    RAISE;
END;
$$ LANGUAGE plpgsql;
GRANT EXECUTE ON FUNCTION public.export_wip_stage_atomic(TEXT, TEXT, NUMERIC, TEXT, TEXT, TEXT, TEXT, TEXT, TEXT) TO anon, authenticated, service_role;
`;

export interface TransferAtomicParams {
  qrId: string;
  partCode: string;
  fromStage: string;
  toStage: string;
  quantity: number;
  fromLocation?: string;
  toLocation?: string;
  poId?: string;
  operator?: string;
}

export interface ExportAtomicParams {
  partCode: string;
  stage: string;
  quantity: number;
  sourceLocation?: string;
  targetStage?: string;
  poId?: string;
  qrId?: string;
  qrData?: string;
  operator?: string;
}

/**
 * Client RPC Caller: transfer_wip_stage_atomic
 * Executes stage transfer atomically in a single Postgres transaction.
 */
export async function transferWipStageAtomicRpc(params: TransferAtomicParams): Promise<{
  success: boolean;
  data?: any;
  error?: string;
  rpcMissing?: boolean;
}> {
  if (!isSupabaseConfigured) {
    return { success: false, error: 'Supabase chưa được cấu hình' };
  }

  try {
    const { data, error } = await supabase.rpc('transfer_wip_stage_atomic', {
      p_qr_id: params.qrId,
      p_part_code: params.partCode,
      p_from_stage: params.fromStage,
      p_to_stage: params.toStage,
      p_qty: params.quantity,
      p_from_location: params.fromLocation || 'OUT',
      p_to_location: params.toLocation || 'IN',
      p_po_id: params.poId || null,
      p_operator: params.operator || 'SYSTEM'
    });

    if (error) {
      const isMissing = error.code === 'PGRST202' || error.message?.includes('not found') || error.message?.includes('schema cache');
      return { success: false, error: error.message, rpcMissing: isMissing };
    }

    return { success: true, data };
  } catch (err: any) {
    return { success: false, error: err?.message || String(err) };
  }
}

/**
 * Client RPC Caller: export_wip_stage_atomic
 * Atomically deducts inventory at source stage upon QR code generation / stage OUT.
 */
export async function exportWipStageAtomicRpc(params: ExportAtomicParams): Promise<{
  success: boolean;
  data?: any;
  error?: string;
  rpcMissing?: boolean;
}> {
  if (!isSupabaseConfigured) {
    return { success: false, error: 'Supabase chưa được cấu hình' };
  }

  try {
    const { data, error } = await supabase.rpc('export_wip_stage_atomic', {
      p_part_code: params.partCode,
      p_stage: params.stage,
      p_qty: params.quantity,
      p_source_location: params.sourceLocation || 'OUT',
      p_target_stage: params.targetStage || null,
      p_po_id: params.poId || null,
      p_qr_id: params.qrId || null,
      p_qr_data: params.qrData || null,
      p_operator: params.operator || 'SYSTEM'
    });

    if (error) {
      const isMissing = error.code === 'PGRST202' || error.message?.includes('not found') || error.message?.includes('schema cache');
      return { success: false, error: error.message, rpcMissing: isMissing };
    }

    return { success: true, data };
  } catch (err: any) {
    return { success: false, error: err?.message || String(err) };
  }
}

