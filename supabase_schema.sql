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

-- ==============================================================================
-- 7. Table: wip_balance (Cân đối tồn kho thời gian thực đồng bộ)
-- ==============================================================================
CREATE TABLE IF NOT EXISTS public.wip_balance (
    id VARCHAR(255) PRIMARY KEY, -- part_code || '_' || stage_id || '_' || location
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

DROP TRIGGER IF EXISTS tr_wip_balance_updated_at ON public.wip_balance;
CREATE TRIGGER tr_wip_balance_updated_at BEFORE UPDATE ON public.wip_balance FOR EACH ROW EXECUTE FUNCTION public.handle_updated_at();

-- ==============================================================================
-- Helper Function: normalize_stage_code
-- Chuẩn hóa tên công đoạn tiếng Việt ('Sơn', 'Lắp ráp (DCLR)', 'Hàn', ...) sang mã chuẩn
-- ==============================================================================
CREATE OR REPLACE FUNCTION public.normalize_stage_code(p_stage TEXT)
RETURNS TEXT AS $$
DECLARE
    v_clean TEXT;
BEGIN
    IF p_stage IS NULL THEN
        RETURN 'UNKNOWN';
    END IF;
    v_clean := UPPER(TRIM(p_stage));
    
    -- Sơn / Kho Sơn
    IF v_clean LIKE '%SƠN%' OR v_clean LIKE '%SON%' OR v_clean = 'PAINTING' OR v_clean = 'KHO SƠN' OR v_clean = 'KHO_SON' THEN
        RETURN 'PAINTING';
    END IF;
    
    -- Lắp ráp (DCLR)
    IF v_clean LIKE '%DCLR%' OR v_clean LIKE '%LẮP RÁP%' OR v_clean LIKE '%LAP RAP%' THEN
        RETURN 'DCLR';
    END IF;
    
    -- Cắt Laser
    IF v_clean LIKE '%LASER%' OR v_clean LIKE '%CẮT%' OR v_clean LIKE '%CAT%' THEN
        RETURN 'LASER';
    END IF;
    
    -- Chấn / Dập
    IF v_clean LIKE '%CHẤN%' OR v_clean LIKE '%CHAN%' OR v_clean LIKE '%DẬP%' OR v_clean LIKE '%DAP%' OR v_clean = 'BENDING' THEN
        RETURN 'BENDING';
    END IF;
    
    -- Hàn
    IF v_clean LIKE '%HÀN%' OR v_clean LIKE '%HAN%' OR v_clean = 'WELDING' THEN
        RETURN 'WELDING';
    END IF;
    
    -- Dán Kính
    IF v_clean LIKE '%DÁN KÍNH%' OR v_clean LIKE '%DAN KINH%' OR v_clean LIKE '%KÍNH%' OR v_clean = 'GLAZING' THEN
        RETURN 'GLAZING';
    END IF;

    RETURN v_clean;
END;
$$ LANGUAGE plpgsql IMMUTABLE;

GRANT EXECUTE ON FUNCTION public.normalize_stage_code(TEXT) TO anon, authenticated, service_role;

-- ==============================================================================
-- 8. Stored Procedure (RPC): transfer_wip_stage_atomic
-- Xử lý toàn bộ quá trình chuyển kho (Kho Sơn -> Lắp ráp DCLR hoặc các công đoạn)
-- trong đúng 1 Transaction duy nhất:
-- a. Khóa dòng kiểm tra tồn kho nguồn (FOR UPDATE)
-- b. Trừ tồn kho nguồn (wip_inventory & wip_balance)
-- c. Cộng tồn kho đích (wip_inventory & wip_balance)
-- d. Cập nhật nhãn QR (qr_labels) sang trạng thái TRANSFERRED
-- e. Ghi nhật ký wip_transactions
-- f. Trả về JSON thành công hoặc Rollback toàn bộ nếu có lỗi
-- ==============================================================================
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
    v_from_inv_id TEXT;
    v_to_inv_id TEXT;
    v_from_bal_id TEXT;
    v_to_bal_id TEXT;
    v_new_source_stock NUMERIC := 0;
    v_new_target_stock NUMERIC := 0;
    v_tx_id TEXT;
    v_timestamp_ms BIGINT;
    v_label_row RECORD;
BEGIN
    -- 1. Validate đầu vào
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

    -- 2. Khóa dòng kiểm tra trạng thái nhãn QR trong qr_labels (FOR UPDATE)
    IF p_qr_id IS NOT NULL AND TRIM(p_qr_id) <> '' THEN
        SELECT * INTO v_label_row 
        FROM public.qr_labels 
        WHERE id = p_qr_id OR qr_data LIKE ('%' || p_qr_id || '%')
        FOR UPDATE;

        IF FOUND THEN
            -- Kiểm tra xem nhãn này đã được quét chuyển kho trước đó chưa
            IF (v_label_row.metadata->>'status') = 'TRANSFERRED' OR (v_label_row.metadata->>'status') = 'COMPLETED' THEN
                RAISE EXCEPTION 'QR_ALREADY_USED: Mã QR [%] đã được quét chuyển kho trước đó vào lúc %', 
                    p_qr_id, COALESCE(v_label_row.metadata->>'transferred_at', 'trước đó');
            END IF;
        END IF;
    END IF;

    -- 3. Khóa và kiểm tra tồn kho nguồn (Kho Sơn / v_from_stage tại p_from_location)
    SELECT COALESCE(SUM(quantity), 0) INTO v_current_stock
    FROM public.wip_inventory
    WHERE (part_id = v_clean_part OR part_id ILIKE (v_clean_part || '%'))
      AND stage_id = v_from_stage 
      AND location = p_from_location
    FOR UPDATE;

    -- Nếu wip_inventory không có, kiểm tra thêm trong wip_balance
    IF v_current_stock = 0 THEN
        SELECT COALESCE(SUM(quantity), 0) INTO v_current_stock
        FROM public.wip_balance
        WHERE (part_code = v_clean_part OR part_code ILIKE (v_clean_part || '%'))
          AND stage_id = v_from_stage 
          AND location = p_from_location
        FOR UPDATE;
    END IF;

    -- KIỂM TRA TỒN KHO NGUỒN: Nếu không đủ -> Báo lỗi & TỰ ĐỘNG ROLLBACK toàn bộ transaction!
    IF v_current_stock < p_qty THEN
        RAISE EXCEPTION 'INSUFFICIENT_STOCK: Kho nguồn [%] (vị trí %) của linh kiện [%] không đủ số lượng để chuyển! Cần %, nhưng tồn kho chỉ còn %.',
            v_from_stage, p_from_location, v_clean_part, p_qty, v_current_stock;
    END IF;

    -- 4. TRỪ TỒN KHO NGUỒN (ATOMIC DEDUCTION)
    -- Trừ trong wip_inventory
    UPDATE public.wip_inventory
    SET quantity = GREATEST(0, quantity - p_qty),
        updated_at = NOW()
    WHERE (part_id = v_clean_part OR part_id ILIKE (v_clean_part || '%'))
      AND stage_id = v_from_stage 
      AND location = p_from_location;

    -- Trừ trong wip_balance
    v_from_bal_id := v_clean_part || '_' || v_from_stage || '_' || p_from_location;
    INSERT INTO public.wip_balance (id, part_code, stage_id, location, quantity, updated_at)
    VALUES (v_from_bal_id, v_clean_part, v_from_stage, p_from_location, 0, NOW())
    ON CONFLICT (id) DO UPDATE
    SET quantity = GREATEST(0, public.wip_balance.quantity - p_qty),
        updated_at = NOW();

    -- Lấy số lượng tồn kho còn lại ở nguồn
    SELECT COALESCE(SUM(quantity), 0) INTO v_new_source_stock
    FROM public.wip_inventory
    WHERE (part_id = v_clean_part OR part_id ILIKE (v_clean_part || '%'))
      AND stage_id = v_from_stage 
      AND location = p_from_location;

    -- 5. CỘNG HOẶC TẠO MỚI TỒN KHO ĐÍCH (ATOMIC ADDITION VÀO DCLR / v_to_stage)
    -- Cộng vào wip_inventory
    v_to_inv_id := v_clean_part || '_' || v_to_stage || '_' || p_to_location || '_';
    INSERT INTO public.wip_inventory (id, part_id, original_part_id, stage_id, location, quantity, updated_at)
    VALUES (v_to_inv_id, v_clean_part, '', v_to_stage, p_to_location, p_qty, NOW())
    ON CONFLICT (id) DO UPDATE
    SET quantity = public.wip_inventory.quantity + p_qty,
        updated_at = NOW();

    -- Cộng vào wip_balance
    v_to_bal_id := v_clean_part || '_' || v_to_stage || '_' || p_to_location;
    INSERT INTO public.wip_balance (id, part_code, stage_id, location, quantity, updated_at)
    VALUES (v_to_bal_id, v_clean_part, v_to_stage, p_to_location, p_qty, NOW())
    ON CONFLICT (id) DO UPDATE
    SET quantity = public.wip_balance.quantity + p_qty,
        updated_at = NOW();

    -- Lấy số lượng tồn kho mới ở đích
    SELECT COALESCE(SUM(quantity), 0) INTO v_new_target_stock
    FROM public.wip_inventory
    WHERE (part_id = v_clean_part OR part_id ILIKE (v_clean_part || '%'))
      AND stage_id = v_to_stage 
      AND location = p_to_location;

    -- 6. CẬP NHẬT TRẠNG THÁI NHÃN QR TRONG qr_labels
    IF p_qr_id IS NOT NULL AND TRIM(p_qr_id) <> '' THEN
        INSERT INTO public.qr_labels (id, part_id, stage_id, quantity, timestamp, qr_data, metadata, created_at)
        VALUES (
            p_qr_id,
            v_clean_part,
            v_to_stage,
            p_qty,
            v_timestamp_ms,
            p_qr_id,
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

    -- 7. GHI NHẬN LỊCH SỬ GIAO DỊCH (wip_transactions)
    INSERT INTO public.wip_transactions (
        id, type, part_id, quantity, stage_id, location,
        timestamp, qr_data, source_stage_id, target_stage_id,
        po_id, created_at
    ) VALUES (
        v_tx_id,
        'STAGE_IN',
        v_clean_part,
        p_qty,
        v_to_stage,
        p_to_location,
        v_timestamp_ms,
        p_qr_id,
        v_from_stage,
        v_to_stage,
        p_po_id,
        NOW()
    );

    -- 8. TRẢ VỀ KẾT QUẢ JSON THÀNH CÔNG NGUYÊN TỬ
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
    -- Khi có lỗi, rollback toàn bộ
    RAISE;
END;
$$ LANGUAGE plpgsql;

GRANT EXECUTE ON FUNCTION public.transfer_wip_stage_atomic(TEXT, TEXT, TEXT, TEXT, NUMERIC, TEXT, TEXT, TEXT, TEXT) TO anon, authenticated, service_role;

-- ==============================================================================
-- 9. Stored Procedure (RPC): export_wip_stage_atomic
-- Trừ kho nguồn ngay khi xuất tạo mã QR OUT để giải quyết triệt để lỗi
-- "xuất tạo mã QR rồi nhưng tồn ở kho xuất vẫn còn"
-- ==============================================================================
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
    IF p_qty IS NULL OR p_qty <= 0 THEN
        RAISE EXCEPTION 'INVALID_QUANTITY: Số lượng xuất phải lớn hơn 0';
    END IF;

    v_stage := public.normalize_stage_code(p_stage);
    v_clean_part := TRIM(p_part_code);
    v_timestamp_ms := (EXTRACT(EPOCH FROM NOW()) * 1000)::BIGINT;
    v_tx_id := COALESCE(p_qr_id, 'EX-' || UPPER(SUBSTRING(MD5(RANDOM()::TEXT || NOW()::TEXT) FROM 1 FOR 8)));

    -- 1. Kiểm tra tồn kho tại vị trí xuất
    SELECT COALESCE(SUM(quantity), 0) INTO v_current_stock
    FROM public.wip_inventory
    WHERE (part_id = v_clean_part OR part_id ILIKE (v_clean_part || '%')) 
      AND stage_id = v_stage 
      AND location = p_source_location
    FOR UPDATE;

    IF v_current_stock = 0 THEN
        SELECT COALESCE(SUM(quantity), 0) INTO v_current_stock
        FROM public.wip_balance
        WHERE (part_code = v_clean_part OR part_code ILIKE (v_clean_part || '%')) 
          AND stage_id = v_stage 
          AND location = p_source_location
        FOR UPDATE;
    END IF;

    IF v_current_stock < p_qty THEN
        RAISE EXCEPTION 'INSUFFICIENT_STOCK: Tồn kho tại công đoạn [%] (vị trí %) của [%] không đủ! Hiện có %, yêu cầu xuất %',
            v_stage, p_source_location, v_clean_part, v_current_stock, p_qty;
    END IF;

    -- 2. Trừ tồn kho tại vị trí xuất
    UPDATE public.wip_inventory
    SET quantity = GREATEST(0, quantity - p_qty), updated_at = NOW()
    WHERE (part_id = v_clean_part OR part_id ILIKE (v_clean_part || '%')) 
      AND stage_id = v_stage 
      AND location = p_source_location;

    UPDATE public.wip_balance
    SET quantity = GREATEST(0, quantity - p_qty), updated_at = NOW()
    WHERE (part_code = v_clean_part OR part_code ILIKE (v_clean_part || '%')) 
      AND stage_id = v_stage 
      AND location = p_source_location;

    SELECT COALESCE(SUM(quantity), 0) INTO v_new_stock
    FROM public.wip_inventory
    WHERE (part_id = v_clean_part OR part_id ILIKE (v_clean_part || '%')) 
      AND stage_id = v_stage 
      AND location = p_source_location;

    -- 3. Tạo nhãn QR trong qr_labels
    INSERT INTO public.qr_labels (id, part_id, stage_id, quantity, timestamp, qr_data, metadata, created_at)
    VALUES (
        v_tx_id,
        v_clean_part,
        v_stage,
        p_qty,
        v_timestamp_ms,
        COALESCE(p_qr_data, v_tx_id),
        jsonb_build_object(
            'status', 'ACTIVE',
            'exported_from_stage', v_stage,
            'target_stage', p_target_stage,
            'source_location', p_source_location,
            'po_id', p_po_id,
            'operator', p_operator
        ),
        NOW()
    )
    ON CONFLICT (id) DO UPDATE
    SET stage_id = v_stage,
        quantity = p_qty,
        qr_data = COALESCE(p_qr_data, EXCLUDED.qr_data);

    -- 4. Ghi giao dịch STAGE_OUT
    INSERT INTO public.wip_transactions (
        id, type, part_id, quantity, stage_id, location,
        timestamp, qr_data, target_stage_id, po_id, created_at
    ) VALUES (
        v_tx_id,
        'STAGE_OUT',
        v_clean_part,
        p_qty,
        v_stage,
        p_source_location,
        v_timestamp_ms,
        COALESCE(p_qr_data, v_tx_id),
        p_target_stage,
        p_po_id,
        NOW()
    );

    -- 5. Cập nhật số lượng exportedQuantity trong production_orders nếu có PO
    IF p_po_id IS NOT NULL AND TRIM(p_po_id) <> '' THEN
        UPDATE public.production_orders
        SET exported_quantity = COALESCE(exported_quantity, 0) + p_qty,
            updated_at = NOW()
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

