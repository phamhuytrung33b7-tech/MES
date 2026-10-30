/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import * as XLSX from 'xlsx-js-style';
import { 
  LayoutDashboard, 
  PackagePlus, 
  QrCode, 
  History, 
  Settings,
  Edit2,
  Trash2,
  ArrowRight,
  ArrowUpRight,
  ArrowDownLeft,
  Boxes,
  CheckCircle2, 
  AlertCircle,
  Printer,
  Sparkles,
  X,
  RotateCcw,
  ChevronRight,
  ChevronDown,
  Search,
  Package,
  Monitor,
  Menu,
  FileUp,
  Upload,
  Layers,
  Flame,
  FileText,
  ClipboardList,
  Clock,
  Save,
  Plus,
  Users,
  FileSpreadsheet,
  FileDown,
  Square,
  Check,
  Download,
  Database,
  HardDrive,
  RefreshCw,
  Copy,
  FileJson,
  Calendar,
  Cloud,
  CloudOff,
  Server,
  Key,
  Lock,
  LogOut,
  ShieldCheck,
  UserCheck,
  Filter,
  Camera,
  Scissors,
  Scale,
  Percent
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { QRCodeSVG } from 'qrcode.react';
import { format, startOfDay } from 'date-fns';
import { 
  BarChart, 
  ComposedChart,
  Line,
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  Legend, 
  ResponsiveContainer,
  Cell,
  ReferenceLine
} from 'recharts';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

import { STAGES, INITIAL_PARTS, StageId, Part, InventoryItem, Transaction, BOMDefinition, BOMDefinitionV2, ProductionOrder, ModelBOMDefinition, ProductivityNorm, LaserNesting, ShiftConfig, PartTransformation, DEFECT_REASONS, StagePlannedWorkers, UserAccount, View, getUserDefaultStage, DeletedPOBackup, POSnapshot } from './types';
import { storageService, HOURLY_DAY_SLOTS, HOURLY_NIGHT_SLOTS, ALL_HOURLY_TIME_SLOTS } from './storage';
import { supabase, SUPABASE_URL, SUPABASE_SQL_SCHEMA, testSupabaseConnection, isSupabaseConfigured } from './supabaseClient';
import { exportPreparationReport } from './exportPreparationReport';
import { exportDetailedProductionPlan } from './exportDetailedPlan';
import { importExcelToPO, downloadSamplePOExcelTemplate, MissingBomModelInfo } from './importExcelToPO';
import { downloadNormTemplate } from './downloadNormTemplate';
import { MobileQRScannerModal } from './MobileQRScanner';
import { importExcelToGlazingPlan, MissingNormModelInfo } from './importExcelToGlazingPlan';
import { UserManagementView, LoginModal, ChangePasswordModal, AdminActionPasswordModal } from './UserManagement';
import { smartFuzzySearch, normalizeVietnamese, buildSearchIndex, searchIndexedItems } from './search';

// Utility for tailwind classes
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

export function isLaserScrapItem(partId?: string, partName?: string): boolean {
  if (!partId) return false;
  const pId = partId.trim().toUpperCase();
  if (pId === 'PL-TON-SX' || pId === 'PL-TON' || pId.startsWith('PL-TON')) return true;
  const pName = (partName || '').toLowerCase();
  if (pName.includes('phế liệu tôn') || pName.includes('phế liệu') || pName.includes('ton-sx')) return true;
  return false;
}

function formatInventoryAge(timestamp: number): { text: string; colorClass: string } {
  const diffMs = Date.now() - timestamp;
  const hours = diffMs / (1000 * 60 * 60);
  const days = Math.floor(hours / 24);
  
  let colorClass = "text-gray-500";
  if (days >= 30) {
    colorClass = "text-red-600";
  } else if (days >= 14) {
    colorClass = "text-yellow-600";
  }

  if (hours < 72) {
    return { text: `Tồn ${Math.floor(hours)} giờ`, colorClass };
  } else if (days > 35) {
    return { text: `Tồn > 35 ngày`, colorClass };
  } else {
    return { text: `Tồn ${days} ngày`, colorClass };
  }
}

export interface AgeBatch {
  quantity: number;
  timestamp: number;
  ageInfo: { text: string; colorClass: string };
}

function cleanPartIdStr(s?: string): string {
  if (!s) return '';
  return s.trim().toLowerCase().replace(/\s*-\s*(cd|h)$/i, '');
}

function matchesInventoryItem(t: Transaction, item: InventoryItem): boolean {
  if (t.stageId !== item.stageId) return false;

  const tPart = cleanPartIdStr(t.partId);
  const itemPart = cleanPartIdStr(item.partId);
  const tOrig = cleanPartIdStr(t.originalPartId);
  const itemOrig = cleanPartIdStr(item.originalPartId);

  const partMatch = 
    (tPart && itemPart && tPart === itemPart) ||
    (tOrig && itemOrig && tOrig === itemOrig) ||
    (tPart && itemOrig && tPart === itemOrig) ||
    (tOrig && itemPart && tOrig === itemPart);

  if (!partMatch) return false;

  const tLoc = t.location || 'IN';
  if (item.location === 'IN') {
    return (t.type === 'STAGE_IN' && tLoc === 'IN') || t.type === 'STAGE_IN';
  } else if (item.location === 'OUT') {
    return (t.type === 'STAGE_OUT' && tLoc === 'IN') || (t.type === 'STAGE_IN' && tLoc === 'OUT') || t.type === 'STAGE_OUT';
  } else if (item.location === 'DEFECT' || (item.location as any) === 'NG') {
    return t.type === 'DEFECT';
  }
  return true;
}

function getInventoryBatches(
  item: InventoryItem,
  transactions: Transaction[]
): AgeBatch[] {
  const targetQty = item.quantity;

  if (targetQty <= 0) return [];

  const additions = transactions
    .filter(t => matchesInventoryItem(t, item))
    .sort((a, b) => b.timestamp - a.timestamp);

  if (additions.length === 0) {
    const anyTx = transactions
      .filter(t => {
        const tPart = cleanPartIdStr(t.partId);
        const itemPart = cleanPartIdStr(item.partId);
        const tOrig = cleanPartIdStr(t.originalPartId);
        const itemOrig = cleanPartIdStr(item.originalPartId);
        return (tPart && itemPart && tPart === itemPart) || (tOrig && itemOrig && tOrig === itemOrig);
      })
      .sort((a, b) => a.timestamp - b.timestamp)[0];

    const ts = anyTx ? anyTx.timestamp : (item.lastUpdated || Date.now());
    return [{
      quantity: targetQty,
      timestamp: ts,
      ageInfo: formatInventoryAge(ts)
    }];
  }

  let accumulated = 0;
  const batches: AgeBatch[] = [];

  for (const add of additions) {
    const remaining = targetQty - accumulated;
    if (remaining <= 0) break;
    const qtyToTake = Math.min(add.quantity, remaining);
    batches.push({
      quantity: qtyToTake,
      timestamp: add.timestamp,
      ageInfo: formatInventoryAge(add.timestamp)
    });
    accumulated += qtyToTake;
  }

  if (accumulated < targetQty) {
    const fallbackTs = item.lastUpdated || additions[additions.length - 1].timestamp;
    batches.push({
      quantity: targetQty - accumulated,
      timestamp: fallbackTs,
      ageInfo: formatInventoryAge(fallbackTs)
    });
  }

  const groupedBatches: AgeBatch[] = [];
  for (const b of batches) {
    if (b.ageInfo.text === 'Tồn > 35 ngày') {
      const existing = groupedBatches.find(x => x.ageInfo.text === 'Tồn > 35 ngày');
      if (existing) {
        existing.quantity += b.quantity;
      } else {
        groupedBatches.push(b);
      }
    } else {
      groupedBatches.push(b);
    }
  }

  return groupedBatches;
}

function getProcessValue(value: string | undefined, part: Part | undefined, stageId: StageId, location: 'IN' | 'OUT') {
  if (!value || !part) return value || '';
  
  // Logic for suffixes
  let suffix = '';
  
  if (stageId === 'BENDING' && location === 'OUT') suffix = 'CD';
  
  if (stageId === 'WELDING') {
    if (location === 'IN' && !part.skipBending) suffix = 'CD';
    if (location === 'OUT') suffix = 'H';
  }
  
  if (stageId === 'PAINTING' && location === 'IN') {
    if (!part.skipWelding) suffix = 'H';
    else if (!part.skipBending) suffix = 'CD';
  }

  // Return original if no suffix logic applies (e.g. Painting OUT, DCLR, Laser)
  if (!suffix) return value;
  
  // Consistently use " - [SUFFIX]"
  return `${value} - ${suffix}`;
}

function SearchableSelect({ 
  options, 
  value, 
  onChange, 
  placeholder = "Gõ từ khóa tìm linh kiện / PO...", 
  className = "" 
}: { 
  options: { id: string, label: string }[], 
  value: string, 
  onChange: (val: string) => void,
  placeholder?: string,
  className?: string
}) {
  const [isOpen, setIsOpen] = useState(false);
  const [search, setSearch] = useState("");
  const [debouncedSearch, setDebouncedSearch] = useState("");
  const [highlightedIndex, setHighlightedIndex] = useState(0);
  const containerRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  // 1. Debounce mechanism (200ms delay) to prevent UI lag during typing
  useEffect(() => {
    const timer = setTimeout(() => {
      setDebouncedSearch(search);
    }, 200);
    return () => clearTimeout(timer);
  }, [search]);

  // 2. Pre-index options in RAM memory ONCE when options change (0ms string allocation per keystroke)
  const preindexedOptions = useMemo(() => {
    const searchableItems = options.map(opt => ({
      id: opt.id,
      label: opt.label,
      code: opt.id
    }));
    return buildSearchIndex(searchableItems, ['id', 'label', 'code']);
  }, [options]);

  // 3. Fast search lookup over pre-indexed RAM cache
  const filteredOptions = useMemo(() => {
    if (!debouncedSearch.trim()) return options;
    const matchedItems = searchIndexedItems(debouncedSearch, preindexedOptions);
    return matchedItems.map(item => ({ id: item.id, label: item.label }));
  }, [options, preindexedOptions, debouncedSearch]);

  const selectedOption = useMemo(() => options.find(opt => opt.id === value), [options, value]);
  const displayedOptions = useMemo(() => filteredOptions.slice(0, 80), [filteredOptions]);

  useEffect(() => {
    setHighlightedIndex(0);
  }, [search]);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const handleSelect = (optId: string) => {
    onChange(optId);
    setIsOpen(false);
    setSearch("");
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (!isOpen) {
      if (e.key === 'ArrowDown' || e.key === 'Enter') {
        setIsOpen(true);
      }
      return;
    }

    if (e.key === 'ArrowDown') {
      e.preventDefault();
      setHighlightedIndex(prev => (prev + 1) % Math.max(1, displayedOptions.length));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setHighlightedIndex(prev => (prev - 1 + displayedOptions.length) % Math.max(1, displayedOptions.length));
    } else if (e.key === 'Enter') {
      if (displayedOptions.length > 0) {
        e.preventDefault();
        handleSelect(displayedOptions[highlightedIndex]?.id || displayedOptions[0].id);
      }
    } else if (e.key === 'Escape') {
      setIsOpen(false);
    }
  };

  return (
    <div className={cn("relative text-left", className)} ref={containerRef}>
      <div 
        onClick={() => {
          setIsOpen(!isOpen);
          if (!isOpen) {
            setTimeout(() => inputRef.current?.focus(), 60);
          }
        }}
        className="w-full p-3.5 sm:p-5 rounded-xl sm:rounded-2xl border-2 border-gray-200 hover:border-blue-500 font-bold text-sm sm:text-lg focus:border-blue-600 outline-none bg-white cursor-pointer flex justify-between items-center transition-all shadow-xs active:scale-[0.99]"
      >
        <div className="flex items-center gap-2.5 truncate min-w-0 pr-2">
          <Search size={18} className="text-blue-600 shrink-0" />
          <span className={selectedOption ? "text-gray-900 truncate" : "text-gray-400 font-normal truncate"}>
            {selectedOption ? selectedOption.label : placeholder}
          </span>
        </div>
        <ChevronDown size={20} className={cn("text-gray-400 transition-transform shrink-0", isOpen && "rotate-180")} />
      </div>

      <AnimatePresence>
        {isOpen && (
          <motion.div 
            initial={{ opacity: 0, y: -10, scale: 0.98 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, y: -10, scale: 0.98 }}
            transition={{ duration: 0.15 }}
            className="absolute z-[120] left-0 right-0 mt-2 bg-white border-2 border-blue-200 rounded-2xl shadow-2xl overflow-hidden flex flex-col max-h-[380px] sm:max-h-[420px]"
          >
            {/* Search Input Header */}
            <div className="p-3 border-b border-gray-100 bg-gray-50/90 backdrop-blur-xs sticky top-0 z-10">
              <div className="relative flex items-center">
                <Search size={18} className="absolute left-3.5 text-blue-600 shrink-0" />
                <input 
                  ref={inputRef}
                  autoFocus
                  type="text"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  onKeyDown={handleKeyDown}
                  placeholder="Gõ từ khóa không dấu (VD: 'sha 766', 'tan tu', '001')..."
                  className="w-full pl-10 pr-9 py-3 rounded-xl border-2 border-blue-400 bg-white text-sm sm:text-base font-bold outline-none focus:ring-4 ring-blue-500/20 shadow-xs placeholder:text-gray-400 placeholder:font-normal"
                />
                {search && (
                  <button
                    type="button"
                    onClick={() => setSearch('')}
                    className="absolute right-2.5 p-1 text-gray-400 hover:text-gray-600 rounded-lg hover:bg-gray-100 cursor-pointer"
                  >
                    <X size={16} />
                  </button>
                )}
              </div>
              <div className="flex justify-between items-center mt-2 px-1 text-[11px] font-bold text-gray-500">
                <span>Gợi ý ({filteredOptions.length} kết quả)</span>
                <span className="text-blue-600 font-mono text-[10px]">Tự động sắp xếp theo độ khớp</span>
              </div>
            </div>

            {/* Suggestions List */}
            <div className="overflow-y-auto flex-1 divide-y divide-gray-100">
              {displayedOptions.length > 0 ? (
                displayedOptions.map((opt, idx) => {
                  const isSelected = value === opt.id;
                  const isHighlighted = idx === highlightedIndex;

                  return (
                    <div 
                      key={opt.id + '_' + idx}
                      onClick={() => handleSelect(opt.id)}
                      onMouseEnter={() => setHighlightedIndex(idx)}
                      className={cn(
                        "p-3.5 px-4 cursor-pointer transition-colors flex items-center justify-between gap-3 text-left",
                        isSelected
                          ? "bg-blue-50 text-blue-900 font-bold"
                          : isHighlighted
                            ? "bg-gray-100 text-gray-900 font-semibold"
                            : "hover:bg-gray-50 text-gray-800"
                      )}
                    >
                      <div className="min-w-0 flex-1">
                        <div className="text-xs sm:text-sm font-bold text-gray-900 leading-snug">
                          {opt.label}
                        </div>
                        {opt.id !== opt.label && (
                          <div className="text-[11px] font-mono text-gray-500 mt-0.5">
                            Mã: <span className="font-bold text-blue-700">{opt.id}</span>
                          </div>
                        )}
                      </div>

                      {isSelected && (
                        <div className="w-5 h-5 rounded-full bg-blue-600 text-white flex items-center justify-center shrink-0">
                          <Check size={12} />
                        </div>
                      )}
                    </div>
                  );
                })
              ) : (
                <div className="p-8 text-center text-gray-400 text-xs sm:text-sm font-medium flex flex-col items-center justify-center gap-2">
                  <AlertCircle size={24} className="text-gray-300" />
                  <span>Không tìm thấy kết quả phù hợp với từ khóa "{search}"</span>
                </div>
              )}
              {filteredOptions.length > 80 && (
                <div className="p-2.5 text-center text-[11px] text-gray-500 italic bg-gray-50 border-t border-gray-100 font-medium">
                  Hiển thị 80 / {filteredOptions.length} kết quả khớp nhất. Gõ thêm ký tự để thu hẹp tìm kiếm.
                </div>
              )}
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

type DialogConfig = {
  id: string;
  type: 'alert' | 'confirm' | 'prompt';
  title: string;
  message?: string;
  defaultValue?: string;
  passwordMode?: boolean;
  resolve: (val: any) => void;
};

let dialogListener: ((config: DialogConfig | null) => void) | null = null;
let currentDialogs: DialogConfig[] = [];

export const runDialog = (config: Omit<DialogConfig, 'id' | 'resolve'>): Promise<any> => {
  return new Promise(resolve => {
    const dialog: DialogConfig = {
      ...config,
      id: Math.random().toString(),
      resolve
    };
    currentDialogs.push(dialog);
    if (currentDialogs.length === 1 && dialogListener) {
      dialogListener(currentDialogs[0]);
    }
  });
};

export const customAlert = (message: string) => runDialog({ type: 'alert', title: 'Thông báo', message });
export const customConfirm = (message: string) => runDialog({ type: 'confirm', title: 'Xác nhận', message });
export const customPrompt = (title: string, defaultValue?: string, passwordMode?: boolean) => runDialog({ type: 'prompt', title, defaultValue, passwordMode });

export function GlobalDialogs() {
  const [activeDialog, setActiveDialog] = useState<DialogConfig | null>(null);
  const [inputValue, setInputValue] = useState('');

  useEffect(() => {
    dialogListener = setActiveDialog;
    if (currentDialogs.length > 0) {
       setActiveDialog(currentDialogs[0]);
       setInputValue(currentDialogs[0].defaultValue || '');
    }
    return () => { dialogListener = null; };
  }, []);

  if (!activeDialog) return null;

  const close = (val: any) => {
    activeDialog.resolve(val);
    currentDialogs.shift();
    if (currentDialogs.length > 0) {
       setActiveDialog(currentDialogs[0]);
       setInputValue(currentDialogs[0].defaultValue || '');
    } else {
       setActiveDialog(null);
       setInputValue('');
    }
  };

  return (
    <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-black/60 backdrop-blur-sm">
      <motion.div 
        initial={{ scale: 0.95, opacity: 0 }}
        animate={{ scale: 1, opacity: 1 }}
        exit={{ scale: 0.95, opacity: 0 }}
        className="bg-white rounded-2xl shadow-2xl p-6 max-w-md w-full mx-4 border border-gray-100"
      >
        <h3 className="text-xl font-bold text-gray-900 mb-2">{activeDialog.title}</h3>
        {activeDialog.message && <p className="text-gray-600 mb-6">{activeDialog.message}</p>}
        
        {activeDialog.type === 'prompt' && (
          <div className="mb-6">
            <input 
              autoFocus
              type={activeDialog.passwordMode || activeDialog.title.toLowerCase().includes('mật khẩu') ? 'password' : 'text'}
              value={inputValue}
              onChange={e => setInputValue(e.target.value)}
              className="w-full px-4 py-3 rounded-lg border-2 border-gray-200 focus:border-blue-500 focus:ring-4 focus:ring-blue-500/20 outline-none transition-all font-medium text-lg"
              placeholder="Nhập giá trị..."
              onKeyDown={e => {
                if (e.key === 'Enter') close(inputValue);
                if (e.key === 'Escape') close(null);
              }}
            />
          </div>
        )}

        <div className="flex justify-end gap-3 rounded-b-xl border-gray-100">
          {activeDialog.type !== 'alert' && (
            <button 
              autoFocus={activeDialog.type === 'confirm'}
              onClick={() => close(activeDialog.type === 'prompt' ? null : false)}
              className="px-5 py-2.5 rounded-lg text-gray-600 font-bold hover:bg-gray-100 transition-colors"
            >
              Hủy
            </button>
          )}
          <button 
            onClick={() => close(activeDialog.type === 'prompt' ? inputValue : true)}
            autoFocus={activeDialog.type === 'alert'}
            className="px-5 py-2.5 rounded-lg bg-blue-600 text-white font-bold hover:bg-blue-700 active:scale-95 transition-all shadow-md shadow-blue-600/20"
          >
            {activeDialog.type === 'alert' ? 'Đóng' : 'Xác nhận'}
          </button>
        </div>
      </motion.div>
    </div>
  );
}

const getGlazingHSQD = (pId: string, dclrNormsMap: Map<string, number>, dclrNorms: ProductivityNorm[], parts: Part[]) => {
  const glazingNorms = storageService.getNorms().filter(n => n.stageId === 'GLAZING');
  const upperId = (pId || '').trim().toUpperCase();

  // 1. Kiểm tra trong tab Dán Kính trước
  for (const g of glazingNorms) {
    if (g.partId.trim().toUpperCase() === upperId && g.secondsPerUnit > 0) return g.secondsPerUnit;
  }
  const cleanGlzId = upperId.replace(/^GLZ-OUT-/i, '').replace(/\s*-\s*DÁN KÍNH$/i, '').trim();
  if (cleanGlzId) {
    for (const g of glazingNorms) {
      if (g.partId.trim().toUpperCase() === cleanGlzId && g.secondsPerUnit > 0) return g.secondsPerUnit;
    }
  }
  const matchedGlzPart = parts.find((p: any) => 
    p.id.toUpperCase() === upperId || 
    p.name.toUpperCase().trim() === upperId ||
    (cleanGlzId && (p.name.toUpperCase().trim() === cleanGlzId || p.id.toUpperCase() === cleanGlzId))
  );
  if (matchedGlzPart) {
    for (const g of glazingNorms) {
      if ((g.partId.toUpperCase().trim() === matchedGlzPart.id.toUpperCase().trim() || g.partId.toUpperCase().trim() === matchedGlzPart.name.toUpperCase().trim()) && g.secondsPerUnit > 0) {
        return g.secondsPerUnit;
      }
    }
  }
  for (const g of glazingNorms) {
    const normPart = parts.find((p: any) => p.id === g.partId);
    if (normPart && (normPart.name.toUpperCase().trim() === cleanGlzId || normPart.name.toUpperCase().trim() === upperId) && g.secondsPerUnit > 0) {
      return g.secondsPerUnit;
    }
  }

  // 2. Dự phòng: kiểm tra tab DCLR
  if (dclrNormsMap.has(upperId)) return dclrNormsMap.get(upperId) || 0;
  if (dclrNormsMap.has(pId)) return dclrNormsMap.get(pId) || 0;

  if (pId.startsWith('GLZ-OUT-')) {
     const cleanId = pId.replace('GLZ-OUT-', '').replace(/\s*-\s*Dán kính$/i, '').trim().toUpperCase();
     
     if (dclrNormsMap.has(cleanId)) return dclrNormsMap.get(cleanId) || 0;

     const matchedPart = parts.find((p: any) => 
       p.name.toUpperCase().trim() === cleanId || 
       p.id.toUpperCase() === cleanId
     );
     if (matchedPart && dclrNormsMap.has(matchedPart.id.toUpperCase().trim())) {
       return dclrNormsMap.get(matchedPart.id.toUpperCase().trim()) || 0;
     }

     for (const norm of dclrNorms) {
       const normPart = parts.find((p: any) => p.id === norm.partId);
       if (normPart && (
         normPart.name.toUpperCase().trim() === cleanId || 
         normPart.name.toUpperCase().trim() === pId.toUpperCase().trim()
       )) {
         return norm.secondsPerUnit;
       }
       if (norm.partId.toUpperCase().trim() === pId.toUpperCase().trim()) {
         return norm.secondsPerUnit;
       }
     }
  }
  return 0;
};

function A7QRLabelCard({ 
  label, 
  parts, 
  labelSettings, 
  isPrintArea = false 
}: { 
  label: Transaction; 
  parts: Part[]; 
  labelSettings: any; 
  isPrintArea?: boolean; 
}) {
  if (!label) return null;

  const part = parts.find(p => p.id === label.partId);
  const isDisposal = label.type === 'DISPOSAL' || label.type === 'DEFECT';

  const partNameDisplay = isDisposal 
    ? (label.partName || part?.name || label.partId)
    : getProcessValue(label.partName || part?.name, part, label.stageId, 'OUT');

  const partIdDisplay = isDisposal 
    ? label.partId 
    : getProcessValue(label.partId, part, label.stageId, 'OUT');

  const sourceStageName = STAGES.find(s => s.id === label.stageId)?.name || label.stageId;
  const rawTargetStage = label.targetStageId || label.qrData?.split('|')?.[5];
  const targetStageName = STAGES.find(s => s.id === rawTargetStage)?.name || 
    (rawTargetStage === 'DCLR' ? 'Lắp ráp (DCLR)' : 'HOÀN THÀNH');

  const po = storageService.getProductionOrders().find(p => p.id === label.poId);
  const masterId = po?.masterPoId || label.qrData?.split('|')?.[7];

  // Responsive font size and QR size scaling based on settings (or defaults for A7)
  const baseFontSize = Math.min(Math.max(labelSettings?.fontSize || 10, 7), 14);
  const qrSize = Math.min(Math.max(labelSettings?.qrSize || 110, 70), 135);

  return (
    <div 
      id={isPrintArea ? undefined : "qr-label-display"}
      className={cn(
        isPrintArea 
          ? "w-full h-full bg-white text-black flex flex-col items-center justify-between box-border overflow-hidden"
          : "w-[300px] h-[400px] bg-white border-2 border-black p-2.5 flex flex-col items-center justify-between relative overflow-hidden rounded-xl shadow-xl text-black shrink-0"
      )}
      style={isPrintArea ? {
        width: `${labelSettings?.width || 75}mm`,
        height: `${labelSettings?.height || 100}mm`,
        padding: '2mm 2.5mm',
        boxSizing: 'border-box'
      } : undefined}
    >
      {/* 1. HEADER / BANNER */}
      {isDisposal ? (
        <div className="w-full bg-black text-white text-center py-0.5 font-black text-[10px] uppercase tracking-wider relative z-10 rounded-sm">
          NHÃN XUẤT HỦY {label.type === 'DEFECT' ? '(DEFECT)' : '(DISPOSAL)'}
        </div>
      ) : (
        <div className="w-full text-center border-b border-black pb-0.5 mb-0.5">
          <h1 className="font-black uppercase tracking-widest text-[11px] text-black leading-none">PHIẾU ĐIỀU CHUYỂN</h1>
        </div>
      )}

      {/* 2. QR CODE */}
      <div className="my-0.5 border border-black p-0.5 bg-white inline-block relative z-10 rounded-sm">
        <QRCodeSVG value={label.qrData || ''} size={qrSize} level="H" includeMargin={false} />
      </div>

      {/* 3. PART NAME & PART CODE */}
      <div className="w-full text-center my-0.5 relative z-10 text-black px-0.5">
        <h2 
          className="font-black uppercase leading-snug tracking-tight break-words line-clamp-2 text-black"
          style={{ fontSize: `${baseFontSize + 3}px` }}
        >
          {partNameDisplay}
        </h2>
        <p 
          className="font-mono font-black text-black mt-0.5 break-words"
          style={{ fontSize: `${baseFontSize}px` }}
        >
          Mã LK: {partIdDisplay}
        </p>
      </div>

      {/* 4. QUANTITY & SOURCE STAGE GRID */}
      <div className="w-full grid grid-cols-2 border-t border-b border-black py-0.5 my-0.5 relative z-10 text-black items-center text-center">
        <div className="border-r border-black px-0.5 flex flex-col justify-center">
          <span className="text-[8px] font-black uppercase text-black">Số lượng</span>
          <span className="font-black leading-tight text-black" style={{ fontSize: `${baseFontSize + 4}px` }}>
            {label.quantity} {part?.unit || 'Cái'}
          </span>
        </div>
        <div className="px-0.5 flex flex-col justify-center">
          <span className="text-[8px] font-black uppercase text-black">
            {isDisposal ? 'Từ kho:' : 'Từ công đoạn:'}
          </span>
          <span className="font-black uppercase leading-tight text-black" style={{ fontSize: `${baseFontSize}px` }}>
            {sourceStageName} {isDisposal && '(NG)'}
          </span>
        </div>
      </div>

      {/* 5. DESTINATION BOX OR DEFECT WARNING */}
      {isDisposal ? (
        <div className="w-full border border-black rounded p-0.5 my-0.5 text-center bg-transparent relative z-10 text-black">
          <span className="text-[10px] font-black uppercase block tracking-tighter text-black leading-none">HÀNG LỖI - CẤM NHẬP KHO</span>
          {label.defectReason && (
            <span className="text-[8.5px] font-bold block text-black truncate mt-0.5">Lý do: {label.defectReason}</span>
          )}
        </div>
      ) : (
        <div className="w-full border border-black rounded p-0.5 my-0.5 text-center text-black">
          <span className="text-[8px] font-black uppercase block text-black leading-none mb-0.5">Đích tiếp theo:</span>
          <div className="flex items-center justify-center gap-1.5 font-black text-[10px] italic text-black">
            <span>{sourceStageName}</span>
            <ArrowRight size={12} strokeWidth={3} className="text-black inline shrink-0" />
            <span>{targetStageName}</span>
          </div>
        </div>
      )}

      {/* 6. GLAZING PLAN TIMES (IF ANY) */}
      {(label as any).planId && (() => {
        const plan = storageService.getGlazingPlans().find(p => p.id === (label as any).planId);
        if (plan) {
          return (
            <div className="w-full flex justify-between font-mono font-bold text-black border-t border-b border-black py-0.5 my-0.5 text-[8px]">
               <div className="flex flex-col text-left">
                 <span className="uppercase font-black text-[7.5px]">HT Dự Kiến</span>
                 <span>{plan.expectedCompletionTime ? format(plan.expectedCompletionTime, 'HH:mm dd/MM') : '--:--'}</span>
               </div>
               <div className="flex flex-col text-right">
                 <span className="uppercase font-black text-[7.5px]">HT Thực Tế</span>
                 <span>{format(label.timestamp, 'HH:mm dd/MM')}</span>
               </div>
            </div>
          );
        }
        return null;
      })()}

      {/* 7. PO DETAILS SECTION */}
      <div 
        className="w-full space-y-[1.5px] font-mono font-black border-2 border-black p-1 rounded-md text-black leading-tight my-0.5"
        style={{ fontSize: `${baseFontSize - 1}px` }}
      >
        <div className="flex justify-between items-center">
          <span className="uppercase font-black text-black">LOẠI PO:</span>
          <span className="uppercase font-black text-black">
            {po?.masterPoId ? 'PO Con (Sub)' : 'PO Tổng (Master)'}
          </span>
        </div>
        <div className="flex justify-between items-center">
          <span className="uppercase font-black text-black">MÃ PO:</span>
          <span className="font-mono font-black text-black truncate max-w-[170px]">{label.poId || 'N/A'}</span>
        </div>
        <div className="flex justify-between items-center font-black text-black">
          <span className="font-black text-black">KH PO Con:</span>
          <span className="font-black text-black">{label.qrData?.split('|')?.[8] || po?.targetQuantity || 0} linh kiện</span>
        </div>
        {po?.plannedStartTime && (
          <div className="flex justify-between items-center pt-[1px] border-t border-black">
            <span className="uppercase font-black text-black">KH Bắt đầu PO:</span>
            <span className="font-mono font-black text-black">{format(po.plannedStartTime, 'dd/MM HH:mm')}</span>
          </div>
        )}
        {po?.expectedCompletionTime && (
          <div className="flex justify-between items-center">
            <span className="uppercase font-black text-black">KH Kết thúc PO:</span>
            <span className="font-mono font-black text-black">{format(po.expectedCompletionTime, 'dd/MM HH:mm')}</span>
          </div>
        )}
        {masterId && (
          <>
            <div className="flex justify-between items-center pt-[1px] border-t border-black">
              <span className="uppercase font-black text-black">Thuộc PO Tổng:</span>
              <span className="font-mono font-black text-black truncate max-w-[150px]">{masterId}</span>
            </div>
            <div className="flex justify-between items-center font-black text-black">
              <span className="font-black text-black">KH PO Tổng:</span>
              <span className="font-black text-black">{label.qrData?.split('|')?.[9] || storageService.getProductionOrders().find(p => p.id === masterId)?.targetQuantity || 0} máy</span>
            </div>
          </>
        )}
        <div className="flex justify-between items-center pt-[1px] border-t border-black">
          <span className="uppercase font-black text-black">Hoàn thành:</span>
          <span className="font-mono font-black text-black">{format(label.timestamp, 'dd/MM/yyyy HH:mm:ss')}</span>
        </div>
      </div>
    </div>
  );
}

export default function App() {
  const [currentView, setCurrentView] = useState<View>('dashboard');
  const [inventory, setInventory] = useState<InventoryItem[]>([]);
  const [transactions, setTransactions] = useState<Transaction[]>([]);
  const [labels, setLabels] = useState<Transaction[]>([]);
  const [parts, setParts] = useState<Part[]>([]);
  const [productionOrders, setProductionOrders] = useState<ProductionOrder[]>([]);
  const [selectedStage, setSelectedStage] = useState<StageId>(() => getUserDefaultStage(storageService.getCurrentUser()));
  const [scanStage, setScanStage] = useState<StageId>(() => getUserDefaultStage(storageService.getCurrentUser()));
  const [selectedPart, setSelectedPart] = useState<string>('');
  const [quantity, setQuantity] = useState<number>(0);
  const [lastTransaction, setLastTransaction] = useState<Transaction | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);
  const [isImportingNorms, setIsImportingNorms] = useState(false);
  const [syncStatus, setSyncStatus] = useState(() => storageService.getSyncStatus());
  const [showSupabaseModal, setShowSupabaseModal] = useState(false);

  const [labelSettings, setLabelSettings] = useState(storageService.getLabelSettings());
  const [defectModal, setDefectModal] = useState<{ partId: string, stageId: StageId, poId?: string } | null>(null);

  // Authentication & RBAC State
  const [currentUser, setCurrentUser] = useState<UserAccount | null>(() => storageService.getCurrentUser());
  const [showLoginModal, setShowLoginModal] = useState(false);
  const [showChangePasswordModal, setShowChangePasswordModal] = useState(false);
  const [showUserDropdown, setShowUserDropdown] = useState(false);

  // RBAC Permission check
  const canAccessView = useCallback((v: View): boolean => {
    if (!currentUser) return false;
    if (currentUser.role === 'ADMIN') return true;
    if (v === 'users') return false; // Admin only
    return currentUser.allowedViews?.includes(v) ?? false;
  }, [currentUser]);

  // Guard: if currentView is not allowed for the user, redirect to their first permitted tab
  useEffect(() => {
    if (currentUser && !canAccessView(currentView)) {
      const allowed = currentUser.allowedViews?.find(v => v !== 'users') || 'dashboard';
      setCurrentView(allowed);
    }
  }, [currentUser, currentView, canAccessView]);

  // Tự động chọn sẵn công đoạn theo phân quyền tài khoản khi đăng nhập / chuyển đổi user
  useEffect(() => {
    if (currentUser) {
      const defaultStage = getUserDefaultStage(currentUser);
      setSelectedStage(defaultStage);
      setScanStage(defaultStage);
    }
  }, [currentUser]);

  // Khi người dùng chuyển sang tab Xuất kho (produce) hoặc Nhập kho (inbound):
  // Tự động thiết lập sẵn công đoạn theo phân quyền user
  useEffect(() => {
    if (currentUser) {
      const defaultStage = getUserDefaultStage(currentUser);
      if (currentView === 'produce') {
        setSelectedStage(defaultStage);
      } else if (currentView === 'inbound') {
        setScanStage(defaultStage);
      }
    }
  }, [currentView, currentUser]);

  const handleLogout = () => {
    storageService.setCurrentUser(null);
    setCurrentUser(null);
    setShowLoginModal(true);
    setShowUserDropdown(false);
  };

  const handleLoginSuccess = (user: UserAccount) => {
    setCurrentUser(user);
    setShowLoginModal(false);
    setShowUserDropdown(false);
    const defaultStage = getUserDefaultStage(user);
    setSelectedStage(defaultStage);
    setScanStage(defaultStage);
    const firstAllowed = user.role === 'ADMIN' ? 'dashboard' : (user.allowedViews?.[0] || 'dashboard');
    setCurrentView(firstAllowed);
    setSuccess(`Xin chào, ${user.fullName || user.username}!`);
  };

  const pendingOddPOCount = useMemo(() => {
    return (productionOrders || []).filter(o => 
      (o.isOddPO || o.id.startsWith('PO-LE') || (o.note && (o.note.includes('Dư Lẻ') || o.note.includes('Tồn Lẻ')))) && 
      o.status === 'PENDING_APPROVAL'
    ).length;
  }, [productionOrders]);

  const dclrNorms = useMemo(() => storageService.getNorms().filter(n => n.stageId === 'DCLR'), [isImportingNorms, parts]);
  const dclrNormsMap = useMemo(() => new Map(dclrNorms.map(n => [n.partId, n.secondsPerUnit])), [dclrNorms]);

  const glazingHSQD = useCallback((pId: string) => getGlazingHSQD(pId, dclrNormsMap, dclrNorms, parts), [dclrNormsMap, dclrNorms, parts]);

  const aggregatedGlazingOut = useMemo(() => {
    const map = new Map<string, import('./types').InventoryItem>();
    inventory.filter(i => i.stageId === 'GLAZING' && i.location === 'OUT' && i.quantity > 0).forEach(i => {
      // Use case-insensitive trimmed partId as aggregation key
      const canonicalKey = i.partId.trim().toUpperCase();
      const existing = map.get(canonicalKey);
      if (existing) {
        existing.quantity += i.quantity;
      } else {
        map.set(canonicalKey, { ...i });
      }
    });
    return Array.from(map.values());
  }, [inventory]);

  useEffect(() => {
    const existingParts = storageService.getParts();
    if (existingParts.length > 0 && !selectedPart) {
      setSelectedPart(existingParts[0].id);
    }

    // Initial immediate local render (0ms lag)
    refreshData();

    // Async event-driven check on startup / F5:
    // Kiểm tra đúng 1 dòng duy nhất 'data_reset_version'.
    // Nếu version thay đổi hoặc LocalStorage trống -> PULL dữ liệu sạch.
    // Nếu version khớp -> dùng trực tiếp LocalStorage (0 Egress / 0 Request ngầm).
    let isMounted = true;
    storageService.initGlobalSync();

    (async () => {
      try {
        await storageService.init(false);
        if (isMounted) refreshData();
      } catch {}
    })();

    const handleDataUpdated = () => {
      if (isMounted) refreshData();
    };
    window.addEventListener('wip_data_updated', handleDataUpdated);

    const unsub = storageService.onSyncChange(status => {
      if (isMounted) setSyncStatus(status);
    });

    return () => {
      isMounted = false;
      window.removeEventListener('wip_data_updated', handleDataUpdated);
      unsub();
    };
  }, []);

  // Instant local state update (0ms lag, 0 Supabase network queries)
  const refreshData = useCallback(() => {
    setInventory(storageService.getInventory());
    setTransactions(storageService.getTransactions());
    setLabels(storageService.getLabels());
    setProductionOrders(storageService.getProductionOrders());
    const currentParts = storageService.getParts();
    setParts(currentParts);
    setSelectedPart(prev => (!prev && currentParts.length > 0 ? currentParts[0].id : prev));
  }, []);

  // Deliberate manual remote sync (force = true pulls all tables on user click)
  const syncWithRemote = useCallback(async (force: boolean = true) => {
    if (!isSupabaseConfigured) return;
    try {
      const ok = await storageService.init(force);
      refreshData();
      if (ok) {
        setSuccess('Đã đồng bộ dữ liệu mới nhất từ máy chủ Supabase thành công!');
        setTimeout(() => setSuccess(null), 3500);
      } else {
        setError('Không thể đồng bộ dữ liệu từ máy chủ. Vui lòng kiểm tra kết nối mạng!');
        setTimeout(() => setError(null), 4000);
      }
    } catch (err: any) {
      setError(err?.message || 'Lỗi khi đồng bộ dữ liệu từ máy chủ');
      setTimeout(() => setError(null), 4000);
    }
  }, [refreshData]);

  const handleProduce = async (e: React.FormEvent, sourceLocation: 'IN' | 'OUT' = 'IN', targetStageId?: StageId, poId?: string) => {
    e.preventDefault();
    if (quantity <= 0) {
      setError('Vui lòng nhập số lượng hợp lệ');
      return;
    }

    try {
      const tx = storageService.recordStageOut(selectedPart, selectedStage, quantity, sourceLocation, targetStageId, poId);
      
      // Update inventory state immediately
      const updatedInventory = storageService.getInventory();
      setInventory(updatedInventory);
      
      setLastTransaction(tx);
      const part = parts.find(p => p.id === selectedPart);
      const partName = getProcessValue(part?.name, part, selectedStage, 'OUT');
      const locationName = sourceLocation === 'IN' ? 'KHO_IN' : 'KHO_OUT';
      setSuccess(`Đã xuất từ ${locationName} ${quantity} ${partName} tại ${STAGES.find(s => s.id === selectedStage)?.name}`);
      setQuantity(0);
      
      refreshData();

      // Ensure we don't automatically trigger print, just show preview
    } catch (err: any) {
      const errMsg = err instanceof Error ? err.message : (err.message || 'Đã xảy ra lỗi');
      if (errMsg.startsWith('OVER_PO:')) {
        const warningMessage = errMsg.replace('OVER_PO:', '');
        await customAlert(warningMessage);
        setError(warningMessage);
      } else {
        setError(errMsg);
      }
    }
  };

  const onScanSuccess = (decodedText: string, targetLocation: 'IN' | 'OUT' = 'IN', overrideStageId?: StageId) => {
    try {
      storageService.recordStageIn(decodedText, overrideStageId || scanStage, targetLocation);
      setSuccess('Nhập kho thành công!');
      refreshData();
      // Don't switch view automatically, let user see the success detail
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Dữ liệu QR không hợp lệ';
      setError(msg);
      throw err; // Re-throw to let ScanView know it failed
    }
  };

  const handleManualInbound = (partId: string, stageId: StageId, location: 'IN' | 'OUT', qty: number, poId?: string) => {
    try {
      storageService.recordManualInbound(partId, stageId, location, qty, poId);
      setSuccess('Nhập kho thủ công thành công!');
      refreshData();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Đã xảy ra lỗi');
      throw err;
    }
  };

  const clearMessages = () => {
    setError(null);
    setSuccess(null);
  };

  const handlePrint = (label: Transaction) => {
    if (!label) return;
    setLastTransaction(label);
    if (label.id) {
      storageService.setTransactionPrinted(label.id, true);
    }
    refreshData();
    // Auto trigger print when handlePrint is called (except from produce view form submit)
    handlePrintConfirm();
  };

  const handlePreviewOnly = (label: Transaction) => {
    if (!label) return;
    setLastTransaction(label);
    if (label.id) {
      storageService.setTransactionPrinted(label.id, true);
    }
    refreshData();
  };

  const handlePrintConfirm = () => {
    setTimeout(() => {
      window.print();
    }, 200);
  };

  const copyToClipboard = (text: string) => {
    navigator.clipboard.writeText(text);
    setSuccess('Đã sao chép mã QR vào bộ nhớ tạm!');
  };

  if (!currentUser) {
    return (
      <LoginModal 
        canCancel={false} 
        onSuccess={handleLoginSuccess} 
      />
    );
  }

  return (
    <div className="flex h-screen bg-[#F8F9FA] text-gray-900 font-sans overflow-hidden print:bg-white">
      <GlobalDialogs />
      <style>
        {`
          @media print {
            @page {
              size: ${labelSettings.width}mm ${labelSettings.height}mm;
              margin: 0;
            }
            body * {
              visibility: hidden;
            }
            #print-area, #print-area * {
              visibility: visible;
              color: #000000 !important;
              border-color: #000000 !important;
              opacity: 1 !important;
              font-weight: 900 !important;
              -webkit-print-color-adjust: exact !important;
              print-color-adjust: exact !important;
            }
            #print-area {
              position: absolute;
              left: 0;
              top: 0;
              width: ${labelSettings.width}mm;
              height: ${labelSettings.height}mm;
              padding: 0;
              margin: 0;
              display: flex;
              flex-direction: column;
              align-items: center;
              justify-content: center;
              background: #ffffff !important;
              overflow: hidden;
            }
            .no-print {
              display: none !important;
            }
          }
        `}
      </style>

      {/* Hidden Print Area */}
      <div id="print-area" className="hidden print:block">
        {lastTransaction && (
          <A7QRLabelCard label={lastTransaction} parts={parts} labelSettings={labelSettings} isPrintArea={true} />
        )}
      </div>

      {/* Mobile Sidebar Overlay Backdrop */}
      {isSidebarOpen && (
        <div 
          onClick={() => setIsSidebarOpen(false)} 
          className="fixed inset-0 bg-black/60 backdrop-blur-xs z-40 md:hidden animate-in fade-in duration-200" 
        />
      )}

      {/* Sidebar */}
      <aside className={cn(
        "bg-white border-r border-gray-200 transition-all duration-300 flex flex-col z-50 fixed inset-y-0 left-0 md:static md:z-auto shadow-2xl md:shadow-none",
        isSidebarOpen 
          ? "w-64 translate-x-0" 
          : "-translate-x-full md:translate-x-0 md:w-20"
      )}>
        <nav className="flex-1 pt-2 px-3 space-y-1 overflow-y-auto">
          {canAccessView('dashboard') && (
            <SidebarLink 
              active={currentView === 'dashboard'} 
              onClick={() => {
                setCurrentView('dashboard');
                if (window.innerWidth < 768) setIsSidebarOpen(false);
              }}
              icon={<LayoutDashboard size={24} />}
              label="Tổng quan (Dashboard)"
              collapsed={!isSidebarOpen}
            />
          )}
          {canAccessView('produce') && (
            <SidebarLink 
              active={currentView === 'produce'} 
              onClick={() => {
                setCurrentView('produce');
                if (window.innerWidth < 768) setIsSidebarOpen(false);
              }}
              icon={<PackagePlus size={24} />}
              label="Xuất kho (Finish)"
              collapsed={!isSidebarOpen}
            />
          )}
          {canAccessView('inbound') && (
            <SidebarLink 
              active={currentView === 'inbound'} 
              onClick={() => {
                setCurrentView('inbound');
                if (window.innerWidth < 768) setIsSidebarOpen(false);
              }}
              icon={<QrCode size={24} />}
              label="Nhập kho (Quét mã)"
              collapsed={!isSidebarOpen}
            />
          )}
          {canAccessView('laser_inbound') && (
            <SidebarLink 
              active={currentView === 'laser_inbound'} 
              onClick={() => {
                setCurrentView('laser_inbound');
                if (window.innerWidth < 768) setIsSidebarOpen(false);
              }}
              icon={<Layers size={24} />}
              label="Nhập kho Laser"
              collapsed={!isSidebarOpen}
            />
          )}
          {canAccessView('welding_inbound') && (
            <SidebarLink 
              active={currentView === 'welding_inbound'} 
              onClick={() => {
                setCurrentView('welding_inbound');
                if (window.innerWidth < 768) setIsSidebarOpen(false);
              }}
              icon={<Flame size={24} />}
              label="Nhập kho Hàn"
              collapsed={!isSidebarOpen}
            />
          )}
          {canAccessView('glazing') && (
            <SidebarLink 
              active={currentView === 'glazing'} 
              onClick={() => {
                setCurrentView('glazing');
                if (window.innerWidth < 768) setIsSidebarOpen(false);
              }}
              icon={<Square size={24} />}
              label="Công đoạn Dán Kính"
              collapsed={!isSidebarOpen}
            />
          )}
          {canAccessView('labels') && (
            <SidebarLink 
              active={currentView === 'labels'} 
              onClick={() => {
                setCurrentView('labels');
                if (window.innerWidth < 768) setIsSidebarOpen(false);
              }}
              icon={<QrCode size={24} />}
              label="Danh sách nhãn QR"
              collapsed={!isSidebarOpen}
            />
          )}
          {canAccessView('po') && (
            <SidebarLink 
              active={currentView === 'po'} 
              onClick={() => {
                setCurrentView('po');
                if (window.innerWidth < 768) setIsSidebarOpen(false);
              }}
              icon={<ClipboardList size={24} />}
              label="Lệnh sản xuất (PO)"
              collapsed={!isSidebarOpen}
            />
          )}
          {canAccessView('defects') && (
            <SidebarLink 
              active={currentView === 'defects'} 
              onClick={() => {
                setCurrentView('defects');
                if (window.innerWidth < 768) setIsSidebarOpen(false);
              }}
              icon={<AlertCircle size={24} className="text-red-500" />}
              label="Báo cáo lỗi (Defect)"
              collapsed={!isSidebarOpen}
            />
          )}
          {canAccessView('working_hours') && (
            <SidebarLink 
              active={currentView === 'working_hours'} 
              onClick={() => {
                setCurrentView('working_hours');
                if (window.innerWidth < 768) setIsSidebarOpen(false);
              }}
              icon={<Clock size={24} />}
              label="Thời gian Ca & Nghỉ"
              collapsed={!isSidebarOpen}
            />
          )}
          {canAccessView('norms') && (
            <SidebarLink 
              active={currentView === 'norms'} 
              onClick={() => {
                setCurrentView('norms');
                if (window.innerWidth < 768) setIsSidebarOpen(false);
              }}
              icon={<Monitor size={24} />}
              label="Định mức Năng suất"
              collapsed={!isSidebarOpen}
            />
          )}
          {canAccessView('odd_po') && (
            <SidebarLink 
              active={currentView === 'odd_po'} 
              onClick={() => {
                setCurrentView('odd_po');
                if (window.innerWidth < 768) setIsSidebarOpen(false);
              }}
              icon={<Layers size={24} />}
              label="PO dư lẻ"
              collapsed={!isSidebarOpen}
              badge={pendingOddPOCount > 0 ? (
                <span className="px-2 py-0.5 rounded-full text-[11px] font-black bg-amber-500 text-white animate-pulse shadow-sm">
                  {pendingOddPOCount}
                </span>
              ) : undefined}
            />
          )}
          {canAccessView('settings') && (
            <SidebarLink 
              active={currentView === 'settings'} 
              onClick={() => {
                setCurrentView('settings');
                if (window.innerWidth < 768) setIsSidebarOpen(false);
              }}
              icon={<Settings size={24} />}
              label="Cài đặt & Sao lưu dữ liệu"
              collapsed={!isSidebarOpen}
            />
          )}

          {/* Admin User Management Tab */}
          {currentUser?.role === 'ADMIN' && (
            <div className="pt-2 mt-2 border-t border-gray-100">
              <SidebarLink 
                active={currentView === 'users'} 
                onClick={() => {
                  setCurrentView('users');
                  if (window.innerWidth < 768) setIsSidebarOpen(false);
                }}
                icon={<Users size={24} className="text-orange-500" />}
                label="Quản lý Tài khoản"
                collapsed={!isSidebarOpen}
              />
            </div>
          )}

          {/* Mobile System Status & Cloud Tools (Hidden on desktop header, shown in mobile drawer) */}
          <div className="pt-3 mt-3 border-t border-gray-100 space-y-2 md:hidden">
            <div className="px-2 py-1 text-[10px] font-black uppercase text-gray-400 tracking-wider">
              Hệ thống & Đồng bộ
            </div>
            <button
              onClick={() => {
                setShowSupabaseModal(true);
                setIsSidebarOpen(false);
              }}
              className="w-full flex items-center justify-between px-3 py-2 text-xs font-bold rounded-xl border border-gray-200 bg-gray-50 text-gray-700 hover:bg-gray-100 cursor-pointer"
            >
              <div className="flex items-center gap-2">
                <Server size={16} className={syncStatus.connected && !syncStatus.error ? "text-emerald-600" : "text-amber-600"} />
                <span>Supabase Cloud</span>
              </div>
              <span className={cn(
                "px-2 py-0.5 text-[10px] rounded-full font-bold",
                syncStatus.connected && !syncStatus.error ? "bg-emerald-100 text-emerald-800" : "bg-amber-100 text-amber-800"
              )}>
                {syncStatus.connected && !syncStatus.error ? "Online" : "Cấu hình"}
              </span>
            </button>

            <button
              onClick={() => {
                setCurrentView('settings');
                setIsSidebarOpen(false);
              }}
              className="w-full flex items-center gap-2 px-3 py-2 text-xs font-bold rounded-xl border border-gray-200 bg-gray-50 text-gray-700 hover:bg-gray-100 cursor-pointer"
            >
              <Database size={16} className="text-emerald-600" />
              <span>Sao lưu / Nạp JSON</span>
            </button>

            <div className="px-3 py-2 bg-gray-50 rounded-xl border border-gray-200 flex items-center justify-between text-[11px] font-mono text-gray-600">
              <span className="opacity-70">Terminal ID:</span>
              <span className="font-bold text-gray-900">DESKTOP-WS-01</span>
            </div>
          </div>
        </nav>

        <div className="p-4 border-t border-gray-100">
          <button 
            onClick={() => setIsSidebarOpen(!isSidebarOpen)}
            className="w-full flex items-center justify-center p-2 hover:bg-gray-100 rounded-lg transition-colors text-gray-400 cursor-pointer"
          >
            <Menu size={20} />
          </button>
        </div>
      </aside>

      {/* Main Content Area */}
      <div className="flex-1 flex flex-col overflow-hidden min-w-0">
        {/* Top Header */}
        <header className="bg-white border-b border-gray-200 py-2.5 px-3 sm:px-6 flex justify-between items-center shadow-xs z-40 gap-2">
          <div className="flex items-center gap-2 sm:gap-4 min-w-0">
            <button
              onClick={() => setIsSidebarOpen(!isSidebarOpen)}
              className="p-2 -ml-1 hover:bg-gray-100 rounded-xl transition-colors text-gray-700 md:hidden cursor-pointer shrink-0"
              title="Menu Bảng điều khiển"
            >
              <Menu size={22} />
            </button>
            <h2 className="font-bold text-xs sm:text-lg md:text-xl uppercase tracking-tight truncate max-w-[180px] sm:max-w-none">
              {currentView === 'dashboard' && 'Báo cáo tồn kho WIP'}
              {currentView === 'produce' && 'Xuất kho & In nhãn QR'}
              {currentView === 'inbound' && 'Nhập kho (Quét mã QR)'}
              {currentView === 'laser_inbound' && 'Nhập kho Laser (Thủ công)'}
              {currentView === 'welding_inbound' && 'Nhập kho Hàn (Thủ công)'}
              {currentView === 'odd_po' && 'PO dư lẻ'}
              {currentView === 'labels' && 'Danh sách nhãn QR đã xuất'}
              {currentView === 'po' && 'Quản lý Lệnh sản xuất (PO)'}
              {currentView === 'norms' && 'Định mức năng suất sản xuất'}
              {currentView === 'working_hours' && 'Cài đặt Ca làm việc & Nghỉ ngơi'}
              {currentView === 'settings' && 'Cài đặt linh kiện & Sao lưu Local Storage'}
              {currentView === 'users' && 'Quản trị Người dùng & Phân quyền'}
            </h2>
          </div>
          <div className="flex items-center gap-2 sm:gap-3 shrink-0">
            {/* Supabase Cloud Status Button (Hidden on Mobile, shown in Mobile Drawer) */}
            <button
              onClick={() => setShowSupabaseModal(true)}
              title="Quản lý đồng bộ Supabase Cloud Database"
              className={cn(
                "hidden md:flex items-center gap-1.5 sm:gap-2 px-2.5 sm:px-3.5 py-1.5 sm:py-2 rounded-xl text-[11px] sm:text-xs font-bold border transition-all shadow-xs active:scale-95 cursor-pointer",
                syncStatus.syncing 
                  ? "bg-blue-50 text-blue-700 border-blue-200" 
                  : syncStatus.connected && !syncStatus.error
                    ? "bg-emerald-50 text-emerald-700 border-emerald-200 hover:bg-emerald-100"
                    : "bg-amber-50 text-amber-700 border-amber-200 hover:bg-amber-100"
              )}
            >
              {syncStatus.syncing ? (
                <>
                  <RefreshCw size={14} className="animate-spin text-blue-600" />
                  <span className="hidden lg:inline">Đang đồng bộ...</span>
                </>
              ) : syncStatus.connected && !syncStatus.error ? (
                <>
                  <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
                  <Server size={14} className="text-emerald-600" />
                  <span className="hidden lg:inline">Supabase Online</span>
                </>
              ) : (
                <>
                  <CloudOff size={14} className="text-amber-600" />
                  <span className="hidden lg:inline">Cấu hình Supabase</span>
                </>
              )}
            </button>

            {/* Manual Sync Button - User-Initiated Server Sync (Realtime disabled by request) */}
            <button
              onClick={() => syncWithRemote(true)}
              disabled={syncStatus.syncing}
              title="Đồng bộ dữ liệu máy chủ Supabase (bấm để tải dữ liệu mới nhất về máy)"
              className={cn(
                "flex items-center gap-1.5 sm:gap-2 px-2.5 sm:px-3 py-1.5 sm:py-2 rounded-xl text-[11px] sm:text-xs font-bold transition-all shadow-xs active:scale-95 cursor-pointer disabled:opacity-50 shrink-0",
                syncStatus.syncing
                  ? "bg-blue-100 text-blue-700 border border-blue-300"
                  : "bg-blue-600 hover:bg-blue-700 text-white border border-blue-700"
              )}
            >
              <RefreshCw size={14} className={syncStatus.syncing ? "animate-spin text-blue-600" : "text-white"} />
              <span>{syncStatus.syncing ? 'Đang đồng bộ...' : 'Đồng bộ dữ liệu máy chủ'}</span>
            </button>

            {/* Backup/Restore JSON Button (Hidden on Mobile) */}
            <button
              onClick={() => {
                setCurrentView('settings');
              }}
              title="Mở giao diện Sao lưu & Nạp dữ liệu Local Storage"
              className="hidden md:flex items-center gap-1.5 sm:gap-2 px-2.5 sm:px-3.5 py-1.5 sm:py-2 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 border border-emerald-200 rounded-xl text-[11px] sm:text-xs font-bold transition-all shadow-xs active:scale-95 cursor-pointer"
            >
              <Database size={15} className="text-emerald-600" />
              <span className="hidden lg:inline">Sao lưu / Nạp JSON</span>
            </button>

            {/* User Account / Profile Dropdown */}
            <div className="relative">
              {currentUser ? (
                <div>
                  <button
                    onClick={() => setShowUserDropdown(!showUserDropdown)}
                    className="flex items-center gap-1.5 sm:gap-2.5 px-2.5 sm:px-3 py-1.5 bg-gray-50 hover:bg-gray-100 border border-gray-200 rounded-xl transition-all shadow-xs cursor-pointer active:scale-95"
                    title={`Đang đăng nhập: @${currentUser.username}`}
                  >
                    <div className={cn(
                      "w-7 h-7 rounded-lg flex items-center justify-center font-black text-xs uppercase shadow-xs shrink-0",
                      currentUser.role === 'ADMIN' ? "bg-orange-500 text-white" : "bg-blue-600 text-white"
                    )}>
                      {currentUser.fullName ? currentUser.fullName.substring(0, 2) : currentUser.username.substring(0, 2)}
                    </div>
                    <div className="text-left hidden md:block">
                      <div className="text-xs font-bold text-gray-900 leading-tight">
                        {currentUser.fullName || currentUser.username}
                      </div>
                      <div className="text-[10px] font-bold uppercase tracking-wider text-gray-400">
                        {currentUser.role === 'ADMIN' ? 'Quản trị viên' : 'Nhân viên'}
                      </div>
                    </div>
                    <ChevronDown size={14} className="text-gray-400 ml-0.5" />
                  </button>

                  <AnimatePresence>
                    {showUserDropdown && (
                      <>
                        <div 
                          className="fixed inset-0 z-40" 
                          onClick={() => setShowUserDropdown(false)} 
                        />
                        <div className="absolute right-0 top-full mt-2 w-60 bg-white border border-gray-200 rounded-2xl shadow-xl z-50 py-2 divide-y divide-gray-100">
                          <div className="px-4 py-2.5">
                            <p className="text-xs font-bold text-gray-900 truncate">{currentUser.fullName}</p>
                            <p className="text-[11px] text-gray-500 font-mono">@{currentUser.username}</p>
                            <span className={cn(
                              "inline-block mt-1 px-2 py-0.5 text-[10px] font-black rounded-md",
                              currentUser.role === 'ADMIN' ? "bg-orange-100 text-orange-700" : "bg-blue-100 text-blue-700"
                            )}>
                              {currentUser.role === 'ADMIN' ? 'Quản trị viên (Admin)' : 'Nhân viên'}
                            </span>
                          </div>

                          <div className="py-1">
                            <button
                              onClick={() => {
                                syncWithRemote(true);
                                setShowUserDropdown(false);
                              }}
                              disabled={syncStatus.syncing}
                              className="w-full flex items-center gap-2.5 px-4 py-2 text-xs font-bold text-gray-700 hover:bg-blue-50 hover:text-blue-600 transition-colors text-left cursor-pointer md:hidden"
                            >
                              <RefreshCw size={15} className={syncStatus.syncing ? "animate-spin text-blue-600" : "text-blue-600"} />
                              <span>{syncStatus.syncing ? 'Đang đồng bộ...' : 'Đồng bộ dữ liệu máy chủ'}</span>
                            </button>
                            {currentUser.role === 'ADMIN' && (
                              <button
                                onClick={() => {
                                  setCurrentView('users');
                                  setShowUserDropdown(false);
                                }}
                                className="w-full flex items-center gap-2.5 px-4 py-2 text-xs font-bold text-gray-700 hover:bg-orange-50 hover:text-orange-600 transition-colors text-left cursor-pointer"
                              >
                                <Users size={15} />
                                <span>Quản lý Tài khoản</span>
                              </button>
                            )}
                            <button
                              onClick={() => {
                                setShowSupabaseModal(true);
                                setShowUserDropdown(false);
                              }}
                              className="w-full flex items-center gap-2.5 px-4 py-2 text-xs font-bold text-gray-700 hover:bg-blue-50 hover:text-blue-600 transition-colors text-left cursor-pointer md:hidden"
                            >
                              <Server size={15} />
                              <span>Cấu hình Supabase Cloud</span>
                            </button>
                            <button
                              onClick={() => {
                                setCurrentView('settings');
                                setShowUserDropdown(false);
                              }}
                              className="w-full flex items-center gap-2.5 px-4 py-2 text-xs font-bold text-gray-700 hover:bg-emerald-50 hover:text-emerald-600 transition-colors text-left cursor-pointer md:hidden"
                            >
                              <Database size={15} />
                              <span>Sao lưu / Nạp JSON</span>
                            </button>
                            <button
                              onClick={() => {
                                setShowChangePasswordModal(true);
                                setShowUserDropdown(false);
                              }}
                              className="w-full flex items-center gap-2.5 px-4 py-2 text-xs font-bold text-gray-700 hover:bg-gray-50 transition-colors text-left cursor-pointer"
                            >
                              <Key size={15} />
                              <span>Đổi mật khẩu</span>
                            </button>
                            <button
                              onClick={() => {
                                setShowLoginModal(true);
                                setShowUserDropdown(false);
                              }}
                              className="w-full flex items-center gap-2.5 px-4 py-2 text-xs font-bold text-gray-700 hover:bg-gray-50 transition-colors text-left cursor-pointer"
                            >
                              <UserCheck size={15} />
                              <span>Chuyển tài khoản</span>
                            </button>
                          </div>

                          <div className="py-1">
                            <button
                              onClick={handleLogout}
                              className="w-full flex items-center gap-2.5 px-4 py-2 text-xs font-bold text-rose-600 hover:bg-rose-50 transition-colors text-left cursor-pointer"
                            >
                              <LogOut size={15} />
                              <span>Đăng xuất</span>
                            </button>
                          </div>
                        </div>
                      </>
                    )}
                  </AnimatePresence>
                </div>
              ) : (
                <button
                  onClick={() => setShowLoginModal(true)}
                  className="flex items-center gap-2 px-3 py-1.5 bg-orange-500 hover:bg-orange-600 text-white rounded-xl text-xs font-black shadow-md shadow-orange-500/20 active:scale-95 cursor-pointer"
                >
                  <Lock size={14} />
                  <span>Đăng nhập</span>
                </button>
              )}
            </div>

            <div className="h-6 w-px bg-gray-200 hidden md:block" />
            <div className="hidden md:flex flex-col items-end">
              <span className="text-xs font-mono opacity-50 uppercase">Terminal ID</span>
              <span className="text-sm font-mono font-bold">DESKTOP-WS-01</span>
            </div>
            <div className="h-6 w-px bg-gray-200 hidden md:block" />
            <div className="hidden md:flex items-center gap-2.5">
              <div className="w-2.5 h-2.5 rounded-full bg-green-500 animate-pulse" />
              <span className="text-xs font-mono opacity-70 uppercase tracking-widest">Sẵn sàng</span>
            </div>
          </div>
        </header>

        {/* Scrollable Content */}
        <main className="flex-1 overflow-y-auto p-2.5 sm:p-6 md:p-8 bg-[#F8F9FA]">
          <div className="w-full max-w-[1600px] mx-auto">
            <AnimatePresence initial={false}>
              {currentView === 'dashboard' && (
                <DashboardView 
                  key="dashboard" 
                  labels={labels} 
                  transactions={transactions} 
                  inventory={inventory} 
                  parts={parts} 
                  refreshData={refreshData} 
                  setDefectModal={setDefectModal} 
                  dclrNorms={dclrNorms}
                  dclrNormsMap={dclrNormsMap}
                  glazingHSQD={glazingHSQD}
                />
              )}
              {currentView === 'produce' && (
                <ProduceView 
                  key="produce"
                  selectedStage={selectedStage}
                  setSelectedStage={setSelectedStage}
                  selectedPart={selectedPart}
                  setSelectedPart={setSelectedPart}
                  quantity={quantity}
                  setQuantity={setQuantity}
                  handleProduce={handleProduce}
                  lastTransaction={lastTransaction}
                  setLastTransaction={setLastTransaction}
                  inventory={inventory}
                  productionOrders={productionOrders}
                  parts={parts}
                  onPrint={handlePrint}
                  onCopy={copyToClipboard}
                  setDefectModal={setDefectModal}
                  labelSettings={labelSettings}
                />
              )}
              {currentView === 'inbound' && (
                <InboundView 
                  key="inbound"
                  selectedStage={scanStage}
                  setSelectedStage={setScanStage}
                  onScanSuccess={onScanSuccess}
                  parts={parts}
                  setDefectModal={setDefectModal}
                />
              )}
              {currentView === 'laser_inbound' && (
                <LaserInboundView 
                  key="laser_inbound"
                  parts={parts}
                  onManualInbound={handleManualInbound}
                  setDefectModal={setDefectModal}
                  productionOrders={productionOrders}
                />
              )}
              {currentView === 'welding_inbound' && (
                <WeldingInboundView 
                  key="welding_inbound"
                  parts={parts}
                  onManualInbound={handleManualInbound}
                  setDefectModal={setDefectModal}
                  productionOrders={productionOrders}
                />
              )}
              {currentView === 'glazing' && (
                <GlazingView 
                  key="glazing"
                  parts={parts}
                  inventory={inventory}
                  onManualInbound={handleManualInbound}
                  setDefectModal={setDefectModal}
                  refreshData={refreshData}
                  labels={labels}
                  onPrint={handlePrint}
                  aggregatedGlazingOut={aggregatedGlazingOut}
                />
              )}
              {currentView === 'odd_po' && (
                <OddProductionOrderView 
                  key="odd_po"
                  parts={parts}
                  onUpdate={refreshData}
                  productionOrders={productionOrders}
                  currentUser={currentUser}
                />
              )}
              {currentView === 'labels' && (
                <LabelHistoryView 
                  key="labels" 
                  parts={parts} 
                  labels={labels}
                  onPrint={handlePrint}
                  onCopy={copyToClipboard}
                  onRollback={refreshData}
                  onManualInboundQR={(qr, stage) => onScanSuccess(qr, 'IN', stage)}
                  labelSettings={labelSettings}
                  currentUser={currentUser}
                />
              )}
              {currentView === 'po' && (
                <ProductionOrderView 
                  parts={parts} 
                  onUpdate={refreshData} 
                  productionOrders={productionOrders} 
                  onNavigateToOddPO={() => setCurrentView('odd_po')}
                />
              )}
              {currentView === 'defects' && (
                <DefectView key="defects" parts={parts} transactions={transactions} inventory={inventory} refreshData={refreshData} setLastTransaction={setLastTransaction} />
              )}
              {currentView === 'norms' && (
                <NormsView parts={parts} onNormsChange={refreshData} />
              )}
              {currentView === 'working_hours' && (
                <WorkingHoursView />
              )}
              {currentView === 'settings' && (
                <SettingsView 
                  key="settings" 
                  parts={parts} 
                  onPartsChange={refreshData} 
                  labelSettings={labelSettings}
                  onLabelSettingsChange={(s: any) => {
                    setLabelSettings(s);
                    storageService.saveLabelSettings(s);
                  }}
                />
              )}
              {currentView === 'users' && currentUser?.role === 'ADMIN' && (
                <UserManagementView 
                  key="users" 
                  currentUser={currentUser} 
                  onUserUpdated={() => {
                    setCurrentUser(storageService.getCurrentUser());
                    refreshData();
                  }} 
                />
              )}
            </AnimatePresence>
          </div>
        </main>

      </div>


      {/* Defect Modal */}
      {defectModal && (
        <DefectModal 
          data={defectModal} 
          inventory={inventory}
          setLastTransaction={setLastTransaction}
          onClose={() => setDefectModal(null)} 
          onDefectRecorded={() => {
            setSuccess('Đã ghi nhận hàng lỗi thành công!');
            refreshData();
          }} 
        />
      )}

      {/* Supabase Sync & Setup Modal */}
      {showSupabaseModal && (
        <SupabaseModal 
          syncStatus={syncStatus} 
          onClose={() => setShowSupabaseModal(false)} 
          onDataChanged={refreshData} 
        />
      )}

      {/* Login Modal */}
      <AnimatePresence>
        {showLoginModal && (
          <LoginModal 
            canCancel={!!currentUser} 
            onCancel={() => setShowLoginModal(false)} 
            onSuccess={handleLoginSuccess} 
          />
        )}
      </AnimatePresence>

      {/* Change Password Modal */}
      <AnimatePresence>
        {showChangePasswordModal && currentUser && (
          <ChangePasswordModal 
            currentUser={currentUser} 
            onClose={() => setShowChangePasswordModal(false)} 
            onSuccess={() => {
              setShowChangePasswordModal(false);
              setSuccess('Đã đổi mật khẩu thành công!');
            }} 
          />
        )}
      </AnimatePresence>

      {/* Notification Overlay */}
      <AnimatePresence>
        {(error || success) && (
          <motion.div 
            initial={{ opacity: 0, x: 100 }}
            animate={{ opacity: 1, x: 0 }}
            exit={{ opacity: 0, x: 100 }}
            className="fixed top-20 right-8 z-50 w-80"
          >
            <div className={cn(
              "p-4 rounded-lg shadow-2xl border flex items-start justify-between",
              error ? "bg-red-50 border-red-200 text-red-800" : "bg-green-50 border-green-200 text-green-800"
            )}>
              <div className="flex items-start gap-3">
                {error ? <AlertCircle size={20} className="mt-0.5" /> : <CheckCircle2 size={20} className="mt-0.5" />}
                <div>
                  <p className="text-sm font-bold uppercase tracking-tight mb-1">{error ? 'Lỗi hệ thống' : 'Thành công'}</p>
                  <p className="text-xs opacity-90">{error || success}</p>
                </div>
              </div>
              <button onClick={clearMessages} className="p-1 hover:bg-black/5 rounded">
                <X size={16} />
              </button>
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

function SupabaseModal({
  syncStatus,
  onClose,
  onDataChanged
}: {
  syncStatus: { connected: boolean; syncing: boolean; lastSynced: number | null; error: string | null };
  onClose: () => void;
  onDataChanged: () => void;
}) {
  const [activeTab, setActiveTab] = useState<'overview' | 'sql'>('overview');
  const [testResult, setTestResult] = useState<{ ok: boolean; message: string; details?: any } | null>(null);
  const [isTesting, setIsTesting] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [feedback, setFeedback] = useState<{ type: 'success' | 'error'; message: string } | null>(null);
  const [copiedSql, setCopiedSql] = useState(false);

  const handleTestConnection = async () => {
    setIsTesting(true);
    setTestResult(null);
    try {
      const res = await testSupabaseConnection();
      setTestResult(res);
    } catch (e: any) {
      setTestResult({ ok: false, message: e?.message || 'Lỗi kiểm tra kết nối' });
    } finally {
      setIsTesting(false);
    }
  };

  const handlePullFromSupabase = async () => {
    setIsProcessing(true);
    setFeedback(null);
    try {
      const success = await storageService.refreshFromSupabase();
      if (success) {
        onDataChanged();
        setFeedback({ type: 'success', message: 'Đã tải thành công dữ liệu mới nhất từ Supabase về RAM & Local Storage!' });
      } else {
        setFeedback({ type: 'error', message: 'Không thể tải dữ liệu từ Supabase. Vui lòng kiểm tra các bảng đã được tạo trên Supabase chưa.' });
      }
    } catch (err: any) {
      setFeedback({ type: 'error', message: err?.message || 'Lỗi khi tải dữ liệu' });
    } finally {
      setIsProcessing(false);
    }
  };

  const handlePushToSupabase = async () => {
    setIsProcessing(true);
    setFeedback(null);
    try {
      const res = await storageService.syncLocalToSupabase();
      if (res.success) {
        onDataChanged();
        setFeedback({ type: 'success', message: res.message });
      } else {
        setFeedback({ type: 'error', message: res.message });
      }
    } catch (err: any) {
      setFeedback({ type: 'error', message: err?.message || 'Lỗi khi đẩy dữ liệu' });
    } finally {
      setIsProcessing(false);
    }
  };

  const handleDbCleanup = async () => {
    setIsProcessing(true);
    setFeedback(null);
    try {
      const res = await storageService.autoCleanupOldData();
      onDataChanged();
      setFeedback({
        type: 'success',
        message: `Đã dọn dẹp DB Supabase & máy: xóa ${res.labelsRemoved} nhãn QR (>24h), ${res.posRemoved} PO hoàn thành (>35 ngày), ${res.txsRemoved} logs (>35 ngày).`
      });
    } catch (err: any) {
      setFeedback({ type: 'error', message: err?.message || 'Lỗi khi dọn dẹp DB' });
    } finally {
      setIsProcessing(false);
    }
  };

  const handleCopySql = () => {
    navigator.clipboard.writeText(SUPABASE_SQL_SCHEMA);
    setCopiedSql(true);
    setTimeout(() => setCopiedSql(false), 3000);
  };

  return (
    <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-2xl max-w-3xl w-full max-h-[90vh] flex flex-col overflow-hidden border border-gray-100">
        {/* Header */}
        <div className="p-5 border-b border-gray-100 flex items-center justify-between bg-gray-50/50">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-blue-600 text-white flex items-center justify-center shadow-md shadow-blue-500/20">
              <Server size={20} />
            </div>
            <div>
              <h3 className="font-bold text-lg text-gray-900">Quản lý Cơ sở dữ liệu Supabase Online</h3>
              <p className="text-xs text-gray-500 font-mono">{SUPABASE_URL}</p>
            </div>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-gray-200 rounded-lg text-gray-400 hover:text-gray-600 transition cursor-pointer">
            <X size={18} />
          </button>
        </div>

        {/* Tabs */}
        <div className="flex border-b border-gray-200 bg-gray-50/30 px-6">
          <button
            onClick={() => setActiveTab('overview')}
            className={cn(
              "py-3 px-4 text-xs font-bold uppercase tracking-wider border-b-2 transition flex items-center gap-2 cursor-pointer",
              activeTab === 'overview'
                ? "border-blue-600 text-blue-600 bg-white"
                : "border-transparent text-gray-500 hover:text-gray-800"
            )}
          >
            <RefreshCw size={14} />
            <span>Tổng quan & Đồng bộ</span>
          </button>
          <button
            onClick={() => setActiveTab('sql')}
            className={cn(
              "py-3 px-4 text-xs font-bold uppercase tracking-wider border-b-2 transition flex items-center gap-2 cursor-pointer",
              activeTab === 'sql'
                ? "border-blue-600 text-blue-600 bg-white"
                : "border-transparent text-gray-500 hover:text-gray-800"
            )}
          >
            <FileText size={14} />
            <span>Mã SQL tạo 6 bảng</span>
          </button>
        </div>

        {/* Content */}
        <div className="p-6 overflow-y-auto space-y-6 flex-1">
          {activeTab === 'overview' ? (
            <>
              {/* Status Box */}
              <div className={cn(
                "p-4 rounded-xl border flex items-start gap-3",
                syncStatus.connected && !syncStatus.error
                  ? "bg-emerald-50/70 border-emerald-200 text-emerald-900"
                  : "bg-amber-50/70 border-amber-200 text-amber-900"
              )}>
                {syncStatus.connected && !syncStatus.error ? (
                  <CheckCircle2 size={20} className="text-emerald-600 shrink-0 mt-0.5" />
                ) : (
                  <AlertCircle size={20} className="text-amber-600 shrink-0 mt-0.5" />
                )}
                <div className="flex-1 text-xs">
                  <div className="font-bold text-sm mb-1">
                    {syncStatus.connected && !syncStatus.error
                      ? "Kết nối cơ sở dữ liệu Supabase hoạt động tốt"
                      : "Trạng thái Supabase cần chú ý"}
                  </div>
                  {syncStatus.error ? (
                    <div className="text-amber-800 font-medium">{syncStatus.error}</div>
                  ) : (
                    <div>Hệ thống đang lưu trữ và đồng bộ hóa trực tuyến trên Cloud Supabase. Đọc/ghi kết hợp bộ nhớ đệm RAM tối ưu gói Free.</div>
                  )}
                  {syncStatus.lastSynced && (
                    <div className="mt-2 text-gray-500 font-mono text-[11px]">
                      Đồng bộ gần nhất: {format(syncStatus.lastSynced, 'dd/MM/yyyy HH:mm:ss')}
                    </div>
                  )}
                </div>
              </div>

              {/* Feedback Alert */}
              {feedback && (
                <div className={cn(
                  "p-3.5 rounded-xl border text-xs flex items-center justify-between",
                  feedback.type === 'success'
                    ? "bg-green-50 border-green-200 text-green-800"
                    : "bg-red-50 border-red-200 text-red-800"
                )}>
                  <span>{feedback.message}</span>
                  <button onClick={() => setFeedback(null)} className="text-gray-400 hover:text-gray-600 cursor-pointer">
                    <X size={14} />
                  </button>
                </div>
              )}

              {/* Actions Grid */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <button
                  disabled={isTesting}
                  onClick={handleTestConnection}
                  className="p-4 rounded-xl border border-gray-200 bg-white hover:border-blue-400 hover:bg-blue-50/30 transition text-left flex items-start gap-3 shadow-xs active:scale-98 disabled:opacity-50 cursor-pointer"
                >
                  <div className="p-2 rounded-lg bg-blue-100 text-blue-700">
                    <RefreshCw size={16} className={isTesting ? "animate-spin" : ""} />
                  </div>
                  <div>
                    <div className="font-bold text-xs text-gray-900">Kiểm tra kết nối</div>
                    <div className="text-[11px] text-gray-500 mt-0.5">Ping thử cơ sở dữ liệu Supabase</div>
                  </div>
                </button>

                <button
                  disabled={isProcessing}
                  onClick={handlePullFromSupabase}
                  className="p-4 rounded-xl border border-gray-200 bg-white hover:border-emerald-400 hover:bg-emerald-50/30 transition text-left flex items-start gap-3 shadow-xs active:scale-98 disabled:opacity-50 cursor-pointer"
                >
                  <div className="p-2 rounded-lg bg-emerald-100 text-emerald-700">
                    <Download size={16} className={isProcessing ? "animate-bounce" : ""} />
                  </div>
                  <div>
                    <div className="font-bold text-xs text-gray-900">Tải dữ liệu từ Supabase</div>
                    <div className="text-[11px] text-gray-500 mt-0.5">Lấy bản ghi mới nhất từ Cloud vào RAM</div>
                  </div>
                </button>

                <button
                  disabled={isProcessing}
                  onClick={handlePushToSupabase}
                  className="p-4 rounded-xl border border-gray-200 bg-white hover:border-indigo-400 hover:bg-indigo-50/30 transition text-left flex items-start gap-3 shadow-xs active:scale-98 disabled:opacity-50 cursor-pointer"
                >
                  <div className="p-2 rounded-lg bg-indigo-100 text-indigo-700">
                    <Upload size={16} className={isProcessing ? "animate-bounce" : ""} />
                  </div>
                  <div>
                    <div className="font-bold text-xs text-gray-900">Đẩy toàn bộ lên Supabase</div>
                    <div className="text-[11px] text-gray-500 mt-0.5">Batch Upsert dữ liệu hiện tại lên Cloud</div>
                  </div>
                </button>

                <button
                  disabled={isProcessing}
                  onClick={handleDbCleanup}
                  className="p-4 rounded-xl border border-gray-200 bg-white hover:border-amber-400 hover:bg-amber-50/30 transition text-left flex items-start gap-3 shadow-xs active:scale-98 disabled:opacity-50 cursor-pointer"
                >
                  <div className="p-2 rounded-lg bg-amber-100 text-amber-700">
                    <Trash2 size={16} />
                  </div>
                  <div>
                    <div className="font-bold text-xs text-gray-900">Dọn dẹp DB (&gt;35 ngày)</div>
                    <div className="text-[11px] text-gray-500 mt-0.5">Xóa trực tiếp logs cũ trên DB Supabase</div>
                  </div>
                </button>
              </div>

              {/* Test result detail */}
              {testResult && (
                <div className={cn(
                  "p-4 rounded-xl border text-xs",
                  testResult.ok ? "bg-emerald-50 border-emerald-200 text-emerald-800" : "bg-red-50 border-red-200 text-red-800"
                )}>
                  <div className="font-bold mb-1 flex items-center gap-2">
                    {testResult.ok ? <CheckCircle2 size={16} /> : <AlertCircle size={16} />}
                    <span>{testResult.ok ? 'Kết nối thành công!' : 'Kết nối thất bại hoặc chưa có bảng'}</span>
                  </div>
                  <div className="opacity-90 leading-relaxed">{testResult.message}</div>
                </div>
              )}
            </>
          ) : (
            <div className="space-y-4">
              <div className="flex items-center justify-between">
                <div>
                  <h4 className="font-bold text-sm text-gray-900">Mã lệnh SQL tạo 6 bảng trên Supabase</h4>
                  <p className="text-xs text-gray-500 mt-0.5">Bao gồm system_configs, production_orders, wip_inventory, wip_transactions, glazing_plans, qr_labels và chính sách RLS.</p>
                </div>
                <button
                  onClick={handleCopySql}
                  className={cn(
                    "px-4 py-2 rounded-xl text-xs font-bold transition flex items-center gap-2 cursor-pointer shadow-sm active:scale-95",
                    copiedSql ? "bg-emerald-600 text-white" : "bg-blue-600 text-white hover:bg-blue-700"
                  )}
                >
                  {copiedSql ? <Check size={14} /> : <Copy size={14} />}
                  <span>{copiedSql ? 'Đã sao chép SQL!' : 'Sao chép toàn bộ SQL'}</span>
                </button>
              </div>

              <div className="p-3 bg-blue-50 border border-blue-200 rounded-xl text-xs text-blue-900 flex items-start gap-2.5">
                <AlertCircle size={16} className="text-blue-600 shrink-0 mt-0.5" />
                <div className="space-y-1">
                  <p className="font-semibold">Hướng dẫn kích hoạt Supabase:</p>
                  <ol className="list-decimal pl-4 space-y-1 text-blue-800">
                    <li>Nhấn nút <b>"Sao chép toàn bộ SQL"</b> ở trên.</li>
                    <li>Mở trình soạn thảo SQL Supabase: <a href="https://supabase.com/dashboard/project/kopsoonfrmtmvpbjyhdu/sql" target="_blank" rel="noreferrer" className="underline font-bold text-blue-700 hover:text-blue-900">Mở Supabase SQL Editor</a>.</li>
                    <li>Dán mã lệnh SQL vào và nhấn <b>RUN</b>.</li>
                    <li>Quay lại ứng dụng và nhấn <b>"Kiểm tra kết nối"</b> rồi <b>"Đẩy toàn bộ lên Supabase"</b> để hoàn tất!</li>
                  </ol>
                </div>
              </div>

              <div className="relative">
                <pre className="p-4 bg-gray-900 text-gray-100 rounded-xl text-xs font-mono overflow-x-auto max-h-72 border border-gray-800 leading-relaxed select-all">
                  {SUPABASE_SQL_SCHEMA}
                </pre>
              </div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="p-4 border-t border-gray-100 bg-gray-50/50 flex justify-end">
          <button
            onClick={onClose}
            className="px-5 py-2.5 bg-gray-900 text-white hover:bg-gray-800 rounded-xl text-xs font-bold transition shadow-sm active:scale-95 cursor-pointer"
          >
            Đóng
          </button>
        </div>
      </div>
    </div>
  );
}

function SidebarLink({ active, onClick, icon, label, collapsed, badge }: { active: boolean, onClick: () => void, icon: React.ReactNode, label: string, collapsed: boolean, badge?: React.ReactNode }) {
  return (
    <button 
      onClick={onClick}
      className={cn(
        "w-full flex items-center justify-between p-3 rounded-lg transition-all duration-200 group relative",
        active 
          ? "bg-[#F27D26] text-white shadow-lg shadow-[#F27D26]/20" 
          : "text-gray-500 hover:text-gray-900 hover:bg-gray-100"
      )}
    >
      <div className="flex items-center gap-3">
        <div className={cn("transition-transform duration-200", active ? "scale-110" : "group-hover:scale-110")}>
          {icon}
        </div>
        {!collapsed && <span className="text-lg font-bold tracking-tight whitespace-nowrap">{label}</span>}
      </div>
      {!collapsed && badge}
      {collapsed && badge && (
        <div className="absolute top-1.5 right-1.5">
          {badge}
        </div>
      )}
    </button>
  );
}

// --- VIEW COMPONENTS ---

interface DashboardProps {
  inventory: InventoryItem[];
  parts: Part[];
  transactions: Transaction[];
  labels: Transaction[];
  refreshData: () => void;
  key?: string;
}

// ==============================================================================
// MODAL: KHÔI PHỤC & ROLLBACK LỆNH SẢN XUẤT (PO RECYCLE BIN & SNAPSHOTS)
// ==============================================================================
interface PORollbackModalProps {
  parts: Part[];
  currentUser?: UserAccount | null;
  onClose: () => void;
  onRestored: () => void;
}

function PORollbackModal({ parts, currentUser, onClose, onRestored }: PORollbackModalProps) {
  const [activeTab, setActiveTab] = useState<'TRASH' | 'SNAPSHOTS'>('TRASH');
  const [backups, setBackups] = useState<DeletedPOBackup[]>(() => storageService.getDeletedPOBackups());
  const [snapshots, setSnapshots] = useState<POSnapshot[]>(() => storageService.getPOSnapshots());
  const [searchTerm, setSearchTerm] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [newSnapshotLabel, setNewSnapshotLabel] = useState('');
  const [showCreateSnapshot, setShowCreateSnapshot] = useState(false);

  const reloadData = () => {
    setBackups(storageService.getDeletedPOBackups());
    setSnapshots(storageService.getPOSnapshots());
  };

  const handleRestoreOne = async (backup: DeletedPOBackup) => {
    const pwd = await customPrompt(`Nhập mật khẩu Admin để khôi phục PO ${backup.masterPo.id}:`, '', true);
    if (!storageService.verifyAdminPassword(pwd, currentUser)) {
      if (pwd !== null) await customAlert('Mật khẩu quản trị không chính xác!');
      return;
    }

    setIsProcessing(true);
    try {
      const res = storageService.restoreDeletedPO(backup.id);
      if (res.success) {
        await customAlert(res.message);
        reloadData();
        onRestored();
      } else {
        await customAlert(res.message);
      }
    } finally {
      setIsProcessing(false);
    }
  };

  const handleRestoreAll = async () => {
    if (backups.length === 0) return;
    const pwd = await customPrompt(`BẠN ĐANG CHỌN KHÔI PHỤC TẤT CẢ ${backups.length} PO TỪ THÙNG RÁC.\nNhập mật khẩu quản trị để tiếp tục:`, '', true);
    if (!storageService.verifyAdminPassword(pwd, currentUser)) {
      if (pwd !== null) await customAlert('Mật khẩu quản trị không chính xác!');
      return;
    }

    setIsProcessing(true);
    try {
      const res = storageService.restoreAllDeletedPOs();
      await customAlert(res.message);
      reloadData();
      onRestored();
    } finally {
      setIsProcessing(false);
    }
  };

  const handleClearTrash = async () => {
    if (backups.length === 0) return;
    const pwd = await customPrompt('CẢNH BÁO: Dọn sạch thùng rác sẽ xóa vĩnh viễn các bản ghi sao lưu!\nNhập mật khẩu Admin để xóa vĩnh viễn:', '', true);
    if (!storageService.verifyAdminPassword(pwd, currentUser)) {
      if (pwd !== null) await customAlert('Mật khẩu quản trị không chính xác!');
      return;
    }
    storageService.clearDeletedPOBackups();
    reloadData();
    await customAlert('Đã làm sạch thùng rác sao lưu PO!');
  };

  const handleCreateSnapshot = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newSnapshotLabel.trim()) return;
    const creator = currentUser?.fullName || currentUser?.username || 'Quản trị viên';
    storageService.takePOSnapshot(newSnapshotLabel.trim(), creator);
    setNewSnapshotLabel('');
    setShowCreateSnapshot(false);
    reloadData();
    await customAlert('Đã tạo điểm chụp an toàn (Snapshot) cho các lệnh sản xuất hiện tại thành công!');
  };

  const handleRollbackSnapshot = async (snap: POSnapshot) => {
    const pwd = await customPrompt(`BẠN ĐANG CHỌN ROLLBACK HỆ THỐNG VỀ BẢN CHỤP: "${snap.label}".\nTất cả PO hiện tại sẽ được thay thế bằng danh sách trong bản chụp này (${snap.poCount} PO).\nNhập mật khẩu quản trị để xác nhận:`, '', true);
    if (!storageService.verifyAdminPassword(pwd, currentUser)) {
      if (pwd !== null) await customAlert('Mật khẩu quản trị không chính xác!');
      return;
    }

    setIsProcessing(true);
    try {
      const res = storageService.rollbackPOSnapshot(snap.id);
      if (res.success) {
        await customAlert(res.message);
        reloadData();
        onRestored();
      } else {
        await customAlert(res.message);
      }
    } finally {
      setIsProcessing(false);
    }
  };

  const filteredBackups = useMemo(() => {
    return backups.filter(b => {
      const query = searchTerm.toLowerCase();
      const pName = (parts.find(p => p.id === b.masterPo.partId)?.name || '').toLowerCase();
      return b.masterPo.id.toLowerCase().includes(query) ||
             b.masterPo.partId.toLowerCase().includes(query) ||
             pName.includes(query) ||
             (b.deletedBy && b.deletedBy.toLowerCase().includes(query));
    });
  }, [backups, searchTerm, parts]);

  return (
    <div className="fixed inset-0 bg-black/60 backdrop-blur-xs z-[100] flex items-center justify-center p-3 sm:p-6">
      <motion.div
        initial={{ opacity: 0, scale: 0.95 }}
        animate={{ opacity: 1, scale: 1 }}
        exit={{ opacity: 0, scale: 0.95 }}
        className="bg-white rounded-3xl max-w-4xl w-full max-h-[90vh] flex flex-col shadow-2xl border border-gray-100 overflow-hidden"
      >
        {/* Header */}
        <div className="p-6 bg-gradient-to-r from-gray-900 via-rose-950 to-gray-900 text-white flex items-center justify-between shrink-0">
          <div className="flex items-center gap-3">
            <div className="w-12 h-12 rounded-2xl bg-rose-500/20 border border-rose-500/40 text-rose-400 flex items-center justify-center">
              <RotateCcw size={24} />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h2 className="text-xl font-black">Khôi Phục & Rollback Lệnh Sản Xuất (PO)</h2>
                <span className="bg-rose-500/30 text-rose-300 text-xs font-bold px-2.5 py-0.5 rounded-full border border-rose-400/30">
                  An toàn chống xóa nhầm
                </span>
              </div>
              <p className="text-xs text-gray-300 mt-1">
                Tự động lưu trữ các PO đã xóa. Khi cần có thể Rollback hoặc khôi phục 1 chạm về hệ thống.
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-2.5 text-gray-400 hover:text-white rounded-xl hover:bg-white/10 transition-colors cursor-pointer"
          >
            <X size={22} />
          </button>
        </div>

        {/* Tab switcher */}
        <div className="bg-gray-50 px-6 py-3 border-b border-gray-200 flex flex-wrap items-center justify-between gap-3 shrink-0">
          <div className="flex items-center gap-2">
            <button
              onClick={() => setActiveTab('TRASH')}
              className={cn(
                "px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-2 cursor-pointer",
                activeTab === 'TRASH' ? "bg-rose-600 text-white shadow-sm" : "bg-white text-gray-700 hover:bg-gray-100 border border-gray-200"
              )}
            >
              <Trash2 size={15} />
              <span>Thùng rác PO đã xóa ({backups.length})</span>
            </button>
            <button
              onClick={() => setActiveTab('SNAPSHOTS')}
              className={cn(
                "px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-2 cursor-pointer",
                activeTab === 'SNAPSHOTS' ? "bg-blue-600 text-white shadow-sm" : "bg-white text-gray-700 hover:bg-gray-100 border border-gray-200"
              )}
            >
              <Database size={15} />
              <span>Bản chụp an toàn (Snapshots) ({snapshots.length})</span>
            </button>
          </div>

          {activeTab === 'TRASH' && backups.length > 0 && (
            <div className="flex items-center gap-2">
              <button
                onClick={handleRestoreAll}
                disabled={isProcessing}
                className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-xl transition-all shadow-sm flex items-center gap-1.5 cursor-pointer disabled:opacity-50"
              >
                <RotateCcw size={14} />
                <span>Khôi phục tất cả ({backups.length})</span>
              </button>
              <button
                onClick={handleClearTrash}
                disabled={isProcessing}
                className="px-3 py-2 bg-gray-200 hover:bg-rose-100 hover:text-rose-700 text-gray-700 text-xs font-bold rounded-xl transition-all flex items-center gap-1.5 cursor-pointer disabled:opacity-50"
                title="Xóa vĩnh viễn thùng rác"
              >
                <Trash2 size={14} />
                <span>Dọn sạch</span>
              </button>
            </div>
          )}

          {activeTab === 'SNAPSHOTS' && (
            <button
              onClick={() => setShowCreateSnapshot(!showCreateSnapshot)}
              className="px-3.5 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-xl transition-all shadow-sm flex items-center gap-1.5 cursor-pointer"
            >
              <Plus size={14} />
              <span>Tạo bản chụp mới</span>
            </button>
          )}
        </div>

        {/* Content body */}
        <div className="flex-1 overflow-y-auto p-6 space-y-4">
          {activeTab === 'TRASH' ? (
            <>
              {backups.length > 0 && (
                <div className="relative">
                  <Search size={16} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-gray-400" />
                  <input
                    type="text"
                    placeholder="Tìm theo mã PO, tên model hoặc người xóa..."
                    value={searchTerm}
                    onChange={e => setSearchTerm(e.target.value)}
                    className="w-full pl-10 pr-4 py-2 bg-gray-50 border border-gray-200 rounded-xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-rose-500/20 focus:border-rose-500"
                  />
                </div>
              )}

              {filteredBackups.length === 0 ? (
                <div className="py-16 text-center text-gray-400">
                  <Trash2 size={48} className="mx-auto mb-3 opacity-30 text-rose-500" />
                  <p className="font-bold text-gray-600">Thùng rác PO hiện đang trống</p>
                  <p className="text-xs text-gray-400 mt-1">Khi bạn xóa bất kỳ PO tổng hoặc PO lẻ nào, hệ thống sẽ lưu vào đây để bảo toàn dữ liệu.</p>
                </div>
              ) : (
                <div className="space-y-3">
                  {filteredBackups.map(backup => {
                    const master = backup.masterPo;
                    const partName = parts.find(p => p.id === master.partId)?.name || master.partId;
                    const childCount = backup.childPos?.length || 0;

                    return (
                      <div
                        key={backup.id}
                        className="p-4 bg-white border border-gray-200 hover:border-rose-300 rounded-2xl shadow-xs transition-all flex flex-col md:flex-row md:items-center justify-between gap-4"
                      >
                        <div className="space-y-1 flex-1">
                          <div className="flex items-center gap-2 flex-wrap">
                            <span className="font-mono font-bold text-sm text-gray-900 bg-gray-100 px-2.5 py-1 rounded-lg">
                              {master.id}
                            </span>
                            {backup.isOddPO ? (
                              <span className="bg-amber-100 text-amber-800 text-[10px] font-bold px-2 py-0.5 rounded uppercase">
                                PO Dư Lẻ
                              </span>
                            ) : (
                              <span className="bg-blue-100 text-blue-800 text-[10px] font-bold px-2 py-0.5 rounded uppercase">
                                PO Kế Hoạch Tổng
                              </span>
                            )}
                            <span className="text-xs text-gray-500 flex items-center gap-1">
                              <Clock size={13} className="text-gray-400" />
                              <span>Xóa lúc: <b>{backup.deletedAtStr}</b></span>
                            </span>
                            {backup.deletedBy && (
                              <span className="text-xs text-gray-500">
                                bởi <b>{backup.deletedBy}</b>
                              </span>
                            )}
                          </div>

                          <div className="text-sm font-bold text-gray-800">
                            {master.partId} - <span className="font-normal text-gray-600">{partName}</span>
                          </div>

                          <div className="flex items-center gap-4 text-xs text-gray-500">
                            <span>Số lượng yêu cầu: <b className="text-gray-900 font-mono">{master.targetQuantity}</b></span>
                            <span>Đã sản xuất: <b className="text-blue-600 font-mono">{master.producedQuantity || 0}</b></span>
                            <span>Công đoạn con đi kèm: <b className="text-orange-600 font-mono">{childCount}</b> PO con</span>
                          </div>
                        </div>

                        <div className="flex items-center gap-2 shrink-0 self-end md:self-center">
                          <button
                            onClick={() => handleRestoreOne(backup)}
                            disabled={isProcessing}
                            className="px-4 py-2.5 bg-rose-600 hover:bg-rose-700 text-white rounded-xl text-xs font-bold transition-all shadow-sm flex items-center gap-1.5 cursor-pointer disabled:opacity-50 active:scale-95"
                          >
                            <RotateCcw size={14} />
                            <span>Khôi phục PO này</span>
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </>
          ) : (
            <>
              {showCreateSnapshot && (
                <form onSubmit={handleCreateSnapshot} className="p-4 bg-blue-50 border border-blue-200 rounded-2xl space-y-3">
                  <div className="flex items-center justify-between">
                    <h4 className="text-xs font-bold text-blue-900 uppercase tracking-wider">Tạo bản chụp an toàn (Snapshot PO)</h4>
                    <button type="button" onClick={() => setShowCreateSnapshot(false)} className="text-blue-500 hover:text-blue-800">
                      <X size={16} />
                    </button>
                  </div>
                  <input
                    type="text"
                    required
                    placeholder="Ghi chú bản chụp (Ví dụ: Trước khi chốt kế hoạch tuần 44...)"
                    value={newSnapshotLabel}
                    onChange={e => setNewSnapshotLabel(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-white border border-blue-200 rounded-xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-blue-500/20"
                  />
                  <div className="flex justify-end gap-2">
                    <button
                      type="button"
                      onClick={() => setShowCreateSnapshot(false)}
                      className="px-3 py-1.5 bg-white border border-gray-200 text-xs font-bold rounded-lg text-gray-600"
                    >
                      Hủy
                    </button>
                    <button
                      type="submit"
                      className="px-4 py-1.5 bg-blue-600 text-white text-xs font-bold rounded-lg shadow-sm hover:bg-blue-700"
                    >
                      Lưu bản chụp
                    </button>
                  </div>
                </form>
              )}

              {snapshots.length === 0 ? (
                <div className="py-16 text-center text-gray-400">
                  <Database size={48} className="mx-auto mb-3 opacity-30 text-blue-500" />
                  <p className="font-bold text-gray-600">Chưa có bản chụp an toàn nào</p>
                  <p className="text-xs text-gray-400 mt-1">Bạn có thể tạo bản chụp để lưu toàn bộ trạng thái danh sách PO tại một thời điểm.</p>
                </div>
              ) : (
                <div className="space-y-3">
                  {snapshots.map(snap => (
                    <div
                      key={snap.id}
                      className="p-4 bg-white border border-gray-200 hover:border-blue-300 rounded-2xl shadow-xs transition-all flex flex-col sm:flex-row sm:items-center justify-between gap-3"
                    >
                      <div className="space-y-1">
                        <div className="flex items-center gap-2">
                          <span className="font-bold text-sm text-gray-900">{snap.label}</span>
                          <span className="bg-blue-100 text-blue-800 text-[10px] font-bold px-2 py-0.5 rounded-full font-mono">
                            {snap.poCount} PO
                          </span>
                        </div>
                        <div className="text-xs text-gray-500 flex items-center gap-2">
                          <span>Thời gian: <b>{snap.createdAtStr}</b></span>
                          <span>•</span>
                          <span>Tạo bởi: <b>{snap.creator || 'Hệ thống'}</b></span>
                        </div>
                      </div>

                      <button
                        onClick={() => handleRollbackSnapshot(snap)}
                        disabled={isProcessing}
                        className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-xl transition-all shadow-sm flex items-center gap-1.5 cursor-pointer disabled:opacity-50 self-end sm:self-auto"
                      >
                        <RotateCcw size={14} />
                        <span>Rollback về bản này</span>
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </>
          )}
        </div>

        {/* Footer */}
        <div className="p-4 bg-gray-50 border-t border-gray-100 flex justify-end shrink-0">
          <button
            onClick={onClose}
            className="px-5 py-2.5 bg-gray-900 hover:bg-gray-800 text-white text-xs font-bold rounded-xl transition-all shadow-sm cursor-pointer"
          >
            Đóng
          </button>
        </div>
      </motion.div>
    </div>
  );
}

function ProductionOrderView({ parts, onUpdate, productionOrders, onNavigateToOddPO, currentUser }: { parts: Part[], onUpdate: () => void, productionOrders: ProductionOrder[], onNavigateToOddPO?: () => void, currentUser?: UserAccount | null }) {
  const [orders, setOrders] = useState<ProductionOrder[]>(productionOrders);
  const [activeTab, setActiveTab] = useState<'STANDARD' | 'ODD' | 'REPAIR'>('STANDARD');
  const [selectedPart, setSelectedPart] = useState("");
  const [quantity, setQuantity] = useState(0);
  const [targetCompletion, setTargetCompletion] = useState("");
  const [searchTerm, setSearchTerm] = useState("");
  const [statusFilter, setStatusFilter] = useState<'ALL' | 'PENDING' | 'IN_PROGRESS' | 'PAUSED' | 'COMPLETED'>('ALL');
  const [estimatedStart, setEstimatedStart] = useState<number | null>(null);
  const [showDeleteModal, setShowDeleteModal] = useState<string | null>(null);
  const [showEditModal, setShowEditModal] = useState<{ id: string, qty: number } | null>(null);
  const [showEditTimeModal, setShowEditTimeModal] = useState<{ id: string, timeStr: string } | null>(null);
  const [showRollbackModal, setShowRollbackModal] = useState(false);
  const [deletedBackupsCount, setDeletedBackupsCount] = useState<number>(() => storageService.getDeletedPOBackups().length);
  const [password, setPassword] = useState("");
  const [expandedMasterPos, setExpandedMasterPos] = useState<Set<string>>(new Set());
  const [expandedModels, setExpandedModels] = useState<Set<string>>(new Set());
  const [showExportPlanModal, setShowExportPlanModal] = useState(false);
  const [exportPlanDate, setExportPlanDate] = useState(() => format(new Date(), 'yyyy-MM-dd'));
  const [excelImportList, setExcelImportList] = useState<ProductionOrder[] | null>(null);
  const [missingBomList, setMissingBomList] = useState<MissingBomModelInfo[]>([]);
  const [excelImportStats, setExcelImportStats] = useState<{ grandTotal: number; validTotal: number; missingTotal: number } | null>(null);
  const [modalTab, setModalTab] = useState<'VALID' | 'MISSING_BOM'>('VALID');
  const [isImportingExcel, setIsImportingExcel] = useState(false);
  const [isBatchCreating, setIsBatchCreating] = useState(false);
  const [poFeedback, setPoFeedback] = useState<{ message: string; type: 'success' | 'error' } | null>(null);

  const handleExcelPOUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsImportingExcel(true);
    setModalTab('VALID');
    try {
      // 1. Lọc theo Model đã có trong Định mức Model (Model BOM) của hệ thống
      const systemModels = Array.from(new Set(storageService.getModelBOM().map(b => b.modelId)));
      const result = await importExcelToPO(file, systemModels);
      
      const missing = result.missingBomModels || [];
      setMissingBomList(missing);
      setExcelImportList(result.orders);
      setExcelImportStats({
        grandTotal: result.grandTotalFileQty,
        validTotal: result.validTotalQty,
        missingTotal: result.missingTotalQty
      });

      // Nếu có bất kỳ model nào thiếu định mức BOM, tự động chuyển sang tab MISSING_BOM và cảnh báo
      if (missing.length > 0) {
        setModalTab('MISSING_BOM');
        alert(`CẢNH BÁO: Phát hiện ${missing.length} mã model trong file Excel CHƯA CÓ ĐỊNH MỨC BOM!\n\nTheo quy định hệ thống: Phải khai báo định mức BOM đầy đủ cho 100% các model có trong file mới được phép tạo Lệnh Sản Xuất.`);
      } else if (result.orders.length === 0) {
        alert('Không tìm thấy dữ liệu số lượng kế hoạch sản xuất (> 0) trong các cột ngày của file Excel!');
      }
    } catch (err: any) {
      alert(err?.message || 'Lỗi khi đọc file Excel!');
    } finally {
      setIsImportingExcel(false);
      e.target.value = '';
    }
  };

  const handleConfirmBatchCreatePO = () => {
    if (!excelImportList || excelImportList.length === 0) return;

    // 1. BẮT BUỘC: Nếu có bất kỳ model nào thiếu định mức thì KHÔNG CHO TẠO, phải đủ mới cho tạo
    if (missingBomList && missingBomList.length > 0) {
      setModalTab('MISSING_BOM');
      alert(`KHÔNG THỂ TẠO LỆNH SẢN XUẤT!\n\nHiện tại có ${missingBomList.length} mã Model chưa có định mức BOM (${missingBomList.map(m => m.modelId).slice(0, 5).join(', ')}${missingBomList.length > 5 ? '...' : ''}).\n\nYêu cầu bắt buộc: Phải khai báo đủ định mức Model (BOM) cho tất cả các mã trước khi tạo lệnh sản xuất!`);
      return;
    }

    // 2. CHỐNG NHÂN BẢN / TRÙNG LẶP: Kiểm tra PO tổng kế hoạch cùng ngày cùng tháng, cùng số lượng
    const currentOrders = storageService.getProductionOrders();
    const existingMasterOrders = currentOrders.filter(o => !o.masterPoId && !o.isOddPO);

    const duplicateItems: { modelId: string; dateStr: string; qty: number; existingId: string }[] = [];

    const getDayMonthKey = (ts?: number) => {
      if (!ts) return '';
      const d = new Date(ts);
      return `${d.getDate()}/${d.getMonth() + 1}/${d.getFullYear()}`;
    };

    for (const po of excelImportList) {
      const planDayMonth = getDayMonthKey(po.excelRequiredDate || po.expectedCompletionTime);
      const isDuplicated = existingMasterOrders.find(existPo => {
        // Cùng model
        if (existPo.partId !== po.partId) return false;
        // Cùng số lượng
        if (existPo.targetQuantity !== po.targetQuantity) return false;
        // Cùng ngày cùng tháng
        const existDayMonth = getDayMonthKey(existPo.excelRequiredDate || existPo.expectedCompletionTime);
        return existDayMonth === planDayMonth;
      });

      if (isDuplicated) {
        duplicateItems.push({
          modelId: po.partId,
          dateStr: planDayMonth,
          qty: po.targetQuantity,
          existingId: isDuplicated.id
        });
      }
    }

    if (duplicateItems.length > 0) {
      const sampleDuplicates = duplicateItems.slice(0, 3).map(d => `- Model ${d.modelId} (Ngày ${d.dateStr}, SL: ${d.qty}) trùng với [${d.existingId}]`).join('\n');
      const moreMsg = duplicateItems.length > 3 ? `\nvà ${duplicateItems.length - 3} lệnh khác...` : '';
      alert(`CẢNH BÁO TRÙNG LẶP / NHÂN BẢN!\n\nPhát hiện ${duplicateItems.length} kế hoạch sản xuất trong file đã tồn tại trên hệ thống (cùng ngày/tháng, cùng mã model và cùng số lượng):\n\n${sampleDuplicates}${moreMsg}\n\nHệ thống đã chặn tạo trùng lặp để tránh nhân đôi lệnh xóa rất mất công! Vui lòng kiểm tra lại danh sách hoặc xóa lệnh cũ trước nếu muốn tạo lại.`);
      return;
    }

    setIsBatchCreating(true);
    setTimeout(async () => {
      try {
        const batchItems = excelImportList.map(po => ({
          modelId: po.partId,
          quantity: po.targetQuantity,
          expectedCompletionTime: po.expectedCompletionTime || Date.now(),
          customNote: po.note,
          excelRequiredDate: po.excelRequiredDate
        }));

        const createdMasters = await storageService.createBatchMasterPOs(batchItems);
        onUpdate();
        setExcelImportList(null);
        setPoFeedback({
          message: `✓ Đã nhập và tạo thành công ${createdMasters.length} Lệnh Sản Xuất (PO) từ Excel kèm toàn bộ PO con đã sinh ra đầy đủ và tự động đồng bộ siêu tốc lên Supabase!`,
          type: 'success'
        });
        setTimeout(() => setPoFeedback(null), 6000);
      } catch (err: any) {
        setPoFeedback({
          message: err?.message || 'Có lỗi xảy ra khi tạo các Lệnh Sản Xuất!',
          type: 'error'
        });
      } finally {
        setIsBatchCreating(false);
      }
    }, 50);
  };

  useEffect(() => {
    setOrders(productionOrders);
  }, [productionOrders]);

  useEffect(() => {
    setTargetCompletion(format(new Date(), "yyyy-MM-dd'T'HH:mm"));
  }, []);

  // Update estimatedStart directly based on storageService calculation
  useEffect(() => {
    if (selectedPart && quantity > 0 && targetCompletion) {
      const endTs = new Date(targetCompletion).getTime();
      const calculatedStartTs = storageService.previewMasterPOStart(selectedPart, quantity, endTs);
      setEstimatedStart(calculatedStartTs);
    } else {
      setEstimatedStart(null);
    }
  }, [selectedPart, quantity, targetCompletion]);

  const toggleExpand = (masterPoId: string) => {
    const newExpanded = new Set(expandedMasterPos);
    if (newExpanded.has(masterPoId)) {
      newExpanded.delete(masterPoId);
    } else {
      newExpanded.add(masterPoId);
    }
    setExpandedMasterPos(newExpanded);
  };

  const handleUpdateSubPoQty = () => {
    if (storageService.verifyAdminPassword(password, currentUser)) {
      if (showEditModal) {
        storageService.updateSubPoQty(showEditModal.id, showEditModal.qty);
        onUpdate();
        setShowEditModal(null);
        setPassword("");
      }
    } else {
      alert('Mật khẩu quản trị không chính xác!');
    }
  };

  const handleUpdateSubPoTime = () => {
    if (storageService.verifyAdminPassword(password, currentUser)) {
      if (showEditTimeModal && showEditTimeModal.timeStr) {
        storageService.updateSubPoTime(showEditTimeModal.id, new Date(showEditTimeModal.timeStr).getTime());
        onUpdate();
        setShowEditTimeModal(null);
        setPassword("");
      }
    } else {
      alert('Mật khẩu quản trị không chính xác!');
    }
  };

  const handleCompleteMasterPO = (masterPoId: string) => {
    storageService.forceCompleteMasterPO(masterPoId);
    onUpdate();
  };

  const handleCreatePO = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedPart || quantity <= 0) return;
    
    try {
      const endTs = targetCompletion ? new Date(targetCompletion).getTime() : Date.now();
      const createdMaster = await storageService.createMasterPO(selectedPart, quantity, endTs);
      // OPTIMISTIC INSTANT UPDATE: Cập nhật giao diện ngay lập tức 0ms lag
      onUpdate();
      setSelectedPart("");
      setQuantity(0);
      setPoFeedback({
        message: `✓ Đã tạo thành công Lệnh PO [${createdMaster.id}] với toàn bộ PO con sinh ra đầy đủ! Đang đồng bộ siêu tốc lên Supabase.`,
        type: 'success'
      });
      setTimeout(() => setPoFeedback(null), 6000);
    } catch (err) {
      setPoFeedback({
        message: err instanceof Error ? err.message : 'Lỗi khi tạo PO',
        type: 'error'
      });
    }
  };

  const handleDeletePO = () => {
    if (storageService.verifyAdminPassword(password, currentUser)) {
      if (showDeleteModal) {
        const deletedBy = currentUser?.fullName || currentUser?.username || 'Admin';
        storageService.deletePO(showDeleteModal, deletedBy);
        setDeletedBackupsCount(storageService.getDeletedPOBackups().length);
        onUpdate();
        setShowDeleteModal(null);
        setPassword("");
        alert('Đã xóa PO thành công! Lệnh sản xuất đã được tự động sao lưu vào Thùng rác (bạn có thể Khôi phục / Rollback bất cứ lúc nào nếu cần).');
      }
    } else {
      alert('Mật khẩu quản trị không chính xác!');
    }
  };

  const subOrdersByMasterId = useMemo(() => {
    const map = new Map<string, ProductionOrder[]>();
    for (const o of orders) {
      if (o.masterPoId) {
        const list = map.get(o.masterPoId) || [];
        list.push(o);
        map.set(o.masterPoId, list);
      }
    }
    return map;
  }, [orders]);

  // Group orders by masterPoId with optimized filtering and searching
  const masterOrders = useMemo(() => {
    const partsMap = new Map(parts.map(p => [p.id, p]));
    const searchLower = searchTerm.toLowerCase();
    
    return orders
      .filter(o => !o.masterPoId)
      .filter(master => {
        const part = partsMap.get(master.partId);
        const partName = part?.name || "";
        
        return (
          master.id.toLowerCase().includes(searchLower) ||
          master.partId.toLowerCase().includes(searchLower) ||
          partName.toLowerCase().includes(searchLower)
        );
      });
  }, [orders, parts, searchTerm]);

  const statusCounts = useMemo(() => {
    const counts = { ALL: 0, PENDING: 0, IN_PROGRESS: 0, PAUSED: 0, COMPLETED: 0 };
    masterOrders.filter(o => activeTab === 'REPAIR' ? o.id.startsWith('REPAIR') : !o.id.startsWith('REPAIR')).forEach(master => {
      let overallStatus = master.status;
      if (master.id.startsWith('REPAIR')) {
          const subOrders = subOrdersByMasterId.get(master.id) || [];
          const allCompleted = subOrders.length > 0 && subOrders.every(so => so.status === 'COMPLETED');
          const hasInProgress = subOrders.some(so => so.status === 'IN_PROGRESS');
          overallStatus = allCompleted ? 'COMPLETED' : (hasInProgress ? 'IN_PROGRESS' : 'PENDING');
      }
      counts[overallStatus as keyof typeof counts] = (counts[overallStatus as keyof typeof counts] || 0) + 1;
      counts.ALL += 1;
    });
    return counts;
  }, [masterOrders, activeTab, subOrdersByMasterId]);

  const filteredMasterOrders = useMemo(() => {
    return masterOrders.filter(o => {
      if (activeTab === 'REPAIR') return o.id.startsWith('REPAIR');
      if (activeTab === 'ODD') return Boolean(o.isOddPO || o.id.startsWith('PO-LE') || (o.note && (o.note.includes('Tồn Lẻ') || o.note.includes('Dư Lẻ'))));
      return !o.id.startsWith('REPAIR') && !o.isOddPO && !o.id.startsWith('PO-LE') && !(o.note && (o.note.includes('Tồn Lẻ') || o.note.includes('Dư Lẻ')));
    }).filter(master => {
      if (statusFilter === 'ALL') return true;
      const subOrders = subOrdersByMasterId.get(master.id) || [];
      let overallStatus = master.status;
      if (master.id.startsWith('REPAIR')) {
          const allCompleted = subOrders.length > 0 && subOrders.every(so => so.status === 'COMPLETED');
          const hasInProgress = subOrders.some(so => so.status === 'IN_PROGRESS');
          overallStatus = allCompleted ? 'COMPLETED' : (hasInProgress ? 'IN_PROGRESS' : 'PENDING');
      }
      return overallStatus === statusFilter;
    });
  }, [masterOrders, activeTab, statusFilter, subOrdersByMasterId]);

  const groupedOrders = useMemo(() => {
    const grouped = new Map<string, ProductionOrder[]>();
    filteredMasterOrders.forEach(o => {
      const group = grouped.get(o.partId) || [];
      group.push(o);
      grouped.set(o.partId, group);
    });
    return Array.from(grouped.entries()).map(([partId, masters]) => ({
      partId,
      masters,
    }));
  }, [filteredMasterOrders]);

  const toggleExpandModel = (partId: string) => {
    const newExpanded = new Set(expandedModels);
    if (newExpanded.has(partId)) newExpanded.delete(partId);
    else newExpanded.add(partId);
    setExpandedModels(newExpanded);
  };
  
  return (
    <div className="space-y-8">
      {/* Create PO Form */}
      <div className="bg-white p-8 rounded-3xl border border-gray-200 shadow-sm">
        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 mb-6">
          <h2 className="text-2xl font-bold flex items-center gap-3">
            <ClipboardList className="text-blue-600" />
            Tạo Lệnh Sản Xuất (PO Tổng)
          </h2>
          <div className="flex items-center gap-3 flex-wrap">
            {onNavigateToOddPO && (
              <button
                type="button"
                onClick={onNavigateToOddPO}
                className="flex items-center gap-2 bg-amber-600 hover:bg-amber-700 text-white px-5 py-3 rounded-2xl font-bold text-sm cursor-pointer shadow-md shadow-amber-200 transition-all active:scale-95"
                title="Tạo PO Dư Lẻ cho từng công đoạn"
              >
                <Layers size={18} />
                <span>Tạo PO Dư Lẻ</span>
              </button>
            )}

            <button
              type="button"
              onClick={() => {
                setDeletedBackupsCount(storageService.getDeletedPOBackups().length);
                setShowRollbackModal(true);
              }}
              className="flex items-center gap-2 bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 px-4 py-3 rounded-2xl font-bold text-sm cursor-pointer transition-all active:scale-95 relative"
              title="Khôi phục hoặc rollback lại các PO bị xóa nhầm"
            >
              <RotateCcw size={18} />
              <span>Khôi Phục PO (Rollback)</span>
              {deletedBackupsCount > 0 && (
                <span className="bg-rose-600 text-white text-[11px] font-black px-2 py-0.5 rounded-full shadow-xs">
                  {deletedBackupsCount}
                </span>
              )}
            </button>

            <button
              type="button"
              onClick={() => {
                const systemModels = Array.from(new Set(storageService.getModelBOM().map(b => b.modelId)));
                downloadSamplePOExcelTemplate(systemModels);
              }}
              className="flex items-center gap-2 bg-blue-50 hover:bg-blue-100 text-blue-700 border border-blue-200 px-4 py-3 rounded-2xl font-bold text-sm cursor-pointer transition-all active:scale-95"
              title="Tải file Excel mẫu đúng chuẩn cấu trúc để nhập liệu"
            >
              <Download size={18} />
              <span>Tải File Mẫu Excel</span>
            </button>

            <label className="flex items-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white px-5 py-3 rounded-2xl font-bold text-sm cursor-pointer shadow-md shadow-emerald-200 transition-all active:scale-95">
              <FileSpreadsheet size={18} />
              <span>{isImportingExcel ? 'Đang đọc file Excel...' : 'Upload Excel Kế Hoạch PO'}</span>
              <input 
                type="file" 
                accept=".xlsx, .xls" 
                className="hidden" 
                onChange={handleExcelPOUpload}
                disabled={isImportingExcel}
              />
            </label>
          </div>
        </div>

        {poFeedback && (
          <div className={`mb-6 p-4 rounded-2xl border flex items-center justify-between gap-3 text-sm font-bold shadow-sm transition-all ${
            poFeedback.type === 'success'
              ? 'bg-emerald-50 text-emerald-800 border-emerald-200'
              : 'bg-red-50 text-red-800 border-red-200'
          }`}>
            <div className="flex items-center gap-2">
              <span>{poFeedback.type === 'success' ? '✓' : '⚠️'}</span>
              <span>{poFeedback.message}</span>
            </div>
            <button 
              type="button" 
              onClick={() => setPoFeedback(null)} 
              className="text-gray-400 hover:text-gray-600 font-bold ml-2 cursor-pointer"
            >
              ✕
            </button>
          </div>
        )}

        <form onSubmit={handleCreatePO} className="space-y-6">
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6 items-end">
            <div className="space-y-2">
              <label className="text-sm font-bold uppercase opacity-50">Chọn Model</label>
              <SearchableSelect 
                options={Array.from(new Set(storageService.getModelBOM().map(b => b.modelId))).map(id => ({ 
                  id, 
                  label: parts.find(p => p.id === id)?.name || id 
                }))}
                value={selectedPart}
                onChange={(val) => {
                  setSelectedPart(val);
                }}
                placeholder="Chọn model..."
              />
            </div>
            <div className="space-y-2">
              <label className="text-sm font-bold uppercase opacity-50">Thời gian mong muốn hoàn thành</label>
              <input 
                type="datetime-local"
                value={targetCompletion}
                onChange={e => setTargetCompletion(e.target.value)}
                className="w-full p-4 p-y-5 rounded-xl border-2 border-gray-100 font-bold text-lg focus:border-blue-600 outline-none"
              />
            </div>
            <div className="space-y-2 col-span-1 md:col-span-2 lg:col-span-1">
              <label className="text-sm font-bold uppercase opacity-50">Số lượng (pcs)</label>
              <input 
                type="number"
                step="any"
                value={quantity || ''}
                onChange={e => setQuantity(parseFloat(e.target.value) || 0)}
                className="w-full p-5 rounded-xl border-2 border-gray-100 font-bold text-lg focus:border-blue-600 outline-none"
                placeholder="Nhập số lượng..."
              />
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6 items-center pt-4 border-t border-gray-50">
            <div className="space-y-1">
              <label className="text-xs font-bold uppercase opacity-30">Dự kiến bắt đầu cơ khí</label>
              <div className="flex items-center gap-4">
                <div className="bg-gray-50 px-6 py-3 rounded-xl border border-gray-100 flex items-center gap-3">
                  <Clock className="text-blue-600" size={20} />
                  {estimatedStart ? (
                    <span className="text-xl font-bold text-blue-700">
                      {format(new Date(estimatedStart), 'HH:mm - dd/MM/yyyy')}
                    </span>
                  ) : (
                    <span className="text-xl font-bold text-gray-300">--:--</span>
                  )}
                </div>
                {estimatedStart && (
                  <p className="text-sm font-medium text-gray-500 italic">
                    * Tính toán tự động theo định mức sản xuất
                  </p>
                )}
              </div>
            </div>
            <div className="flex justify-end">
              <button 
                type="submit"
                disabled={!selectedPart || quantity <= 0}
                className="bg-blue-600 text-white px-12 py-5 rounded-xl font-bold uppercase tracking-widest hover:bg-blue-700 transition-all shadow-lg shadow-blue-200 disabled:opacity-50 disabled:shadow-none"
              >
                TẠO LỆNH PO
              </button>
            </div>
          </div>
        </form>
      </div>

      {/* PO List */}
      <div className="bg-white rounded-3xl border border-gray-200 shadow-sm overflow-hidden">
        <div className="p-8 border-b border-gray-100 flex flex-col md:flex-row justify-between items-center gap-6">
          <div className="flex flex-col gap-4">
            <h2 className="text-2xl font-bold">Danh sách Lệnh Sản Xuất</h2>
            <div className="flex bg-gray-100 p-1 rounded-xl w-full sm:w-fit overflow-x-auto no-scrollbar whitespace-nowrap">
              <button
                onClick={() => setActiveTab('STANDARD')}
                className={cn(
                  "px-3 sm:px-6 py-1.5 sm:py-2 rounded-lg font-bold text-xs sm:text-sm uppercase transition-all shrink-0 cursor-pointer",
                  activeTab === 'STANDARD' 
                    ? "bg-white text-blue-600 shadow-sm" 
                    : "text-gray-500 hover:text-gray-700"
                )}
              >
                Sản xuất chính
              </button>
              <button
                onClick={() => setActiveTab('ODD')}
                className={cn(
                  "px-3 sm:px-6 py-1.5 sm:py-2 rounded-lg font-bold text-xs sm:text-sm uppercase transition-all shrink-0 cursor-pointer",
                  activeTab === 'ODD' 
                    ? "bg-white text-amber-600 shadow-sm" 
                    : "text-gray-500 hover:text-gray-700"
                )}
              >
                PO Dư Lẻ
              </button>
              <button
                onClick={() => setActiveTab('REPAIR')}
                className={cn(
                  "px-3 sm:px-6 py-1.5 sm:py-2 rounded-lg font-bold text-xs sm:text-sm uppercase transition-all shrink-0 cursor-pointer",
                  activeTab === 'REPAIR' 
                    ? "bg-white text-red-600 shadow-sm" 
                    : "text-gray-500 hover:text-gray-700"
                )}
              >
                Kế hoạch bù (Repair)
              </button>
            </div>
            {/* Status Filter */}
            <div className="flex gap-2 flex-wrap text-sm font-bold mt-2">
              {(['ALL', 'PENDING', 'IN_PROGRESS', 'PAUSED', 'COMPLETED'] as const).map(status => (
                <button
                  key={status}
                  onClick={() => setStatusFilter(status)}
                  className={cn(
                    "px-4 py-1.5 rounded-full border transition-all uppercase tracking-tight flex items-center gap-2",
                    statusFilter === status
                      ? "bg-gray-800 text-white border-gray-800"
                      : "bg-white text-gray-500 border-gray-200 hover:border-gray-400"
                  )}
                >
                  <span>
                    {status === 'ALL' ? 'Tất cả' :
                      status === 'PENDING' ? 'Chờ sản xuất' :
                      status === 'IN_PROGRESS' ? 'Đang chạy' :
                      status === 'PAUSED' ? 'Tạm dừng' : 'Hoàn thành'
                    }
                  </span>
                  <span className={cn(
                    "px-1.5 py-0.5 rounded text-[10px] bg-white/20",
                    statusFilter !== status && "bg-gray-100 text-gray-600"
                  )}>
                    {statusCounts[status]}
                  </span>
                </button>
              ))}
            </div>
          </div>
          <div className="flex flex-col sm:flex-row items-center gap-3 w-full md:w-auto">
            <button
              onClick={() => setShowExportPlanModal(true)}
              className="w-full sm:w-auto flex items-center justify-center gap-2 px-4 py-3 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-bold uppercase text-xs md:text-sm transition-all shadow-md shadow-emerald-200 active:scale-95 whitespace-nowrap cursor-pointer"
              title="Xuất file Excel kế hoạch chi tiết (KHSX Laser, Chấn Dập, Hàn, Sơn, Dán Kính)"
            >
              <FileSpreadsheet size={18} />
              <span>Xuất Kế Hoạch Chi Tiết</span>
            </button>
            <div className="relative w-full sm:w-72">
              <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
              <input 
                type="text"
                placeholder="Tìm kiếm mã PO, Model..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full pl-11 pr-4 py-3 bg-gray-50 border border-gray-200 rounded-xl font-medium focus:bg-white focus:border-blue-600 outline-none transition-all"
              />
            </div>
          </div>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead className="bg-gray-50 text-sm font-bold uppercase tracking-wider text-gray-500">
              <tr>
                <th className="px-8 py-5">Mã PO</th>
                <th className="px-8 py-5">Linh kiện / Thành phẩm</th>
                <th className="px-8 py-5">Công đoạn</th>
                <th className="px-8 py-5 text-right">Kế hoạch</th>
                <th className="px-8 py-5 text-right">Thực tế</th>
                <th className="px-8 py-5 text-right">Đã xuất</th>
                <th className="px-8 py-5">Tiến độ</th>
                <th className="px-8 py-5">Bắt đầu</th>
                <th className="px-8 py-5">Dự kiến</th>
                <th className="px-8 py-5">Trạng thái</th>
                <th className="px-8 py-5 text-center">Thao tác</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {groupedOrders.map((group, groupIdx) => {
                const modelPartName = parts.find(p => p.id === group.partId)?.name || group.partId;
                const isModelExpanded = expandedModels.has(group.partId);

                return (
                  <React.Fragment key={`group-${group.partId}-${groupIdx}`}>
                    <tr 
                      className="bg-gray-50 hover:bg-gray-100 cursor-pointer transition-colors group border-b-2 border-gray-200"
                      onClick={() => toggleExpandModel(group.partId)}
                    >
                      <td colSpan={11} className="px-8 py-4">
                        <div className="flex items-center gap-4">
                          <div className={cn("transition-transform duration-200", isModelExpanded ? "rotate-90" : "rotate-0")}>
                            <ChevronRight size={20} className="text-gray-400" />
                          </div>
                          <span className="font-black text-gray-900 text-lg uppercase">{modelPartName} <span className="text-gray-400 text-sm ml-2 font-mono">({group.partId})</span></span>
                          <span className="bg-white px-3 py-1 rounded-full text-xs font-bold text-gray-600 shadow-sm border border-gray-200">{group.masters.length} lệnh PO</span>
                        </div>
                      </td>
                    </tr>
                    
                    {isModelExpanded && group.masters.map((master, idx) => {
                      const subOrders = subOrdersByMasterId.get(master.id) || [];
                      const totalTarget = subOrders.reduce((sum, o) => sum + o.targetQuantity, 0);
                      const totalProduced = subOrders.reduce((sum, o) => sum + o.producedQuantity, 0);
                      let overallProgress = totalTarget > 0 ? (totalProduced / totalTarget) * 100 : 0;
                      const isExpanded = expandedMasterPos.has(master.id);
                      
                      const allCompleted = subOrders.length > 0 && subOrders.every(so => so.status === 'COMPLETED');
                      const hasInProgress = subOrders.some(so => so.status === 'IN_PROGRESS');
                      const overallStatus: string = (master.status === 'COMPLETED' || allCompleted) ? 'COMPLETED' : (hasInProgress ? 'IN_PROGRESS' : 'PENDING');
                      
                      if (overallStatus === 'COMPLETED') {
                        overallProgress = 100;
                      }
                      
                      let currentStageText = 'MODEL';
                      let displayStatusText = overallStatus === 'COMPLETED' ? 'HOÀN THÀNH' : overallStatus === 'PENDING' ? 'CHỜ SẢN XUẤT' : overallStatus === 'PAUSED' ? 'TẠM DỪNG' : 'ĐANG CHẠY';
                      let statusBadgeClass = overallStatus === 'COMPLETED' ? "bg-green-100 text-green-700" : overallStatus === 'PENDING' ? "bg-gray-100 text-gray-700" : overallStatus === 'PAUSED' ? "bg-red-100 text-red-700" : "bg-yellow-100 text-yellow-700";

                      if (master.id.startsWith('REPAIR')) {
                        const sortedSubOrders = [...subOrders].sort((a, b) => {
                          return STAGES.findIndex(s => s.id === a.stageId) - STAGES.findIndex(s => s.id === b.stageId);
                        });
                        const activeSubOrder = sortedSubOrders.find(o => o.status !== 'COMPLETED');
                        if (activeSubOrder) {
                          currentStageText = `Đang làm: ${STAGES.find(s => s.id === activeSubOrder.stageId)?.name}`;
                        } else if (sortedSubOrders.length > 0 && sortedSubOrders.every(o => o.status === 'COMPLETED')) {
                          currentStageText = 'Hoàn thành';
                        } else {
                          currentStageText = 'Đang chờ';
                        }
                      }

                      return (
                        <React.Fragment key={`${master.id}-${idx}`}>
                          <tr 
                            className={cn(
                              "bg-blue-50/10 font-bold cursor-pointer hover:bg-blue-50/50 transition-colors group",
                              isExpanded && "bg-blue-50/80"
                            )}
                            onClick={() => toggleExpand(master.id)}
                          >
                      <td className="px-8 py-5 font-mono text-blue-700 flex items-center gap-3 text-lg">
                        <div className={cn("transition-transform duration-200", isExpanded ? "rotate-90" : "rotate-0")}>
                          <ChevronRight size={18} />
                        </div>
                        {master.id}
                        {master.id.startsWith('REPAIR') && (
                          <span className="bg-red-100 text-red-600 px-2 py-0.5 rounded text-[10px] font-black uppercase tracking-tighter shadow-sm animate-pulse">REPAIR</span>
                        )}
                      </td>
                      <td className="px-8 py-5 text-lg">{parts.find(p => p.id === master.partId)?.name || master.partId}</td>
                      <td className="px-8 py-5">
                        {master.id.startsWith('REPAIR') ? (
                          <span className="px-2 py-1 bg-red-100 text-red-700 rounded text-xs font-bold">{currentStageText}</span>
                        ) : (
                          <span className="px-2 py-1 bg-blue-600 text-white rounded text-xs font-bold">MODEL</span>
                        )}
                      </td>
                      <td className="px-8 py-5 text-right text-xl">{master.targetQuantity}</td>
                      <td className="px-8 py-5 text-right text-xl text-blue-600">-</td>
                      <td className="px-8 py-5 text-right text-xl text-orange-600">-</td>
                      <td className="px-8 py-5">
                        <div className="w-full bg-gray-200 rounded-full h-2.5">
                          <div className="bg-blue-600 h-2.5 rounded-full" style={{ width: `${overallProgress}%` }}></div>
                        </div>
                        <span className="text-xs opacity-50 font-bold">{overallProgress.toFixed(1)}%</span>
                      </td>
                      <td className="px-8 py-5">
                        {master.plannedStartTime ? (
                          <div className="flex flex-col">
                            <span className="font-mono text-sm font-bold text-gray-600">
                              {format(new Date(master.plannedStartTime), 'HH:mm')}
                            </span>
                            <span className="text-xs opacity-50">
                              {format(new Date(master.plannedStartTime), 'dd/MM')}
                            </span>
                          </div>
                        ) : (
                          <span className="text-xs font-mono opacity-50 uppercase tracking-tighter">-</span>
                        )}
                      </td>
                      <td className="px-8 py-5">
                        {master.expectedCompletionTime ? (
                          <div className="flex flex-col">
                            <span className="font-mono text-sm font-bold text-blue-600">
                              {format(new Date(master.expectedCompletionTime), 'HH:mm')}
                            </span>
                            <span className="text-xs opacity-50">
                              {format(new Date(master.expectedCompletionTime), 'dd/MM')}
                            </span>
                          </div>
                        ) : (
                          <span className="text-xs font-mono opacity-50 uppercase tracking-tighter">-</span>
                        )}
                      </td>
                      <td className="px-8 py-5">
                        <span className={cn(
                          "px-3 py-1 rounded-full text-xs font-bold",
                          statusBadgeClass
                        )}>
                          {displayStatusText}
                        </span>
                      </td>
                      <td className="px-8 py-5 text-center flex items-center justify-center gap-2">
                        <button 
                          onClick={async (e) => {
                            e.stopPropagation();
                            const pwd = await customPrompt("Nhập mật khẩu để hoàn thành PO tổng này:", "", true);
                            if (storageService.verifyAdminPassword(pwd, currentUser)) {
                              if (await customConfirm("Bạn có muốn hoàn thành kế hoạch này không?")) {
                                handleCompleteMasterPO(master.id);
                              }
                            } else if (pwd !== null) {
                              await customAlert("Mật khẩu không đúng");
                            }
                          }}
                          className="p-2 text-green-600 hover:text-green-800 hover:bg-green-50 rounded-lg transition-all"
                          title="Hoàn thành PO"
                        >
                          <Check size={18} />
                        </button>
                        <button 
                          onClick={(e) => {
                            e.stopPropagation();
                            setShowDeleteModal(master.id);
                          }}
                          className="p-2 text-red-400 hover:text-red-600 hover:bg-red-50 rounded-lg transition-all"
                        >
                          <Trash2 size={18} />
                        </button>
                      </td>
                    </tr>
                    {isExpanded && subOrders.sort((a, b) => {
                      const orderA = STAGES.findIndex(s => s.id === a.stageId);
                      const orderB = STAGES.findIndex(s => s.id === b.stageId);
                      return orderB - orderA; // order from latest stage (Painting) down to earliest (Mechanics like Laser)
                    }).map((sub, subIdx) => {
                      const progress = sub.status === 'COMPLETED' ? 100 : (sub.producedQuantity / sub.targetQuantity) * 100;
                      return (
                        <tr key={`${sub.id}-${subIdx}`} className="text-base text-gray-600 bg-gray-50/50">
                          <td className="px-8 py-4 pl-16 font-mono opacity-50 text-sm">{sub.id}</td>
                          <td className="px-8 py-4 text-xl">
                            {getProcessValue(parts.find(p => p.id === sub.partId)?.name, parts.find(p => p.id === sub.partId), sub.stageId as StageId, 'OUT')}
                          </td>
                          <td className="px-8 py-4">
                            <span className="px-2 py-1 bg-gray-100 text-gray-700 rounded text-xs font-bold">
                              {STAGES.find(s => s.id === sub.stageId)?.name}
                            </span>
                          </td>
                          <td className="px-8 py-4 text-right font-bold">{sub.targetQuantity}</td>
                          <td className="px-8 py-4 text-right font-bold text-blue-600">{sub.producedQuantity}</td>
                          <td className="px-8 py-4 text-right font-bold text-orange-600">{sub.exportedQuantity || 0}</td>
                          <td className="px-8 py-4">
                            <div className="w-full bg-gray-100 rounded-full h-1.5">
                              <div className={cn(
                                "h-1.5 rounded-full",
                                progress >= 100 ? "bg-green-500" : "bg-blue-400"
                              )} style={{ width: `${Math.min(progress, 100)}%` }}></div>
                            </div>
                          </td>
                          <td className="px-8 py-4">
                            {sub.plannedStartTime ? (
                              <div className="flex flex-col">
                                <span className="font-mono text-sm font-bold text-gray-500">
                                  {format(new Date(sub.plannedStartTime), 'HH:mm')}
                                </span>
                                <span className="text-xs opacity-50">
                                  {format(new Date(sub.plannedStartTime), 'dd/MM')}
                                </span>
                              </div>
                            ) : (
                              <span className="text-[10px] opacity-30 italic">-</span>
                            )}
                          </td>
                          <td className="px-8 py-4">
                            {sub.expectedCompletionTime ? (
                              <div className="flex flex-col relative group">
                                <span className="font-mono text-sm font-bold text-blue-600">
                                  {format(new Date(sub.expectedCompletionTime), 'HH:mm')}
                                </span>
                                <span className="text-xs opacity-50">
                                  {format(new Date(sub.expectedCompletionTime), 'dd/MM')}
                                </span>
                                <button
                                  onClick={() => setShowEditTimeModal({ id: sub.id, timeStr: format(new Date(sub.expectedCompletionTime!), "yyyy-MM-dd'T'HH:mm") })}
                                  className="absolute -right-2 top-1/2 -translate-y-1/2 p-1.5 bg-blue-100 text-blue-600 rounded opacity-0 group-hover:opacity-100 transition-opacity"
                                  title="Chỉnh sửa ngày giờ hoàn thành"
                                >
                                  <Edit2 size={12} />
                                </button>
                              </div>
                            ) : (
                              <span className="text-[10px] opacity-30 italic">Chưa có ĐM</span>
                            )}
                          </td>
                            <td className="px-8 py-4">
                              <span className={cn(
                                "px-2 py-1 rounded text-[10px] font-bold uppercase",
                                sub.status === 'COMPLETED' ? "bg-green-100 text-green-700" : sub.status === 'PENDING' ? "bg-gray-100 text-gray-700" : sub.status === 'PAUSED' ? "bg-red-100 text-red-700" : "bg-yellow-100 text-yellow-700"
                              )}>
                                {sub.status === 'COMPLETED' ? 'HOÀN THÀNH' : sub.status === 'PENDING' ? 'CHỜ SX' : sub.status === 'PAUSED' ? 'TẠM DỪNG' : 'ĐANG CHẠY'}
                              </span>
                            </td>
                            <td className="px-8 py-4 text-center">
                              <button
                                  onClick={() => setShowEditModal({ id: sub.id, qty: sub.targetQuantity })}
                                  className="p-2 text-blue-600 hover:bg-blue-50 rounded-lg transition-all"
                                  title="Điều chỉnh số lượng"
                                >
                                <Edit2 size={16} />
                              </button>
                            </td>
                          </tr>
                        );
                      })}
                    </React.Fragment>
                  );
                })}
              </React.Fragment>
            );
          })}
          
          {groupedOrders.length === 0 && (
            <tr>
              <td colSpan={11} className="px-8 py-20 text-center text-gray-400 italic font-bold">Chưa có lệnh sản xuất nào phù hợp với bộ lọc.</td>
            </tr>
          )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Delete Confirmation Modal */}
      <AnimatePresence>
        {showDeleteModal && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="bg-white rounded-2xl p-8 max-w-md w-full shadow-2xl space-y-6"
            >
              <div className="flex items-center gap-4 text-red-600">
                <Trash2 size={32} />
                <h3 className="text-xl font-bold">Xác nhận xóa PO</h3>
              </div>
              <p className="text-sm text-gray-500">Hành động này sẽ xóa PO Tổng và <b>tất cả các PO Con</b> liên quan. Vui lòng nhập mật khẩu để xác nhận.</p>
              <div className="space-y-2">
                <label className="text-xs font-bold uppercase opacity-50">Nhập mật khẩu xác nhận</label>
                <input 
                  type="password"
                  autoComplete="off"
                  value={password}
                  onChange={e => setPassword(e.target.value)}
                  className="w-full p-4 rounded-lg border border-gray-200 focus:border-red-500 outline-none text-lg"
                  placeholder="Nhập mật khẩu..."
                  onKeyDown={(e) => e.key === 'Enter' && handleDeletePO()}
                  autoFocus
                />
              </div>
              <div className="flex gap-3 pt-2">
                <button 
                  onClick={() => { setShowDeleteModal(null); setPassword(''); }}
                  className="flex-1 py-4 bg-gray-100 rounded-lg font-bold text-sm uppercase hover:bg-gray-200 transition-all"
                >
                  Hủy
                </button>
                <button 
                  onClick={handleDeletePO}
                  className="flex-1 py-4 bg-red-600 text-white rounded-lg font-bold text-sm uppercase hover:bg-red-700 transition-all shadow-lg shadow-red-200"
                >
                  Xác nhận xóa
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Edit PO Time Modal */}
      <AnimatePresence>
        {showEditTimeModal && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="bg-white rounded-2xl p-8 max-w-md w-full shadow-2xl space-y-6"
            >
              <div className="flex items-center gap-4 text-blue-600">
                <Edit2 size={32} />
                <h3 className="text-xl font-bold">Sửa ngày giờ hoàn thành PO</h3>
              </div>
              <div className="space-y-4">
                <div className="space-y-2">
                  <label className="text-xs font-bold uppercase opacity-50">Thời gian hoàn thành dự kiến</label>
                  <input 
                    type="datetime-local"
                    value={showEditTimeModal.timeStr}
                    onChange={e => setShowEditTimeModal({...showEditTimeModal, timeStr: e.target.value})}
                    className="w-full p-4 rounded-lg border border-gray-200 focus:border-blue-500 outline-none text-lg font-bold"
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-bold uppercase opacity-50">Mật khẩu xác nhận (Admin)</label>
                  <input 
                    type="password"
                    autoComplete="off"
                    value={password}
                    onChange={e => setPassword(e.target.value)}
                    className="w-full p-4 rounded-lg border border-gray-200 focus:border-blue-500 outline-none text-lg"
                    placeholder="Nhập mật khẩu..."
                    onKeyDown={(e) => e.key === 'Enter' && handleUpdateSubPoTime()}
                  />
                </div>
              </div>
              <div className="flex gap-3 pt-2">
                <button 
                  onClick={() => { setShowEditTimeModal(null); setPassword(''); }}
                  className="flex-1 py-4 bg-gray-100 rounded-lg font-bold text-sm uppercase hover:bg-gray-200 transition-all"
                >
                  Hủy
                </button>
                <button 
                  onClick={handleUpdateSubPoTime}
                  disabled={!showEditTimeModal.timeStr || !password}
                  className="flex-1 py-4 bg-blue-600 text-white rounded-lg font-bold text-sm uppercase hover:bg-blue-700 transition-all shadow-lg disabled:opacity-50"
                >
                  Cập nhật
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Edit PO Quantity Modal */}
      <AnimatePresence>
        {showEditModal && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="bg-white rounded-2xl p-8 max-w-md w-full shadow-2xl space-y-6"
            >
              <div className="flex items-center gap-4 text-blue-600">
                <Edit2 size={32} />
                <h3 className="text-xl font-bold">Sửa số lượng PO Con</h3>
              </div>
              <div className="space-y-4">
                <div className="space-y-2">
                  <label className="text-xs font-bold uppercase opacity-50">Số lượng mới</label>
                  <input 
                    type="number"
                    min="1"
                    value={showEditModal.qty}
                    onChange={e => setShowEditModal({...showEditModal, qty: parseInt(e.target.value) || 0})}
                    className="w-full p-4 rounded-lg border border-gray-200 focus:border-blue-500 outline-none text-lg font-bold"
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-bold uppercase opacity-50">Mật khẩu xác nhận (Admin)</label>
                  <input 
                    type="password"
                    autoComplete="off"
                    value={password}
                    onChange={e => setPassword(e.target.value)}
                    className="w-full p-4 rounded-lg border border-gray-200 focus:border-blue-500 outline-none text-lg"
                    placeholder="Nhập mật khẩu..."
                    onKeyDown={(e) => e.key === 'Enter' && handleUpdateSubPoQty()}
                  />
                </div>
              </div>
              <div className="flex gap-3 pt-2">
                <button 
                  onClick={() => { setShowEditModal(null); setPassword(''); }}
                  className="flex-1 py-4 bg-gray-100 rounded-lg font-bold text-sm uppercase hover:bg-gray-200 transition-all"
                >
                  Hủy
                </button>
                <button 
                  onClick={handleUpdateSubPoQty}
                  disabled={showEditModal.qty <= 0 || !password}
                  className="flex-1 py-4 bg-blue-600 text-white rounded-lg font-bold text-sm uppercase hover:bg-blue-700 transition-all shadow-lg disabled:opacity-50"
                >
                  Cập nhật
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Export Detailed PO Plan Modal */}
      <AnimatePresence>
        {showExportPlanModal && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white rounded-3xl p-8 max-w-lg w-full shadow-2xl space-y-6 relative"
            >
              <button 
                onClick={() => setShowExportPlanModal(false)}
                className="absolute top-5 right-5 text-gray-400 hover:text-gray-700 p-1.5 rounded-full hover:bg-gray-100 transition-colors cursor-pointer"
              >
                <X size={20} />
              </button>

              <div className="flex items-center gap-3 text-emerald-600">
                <FileSpreadsheet size={32} />
                <div>
                  <h3 className="text-xl font-bold uppercase text-gray-900">Xuất Kế Hoạch Chi Tiết</h3>
                  <p className="text-xs text-gray-500 font-medium">Báo cáo kế hoạch sản xuất các công đoạn (Laser, Chấn, Hàn, Sơn, Dán Kính)</p>
                </div>
              </div>

              <div className="space-y-4 bg-gray-50 p-5 rounded-2xl border border-gray-100">
                <div>
                  <label className="block text-xs font-bold uppercase text-gray-600 mb-2">
                    Ngày mốc xuất báo cáo (T0)
                  </label>
                  <input 
                    type="date"
                    value={exportPlanDate}
                    onChange={(e) => setExportPlanDate(e.target.value)}
                    className="w-full bg-white border border-gray-200 p-3.5 rounded-xl font-bold text-gray-800 outline-none focus:border-emerald-500 font-mono text-base cursor-pointer"
                  />
                  <p className="text-xs text-gray-500 mt-2 font-medium">
                    Hệ thống sẽ xuất kế hoạch chi tiết từ 3 ngày trước đến 30 ngày sau mốc ngày này theo từng sheet công đoạn.
                  </p>
                </div>
              </div>

              <div className="flex gap-3 pt-2">
                <button 
                  onClick={() => setShowExportPlanModal(false)}
                  className="flex-1 py-4 bg-gray-100 rounded-xl font-bold text-sm uppercase hover:bg-gray-200 transition-all text-gray-600 cursor-pointer"
                >
                  Hủy
                </button>
                <button 
                  onClick={() => {
                    exportDetailedProductionPlan(
                      storageService.getProductionOrders(),
                      parts,
                      storageService.getInventory(),
                      storageService.getModelBOM(),
                      storageService.getBOMV2(),
                      { 
                        reportDateStr: exportPlanDate,
                        glazingPlans: storageService.getGlazingPlans(),
                        glazingPlanNorms: storageService.getGlazingPlanNorms()
                      }
                    );
                    setShowExportPlanModal(false);
                  }}
                  className="flex-1 py-4 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-bold text-sm uppercase transition-all shadow-lg shadow-emerald-200 flex items-center justify-center gap-2 active:scale-95 cursor-pointer"
                >
                  <FileSpreadsheet size={18} />
                  <span>Tải File Excel</span>
                </button>
              </div>
            </motion.div>
          </div>
        )}

        {/* Modal Import Excel Kế Hoạch PO */}
        {excelImportList && (
          <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
            <motion.div 
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white rounded-3xl p-6 md:p-8 max-w-5xl w-full shadow-2xl border border-gray-100 flex flex-col max-h-[90vh]"
            >
              {/* Header */}
              <div className="flex items-center justify-between pb-4 border-b border-gray-100">
                <div className="flex items-center gap-3">
                  <div className="w-12 h-12 rounded-2xl bg-emerald-100 text-emerald-700 flex items-center justify-center">
                    <FileSpreadsheet size={24} />
                  </div>
                  <div>
                    <h3 className="text-xl font-bold text-gray-900">Xem Trước Kế Hoạch Sản Xuất Từ Excel</h3>
                    <p className="text-xs text-gray-500 font-medium">Tự động phân tách lệnh PO theo từng ngày và lùi 48 giờ thời gian mong muốn hoàn thành</p>
                  </div>
                </div>
                <button 
                  onClick={() => setExcelImportList(null)}
                  className="p-2 hover:bg-gray-100 rounded-xl text-gray-400 hover:text-gray-700 transition-colors"
                >
                  <X size={20} />
                </button>
              </div>

              {/* Summary Stats */}
              <div className="grid grid-cols-2 md:grid-cols-4 gap-4 my-4">
                <div className="bg-emerald-50/80 border border-emerald-100 p-4 rounded-2xl">
                  <span className="text-xs font-bold uppercase text-emerald-600 block mb-1">Số Lệnh PO Tạo Ra</span>
                  <span className="text-2xl font-black text-emerald-800">{excelImportList.length}</span>
                  <span className="text-xs text-emerald-600 ml-1">lệnh ({new Set(excelImportList.map(o => o.partId)).size} model)</span>
                </div>

                <div className="bg-indigo-50/80 border border-indigo-100 p-4 rounded-2xl">
                  <span className="text-xs font-bold uppercase text-indigo-600 block mb-1">Tổng SL Chạy Trong File</span>
                  <span className="text-2xl font-black text-indigo-800">
                    {(excelImportStats?.grandTotal ?? excelImportList.reduce((sum, o) => sum + o.targetQuantity, 0)).toLocaleString('vi-VN')}
                  </span>
                  <span className="text-xs text-indigo-600 ml-1">PCS</span>
                </div>

                <div className="bg-blue-50/80 border border-blue-100 p-4 rounded-2xl">
                  <span className="text-xs font-bold uppercase text-blue-600 block mb-1">SL Hợp Lệ Sẽ Tạo PO</span>
                  <span className="text-2xl font-black text-blue-800">
                    {(excelImportStats?.validTotal ?? excelImportList.reduce((sum, o) => sum + o.targetQuantity, 0)).toLocaleString('vi-VN')}
                  </span>
                  <span className="text-xs text-blue-600 ml-1">PCS</span>
                </div>

                <div className={cn(
                  "p-4 rounded-2xl border",
                  missingBomList.length > 0 
                    ? "bg-rose-50/80 border-rose-200" 
                    : "bg-emerald-50/80 border-emerald-100"
                )}>
                  <span className={cn(
                    "text-xs font-bold uppercase block mb-1",
                    missingBomList.length > 0 ? "text-rose-600" : "text-emerald-600"
                  )}>
                    {missingBomList.length > 0 ? 'Model Thiếu BOM Cần Bổ Sung' : 'Định Mức BOM'}
                  </span>
                  <span className={cn(
                    "text-2xl font-black",
                    missingBomList.length > 0 ? "text-rose-700" : "text-emerald-800"
                  )}>
                    {missingBomList.length > 0 ? `${missingBomList.length} model` : '100% Đầy đủ'}
                  </span>
                  <span className={cn(
                    "text-xs ml-1 font-bold",
                    missingBomList.length > 0 ? "text-rose-600" : "text-emerald-600"
                  )}>
                    {missingBomList.length > 0 ? `(${excelImportStats?.missingTotal.toLocaleString('vi-VN')} PCS)` : 'hợp lệ'}
                  </span>
                </div>
              </div>

              {/* View Mode Switch (Tabs) */}
              <div className="flex items-center justify-between mb-3">
                <div className="flex bg-gray-100 p-1 rounded-xl gap-1">
                  <button
                    onClick={() => setModalTab('VALID')}
                    className={cn(
                      "px-4 py-2 rounded-lg font-bold text-xs uppercase transition-all flex items-center gap-2 cursor-pointer",
                      modalTab === 'VALID' ? "bg-white text-emerald-700 shadow-sm" : "text-gray-500 hover:text-gray-700"
                    )}
                  >
                    <CheckCircle2 size={15} className="text-emerald-600" />
                    <span>Lệnh PO Hợp Lệ ({excelImportList.length})</span>
                  </button>

                  <button
                    onClick={() => setModalTab('MISSING_BOM')}
                    className={cn(
                      "px-4 py-2 rounded-lg font-bold text-xs uppercase transition-all flex items-center gap-2 cursor-pointer",
                      modalTab === 'MISSING_BOM' 
                        ? (missingBomList.length > 0 ? "bg-rose-600 text-white shadow-sm" : "bg-white text-blue-700 shadow-sm") 
                        : (missingBomList.length > 0 ? "text-rose-600 hover:text-rose-800 bg-rose-50" : "text-gray-500 hover:text-gray-700")
                    )}
                  >
                    <AlertCircle size={15} className={missingBomList.length > 0 ? (modalTab === 'MISSING_BOM' ? "text-white" : "text-rose-600") : "text-gray-400"} />
                    <span>Model Chưa Khai Báo Định Mức BOM ({missingBomList.length})</span>
                  </button>
                </div>
              </div>

              {/* Main Content Area */}
              <div className="flex-1 overflow-y-auto border border-gray-100 rounded-2xl min-h-[220px]">
                {modalTab === 'MISSING_BOM' ? (
                  <div className="p-4 space-y-4">
                    {missingBomList.length === 0 ? (
                      <div className="bg-emerald-50 border border-emerald-200 rounded-2xl p-6 text-center">
                        <CheckCircle2 size={40} className="text-emerald-600 mx-auto mb-2" />
                        <h4 className="text-base font-bold text-emerald-800">Tất cả Model đều hợp lệ!</h4>
                        <p className="text-xs text-emerald-600 mt-1">
                          100% Model trong file Excel đều đã được khai báo Định mức Model (Model BOM) đầy đủ trong hệ thống.
                        </p>
                      </div>
                    ) : (
                      <>
                        <div className="bg-rose-50 border border-rose-200 rounded-2xl p-4 flex items-start gap-3">
                          <AlertCircle size={24} className="text-rose-600 shrink-0 mt-0.5" />
                          <div>
                            <h4 className="text-sm font-bold text-rose-800">
                              Phát hiện {missingBomList.length} Model trong file Excel chưa có Định mức Model (Model BOM)
                            </h4>
                            <p className="text-xs text-rose-600 mt-1">
                              Các model này sẽ <strong>tạm thời không được tạo Lệnh PO</strong> để tránh lỗi khi phân rã chi tiết linh kiện cho các công đoạn Laser, Chấn, Hàn, Sơn. Vui lòng vào mục <strong>Cài Đặt &gt; Định mức Model</strong> để khai báo bổ sung.
                            </p>
                          </div>
                        </div>

                        <table className="w-full text-left text-sm">
                          <thead className="bg-rose-50/50 text-xs font-bold uppercase text-rose-700 sticky top-0 border-b border-rose-200">
                            <tr>
                              <th className="p-3.5 text-center w-12">STT</th>
                              <th className="p-3.5">Mã Model (Cột B)</th>
                              <th className="p-3.5">Tên Sản Phẩm</th>
                              <th className="p-3.5 text-right">SL Trong File (PCS)</th>
                              <th className="p-3.5">Nguyên Nhân Lỗi / Khắc Phục</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-gray-100 font-medium">
                            {missingBomList.map((m, idx) => (
                              <tr key={m.modelId || idx} className="hover:bg-rose-50/30 transition-colors">
                                <td className="p-3.5 text-center text-gray-400 font-mono">{idx + 1}</td>
                                <td className="p-3.5 font-mono font-bold text-rose-700">{m.modelId}</td>
                                <td className="p-3.5 text-gray-700">{m.name || '-'}</td>
                                <td className="p-3.5 text-right font-bold text-gray-900 font-mono">
                                  {m.totalQty > 0 ? m.totalQty.toLocaleString('vi-VN') : '-'}
                                </td>
                                <td className="p-3.5 text-xs text-rose-600">
                                  <span className="bg-rose-100 text-rose-800 px-2.5 py-1 rounded-md font-semibold">
                                    {m.issueReason}
                                  </span>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </>
                    )}
                  </div>
                ) : (
                  <table className="w-full text-left text-sm">
                    <thead className="bg-gray-50 text-xs font-bold uppercase text-gray-500 sticky top-0 border-b border-gray-200">
                      <tr>
                        <th className="p-3.5 text-center w-12">STT</th>
                        <th className="p-3.5">Mã Lệnh PO</th>
                        <th className="p-3.5">Mã Model (partId)</th>
                        <th className="p-3.5 text-right">Số Lượng (PCS)</th>
                        <th className="p-3.5">Thời Gian Bắt Đầu</th>
                        <th className="p-3.5">TG Mong Muốn Hoàn Thành</th>
                        <th className="p-3.5">Ghi Chú Đợt Chạy</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100 font-medium">
                      {excelImportList.map((po, index) => (
                        <tr key={po.id || index} className="hover:bg-blue-50/40 transition-colors">
                          <td className="p-3.5 text-center text-gray-400 font-mono">{index + 1}</td>
                          <td className="p-3.5 font-mono font-bold text-blue-700">{po.id}</td>
                          <td className="p-3.5">
                            <span className="font-bold text-gray-900">{po.partId}</span>
                            {parts.find(p => p.id === po.partId)?.name && (
                              <span className="text-xs text-gray-500 block">
                                {parts.find(p => p.id === po.partId)?.name}
                              </span>
                            )}
                          </td>
                          <td className="p-3.5 text-right font-black text-emerald-700 text-base">
                            {po.targetQuantity.toLocaleString('vi-VN')}
                          </td>
                          <td className="p-3.5 text-xs font-mono text-gray-600">
                            {po.plannedStartTime ? format(new Date(po.plannedStartTime), 'dd/MM/yyyy HH:mm') : '-'}
                          </td>
                          <td className="p-3.5 text-xs font-mono text-gray-800 font-bold">
                            {po.expectedCompletionTime ? format(new Date(po.expectedCompletionTime), 'dd/MM/yyyy HH:mm') : '-'}
                          </td>
                          <td className="p-3.5 text-xs text-gray-600">
                            <span className="bg-gray-100 text-gray-700 px-2.5 py-1 rounded-md font-medium">
                              {po.note || '-'}
                            </span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>

              {/* Action Buttons */}
              {missingBomList.length > 0 && (
                <div className="bg-rose-50 border border-rose-300 rounded-xl p-3 text-xs text-rose-800 font-bold flex items-center gap-2 mt-3">
                  <AlertCircle size={18} className="text-rose-600 shrink-0" />
                  <span>
                    Chặn tạo lệnh: Có {missingBomList.length} mã model chưa khai báo định mức BOM. Cần bổ sung đầy đủ định mức 100% các model trước khi tạo Lệnh Sản Xuất!
                  </span>
                </div>
              )}
              <div className="flex gap-3 pt-4 border-t border-gray-100 mt-3">
                <button
                  onClick={() => setExcelImportList(null)}
                  className="flex-1 py-3.5 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-xl font-bold text-sm uppercase transition-all cursor-pointer"
                >
                  Đóng
                </button>
                <button
                  onClick={handleConfirmBatchCreatePO}
                  disabled={isBatchCreating || excelImportList.length === 0 || missingBomList.length > 0}
                  className={cn(
                    "flex-2 py-3.5 rounded-xl font-bold text-sm uppercase transition-all shadow-lg flex items-center justify-center gap-2",
                    missingBomList.length > 0 
                      ? "bg-gray-300 text-gray-500 cursor-not-allowed shadow-none" 
                      : "bg-emerald-600 hover:bg-emerald-700 text-white shadow-emerald-200 active:scale-95 cursor-pointer disabled:opacity-50"
                  )}
                  title={missingBomList.length > 0 ? "Chưa thể tạo: Cần bổ sung đầy đủ định mức BOM cho các model" : undefined}
                >
                  <CheckCircle2 size={18} />
                  <span>
                    {isBatchCreating 
                      ? 'Đang tạo lệnh...' 
                      : missingBomList.length > 0 
                        ? `Chưa Đủ Định Mức BOM (${missingBomList.length} Model Thiếu)` 
                        : `Xác Nhận Tạo Hàng Loạt (${excelImportList.length} PO)`}
                  </span>
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* PO Rollback & Recycle Bin Modal */}
      <AnimatePresence>
        {showRollbackModal && (
          <PORollbackModal
            parts={parts}
            currentUser={currentUser}
            onClose={() => setShowRollbackModal(false)}
            onRestored={() => {
              setDeletedBackupsCount(storageService.getDeletedPOBackups().length);
              onUpdate();
            }}
          />
        )}
      </AnimatePresence>
    </div>
  );
}

function LabelHistoryView({ 
  parts, 
  labels: initialLabels, 
  onPrint, 
  onCopy, 
  onRollback, 
  onManualInboundQR, 
  labelSettings, 
  currentUser 
}: { 
  parts: Part[], 
  labels: Transaction[], 
  onPrint: (l: Transaction) => void, 
  onCopy: (t: string) => void, 
  onRollback: () => void, 
  onManualInboundQR: (qrData: string, targetStageId: StageId) => void, 
  labelSettings: any, 
  currentUser?: UserAccount | null, 
  key?: string 
}) {
  const [labels, setLabels] = useState<Transaction[]>(initialLabels);
  const [scannedIds, setScannedIds] = useState<Set<string>>(new Set());
  const [printedIds, setPrintedIds] = useState<Set<string>>(new Set());
  const [showDeleteModal, setShowDeleteModal] = useState<string | null>(null);
  const [showInboundModal, setShowInboundModal] = useState<string | null>(null);
  const [password, setPassword] = useState("");
  const [selectedLabel, setSelectedLabel] = useState<Transaction | null>(null);
  const [isDeleting, setIsDeleting] = useState(false);
  const getPastDateStr = (days: number) => {
    const d = new Date();
    d.setDate(d.getDate() - days);
    const offset = d.getTimezoneOffset();
    const local = new Date(d.getTime() - (offset*60*1000));
    return local.toISOString().split('T')[0];
  };

  const [dateFilter, setDateFilter] = useState(() => getPastDateStr(0));
  const [searchQuery, setSearchQuery] = useState("");
  const [limit, setLimit] = useState(50);
  const [activeTab, setActiveTab] = useState<'PENDING' | 'FINISHED' | 'PAINTING' | 'DISPOSAL' | 'GLAZING'>('PENDING');
  const [adminStageFilter, setAdminStageFilter] = useState<string>('ALL');
  const [stageDirectionFilter, setStageDirectionFilter] = useState<'ALL' | 'INBOUND' | 'OUTBOUND'>('ALL');

  const isAdmin = currentUser?.role === 'ADMIN';
  const canViewAllStages = isAdmin || Boolean(currentUser?.canViewAllStageLabels);
  const canViewOwnStageExported = isAdmin || canViewAllStages || currentUser?.canViewOwnStageExportedLabels !== false;
  const userDefaultStage = useMemo(() => getUserDefaultStage(currentUser), [currentUser]);

  const getLabelTargetStage = (label: Transaction): string => {
    if (label.targetStageId) return label.targetStageId;
    if (label.qrData) {
      const qrParts = label.qrData.split('|');
      if (qrParts.length > 5 && qrParts[5]) {
        return qrParts[5];
      }
    }
    // Check if this component is specifically configured to route to GLAZING from label.stageId
    const cleanId = (label.partId || '').normalize('NFC').trim().toUpperCase().replace(/\s*-\s*(H|CD|S|C|GLZ|OUT|IN)$/i, '');
    const glazingConfigs = storageService.getGlazingConfigs();
    const isGlzTarget = glazingConfigs.some(c => 
      c.sourceStageId === label.stageId && 
      c.partId.normalize('NFC').trim().toUpperCase().replace(/\s*-\s*(H|CD|S|C|GLZ|OUT|IN)$/i, '') === cleanId
    );
    if (isGlzTarget) return 'GLAZING';

    // Flow tiêu chuẩn công đoạn:
    if (label.stageId === 'LASER') return 'BENDING';
    if (label.stageId === 'BENDING') return 'WELDING';
    if (label.stageId === 'WELDING') return 'PAINTING';
    if (label.stageId === 'PAINTING') return 'DCLR';
    if (label.stageId === 'GLAZING') return 'DCLR';
    return '';
  };

  const getStageLabel = (stageId?: string) => {
    if (!stageId) return '';
    const match = STAGES.find(s => s.id === stageId);
    return match ? match.name : stageId;
  };

  const isLabelRelatedToStage = (
    label: Transaction, 
    stage: string, 
    direction: 'ALL' | 'INBOUND' | 'OUTBOUND', 
    allowExported: boolean
  ): boolean => {
    const targetStage = getLabelTargetStage(label);
    const isIncoming = targetStage === stage;
    const isOutgoing = allowExported && label.stageId === stage;

    if (direction === 'INBOUND') return isIncoming;
    if (direction === 'OUTBOUND') return isOutgoing;
    return isIncoming || isOutgoing;
  };

  useEffect(() => {
    setLabels(initialLabels);
  }, [initialLabels]);

  useEffect(() => {
    const allTransactions = storageService.getTransactions();
    
    // Identify which labels have been scanned and printed
    const scanned = new Set<string>();
    const printed = new Set<string>();

    allTransactions.forEach(tx => {
      if (tx.type === 'STAGE_IN' && tx.qrData) {
        const qrParts = tx.qrData.split('|');
        if (qrParts.length >= 5) {
          scanned.add(qrParts[4]);
        }
      }
      if (tx.printed) {
        printed.add(tx.id);
      }
    });

    labels.forEach(l => {
      if (l.printed) printed.add(l.id);
    });
    
    setScannedIds(scanned);
    setPrintedIds(printed);
  }, [labels]);

  const handlePrintLabel = (labelToPrint: Transaction) => {
    if (!labelToPrint || !labelToPrint.id) return;
    setPrintedIds(prev => new Set(prev).add(labelToPrint.id));
    setLabels(prev => prev.map(l => l.id === labelToPrint.id ? { ...l, printed: true } : l));
    if (selectedLabel && selectedLabel.id === labelToPrint.id) {
      setSelectedLabel({ ...selectedLabel, printed: true });
    }
    onPrint(labelToPrint);
  };

  const handleRollback = async () => {
    const isAdmin = currentUser?.role === 'ADMIN';
    if (isAdmin || storageService.verifyAdminPassword(password, currentUser)) {
      if (showDeleteModal) {
        setIsDeleting(true);
        try {
          const targetId = showDeleteModal;
          // 1. Rollback locally
          storageService.rollbackTransaction(targetId);

          // 2. Direct Supabase deletion for instant synchronization
          if (isSupabaseConfigured) {
            try {
              const p1 = supabase.from('qr_labels').delete().eq('id', targetId);
              const p2 = supabase.from('wip_transactions').delete().eq('id', targetId);
              await Promise.allSettled([p1, p2]);
            } catch (supaErr) {
              console.warn('[Supabase Sync] Lỗi xóa nhãn trên Cloud:', supaErr);
            }
          }

          // 3. Update local state
          const updatedLabels = storageService.getLabels();
          setLabels(updatedLabels);
          onRollback();
          if (selectedLabel?.id === targetId) setSelectedLabel(null);
          setShowDeleteModal(null);
          setPassword("");
          alert('Đã xóa nhãn QR và đồng bộ dữ liệu với Supabase thành công!');
        } catch (err: any) {
          alert('Lỗi khi xóa nhãn: ' + (err?.message || err));
        } finally {
          setIsDeleting(false);
        }
      }
    } else {
      alert('Mật khẩu Admin không chính xác!');
    }
  };

  const handleManualInbound = () => {
    const isAdmin = currentUser?.role === 'ADMIN';
    if (isAdmin || storageService.verifyAdminPassword(password, currentUser)) {
      if (showInboundModal) {
        const label = labels.find(l => l.id === showInboundModal);
        if (label) {
          const qrParts = label.qrData ? label.qrData.split('|') : [];
          const targetStageId = (label.targetStageId || qrParts[5] || getLabelTargetStage(label)) as StageId;
          if (!targetStageId) {
            alert('Nhãn này không có công đoạn đích được chỉ định. Không thể nhập kho tức thì.');
            return;
          }
          
          try {
            onManualInboundQR(label.qrData || label.id, targetStageId);
            // Update labels status locally
            setScannedIds(prev => {
              const next = new Set(prev);
              next.add(showInboundModal);
              return next;
            });
            setShowInboundModal(null);
            setPassword("");
            alert('Đã nhập kho tức thì thành công!');
          } catch (err: any) {
            alert(err.message || 'Đã xảy ra lỗi khi nhập kho');
          }
        }
      }
    } else {
      alert('Mật khẩu không chính xác!');
    }
  };

  const filteredLabels = labels.filter(label => {
    const isScanned = scannedIds.has(label.id);
    const isPaintingOut = label.stageId === 'PAINTING' && label.type === 'STAGE_OUT';
    const isGlazingOut = label.stageId === 'GLAZING' && label.type === 'STAGE_OUT';
    const isDisposal = label.type === 'DISPOSAL';
    
    // Split logic into tabs
    if (activeTab === 'PENDING') {
      // Pending: Not scanned and NOT specialized labels
      if (isScanned || isPaintingOut || isGlazingOut || isDisposal) return false;

      // Phân quyền theo công đoạn & chiều tem (đích đến hoặc do chính công đoạn xuất đi)
      if (!canViewAllStages) {
        if (!isLabelRelatedToStage(label, userDefaultStage, stageDirectionFilter, canViewOwnStageExported)) {
          return false;
        }
      } else if (adminStageFilter !== 'ALL') {
        if (!isLabelRelatedToStage(label, adminStageFilter, stageDirectionFilter, true)) {
          return false;
        }
      }
    } else if (activeTab === 'PAINTING') {
      // Painting: Only Painting OUT labels
      if (!isPaintingOut) return false;
    } else if (activeTab === 'GLAZING') {
      // Glazing: Only Glazing OUT labels that are NOT draft (already printed)
      if (!isGlazingOut || label.printed === false) return false;
    } else if (activeTab === 'DISPOSAL') {
      // Disposal: Only disposal labels
      if (!isDisposal) return false;
    } else {
      // Finished: Scanned labels (any stage)
      if (!isScanned) return false;
    }

    if (searchQuery) {
      const q = searchQuery.toLowerCase();
      const pName = parts.find(p => p.id === label.partId)?.name?.toLowerCase() || "";
      const matches = label.partId.toLowerCase().includes(q) || 
                      pName.includes(q) || 
                      (label.poId || "").toLowerCase().includes(q) ||
                      label.id.toLowerCase().includes(q);
      if (!matches) return false;
    }

    if (!dateFilter) return true;
    const labelDate = format(label.timestamp, 'yyyy-MM-dd');
    return labelDate === dateFilter;
  });

  // Counts for non-admin user
  const userPendingCounts = useMemo(() => {
    let inbound = 0;
    let outbound = 0;
    let all = 0;

    labels.forEach(label => {
      const isScanned = scannedIds.has(label.id);
      const isPaintingOut = label.stageId === 'PAINTING' && label.type === 'STAGE_OUT';
      const isGlazingOut = label.stageId === 'GLAZING' && label.type === 'STAGE_OUT';
      const isDisposal = label.type === 'DISPOSAL';
      if (isScanned || isPaintingOut || isGlazingOut || isDisposal) return;

      const target = getLabelTargetStage(label);
      const isInc = target === userDefaultStage;
      const isOut = canViewOwnStageExported && label.stageId === userDefaultStage;

      if (isInc) inbound++;
      if (isOut) outbound++;
      if (isInc || isOut) all++;
    });

    return { inbound, outbound, all };
  }, [labels, scannedIds, userDefaultStage, canViewOwnStageExported]);

  // Counts for Admin
  const adminPendingCounts = useMemo(() => {
    let inbound = 0;
    let outbound = 0;
    let all = 0;
    let totalAllPending = 0;

    labels.forEach(label => {
      const isScanned = scannedIds.has(label.id);
      const isPaintingOut = label.stageId === 'PAINTING' && label.type === 'STAGE_OUT';
      const isGlazingOut = label.stageId === 'GLAZING' && label.type === 'STAGE_OUT';
      const isDisposal = label.type === 'DISPOSAL';
      if (isScanned || isPaintingOut || isGlazingOut || isDisposal) return;

      totalAllPending++;

      if (adminStageFilter !== 'ALL') {
        const target = getLabelTargetStage(label);
        const isInc = target === adminStageFilter;
        const isOut = label.stageId === adminStageFilter;

        if (isInc) inbound++;
        if (isOut) outbound++;
        if (isInc || isOut) all++;
      }
    });

    return { inbound, outbound, all, totalAllPending };
  }, [labels, scannedIds, adminStageFilter]);

  const pendingTabCount = useMemo(() => {
    if (!canViewAllStages) {
      if (stageDirectionFilter === 'INBOUND') return userPendingCounts.inbound;
      if (stageDirectionFilter === 'OUTBOUND') return userPendingCounts.outbound;
      return userPendingCounts.all;
    }
    if (adminStageFilter === 'ALL') return adminPendingCounts.totalAllPending;
    if (stageDirectionFilter === 'INBOUND') return adminPendingCounts.inbound;
    if (stageDirectionFilter === 'OUTBOUND') return adminPendingCounts.outbound;
    return adminPendingCounts.all;
  }, [canViewAllStages, stageDirectionFilter, userPendingCounts, adminStageFilter, adminPendingCounts]);

  const displayedLabels = filteredLabels.slice(0, limit);

  return (
    <motion.div 
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, y: -20 }}
      className="grid grid-cols-1 lg:grid-cols-3 gap-8"
    >
      <div className="lg:col-span-1 bg-white rounded-2xl border border-gray-200 shadow-sm overflow-hidden flex flex-col h-[750px]">
        <div className="p-6 border-b border-gray-100 bg-gray-50/50 space-y-4">
          <div className="flex justify-between items-center">
            <div>
              <h2 className="font-bold text-xl tracking-tight">Danh sách nhãn QR</h2>
              <div className="flex items-center gap-1.5 mt-1">
                <p className="text-xs text-gray-400">Tổng số: {labels.length} nhãn</p>
                <span className="text-[10px] bg-amber-50 text-amber-700 px-2 py-0.5 rounded-full font-bold border border-amber-200">
                  Lưu 24h
                </span>
              </div>
            </div>
            <div className="flex flex-col gap-1 items-end">
              <div className="flex items-center gap-1">
                <div className="w-2 h-2 rounded-full bg-green-500" />
                <span className="text-[9px] font-bold opacity-60 uppercase">Đã nhập kho</span>
              </div>
              <div className="flex items-center gap-1">
                <div className="w-2 h-2 rounded-full bg-emerald-500" />
                <span className="text-[9px] font-bold text-emerald-700 uppercase">Đã in nhãn</span>
              </div>
              <div className="flex items-center gap-1">
                <div className="w-2 h-2 rounded-full bg-amber-500" />
                <span className="text-[9px] font-bold opacity-60 uppercase">Chưa in nhãn</span>
              </div>
            </div>
          </div>

          <div className="relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
            <input 
              type="text"
              placeholder="Tìm theo Mã LK, Tên, hoặc Mã PO..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full pl-10 pr-4 py-2 bg-white border border-gray-200 rounded-xl text-sm focus:ring-4 focus:ring-blue-100 focus:border-blue-500 outline-none transition-all"
            />
            {searchQuery && (
              <button 
                onClick={() => setSearchQuery("")}
                className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600"
              >
                <X size={14} />
              </button>
            )}
          </div>

          <div className="flex bg-gray-100 p-1 rounded-xl gap-1 overflow-x-auto no-scrollbar whitespace-nowrap">
            <button 
              onClick={() => setActiveTab('PENDING')}
              className={cn(
                "px-2.5 sm:px-3 py-1.5 sm:py-2 text-[10px] sm:text-xs font-bold uppercase tracking-tight rounded-lg transition-all shrink-0 sm:flex-1 cursor-pointer flex items-center justify-center gap-1",
                activeTab === 'PENDING' ? "bg-white text-blue-600 shadow-sm" : "text-gray-500 hover:text-gray-700"
              )}
            >
              <span>Chờ nhập kho</span>
              <span className={cn(
                "px-1.5 py-0.2 rounded-full text-[9px] font-black",
                activeTab === 'PENDING' ? "bg-blue-100 text-blue-700" : "bg-gray-200 text-gray-600"
              )}>
                {pendingTabCount}
              </span>
            </button>
            <button 
              onClick={() => setActiveTab('FINISHED')}
              className={cn(
                "px-2.5 sm:px-3 py-1.5 sm:py-2 text-[10px] sm:text-xs font-bold uppercase tracking-tight rounded-lg transition-all shrink-0 sm:flex-1 cursor-pointer",
                activeTab === 'FINISHED' ? "bg-white text-green-600 shadow-sm" : "text-gray-500 hover:text-gray-700"
              )}
            >
              Đã nhập kho
            </button>
            <button 
              onClick={() => setActiveTab('PAINTING')}
              className={cn(
                "px-2.5 sm:px-3 py-1.5 sm:py-2 text-[10px] sm:text-xs font-bold uppercase tracking-tight rounded-lg transition-all shrink-0 sm:flex-1 cursor-pointer",
                activeTab === 'PAINTING' ? "bg-white text-purple-600 shadow-sm" : "text-gray-500 hover:text-gray-700"
              )}
            >
              Hàng Sơn
            </button>
            <button 
              onClick={() => setActiveTab('GLAZING')}
              className={cn(
                "px-2.5 sm:px-3 py-1.5 sm:py-2 text-[10px] sm:text-xs font-bold uppercase tracking-tight rounded-lg transition-all shrink-0 sm:flex-1 cursor-pointer",
                activeTab === 'GLAZING' ? "bg-white text-orange-600 shadow-sm" : "text-gray-500 hover:text-gray-700"
              )}
            >
              Dán Kính
            </button>
            <button 
              onClick={() => setActiveTab('DISPOSAL')}
              className={cn(
                "px-2.5 sm:px-3 py-1.5 sm:py-2 text-[10px] sm:text-xs font-bold uppercase tracking-tight rounded-lg transition-all shrink-0 sm:flex-1 cursor-pointer",
                activeTab === 'DISPOSAL' ? "bg-white text-red-600 shadow-sm" : "text-gray-500 hover:text-gray-700"
              )}
            >
              Hàng Hủy
            </button>
          </div>

          {activeTab === 'PENDING' && (
            canViewAllStages ? (
              <div className="bg-slate-50 border border-slate-200 rounded-xl p-2.5 space-y-2">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-1.5 text-[11px] font-bold text-slate-700">
                    <Filter size={13} className="text-slate-500" />
                    <span>{isAdmin ? 'Lọc công đoạn (Admin):' : 'Lọc công đoạn (Toàn quyền tất cả CĐ):'}</span>
                  </div>
                  {adminStageFilter !== 'ALL' && (
                    <button 
                      onClick={() => { setAdminStageFilter('ALL'); setStageDirectionFilter('ALL'); }}
                      className="text-[10px] text-blue-600 font-bold hover:underline cursor-pointer"
                    >
                      Xem tất cả ({adminPendingCounts.totalAllPending})
                    </button>
                  )}
                </div>
                <div className="flex flex-wrap gap-1">
                  {[
                    { id: 'ALL', label: 'Tất cả' },
                    { id: 'BENDING', label: 'Chấn/Dập' },
                    { id: 'WELDING', label: 'Hàn' },
                    { id: 'PAINTING', label: 'Sơn' },
                    { id: 'GLAZING', label: 'Dán Kính' },
                    { id: 'DCLR', label: 'Lắp ráp (DCLR)' },
                    { id: 'LASER', label: 'Cắt Laser' },
                  ].map(item => (
                    <button
                      key={item.id}
                      onClick={() => setAdminStageFilter(item.id)}
                      className={cn(
                        "px-2 py-1 rounded-md text-[10px] font-bold transition-all cursor-pointer",
                        adminStageFilter === item.id 
                          ? "bg-blue-600 text-white shadow-xs" 
                          : "bg-white text-slate-600 hover:bg-slate-200/80 border border-slate-200"
                      )}
                    >
                      {item.label}
                    </button>
                  ))}
                </div>

                {adminStageFilter !== 'ALL' && (
                  <div className="flex flex-wrap items-center gap-1 pt-1.5 border-t border-slate-200/80">
                    <span className="text-[10px] font-bold text-slate-500 mr-1">Hướng tem:</span>
                    <button
                      onClick={() => setStageDirectionFilter('ALL')}
                      className={cn(
                        "px-2 py-0.5 rounded text-[10px] font-bold transition-all cursor-pointer",
                        stageDirectionFilter === 'ALL' 
                          ? "bg-slate-800 text-white shadow-xs" 
                          : "bg-white text-slate-600 hover:bg-slate-100 border border-slate-200"
                      )}
                    >
                      Cả đến & xuất ({adminPendingCounts.all})
                    </button>
                    <button
                      onClick={() => setStageDirectionFilter('INBOUND')}
                      className={cn(
                        "px-2 py-0.5 rounded text-[10px] font-bold transition-all cursor-pointer",
                        stageDirectionFilter === 'INBOUND' 
                          ? "bg-blue-600 text-white shadow-xs" 
                          : "bg-white text-slate-600 hover:bg-slate-100 border border-slate-200"
                      )}
                    >
                      Đích đến ({adminPendingCounts.inbound})
                    </button>
                    <button
                      onClick={() => setStageDirectionFilter('OUTBOUND')}
                      className={cn(
                        "px-2 py-0.5 rounded text-[10px] font-bold transition-all cursor-pointer",
                        stageDirectionFilter === 'OUTBOUND' 
                          ? "bg-emerald-600 text-white shadow-xs" 
                          : "bg-white text-slate-600 hover:bg-slate-100 border border-slate-200"
                      )}
                    >
                      Xuất đi ({adminPendingCounts.outbound})
                    </button>
                  </div>
                )}
              </div>
            ) : (
              <div className="bg-blue-50/80 border border-blue-200 rounded-xl p-2.5 space-y-2">
                <div className="flex items-center justify-between text-xs text-blue-900">
                  <div className="flex items-center gap-1.5 font-medium">
                    <ShieldCheck size={16} className="text-blue-600 shrink-0" />
                    <span>Công đoạn phụ trách:</span>
                    <span className="font-extrabold uppercase bg-blue-600 text-white px-2 py-0.5 rounded text-[11px]">
                      {getStageLabel(userDefaultStage)}
                    </span>
                  </div>
                  <span className="text-[10px] text-blue-600 font-mono font-bold">
                    @{currentUser?.username || 'user'}
                  </span>
                </div>

                {canViewOwnStageExported && (
                  <div className="flex flex-wrap items-center gap-1 pt-1 border-t border-blue-200/60">
                    <button
                      onClick={() => setStageDirectionFilter('ALL')}
                      className={cn(
                        "px-2 py-1 rounded-lg text-[10px] font-bold transition-all cursor-pointer flex items-center gap-1",
                        stageDirectionFilter === 'ALL' 
                          ? "bg-blue-600 text-white shadow-xs" 
                          : "bg-white text-blue-800 border border-blue-200 hover:bg-blue-100/70"
                      )}
                    >
                      <span>Tất cả</span>
                      <span className="opacity-80 font-black font-mono">({userPendingCounts.all})</span>
                    </button>
                    <button
                      onClick={() => setStageDirectionFilter('INBOUND')}
                      className={cn(
                        "px-2 py-1 rounded-lg text-[10px] font-bold transition-all cursor-pointer flex items-center gap-1",
                        stageDirectionFilter === 'INBOUND' 
                          ? "bg-blue-600 text-white shadow-xs" 
                          : "bg-white text-blue-800 border border-blue-200 hover:bg-blue-100/70"
                      )}
                    >
                      <span>Đích đến (Chờ nhập)</span>
                      <span className="opacity-80 font-black font-mono">({userPendingCounts.inbound})</span>
                    </button>
                    <button
                      onClick={() => setStageDirectionFilter('OUTBOUND')}
                      className={cn(
                        "px-2 py-1 rounded-lg text-[10px] font-bold transition-all cursor-pointer flex items-center gap-1",
                        stageDirectionFilter === 'OUTBOUND' 
                          ? "bg-emerald-600 text-white shadow-xs" 
                          : "bg-white text-emerald-800 border border-emerald-200 hover:bg-emerald-100/70"
                      )}
                    >
                      <span>Do mình xuất đi (Kiểm tra in)</span>
                      <span className="opacity-80 font-black font-mono">({userPendingCounts.outbound})</span>
                    </button>
                  </div>
                )}
              </div>
            )
          )}
          
          <div className="relative group">
            <input 
              type="date"
              min={getPastDateStr(6)}
              max={getPastDateStr(0)}
              value={dateFilter}
              onChange={(e) => setDateFilter(e.target.value)}
              className="w-full pl-10 pr-4 py-2 rounded-lg border border-gray-200 focus:border-blue-600 outline-none text-sm font-mono cursor-pointer"
            />
            <div className="absolute -top-10 left-1/2 -translate-x-1/2 whitespace-nowrap bg-gray-800 text-white text-[10px] font-bold py-1 px-3 rounded-lg opacity-0 group-hover:opacity-100 transition-opacity pointer-events-none z-10 block">
              Chỉ xem lịch sử 7 ngày gần nhất
            </div>
            <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
            {dateFilter && (
              <button 
                onClick={() => setDateFilter("")}
                className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600"
              >
                <X size={14} />
              </button>
            )}
          </div>
        </div>
        
        <div className="flex-1 overflow-y-auto divide-y divide-gray-50">
          {filteredLabels.length === 0 ? (
            <div className="p-12 text-center text-gray-400 italic">
              {dateFilter ? `Không tìm thấy nhãn nào trong ngày ${format(new Date(dateFilter), 'dd/MM/yyyy')}` : 'Chưa có nhãn nào được tạo.'}
            </div>
          ) : (
            <>
              {displayedLabels.map(label => {
                const isScanned = scannedIds.has(label.id);
                const isPrinted = printedIds.has(label.id) || Boolean(label.printed);
                const isPaintingOut = label.stageId === 'PAINTING' && !label.targetStageId && label.type === 'STAGE_OUT';
                const targetStage = getLabelTargetStage(label);
                const originName = getStageLabel(label.stageId);
                const targetName = getStageLabel(targetStage);
                const currentStageContext = canViewAllStages ? (adminStageFilter === 'ALL' ? '' : adminStageFilter) : userDefaultStage;
                const isExportedByCurrentStage = currentStageContext && label.stageId === currentStageContext;
                const isArrivingAtCurrentStage = currentStageContext && targetStage === currentStageContext;

                return (
                  <div 
                    key={label.id}
                    onClick={() => setSelectedLabel(label)}
                    className={cn(
                      "p-5 cursor-pointer transition-all hover:bg-gray-50 flex justify-between items-center group relative border-l-4",
                      selectedLabel?.id === label.id && "ring-2 ring-inset ring-blue-600 z-10",
                      isScanned ? "bg-green-50/50 border-green-500" : isPrinted ? "bg-emerald-50/60 border-emerald-500" : "bg-white border-transparent"
                    )}
                  >
                    <div className="flex flex-col gap-1">
                      <div className="flex items-center gap-2">
                        {isScanned ? (
                          <CheckCircle2 size={14} className="text-green-600" />
                        ) : isPrinted ? (
                          <CheckCircle2 size={14} className="text-emerald-600" />
                        ) : label.type === 'DISPOSAL' ? (
                          <Trash2 size={14} className="text-red-600" />
                        ) : null}
                        <span className="font-bold text-sm">{parts.find(p => p.id === label.partId)?.name || label.partId}</span>
                      </div>
                      <div className="flex flex-wrap items-center gap-1.5">
                        {label.poId && (
                          <div className="text-[10px] font-mono font-bold text-blue-600 bg-blue-50 px-1 rounded border border-blue-100">
                            {label.poId}
                          </div>
                        )}
                        {originName && targetName ? (
                          <span className="text-[9px] font-bold px-1.5 py-0.5 bg-indigo-50 text-indigo-700 border border-indigo-200 rounded flex items-center gap-1">
                            <span>{originName}</span>
                            <span className="text-indigo-400 font-black">➔</span>
                            <span className="font-extrabold text-indigo-900">{targetName}</span>
                          </span>
                        ) : targetName ? (
                          <span className="text-[9px] font-bold px-1.5 py-0.5 bg-blue-50 text-blue-700 border border-blue-200 rounded">
                            Đích: {targetName}
                          </span>
                        ) : null}
                        {isExportedByCurrentStage && (
                          <span className="text-[9px] font-black uppercase px-1.5 py-0.5 bg-emerald-100 text-emerald-800 rounded">
                            Xuất đi
                          </span>
                        )}
                        {isArrivingAtCurrentStage && (
                          <span className="text-[9px] font-black uppercase px-1.5 py-0.5 bg-sky-100 text-sky-800 rounded">
                            Chờ nhập
                          </span>
                        )}
                        {isPaintingOut && (
                          <span className="text-[9px] font-bold uppercase px-1.5 py-0.5 bg-purple-100 text-purple-700 rounded">Sơn OUT</span>
                        )}
                        {label.stageId === 'GLAZING' && label.type === 'STAGE_OUT' && (
                          <span className="text-[9px] font-bold uppercase px-1.5 py-0.5 bg-orange-100 text-orange-700 rounded">Dán Kính OUT</span>
                        )}
                        {label.type === 'DISPOSAL' && (
                          <span className="text-[9px] font-bold uppercase px-1.5 py-0.5 bg-red-100 text-red-700 rounded">HÀNG HỦY</span>
                        )}
                      </div>
                      <div className="flex items-center gap-2 text-[10px] font-mono opacity-50">
                        <span>{format(label.timestamp, 'dd/MM HH:mm')}</span>
                        <span>•</span>
                        <span>{label.quantity} {parts.find(p => p.id === label.partId)?.unit}</span>
                      </div>
                    </div>
                    <div className="flex items-center gap-2">
                      {isScanned ? (
                        <span className="text-[9px] font-black uppercase px-2 py-0.5 bg-green-100 text-green-800 border border-green-300 rounded whitespace-nowrap">Đã nhập</span>
                      ) : isPrinted ? (
                        <span className="text-[9px] font-black uppercase px-2 py-0.5 bg-emerald-100 text-emerald-800 border border-emerald-300 rounded whitespace-nowrap flex items-center gap-1 shadow-2xs">
                          <CheckCircle2 size={11} className="text-emerald-600 inline" />
                          Đã in nhãn
                        </span>
                      ) : (
                        <span className="text-[9px] font-bold uppercase px-2 py-0.5 bg-amber-50 text-amber-700 border border-amber-300 rounded whitespace-nowrap">Chưa in</span>
                      )}
                      <button 
                        onClick={(e) => {
                          e.stopPropagation();
                          setShowDeleteModal(label.id);
                        }}
                        className="p-1.5 sm:p-2 text-red-500 hover:text-red-700 hover:bg-red-50 rounded-lg transition-all cursor-pointer opacity-80 sm:opacity-0 group-hover:opacity-100"
                        title="Xóa nhãn QR & hoàn trả kho (Admin)"
                      >
                        <Trash2 size={16} />
                      </button>
                    </div>
                  </div>
                );
              })}
              {limit < filteredLabels.length && (
                <button 
                  onClick={() => setLimit(prev => prev + 50)}
                  className="w-full py-4 text-sm font-bold text-blue-600 hover:bg-blue-50 transition-colors uppercase tracking-widest"
                >
                  Xem thêm ({filteredLabels.length - limit} nhãn còn lại)
                </button>
              )}
            </>
          )}
        </div>
      </div>

      <div className="lg:col-span-2">
        <AnimatePresence mode="wait">
          {selectedLabel ? (
            <motion.div 
              key={selectedLabel.id}
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white p-10 rounded-2xl border-2 border-dashed border-[#141414] flex flex-col items-center space-y-8 h-full shadow-xl"
            >
              <div className="flex justify-between w-full items-center">
                <div className="flex items-center gap-3">
                  <Printer size={24} className="text-[#F27D26]" />
                  <h3 className="font-mono text-sm uppercase font-bold tracking-widest">Xem lại nhãn QR</h3>
                  {(printedIds.has(selectedLabel.id) || selectedLabel.printed) && (
                    <span className="text-[10px] font-black uppercase px-2.5 py-0.5 bg-emerald-100 text-emerald-800 border border-emerald-300 rounded-full flex items-center gap-1 shadow-2xs">
                      <CheckCircle2 size={12} className="text-emerald-600 inline" />
                      ĐÃ IN NHÃN
                    </span>
                  )}
                </div>
                <button onClick={() => setSelectedLabel(null)} className="text-gray-400 hover:text-black transition-colors">
                  <X size={32} />
                </button>
              </div>
              
              <A7QRLabelCard label={selectedLabel} parts={parts} labelSettings={labelSettings} isPrintArea={false} />

              <div className="w-full flex flex-col sm:flex-row gap-3 no-print">
                <button 
                  onClick={() => handlePrintLabel(selectedLabel)}
                  className="flex-1 bg-blue-600 text-white py-4 sm:py-5 rounded-xl flex items-center justify-center gap-2 font-bold text-base uppercase hover:bg-blue-700 transition-all shadow-lg shadow-blue-200 cursor-pointer active:scale-95"
                >
                  <Printer size={22} />
                  <span>{(printedIds.has(selectedLabel.id) || selectedLabel.printed) ? 'In lại nhãn (PDF)' : 'In nhãn (PDF)'}</span>
                </button>
                <button 
                  onClick={() => {
                    const isScanned = scannedIds.has(selectedLabel.id);
                    if (isScanned) {
                      alert('Nhãn này đã được nhập kho rồi!');
                      return;
                    }
                    setShowInboundModal(selectedLabel.id);
                  }}
                  className={cn(
                    "flex-1 py-4 sm:py-5 rounded-xl flex items-center justify-center gap-2 font-bold text-base uppercase transition-all shadow-lg cursor-pointer active:scale-95",
                    scannedIds.has(selectedLabel.id) 
                      ? "bg-gray-100 text-gray-400 cursor-not-allowed shadow-none" 
                      : "bg-green-600 text-white hover:bg-green-700 shadow-green-200"
                  )}
                >
                  <PackagePlus size={22} />
                  <span>Nhập kho ngay</span>
                </button>
                <button 
                  onClick={() => setShowDeleteModal(selectedLabel.id)}
                  className="px-5 py-4 sm:py-5 bg-red-50 text-red-600 hover:bg-red-100 border border-red-200 rounded-xl flex items-center justify-center gap-2 font-bold text-sm uppercase transition-all cursor-pointer active:scale-95"
                  title="Xóa nhãn QR và hoàn trả kho (Admin)"
                >
                  <Trash2 size={20} />
                  <span className="hidden sm:inline">Xóa nhãn (Admin)</span>
                </button>
                <button 
                  onClick={() => onCopy(selectedLabel.qrData || '')}
                  className="px-4 bg-gray-100 text-gray-600 py-4 sm:py-5 rounded-xl flex items-center justify-center gap-2 font-bold text-base uppercase hover:bg-gray-200 transition-all cursor-pointer"
                  title="Sao chép mã QR"
                >
                  <QrCode size={22} />
                </button>
              </div>
            </motion.div>
          ) : (
            <div className="bg-gray-100 rounded-2xl border-2 border-dashed border-gray-300 flex flex-col items-center justify-center h-full p-12 text-center text-gray-400">
              <QrCode size={64} strokeWidth={1} />
              <p className="mt-6 font-medium">Chọn một nhãn từ danh sách bên trái<br />để xem chi tiết và in lại.</p>
            </div>
          )}
        </AnimatePresence>
      </div>

      <AnimatePresence>
        {showDeleteModal && (() => {
          const targetLabel = labels.find(l => l.id === showDeleteModal);
          const targetPart = parts.find(p => p.id === targetLabel?.partId);
          const isAdmin = currentUser?.role === 'ADMIN';

          return (
            <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
              <motion.div 
                initial={{ scale: 0.9, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                exit={{ scale: 0.9, opacity: 0 }}
                className="bg-white rounded-2xl p-6 sm:p-8 max-w-md w-full shadow-2xl space-y-5 border-t-4 border-red-600"
              >
                <div className="flex items-center gap-3.5 text-red-600">
                  <div className="p-3 bg-red-100 text-red-600 rounded-xl">
                    <Trash2 size={28} />
                  </div>
                  <div>
                    <h3 className="text-xl font-bold text-gray-900">Xác nhận xóa nhãn QR</h3>
                    <p className="text-xs text-red-600 font-semibold mt-0.5">Đồng bộ xóa dữ liệu trên Supabase</p>
                  </div>
                </div>

                {targetLabel && (
                  <div className="p-3.5 bg-gray-50 border border-gray-200 rounded-xl text-xs space-y-1.5 font-medium text-gray-700">
                    <div className="flex justify-between">
                      <span className="text-gray-500">Linh kiện:</span>
                      <span className="font-bold text-gray-900">{targetPart?.name || targetLabel.partId}</span>
                    </div>
                    <div className="flex justify-between">
                      <span className="text-gray-500">Mã linh kiện:</span>
                      <span className="font-mono font-bold text-blue-600">{targetLabel.partId}</span>
                    </div>
                    {targetLabel.poId && (
                      <div className="flex justify-between">
                        <span className="text-gray-500">Mã PO:</span>
                        <span className="font-mono font-bold text-gray-800">{targetLabel.poId}</span>
                      </div>
                    )}
                    <div className="flex justify-between">
                      <span className="text-gray-500">Số lượng hoàn kho:</span>
                      <span className="font-mono font-bold text-emerald-600">{targetLabel.quantity} {targetPart?.unit || 'cái'}</span>
                    </div>
                  </div>
                )}

                <div className="text-xs text-gray-600 space-y-1 leading-relaxed bg-amber-50 border border-amber-200 p-3 rounded-xl">
                  <p className="font-bold text-amber-900">Thao tác này sẽ thực hiện:</p>
                  <ul className="list-disc pl-4 space-y-0.5 text-amber-800">
                    <li>Xóa nhãn QR khỏi danh sách.</li>
                    <li><b>Đồng bộ xóa vĩnh viễn trên Supabase</b> (<code className="font-mono text-[11px] bg-white px-1 rounded">qr_labels</code> và <code className="font-mono text-[11px] bg-white px-1 rounded">wip_transactions</code>).</li>
                    <li>Hoàn trả số lượng vào tồn kho ban đầu và cập nhật lại tiến độ PO.</li>
                  </ul>
                </div>

                {isAdmin ? (
                  <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-xl text-xs text-emerald-900 flex items-center gap-2.5">
                    <CheckCircle2 size={18} className="text-emerald-600 shrink-0" />
                    <span>Tài khoản <b>Admin</b> ({currentUser.fullName || currentUser.username}) đã được xác thực quyền xóa.</span>
                  </div>
                ) : (
                  <div className="space-y-2">
                    <label className="text-xs font-bold uppercase text-gray-600">Nhập mật khẩu Admin</label>
                    <input 
                      type="password"
                      value={password}
                      onChange={e => setPassword(e.target.value)}
                      className="w-full p-3.5 rounded-xl border border-gray-200 focus:border-red-500 focus:ring-4 focus:ring-red-100 outline-none text-base"
                      placeholder="Mật khẩu Admin..."
                      onKeyDown={(e) => e.key === 'Enter' && handleRollback()}
                      autoFocus
                    />
                  </div>
                )}

                <div className="flex gap-3 pt-2">
                  <button 
                    type="button"
                    disabled={isDeleting}
                    onClick={() => { setShowDeleteModal(null); setPassword(''); }}
                    className="flex-1 py-3.5 bg-gray-100 rounded-xl font-bold text-sm uppercase hover:bg-gray-200 transition-all cursor-pointer disabled:opacity-50"
                  >
                    Hủy bỏ
                  </button>
                  <button 
                    type="button"
                    disabled={isDeleting}
                    onClick={handleRollback}
                    className="flex-1 py-3.5 bg-red-600 text-white rounded-xl font-bold text-sm uppercase hover:bg-red-700 transition-all shadow-md shadow-red-200 flex items-center justify-center gap-2 cursor-pointer active:scale-95 disabled:opacity-50"
                  >
                    {isDeleting ? 'Đang xóa...' : 'Xác nhận xóa & Đồng bộ'}
                  </button>
                </div>
              </motion.div>
            </div>
          );
        })()}

        {showInboundModal && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="bg-white rounded-2xl p-8 max-w-md w-full shadow-2xl space-y-6 border-t-4 border-green-600"
            >
              <div className="flex items-center gap-4 text-green-600">
                <PackagePlus size={32} />
                <h3 className="text-xl font-bold">Xác nhận nhập kho ngay</h3>
              </div>
              <p className="text-sm text-gray-500">Bạn đang thực hiện nhập kho tức thì dựa trên thông tin nhãn QR mà không cần quét mã. Vui lòng nhập mật khẩu Admin để xác nhận.</p>
              <div className="space-y-2">
                <label className="text-xs font-bold uppercase opacity-50">Nhập mật khẩu Admin</label>
                <input 
                  type="password"
                  value={password}
                  onChange={e => setPassword(e.target.value)}
                  className="w-full p-4 rounded-lg border border-gray-200 focus:border-green-500 outline-none text-lg"
                  placeholder="Vui lòng nhập mật khẩu..."
                  onKeyDown={(e) => e.key === 'Enter' && handleManualInbound()}
                  autoFocus
                />
              </div>
              <div className="flex gap-3 pt-2">
                <button 
                  onClick={() => { setShowInboundModal(null); setPassword(''); }}
                  className="flex-1 py-4 bg-gray-100 rounded-lg font-bold text-sm uppercase hover:bg-gray-200 transition-all"
                >
                  Hủy bỏ
                </button>
                <button 
                  onClick={handleManualInbound}
                  className="flex-1 py-4 bg-green-600 text-white rounded-lg font-bold text-sm uppercase hover:bg-green-700 transition-all"
                >
                  Xác nhận nhập kho
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </motion.div>
  );
}

const exportDailyProductionReport = (transactions: Transaction[], parts: Part[]) => {
  const norms = storageService.getNorms();
  const pos = storageService.getProductionOrders();
  
  const now = new Date();
  const startOfToday = startOfDay(now).getTime();

  // Filter transactions for produced items in the current day
  const todayTransactions = transactions.filter(t => {
    const isProducedAtStage = 
      (t.type === 'STAGE_OUT' && t.location === 'IN') ||
      (t.type === 'STAGE_IN' && t.location === 'OUT');
    return isProducedAtStage && t.timestamp >= startOfToday;
  });

  if (todayTransactions.length === 0) {
    alert('Không có dữ liệu sản xuất xuất kho trong ngày hôm nay.');
    return;
  }

  // Time slot definitions
  const TIME_SLOTS = [
    { label: '7h-8h', start: 7, end: 8 },
    { label: '8h-10h', start: 8, end: 10 },
    { label: '10h-12h', start: 10, end: 12 },
    { label: '13h-15h', start: 13, end: 15 },
    { label: '15h-17h', start: 15, end: 17 },
    { label: '17h-19h', start: 17, end: 19 },
    { label: '19h-20h', start: 19, end: 20 },
  ];

  const getSlot = (timestamp: number) => {
    const d = new Date(timestamp);
    const h = d.getHours();
    for (const slot of TIME_SLOTS) {
      if (h >= slot.start && h < slot.end) return slot.label;
    }
    return 'Khác';
  };

  // Group by Part and Stage
  const processData = new Map<string, {
    partId: string;
    stageId: StageId;
    partName: string;
    targetQty: number;
    producedQty: number;
    norms: number;
    slots: Record<string, number>;
  }>();

  todayTransactions.forEach(t => {
    const key = `${t.partId}_${t.stageId}`;
    if (!processData.has(key)) {
      const part = parts.find(p => p.id === t.partId);
      const stageName = STAGES.find(s => s.id === t.stageId)?.name || t.stageId;
      
      const partPos = pos.filter(po => po.partId === t.partId && po.status === 'IN_PROGRESS' && po.stageId === t.stageId);
      const targetQty = partPos.reduce((sum, po) => sum + po.targetQuantity, 0);

      const norm = norms.find(n => n.partId === t.partId && n.stageId === t.stageId)?.secondsPerUnit || 0;

      processData.set(key, {
        partId: t.partId,
        stageId: t.stageId,
        partName: `${part?.name || t.partId} (${stageName})`,
        targetQty: targetQty || 0,
        producedQty: 0,
        norms: norm,
        slots: {}
      });
    }

    const data = processData.get(key)!;
    data.producedQty += t.quantity;
    
    const slot = getSlot(t.timestamp);
    data.slots[slot] = (data.slots[slot] || 0) + t.quantity;
  });

  const rows = Array.from(processData.values()).map(data => {
    const remaining = data.targetQty > 0 ? Math.max(0, data.targetQty - data.producedQty) : 0;
    const completionRate = data.targetQty > 0 ? (data.producedQty / data.targetQty) : 1;
    
    const row: any = {
      'Model': data.partName,
      'KHSX Ngày': data.targetQty || data.producedQty,
      'Thực Hiện': data.producedQty,
      'Còn Lại': remaining,
      'Tỉ Lệ Hoàn Thành': `${(completionRate * 100).toFixed(1)}%`,
      'HSQĐ': data.norms > 0 ? data.norms : '',
      'KHSX sau QĐ': data.targetQty,
    };

    let totalSlots = 0;
    TIME_SLOTS.forEach(slot => {
      row[slot.label] = data.slots[slot.label] || 0;
      totalSlots += row[slot.label];
    });

    const other = data.slots['Khác'] || 0;
    if (other > 0) {
      row['Khác'] = other;
      totalSlots += other;
    }

    row['Tổng sản lượng'] = data.producedQty;
    row['Ghi chú'] = '';

    return row;
  });

  const totalRow: any = {
    'Model': 'TỔNG CỘNG',
    'KHSX Ngày': rows.reduce((s, r) => s + (r['KHSX Ngày'] || 0), 0),
    'Thực Hiện': rows.reduce((s, r) => s + (r['Thực Hiện'] || 0), 0),
    'Còn Lại': rows.reduce((s, r) => s + (r['Còn Lại'] || 0), 0),
    'Tỉ Lệ Hoàn Thành': '',
    'HSQĐ': '',
    'KHSX sau QĐ': rows.reduce((s, r) => s + (r['KHSX Ngày'] || 0), 0),
  };
  
  TIME_SLOTS.forEach(slot => {
    totalRow[slot.label] = rows.reduce((s, r) => s + (r[slot.label] || 0), 0);
  });
  if (rows.some(r => r['Khác'] !== undefined)) {
    totalRow['Khác'] = rows.reduce((s, r) => s + (r['Khác'] || 0), 0);
  }
  totalRow['Tổng sản lượng'] = totalRow['Thực Hiện'];
  totalRow['Ghi chú'] = '';

  rows.push(totalRow);

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "BaoCaoNgay");
  XLSX.writeFile(wb, `BaoCao_SanLuong_${format(new Date(), 'yyyyMMdd')}.xlsx`);
};

const exportHourlyProductionReport = (transactions: Transaction[], parts: Part[], dateStr: string) => {
  const dclrNorms = storageService.getNorms().filter(n => n.stageId === 'DCLR');
  const dclrNormsMap = new Map(dclrNorms.map(n => [n.partId, n.secondsPerUnit]));
  const glazingHSQD = (partId: string) => getGlazingHSQD(partId, dclrNormsMap, dclrNorms, parts);
  const hourlyPeopleLaser = storageService.getHourlyPeopleLaser(dateStr);
  const hourlyPeoplePainting = storageService.getHourlyPeoplePainting(dateStr);
  const hourlyPeopleGlazing = storageService.getHourlyPeopleGlazing(dateStr);
  const hourlyPeopleBending = storageService.getHourlyPeopleBending(dateStr);
  const hourlyPeopleWelding = storageService.getHourlyPeopleWelding(dateStr);
  
  const bendingWeldingHSQDArray = storageService.getBendingWeldingHSQD();
  const bendingWeldingHSQDMap = new Map(bendingWeldingHSQDArray.map(n => [n.partId.toUpperCase(), n.hsqd]));
  const getBendingWeldingHSQD = (partId: string) => {
    const upper = (partId || '').trim().toUpperCase();
    if (bendingWeldingHSQDMap.has(upper)) return bendingWeldingHSQDMap.get(upper)!;
    const cleanId = upper.replace(/^(BND|WLD)-OUT-/i, '').replace(/\s*-\s*(CHẤN\/DẬP|HÀN|CHẤN|DẬP)$/i, '').trim();
    if (cleanId && bendingWeldingHSQDMap.has(cleanId)) return bendingWeldingHSQDMap.get(cleanId)!;
    const matchedPart = parts.find((p: any) => p.id.toUpperCase() === upper || (cleanId && (p.name.toUpperCase().trim() === cleanId || p.id.toUpperCase() === cleanId)));
    if (matchedPart && bendingWeldingHSQDMap.has(matchedPart.id.toUpperCase())) return bendingWeldingHSQDMap.get(matchedPart.id.toUpperCase())!;
    return 0;
  };

  const laserHSQDArray = storageService.getLaserHSQD();
  const laserHSQDMap = new Map(laserHSQDArray.map(n => [n.partId.toUpperCase(), n.hsqd]));
  const getLaserHSQD = (partId: string) => {
    const upper = (partId || '').trim().toUpperCase();
    if (laserHSQDMap.has(upper)) return laserHSQDMap.get(upper)!;
    const cleanId = upper.replace(/^LSR-OUT-/i, '').replace(/\s*-\s*CẮT LASER$/i, '').trim();
    if (cleanId && laserHSQDMap.has(cleanId)) return laserHSQDMap.get(cleanId)!;
    const matchedPart = parts.find((p: any) => p.id.toUpperCase() === upper || (cleanId && (p.name.toUpperCase().trim() === cleanId || p.id.toUpperCase() === cleanId)));
    if (matchedPart && laserHSQDMap.has(matchedPart.id.toUpperCase())) return laserHSQDMap.get(matchedPart.id.toUpperCase())!;
    return 0;
  };

  const glazingStageHSQDArray = storageService.getGlazingStageHSQD();
  const glazingStageHSQDMap = new Map(glazingStageHSQDArray.map(n => [n.partId.toUpperCase(), n.hsqd]));
  const getGlazingStageHSQD = (partId: string) => {
    const upper = (partId || '').trim().toUpperCase();
    if (glazingStageHSQDMap.has(upper)) return glazingStageHSQDMap.get(upper)!;
    const cleanId = upper.replace(/^GLZ-OUT-/i, '').replace(/\s*-\s*DÁN KÍNH$/i, '').trim();
    if (cleanId && glazingStageHSQDMap.has(cleanId)) return glazingStageHSQDMap.get(cleanId)!;
    const matchedPart = parts.find((p: any) => p.id.toUpperCase() === upper || (cleanId && (p.name.toUpperCase().trim() === cleanId || p.id.toUpperCase() === cleanId)));
    if (matchedPart && glazingStageHSQDMap.has(matchedPart.id.toUpperCase())) return glazingStageHSQDMap.get(matchedPart.id.toUpperCase())!;
    return glazingHSQD(partId);
  };

  const paintingHSQDArray = storageService.getPaintingHSQD();
  const paintingHSQDMap = new Map(paintingHSQDArray.map(n => [n.partId.toUpperCase(), n.hsqd]));
  const getPaintingHSQD = (partId: string) => {
    const upper = (partId || '').trim().toUpperCase();
    if (paintingHSQDMap.has(upper)) return paintingHSQDMap.get(upper)!;
    const cleanId = upper.replace(/^PNT-OUT-/i, '').replace(/\s*-\s*SƠN$/i, '').trim();
    if (cleanId && paintingHSQDMap.has(cleanId)) return paintingHSQDMap.get(cleanId)!;
    const matchedPart = parts.find((p: any) => p.id.toUpperCase() === upper || (cleanId && (p.name.toUpperCase().trim() === cleanId || p.id.toUpperCase() === cleanId)));
    if (matchedPart && paintingHSQDMap.has(matchedPart.id.toUpperCase())) return paintingHSQDMap.get(matchedPart.id.toUpperCase())!;
    return dclrNormsMap.get(partId) || 0;
  };

  const [sY, sM, sD] = dateStr.split('-').map(Number);
  const start = new Date(sY, sM - 1, sD, 0, 0, 0, 0);
  start.setHours(0, 0, 0, 0);
  const startMs = start.getTime();

  const end = new Date(dateStr);
  end.setHours(23, 59, 59, 999);
  const endMs = end.getTime();

  const labels = storageService.getLabels();

  const rangeTx = transactions.filter(t => 
    (t.type === 'STAGE_OUT' && t.location === 'IN') ||
    (t.type === 'STAGE_IN' && t.location === 'OUT') // Năng lực sản xuất includes this
  ).filter(t => t.timestamp >= startMs && t.timestamp <= endMs);

  const wb = XLSX.utils.book_new();

  const stagesToExport = [
    { id: 'LASER', name: 'Cắt Laser' },
    { id: 'BENDING', name: 'Chấn Dập' },
    { id: 'WELDING', name: 'Hàn' },
    { id: 'PAINTING', name: 'Sơn' },
    { id: 'GLAZING', name: 'Dán kính' },
  ];

  const borderStyle = {
    top: { style: "thin", color: { rgb: "000000" } },
    bottom: { style: "thin", color: { rgb: "000000" } },
    left: { style: "thin", color: { rgb: "000000" } },
    right: { style: "thin", color: { rgb: "000000" } }
  };

  const headerStyle = {
    font: { bold: true, sz: 11 },
    fill: { fgColor: { rgb: "FFFF00" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: borderStyle
  };

  const titleStyle = {
    font: { bold: true, sz: 16 },
    alignment: { horizontal: "center", vertical: "center" }
  };

  const dataStyle = {
    font: { sz: 11 },
    alignment: { vertical: "center" },
    border: borderStyle
  };
  

  const daySlots = [
    { label: '6h-8h', start: 6, end: 8, factor: 0.25 },
    { label: '8h-10h', start: 8, end: 10, factor: 0.25 },
    { label: '10h-12h', start: 10, end: 12, factor: 0.25 },
    { label: '13h-15h', start: 13, end: 15, factor: 0.25 },
    { label: '15h-17h', start: 15, end: 17, factor: 0.25 },
    { label: '17h-19h', start: 17, end: 19, factor: 0.25 },
  ];

  const nightSlots = [
    { label: '19h-21h', start: 19, end: 21, factor: 0.25 },
    { label: '21h-23h', start: 21, end: 23, factor: 0.25 },
    { label: '23h-0h', start: 23, end: 24, factor: 0.125 },
    { label: '1h-3h', start: 1, end: 3, factor: 0.25 },
    { label: '3h-5h', start: 3, end: 5, factor: 0.25 },
    { label: '5h-6h', start: 5, end: 6, factor: 0.125 },
  ];

  const shifts = [
    { id: 'DAY', name: 'Ca Ngày', timeSlots: daySlots, hoursDesc: '6h - 19h' },
    { id: 'NIGHT', name: 'Ca Đêm', timeSlots: nightSlots, hoursDesc: '19h - 6h' },
  ];

  const getWorkerCount = (stageId: string, sLabel: string) => {
    if (stageId === 'LASER') return hourlyPeopleLaser[sLabel] !== undefined ? hourlyPeopleLaser[sLabel] : 0;
    if (stageId === 'PAINTING') return hourlyPeoplePainting[sLabel] !== undefined ? hourlyPeoplePainting[sLabel] : 8;
    if (stageId === 'GLAZING') return hourlyPeopleGlazing[sLabel] !== undefined ? hourlyPeopleGlazing[sLabel] : 8;
    if (stageId === 'BENDING') return hourlyPeopleBending[sLabel] !== undefined ? hourlyPeopleBending[sLabel] : 8;
    if (stageId === 'WELDING') return hourlyPeopleWelding[sLabel] !== undefined ? hourlyPeopleWelding[sLabel] : 8;
    return 8;
  };

  const titleDate = `${String(sD).padStart(2, '0')}/${String(sM).padStart(2, '0')}/${sY}`;

  stagesToExport.forEach(stageInfo => {
    const stageTx = rangeTx.filter(t => {
      if (t.stageId !== stageInfo.id) return false;
      if (stageInfo.id === 'LASER' && isLaserScrapItem(t.partId, t.partName)) return false;
      return true;
    });

    const isStandardPOStage = stageInfo.id === 'PAINTING' || stageInfo.id === 'BENDING' || stageInfo.id === 'WELDING' || stageInfo.id === 'LASER';

    shifts.forEach(shift => {
      const timeSlots = shift.timeSlots;
      const isDay = shift.id === 'DAY';

      const shiftTx = stageTx.filter(t => {
        const H = new Date(t.timestamp).getHours();
        return isDay ? (H >= 6 && H < 19) : ((H >= 19 && H < 24) || (H >= 0 && H < 6));
      });

      const partTotals = new Map<string, number>();
      const partSlots = new Map<string, number[]>();

      shiftTx.forEach(t => {
        partTotals.set(t.partId, (partTotals.get(t.partId) || 0) + (t.quantity || 0));
        const H = new Date(t.timestamp).getHours();
        let slotIndex = -1;
        if (isDay) {
          if (H >= 6 && H < 8) slotIndex = 0;
          else if (H >= 8 && H < 10) slotIndex = 1;
          else if (H >= 10 && H < 12) slotIndex = 2;
          else if (H === 12) slotIndex = 2; // Khung nghỉ trưa gộp vào 10h-12h
          else if (H >= 13 && H < 15) slotIndex = 3;
          else if (H >= 15 && H < 17) slotIndex = 4;
          else if (H >= 17 && H < 19) slotIndex = 5;
        } else {
          if (H >= 19 && H < 21) slotIndex = 0;
          else if (H >= 21 && H < 23) slotIndex = 1;
          else if (H >= 23 || H === 0) slotIndex = 2; // 23h-0h và giờ nghỉ đêm 0h-1h gộp vào 23h-0h
          else if (H >= 1 && H < 3) slotIndex = 3;
          else if (H >= 3 && H < 5) slotIndex = 4;
          else if (H >= 5 && H < 6) slotIndex = 5;
        }

        if (slotIndex !== -1) {
          if (!partSlots.has(t.partId)) {
            partSlots.set(t.partId, Array(timeSlots.length).fill(0));
          }
          const arr = partSlots.get(t.partId)!;
          arr[slotIndex] += (t.quantity || 0);
        }
      });

      const activeParts = Array.from(partTotals.keys());

      const sheetData: any[] = [];

      // Title Row
      const titleRowText = `BÁO CÁO SẢN XUẤT THEO CUNG GIỜ ${stageInfo.name.toUpperCase()} - ${shift.name.toUpperCase()} NGÀY ${titleDate}`;
      const titleRow = [titleRowText];

      // Rows 1 and 2 headers
      const row1 = ['Tên linh kiện', 'KHSX Ngày', 'Thực Hiện', 'Còn Lại', 'Tỉ lệ Hoàn\nThành KHSX', 'HSQĐ', 'KHSX sau QĐ', 'Sản lượng theo múi giờ'];
      for (let i = 0; i < timeSlots.length - 1; i++) row1.push('');
      row1.push('Tổng sản\nlượng');

      for (let i = 0; i < row1.length - 1; i++) titleRow.push('');

      const row2 = ['', '', '', '', '', '', ''];
      timeSlots.forEach(s => row2.push(s.label));
      row2.push('');

      sheetData.push(titleRow, row1, row2);

      let sumKHSX = 0;
      let sumTH = 0;
      let sumKHSX_QD = 0;
      let sumSlots = Array(timeSlots.length).fill(0);
      let sumSlots_QD = Array(timeSlots.length).fill(0);
      let totalQD = 0;

      activeParts.sort().forEach(partId => {
        const part = parts.find(p => p.id === partId);
        const partName = part?.name || partId;
        const isLaser = stageInfo.id === 'LASER';
        const isBW = stageInfo.id === 'BENDING' || stageInfo.id === 'WELDING';
        const isGlz = stageInfo.id === 'GLAZING';
        const isPainting = stageInfo.id === 'PAINTING';
        const hsqd = isLaser
          ? getLaserHSQD(partId)
          : (isBW ? getBendingWeldingHSQD(partId) : (isPainting ? getPaintingHSQD(partId) : (isGlz ? getGlazingStageHSQD(partId) : glazingHSQD(partId))));
        const thucHien = partTotals.get(partId) || 0;

        // get KHSX based on Stage
        let khsx = 0;
        if (isStandardPOStage) {
          const orders = storageService.getProductionOrders().filter(o => o.stageId === stageInfo.id && o.partId === partId);
          khsx = orders.filter(o => {
            const dt = o.plannedStartTime || o.createdAt;
            return dt >= startMs && dt <= endMs;
          }).reduce((s, o) => s + (o.targetQuantity || 0), 0);
        } else if (stageInfo.id === 'GLAZING') {
          const glazingPlans = storageService.getGlazingPlans();
          const glazingPlanNorms = storageService.getGlazingPlanNorms();
          const planNorm = glazingPlanNorms.find(n => n.id === partId);
          if (planNorm) {
            khsx = glazingPlans.filter(p => p.modelId === planNorm.appliedModel && (p.plannedStartTime || p.createdAt) >= startMs && (p.plannedStartTime || p.createdAt) <= endMs).reduce((s, p) => s + p.targetQuantity, 0);
          }
        }

        const conLai = khsx - thucHien;
        const khsx_qd = khsx * hsqd;

        sumKHSX += khsx;
        sumTH += thucHien;
        sumKHSX_QD += khsx_qd;

        const r = [
          partName, 
          khsx > 0 ? khsx : '', 
          thucHien > 0 ? thucHien : '',
          conLai !== 0 ? conLai : '',
          khsx > 0 ? (thucHien/khsx) : '',
          hsqd > 0 ? Number(hsqd.toFixed(2)) : '',
          '' // KHSX sau QD de trong for parts
        ];

        const slotsArr = partSlots.get(partId) || Array(timeSlots.length).fill(0);
        slotsArr.forEach((v, i) => {
          r.push(v > 0 ? v : '');
          sumSlots[i] += v;
          sumSlots_QD[i] += (v * hsqd);
        });
        totalQD += (thucHien * hsqd);

        r.push(thucHien > 0 ? thucHien : '');
        sheetData.push(r);
      });

      // Total row
      const totalRow: any[] = ['Tổng số lượng chưa quy đổi', sumKHSX > 0 ? sumKHSX : '', sumTH > 0 ? sumTH : '', sumKHSX - sumTH !== 0 ? sumKHSX - sumTH : '', sumKHSX > 0 ? (sumTH/sumKHSX) : '', '', sumKHSX_QD > 0 ? sumKHSX_QD : ''];
      sumSlots.forEach(s => totalRow.push(s > 0 ? s : ''));
      totalRow.push(sumTH > 0 ? sumTH : '');
      sheetData.push(totalRow);

      // San pham quy doi
      const qdRow: any[] = ['Sản phẩm quy đổi', '', '', '', '', '', ''];
      sumSlots_QD.forEach(s => qdRow.push(Math.round(s)));
      qdRow.push(Math.round(totalQD));
      sheetData.push(qdRow);

      // Target calculations for shift
      const targetQDs = timeSlots.map(s => {
        const workerCount = getWorkerCount(stageInfo.id, s.label);
        const factor = s.factor;
        if (stageInfo.id === 'LASER') {
          return Math.round(workerCount * 250 * factor * 10) / 10;
        }
        if (stageInfo.id === 'BENDING') {
          return Math.round(workerCount * 132.7 * factor * 10) / 10;
        }
        if (stageInfo.id === 'WELDING') {
          return Math.round(workerCount * 77.2 * factor * 10) / 10;
        }
        if (stageInfo.id === 'PAINTING') {
          return Math.round(workerCount * 127.5 * factor * 10) / 10;
        }
        if (stageInfo.id === 'GLAZING') {
          return Math.round(workerCount * 242 * factor * 10) / 10;
        }
        return Math.round(workerCount * 64 * factor * 10) / 10;
      });
      const totalTarget = targetQDs.reduce((a, b) => a + b, 0);

      // Năng suất lao động:
      // Cơ chế mới: Nếu khung giờ đó có ghi người nhưng không tạo ra sản phẩm quy đổi (s == 0)
      // thì tính NSLĐ khung giờ đó bằng 0 (0%).
      // Năng suất lao động của cả ca công đoạn sẽ tính bằng cách trung bình cộng của các khung giờ có ghi người
      // (bao gồm cả các khung giờ 0% nếu có ghi số người).
      const tlRow: any[] = ['Năng suất lao động', '', '', '', '', '', ''];
      const validProductivities: number[] = [];
      sumSlots_QD.forEach((s, i) => {
        const workerCount = getWorkerCount(stageInfo.id, timeSlots[i].label);
        const tgt = targetQDs[i];
        if (workerCount > 0) {
          const prod = tgt > 0 ? (s / tgt) : 0;
          tlRow.push(prod);
          validProductivities.push(prod);
        } else {
          tlRow.push('');
        }
      });
      const avgProductivity = validProductivities.length > 0 ? (validProductivities.reduce((a, b) => a + b, 0) / validProductivities.length) : 0;
      tlRow.push(avgProductivity);
      sheetData.push(tlRow);

      // Nhân sự mỗi giờ
      const nhanSuRow: any[] = ['Nhân sự mỗi giờ', '', '', '', '', '', ''];
      timeSlots.forEach(s => {
        nhanSuRow.push(getWorkerCount(stageInfo.id, s.label));
      });
      nhanSuRow.push('');
      sheetData.push(nhanSuRow);

      // Số lượng sản phẩm quy đổi cần đạt được
      const tgtRow: any[] = ['Số lượng sản phẩm quy đổi cần đạt được', '', '', '', '', '', ''];
      targetQDs.forEach(t => tgtRow.push(t));
      tgtRow.push(totalTarget);
      sheetData.push(tgtRow);

      const ws = XLSX.utils.aoa_to_sheet(sheetData);

      // Formatting
      const mergeArr = [
        { s: {r:0, c:0}, e: {r:0, c:7 + timeSlots.length} }, // Title ROW
        { s: {r:1, c:0}, e: {r:2, c:0} },
        { s: {r:1, c:1}, e: {r:2, c:1} },
        { s: {r:1, c:2}, e: {r:2, c:2} },
        { s: {r:1, c:3}, e: {r:2, c:3} },
        { s: {r:1, c:4}, e: {r:2, c:4} },
        { s: {r:1, c:5}, e: {r:2, c:5} },
        { s: {r:1, c:6}, e: {r:2, c:6} },
        { s: {r:1, c:7}, e: {r:1, c:6 + timeSlots.length} }, // Sản lượng theo múi giờ
        { s: {r:1, c:7 + timeSlots.length}, e: {r:2, c:7 + timeSlots.length} }, // Tổng sản lượng
      ];

      // footer merges
      const frOffset = sheetData.length - 5;
      [1, 2, 3, 4].forEach(i => {
        mergeArr.push({ s: {r: frOffset + i, c: 0}, e: {r: frOffset + i, c: 6}});
      });

      ws['!merges'] = mergeArr;

      const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
      for (let R = 0; R <= range.e.r; ++R) {
        for (let C = 0; C <= range.e.c; ++C) {
          const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
          if (!ws[cellAddress]) ws[cellAddress] = {t: 's', v: ''};

          if (R === 0) {
            ws[cellAddress].s = titleStyle;
          } else if (R === 1 || R === 2) {
            ws[cellAddress].s = headerStyle;
          } else {
            let rowStyle = { ...dataStyle };

            if (R === frOffset) {
              rowStyle = { ...dataStyle, font: { ...dataStyle.font, bold: true, color: { rgb: "0000FF" } } } as any;
            } else if (R >= frOffset + 1 && R <= frOffset + 4) {
              rowStyle = { ...dataStyle, font: { ...dataStyle.font, bold: true, color: { rgb: "0000FF" } } } as any;
              if (R === frOffset + 4) {
                rowStyle = { ...dataStyle, font: { ...dataStyle.font, bold: true } } as any;
              }
              if (R === frOffset + 3) {
                rowStyle = { ...dataStyle, font: { ...dataStyle.font, bold: true }, fill: { fgColor: { rgb: "92D050" } } } as any;
              }
            }

            ws[cellAddress].s = rowStyle;

            if (typeof ws[cellAddress].v === 'number') {
              if (C === 4 && R <= frOffset) {
                ws[cellAddress].z = '0%';
              }
              if (C >= 7 && R === frOffset + 2) {
                ws[cellAddress].z = '0%';
              }
            }
          }
        }
      }

      const cols = [{wch: 35}, {wch: 10}, {wch: 10}, {wch: 10}, {wch: 10}, {wch: 10}, {wch: 12}];
      timeSlots.forEach(() => cols.push({wch: 10}));
      cols.push({wch: 12});
      ws['!cols'] = cols;

      const sheetName = `${stageInfo.name} - ${shift.name}`;
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });
  });

  if (wb.SheetNames.length > 0) {
    XLSX.writeFile(wb, `BaoCao_NangLuc_${format(start, 'yyyyMMdd')}.xlsx`);
  }
};

const exportProductionReportRange = (transactions: Transaction[], parts: Part[], startDateStr: string, endDateStr: string, peoplePerDay: Record<string, number> = {}, mandaysPerDay: Record<string, number> = {}, totalMandays: number = 0, weekName?: string) => {
  const [sYear, sMonth, sDay] = startDateStr.split('-').map(Number);
  const start = new Date(sYear, sMonth - 1, sDay, 0, 0, 0, 0);
  const startMs = start.getTime();

  const [eYear, eMonth, eDay] = endDateStr.split('-').map(Number);
  const end = new Date(eYear, eMonth - 1, eDay, 23, 59, 59, 999);
  const endMs = end.getTime();

  const labels = storageService.getLabels();

  const rangeTx = transactions.filter(t => 
    ((t.type === 'STAGE_OUT' && t.location === 'IN') ||
     (t.type === 'STAGE_IN' && t.location === 'OUT')) && 
    t.timestamp >= startMs && t.timestamp <= endMs
  );

  const wb = XLSX.utils.book_new();

  const dclrNorms = storageService.getNorms().filter(n => n.stageId === 'DCLR');
  const dclrNormsMap = new Map(dclrNorms.map(n => [n.partId, n.secondsPerUnit]));
  const glazingHSQD = (partId: string) => getGlazingHSQD(partId, dclrNormsMap, dclrNorms, parts);

  const stagesToExport = [
    { id: 'LASER', name: 'Cắt laser' },
    { id: 'BENDING', name: 'Chấn dập' },
    { id: 'WELDING', name: 'Hàn' },
    { id: 'PAINTING', name: 'Sơn' },
    { id: 'GLAZING', name: 'Dán kính' },
  ];

  const borderStyle = {
    top: { style: "thin", color: { rgb: "000000" } },
    bottom: { style: "thin", color: { rgb: "000000" } },
    left: { style: "thin", color: { rgb: "000000" } },
    right: { style: "thin", color: { rgb: "000000" } }
  };

  const headerStyle = {
    font: { bold: true, sz: 12, color: { rgb: "FFFFFF" } },
    fill: { fgColor: { rgb: "4F81BD" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: borderStyle
  };

  const titleStyle = {
    font: { bold: true, sz: 16 },
    alignment: { horizontal: "center", vertical: "center" }
  };

  const dataStyle = {
    font: { sz: 11 },
    alignment: { vertical: "center" },
    border: borderStyle
  };
  
  const alignRightStyle = {
    ...dataStyle,
    alignment: { horizontal: "right", vertical: "center" }
  };
  
  const alignCenterStyle = {
    ...dataStyle,
    alignment: { horizontal: "center", vertical: "center" }
  };

  stagesToExport.forEach(stageInfo => {
    const stageTx = rangeTx.filter(t => t.stageId === stageInfo.id);
    
    const partTotals = new Map<string, number>();
    stageTx.forEach(t => {
      partTotals.set(t.partId, (partTotals.get(t.partId) || 0) + t.quantity);
    });

    const rows = Array.from(partTotals.entries()).map(([partId, qty]) => {
      const part = parts.find(p => p.id === partId);
      return {
        'Mã linh kiện': partId,
        'Tên linh kiện': part?.name || partId,
        'Đơn vị': part?.unit || 'cái',
        'Tổng sản lượng': qty
      };
    });

    const sheetData: any[][] = [
      [`Báo cáo sản lượng CÔNG ĐOẠN ${stageInfo.name.toUpperCase()}`],
      [`Từ ngày: ${format(start, 'dd/MM/yyyy')} đến ngày ${format(end, 'dd/MM/yyyy')}`],
      ['STT', 'Mã linh kiện', 'Tên linh kiện', 'Đơn vị', 'Tổng sản lượng']
    ];

    rows.forEach((r, idx) => {
      sheetData.push([
        idx + 1,
        r['Mã linh kiện'],
        r['Tên linh kiện'],
        r['Đơn vị'],
        r['Tổng sản lượng']
      ]);
    });
    
    const grandTotal = rows.reduce((s, r) => s + r['Tổng sản lượng'], 0);
    sheetData.push(['', 'TỔNG CỘNG', '', '', grandTotal]);

    const ws = XLSX.utils.aoa_to_sheet(sheetData);

    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } },
      { s: { r: sheetData.length - 1, c: 1 }, e: { r: sheetData.length - 1, c: 3 } }
    ];

    ws['!cols'] = [
      { wch: 8 },  // STT
      { wch: 30 }, // Mã LK
      { wch: 40 }, // Tên LK
      { wch: 10 }, // Đơn vị
      { wch: 15 }, // Sản lượng
    ];

    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
    for (let R = 0; R <= range.e.r; ++R) {
      for (let C = 0; C <= range.e.c; ++C) {
        const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
        if (!ws[cellAddress]) ws[cellAddress] = { t: 's', v: '' };
        
        if (R === 0) {
          ws[cellAddress].s = titleStyle;
        } else if (R === 1) {
          ws[cellAddress].s = { alignment: { horizontal: "center" }, font: { italic: true } };
        } else if (R === 2) {
          ws[cellAddress].s = headerStyle;
        } else if (R > 2) {
           if (R === sheetData.length - 1) {
             const fontStyle = { bold: true, sz: 12, color: { rgb: "000000" } };
             ws[cellAddress].s = { 
               font: fontStyle, 
               fill: { fgColor: { rgb: "E2E8F0" } },
               alignment: C === 4 ? { horizontal: "right", vertical: "center" } : { horizontal: "center", vertical: "center" },
               border: borderStyle
             };
           } else {
             if (C === 0 || C === 3) ws[cellAddress].s = alignCenterStyle;
             else if (C === 4) ws[cellAddress].s = alignRightStyle;
             else ws[cellAddress].s = dataStyle;
           }
        }
      }
    }

    const sheetName = stageInfo.name.replace(/[^a-zA-Z0-9\u00C0-\u024F\u1E00-\u1EFF ]/g, '').substring(0, 31).trim();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  });

  // ========== BÁO CÁO LẮP RÁP ==========

  const masterOrders = storageService.getProductionOrders().filter(o => !o.masterPoId && !o.id.startsWith('REPAIR'));
  
  const allLapRapParts = new Set<string>();
  const outConfigs = storageService.getGlazingOutConfigs();
  const excludedGlazingParts = new Set(
    outConfigs
      .filter(c => c.finalPartName.toLowerCase().endsWith('dán kính') || c.finalPartName.toLowerCase().endsWith('dán kính '))
      .map(c => 'GLZ-OUT-' + c.finalPartName.toUpperCase().trim())
  );

  dclrNorms.forEach(n => {
    if (n.partId && !excludedGlazingParts.has(n.partId)) allLapRapParts.add(n.partId);
  });
  
  const dayLabels: string[] = [];
  const daysStartMs: number[] = [];
  const daysEndMs: number[] = [];
  let curr = new Date(startMs);
  while (curr.getTime() <= endMs) {
    const dStart = new Date(curr); dStart.setHours(0,0,0,0);
    const dEnd = new Date(curr); dEnd.setHours(23,59,59,999);
    dayLabels.push(`${format(dStart, 'EEE').replace('Mon','Thứ 2').replace('Tue','Thứ 3').replace('Wed','Thứ 4').replace('Thu','Thứ 5').replace('Fri','Thứ 6').replace('Sat','Thứ 7').replace('Sun','Chủ Nhật')} ${format(dStart, 'dd/MM')}`);
    daysStartMs.push(dStart.getTime());
    daysEndMs.push(dEnd.getTime());
    curr.setDate(curr.getDate() + 1);
  }

  const paintingOrders = storageService.getProductionOrders().filter(o => o.stageId === 'PAINTING' && o.status !== 'COMPLETED'); // or all statuses? Wait, all statues.
  // Actually, we want KHSX for any PO created in the day, regardless of completeness.
  const allPaintingOrders = storageService.getProductionOrders().filter(o => o.stageId === 'PAINTING');
  const glazingPlans = storageService.getGlazingPlans();
  const glazingPlanNorms = storageService.getGlazingPlanNorms();

  allPaintingOrders.forEach(po => {
    const time = po.plannedStartTime || po.createdAt;
    if (time >= startMs && time <= endMs) {
      if (!excludedGlazingParts.has(po.partId)) allLapRapParts.add(po.partId);
    }
  });

  glazingPlans.forEach(plan => {
    const time = plan.plannedStartTime || plan.createdAt;
    if (time >= startMs && time <= endMs) {
      const relatedNorms = glazingPlanNorms.filter(n => n.appliedModel === plan.modelId);
      relatedNorms.forEach(n => {
        if (!excludedGlazingParts.has(n.id)) allLapRapParts.add(n.id);
      });
    }
  });

  transactions.filter(t => t.timestamp >= startMs && t.timestamp <= endMs)
    .filter(t => t.type === 'STAGE_OUT')
    .filter(t => t.stageId === 'PAINTING')
    .forEach(t => {
       if (!excludedGlazingParts.has(t.partId)) allLapRapParts.add(t.partId);
    });

  labels.filter(l => l.timestamp >= startMs && l.timestamp <= endMs)
    .filter(l => l.type === 'STAGE_OUT' && l.stageId === 'GLAZING')
    .forEach(l => {
       if (!excludedGlazingParts.has(l.partId)) allLapRapParts.add(l.partId);
    });

  const cleanWeekName = weekName?.replace(/^tuần\s*/i, '') || '';
  const titleText = weekName ? `BÁO CÁO SẢN XUẤT TUẦN ${cleanWeekName.toUpperCase()}` : `BÁO CÁO SẢN XUẤT TỪ ${format(start, 'dd/MM/yyyy')} ĐẾN ${format(end, 'dd/MM/yyyy')}`;
  const lrSheetData: any[][] = [];
  const lrTitleRow: any[] = [titleText];
  const lrRow1: any[] = ['Mã thành phẩm', 'Tên thành phẩm', 'HSQĐ'];
  const lrRow2: any[] = ['', '', ''];
  
  dayLabels.forEach(dl => {
     lrRow1.push(dl, ''); 
     lrRow2.push('KHSX', 'TH');
  });
  
  lrRow1.push('KHSX', 'Tổng TH', 'Chênh lệch');
  lrRow2.push('', '', '');
  
  for(let i=0; i<lrRow1.length-1; i++) lrTitleRow.push('');
  
  lrSheetData.push(lrTitleRow, lrRow1, lrRow2);

  const lapRapRows = Array.from(allLapRapParts).sort().map(partId => {
    const part = parts.find(p => p.id === partId);
    const partName = part?.name || partId;
    const hsqd = Number(glazingHSQD(partId).toFixed(2));
    
    const rowData: any[] = [partId, partName, hsqd];
    let totalKHSX = 0;
    let totalTH = 0;
    
    for (let i = 0; i < daysStartMs.length; i++) {
        const dStart = daysStartMs[i];
        const dEnd = daysEndMs[i];
        
        const dayKHSX_Painting = allPaintingOrders.filter(o => o.partId === partId && (o.plannedStartTime || o.createdAt) >= dStart && (o.plannedStartTime || o.createdAt) <= dEnd).reduce((sum, o) => sum + o.targetQuantity, 0);
        const modelsForPart = glazingPlanNorms.filter(n => n.id === partId).map(n => n.appliedModel);
        const dayKHSX_Glazing = modelsForPart.length > 0 
           ? glazingPlans.filter(p => modelsForPart.includes(p.modelId) && (p.plannedStartTime || p.createdAt) >= dStart && (p.plannedStartTime || p.createdAt) <= dEnd).reduce((sum, p) => sum + p.targetQuantity, 0)
           : 0;

        const dayKHSX = dayKHSX_Painting + dayKHSX_Glazing;
        
        // Lấy số lượng thực hiện của các công đoạn Lắp Ráp (Sơn, Dán Kính) khi hoàn thành chi tiết đưa vào kho OUT
        const dayTH = transactions.filter(t => 
           t.timestamp >= dStart && t.timestamp <= dEnd && 
           t.partId === partId && t.type === 'STAGE_OUT' && t.location === 'IN' && 
           (t.stageId === 'PAINTING' || t.stageId === 'GLAZING')
        ).reduce((sum, t) => sum + Math.abs(t.quantity), 0);
        
        rowData.push(dayKHSX > 0 ? dayKHSX : '', dayTH > 0 ? dayTH : '');
        totalKHSX += dayKHSX;
        totalTH += dayTH;
    }
    
    const diff = totalTH - totalKHSX;
    let diffDisplay: any = '';
    if (diff < 0) {
      diffDisplay = `(${Math.abs(diff)})`;
    } else if (diff > 0) {
      diffDisplay = diff;
    }
    
    rowData.push(totalKHSX > 0 ? totalKHSX : '', totalTH > 0 ? totalTH : '', diffDisplay);
    return { rowData, hsqd, totalKHSX, totalTH, diff };
  });

  const activeLapRapRows = lapRapRows.filter(r => r.totalKHSX > 0 || r.totalTH > 0);
  activeLapRapRows.forEach(r => lrSheetData.push(r.rowData));

  const footerUnconverted: any[] = ['Tổng sản phẩm chưa quy đổi', '', ''];
  const footerConverted: any[] = ['Tổng sản phẩm quy đổi/ngày', '', ''];
  const footerPeople: any[] = ['Tổng số người', '', ''];
  const footerCapacity: any[] = ['Năng lực sản xuất dây chuyền (8h)', '', ''];
  const footerRatio: any[] = ['KHSX / NLSX', '', ''];
  const footerDailyProductivity: any[] = ['Năng suất lao động ngày', '', ''];

  let grandKHSX_un = 0;
  let grandTH_un = 0;
  let grandKHSX_conv = 0;
  let grandTH_conv = 0;

  for (let i = 0; i < daysStartMs.length; i++) {
      let dayKHSX_un = 0; let dayTH_un = 0;
      let dayKHSX_conv = 0; let dayTH_conv = 0;
      
      activeLapRapRows.forEach(r => {
         const k = Number(r.rowData[3 + i*2]) || 0;
         const th = Number(r.rowData[3 + i*2 + 1]) || 0;
         dayKHSX_un += k;
         dayTH_un += th;
         dayKHSX_conv += k * r.hsqd;
         dayTH_conv += th * r.hsqd;
      });
      
      const dayDateStr = format(new Date(daysStartMs[i]), 'yyyy-MM-dd');
      let people = peoplePerDay[dayDateStr];
      if (people === undefined) {
        const avg = storageService.getDailyTotalPeopleAverage(dayDateStr);
        people = avg !== null ? avg : 40;
      }
      let dayMandays = mandaysPerDay[dayDateStr];
      if (dayMandays === undefined || dayMandays === 0) {
        const autoM = storageService.getDailyTotalMandays(dayDateStr);
        dayMandays = autoM > 0 ? autoM : (people > 0 ? people : 0);
      }
      // Năng lực sản xuất dây chuyền (8h) = 23.8 * số người của ngày đó
      const capacity = Math.round(people * 23.8);

      footerUnconverted.push(dayKHSX_un > 0 ? dayKHSX_un : '-', dayTH_un > 0 ? dayTH_un : '-');
      footerConverted.push(dayKHSX_conv > 0 ? Math.round(dayKHSX_conv) : '-', dayTH_conv > 0 ? Math.round(dayTH_conv) : '-');
      // Dòng tổng số người: tổng người của tất cả các công đoạn
      footerPeople.push(people, people);
      footerCapacity.push(capacity, capacity);
      footerRatio.push(dayKHSX_conv > 0 ? (capacity > 0 ? Math.round((dayKHSX_conv / capacity)*100) + '%' : '0%') : '0%', ''); 
      
      // Năng suất lao động của ngày = Tổng số sản phẩm quy đổi của ngày hôm đó / Tổng công của ngày hôm đó / 23,8
      const dailyProductivityVal = (dayMandays > 0 && dayTH_conv > 0)
        ? Math.round((dayTH_conv / dayMandays / 23.8) * 100)
        : 0;
      footerDailyProductivity.push('', dailyProductivityVal > 0 ? `${dailyProductivityVal}%` : '');
      
      grandKHSX_un += dayKHSX_un;
      grandTH_un += dayTH_un;
      grandKHSX_conv += dayKHSX_conv;
      grandTH_conv += dayTH_conv;
  }
  
  let total_people = 0;
  let total_mandays_range = 0;
  for (let i = 0; i < daysStartMs.length; i++) {
      const dayDateStr = format(new Date(daysStartMs[i]), 'yyyy-MM-dd');
      let people = peoplePerDay[dayDateStr];
      if (people === undefined) {
        const avg = storageService.getDailyTotalPeopleAverage(dayDateStr);
        people = avg !== null ? avg : 40;
      }
      total_people += people;

      let dMandays = mandaysPerDay[dayDateStr];
      if (dMandays === undefined || dMandays === 0) {
        const autoM = storageService.getDailyTotalMandays(dayDateStr);
        dMandays = autoM > 0 ? autoM : (people > 0 ? people : 0);
      }
      total_mandays_range += (dMandays > 0 ? dMandays : 0);
  }
  // Kế hoạch: 23.8
  // Thực hiện = tổng số lượng sản phẩm quy đổi của quãng ngày đã chọn chia cho tổng số công của quãng ngày đã chọn
  // % thực hiện = lấy thực hiện chia cho kế hoạch
  const targetPlan = 23.8;
  const mandaysTotal = totalMandays > 0 ? totalMandays : (total_mandays_range > 0 ? total_mandays_range : (total_people || 1));
  const thucHienVal = mandaysTotal > 0 ? (grandTH_conv / mandaysTotal) : 0;
  const phanTramVal = targetPlan > 0 ? (thucHienVal / targetPlan) * 100 : 0;

  const grandUnDiff = grandTH_un - grandKHSX_un;
  const grandUnDiffDisplay = grandUnDiff < 0 ? `(${Math.abs(grandUnDiff)})` : grandUnDiff;

  const grandConvDiff = Math.round(grandTH_conv - grandKHSX_conv);
  const grandConvDiffDisplay = grandConvDiff < 0 ? `(${Math.abs(grandConvDiff)})` : grandConvDiff;

  footerUnconverted.push(grandKHSX_un > 0 ? grandKHSX_un : '-', grandTH_un > 0 ? grandTH_un : '-', grandUnDiffDisplay);
  footerConverted.push(Math.round(grandKHSX_conv), Math.round(grandTH_conv), grandConvDiffDisplay);
  footerPeople.push('Năng suất Tuần', '', '');
  footerCapacity.push('Kế hoạch', 'Thực hiện', '% thực hiện');
  footerRatio.push(23.8, Number(thucHienVal.toFixed(1)), Math.round(phanTramVal) + '%'); 
  footerDailyProductivity.push('', '', '');


  lrSheetData.push(footerUnconverted, footerConverted, footerPeople, footerCapacity, footerRatio, footerDailyProductivity);

  const wsLR = XLSX.utils.aoa_to_sheet(lrSheetData);

  const numDays = daysStartMs.length;
  const diffCol = 3 + numDays * 2 + 2;
  const mergesLR = [
    { s: {r:0, c:0}, e: {r:0, c: 3 + numDays*2 + 2} }, // Title ROW
    { s: {r:1, c:0}, e: {r:2, c:0} }, 
    { s: {r:1, c:1}, e: {r:2, c:1} }, 
    { s: {r:1, c:2}, e: {r:2, c:2} }
  ];
  for (let i = 0; i < numDays; i++) {
    mergesLR.push({ s: {r:1, c: 3 + i*2}, e: {r:1, c: 3 + i*2 + 1} }); 
  }
  mergesLR.push(
    { s: {r:1, c: 3 + numDays*2}, e: {r:2, c: 3 + numDays*2} }, 
    { s: {r:1, c: 3 + numDays*2 + 1}, e: {r:2, c: 3 + numDays*2 + 1} }, 
    { s: {r:1, c: 3 + numDays*2 + 2}, e: {r:2, c: 3 + numDays*2 + 2} }  
  );
  
  const footerStartR = lrSheetData.length - 6;
  for (let i = 0; i < 6; i++) {
     mergesLR.push({ s: {r:footerStartR + i, c:0}, e: {r:footerStartR + i, c:2} });
  }

  const nangSuatTuanRow = footerStartR + 2;
  const nangSuatTuanCol = 3 + numDays * 2;
  mergesLR.push({ s: { r: nangSuatTuanRow, c: nangSuatTuanCol }, e: { r: nangSuatTuanRow, c: nangSuatTuanCol + 2 } });

  wsLR['!merges'] = mergesLR;

  const rangeLR = XLSX.utils.decode_range(wsLR['!ref'] || 'A1:A1');
  for (let R = 0; R <= rangeLR.e.r; ++R) {
    for (let C = 0; C <= rangeLR.e.c; ++C) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      if (!wsLR[cellAddress]) wsLR[cellAddress] = { t: 's', v: '' };
      
      const isTitle = R === 0;
      const isHeader = R === 1 || R === 2;
      const isFooter = R >= footerStartR;
      const isDiffCol = C === diffCol;
      
      const cellVal = wsLR[cellAddress].v;
      let isNegative = false;
      if (typeof cellVal === 'number' && cellVal < 0) {
        isNegative = true;
        if (isDiffCol) {
          wsLR[cellAddress].v = `(${Math.abs(cellVal)})`;
          wsLR[cellAddress].t = 's';
        }
      } else if (typeof cellVal === 'string' && cellVal.startsWith('(') && cellVal.endsWith(')')) {
        isNegative = true;
      }
      
      if (isTitle) {
         wsLR[cellAddress].s = {
            font: { bold: true, sz: 16 },
            alignment: { horizontal: "center", vertical: "center" }
         };
      } else if (isHeader) {
         wsLR[cellAddress].s = {
            ...headerStyle,
            font: { bold: true, sz: 10, color: { rgb: "FFFFFF" } },
         };
      } else if (isFooter) {
         wsLR[cellAddress].s = {
            ...dataStyle,
            font: { 
              bold: true, 
              sz: 11,
              color: (isDiffCol && isNegative) ? { rgb: "DC2626" } : undefined
            },
            fill: { fgColor: { rgb: R === footerStartR ? "E2E8F0" : R === footerStartR + 1 ? "FED7AA" : R === footerStartR + 4 ? "F3F4F6" : R === footerStartR + 5 ? "FFFF00" : "FEF3C7" } }
         };
         if (C > 2) wsLR[cellAddress].s.alignment = { horizontal: "center", vertical: "center" };
      } else {
         wsLR[cellAddress].s = { ...dataStyle, alignment: (C > 1 && C !== 2) ? { horizontal: "center", vertical: "center" } : { horizontal: "left", vertical: "center" } };
         if (C >= 3 && C < 3 + numDays*2) {
             const isKHSX = (C - 3) % 2 === 0;
             wsLR[cellAddress].s.font = { ...wsLR[cellAddress].s.font, color: { rgb: isKHSX ? "DC2626" : "2563EB" }, bold: true };
         } else if (isDiffCol && isNegative) {
             wsLR[cellAddress].s.font = { ...wsLR[cellAddress].s.font, color: { rgb: "DC2626" }, bold: true };
         }
      }
    }
  }

  const colsLR = [
    { wch: 30 }, // Mã TP
    { wch: 45 }, // Tên TP
    { wch: 8 },  // HSQĐ
  ];
  for (let i = 0; i < numDays; i++) {
    colsLR.push({ wch: 6 }, { wch: 6 }); // KHSX, TH
  }
  colsLR.push({ wch: 8 }, { wch: 8 }, { wch: 10 }); // Total KHSX, TH, Chênh lệch
  wsLR['!cols'] = colsLR;

  if (activeLapRapRows.length > 0) {
    XLSX.utils.book_append_sheet(wb, wsLR, "Lắp Ráp");
  }



  XLSX.writeFile(wb, `BaoCao_SanLuong_${format(start, 'ddMMyy')}-${format(end, 'ddMMyy')}.xlsx`);
};

const STAGE_CHART_CONFIGS: Record<string, {
  name: string;
  stageName: string;
  normText: string;
  normPerShift: number;
  hourlyNorm: number;
  color: string;
  lineColor: string;
  barKey: string;
  lineKey: string;
  convertedKey: string;
  targetKey: string;
  workerKey: string;
  hsqdSource: string;
}> = {
  LASER: {
    name: 'Cắt Laser',
    stageName: 'Cắt Laser',
    normText: '250 SPQĐ / người / ca 8h (31.25 SPQĐ / giờ)',
    normPerShift: 250,
    hourlyNorm: 31.25,
    color: '#3B82F6',
    lineColor: '#2563EB',
    barKey: 'Cắt Laser',
    lineKey: 'NSLD_Laser',
    convertedKey: 'laserConverted',
    targetKey: 'targetQDLaser',
    workerKey: 'workerCountLaser',
    hsqdSource: 'Tab HSQĐ Laser',
  },
  BENDING: {
    name: 'Chấn / Dập',
    stageName: 'Chấn/Dập',
    normText: '132.7 SPQĐ / người / ca 8h (16.59 SPQĐ / giờ)',
    normPerShift: 132.7,
    hourlyNorm: 16.59,
    color: '#9333EA',
    lineColor: '#7E22CE',
    barKey: 'Chấn/Dập',
    lineKey: 'NSLD_ChanDap',
    convertedKey: 'chanDapConverted',
    targetKey: 'targetQDChanDap',
    workerKey: 'workerCountChanDap',
    hsqdSource: 'Tab HSQĐ Chấn / Hàn',
  },
  WELDING: {
    name: 'Hàn',
    stageName: 'Hàn',
    normText: '77.2 SPQĐ / người / ca 8h (9.65 SPQĐ / giờ)',
    normPerShift: 77.2,
    hourlyNorm: 9.65,
    color: '#10B981',
    lineColor: '#059669',
    barKey: 'Hàn',
    lineKey: 'NSLD_Han',
    convertedKey: 'hanConverted',
    targetKey: 'targetQDHan',
    workerKey: 'workerCountHan',
    hsqdSource: 'Tab HSQĐ Chấn / Hàn',
  },
  PAINTING: {
    name: 'Sơn',
    stageName: 'Sơn',
    normText: '127.5 SPQĐ / người / ca 8h (15.94 SPQĐ / giờ)',
    normPerShift: 127.5,
    hourlyNorm: 15.94,
    color: '#EAB308',
    lineColor: '#CA8A04',
    barKey: 'Sơn',
    lineKey: 'NSLD_Son',
    convertedKey: 'sonConverted',
    targetKey: 'targetQDSon',
    workerKey: 'workerCountSon',
    hsqdSource: 'Tab HSQĐ Sơn',
  },
  GLAZING: {
    name: 'Dán Kính',
    stageName: 'Dán Kính',
    normText: '242 SPQĐ / người / ca 8h (30.25 SPQĐ / giờ)',
    normPerShift: 242,
    hourlyNorm: 30.25,
    color: '#EF4444',
    lineColor: '#DC2626',
    barKey: 'Dán Kính',
    lineKey: 'NSLD_DanKinh',
    convertedKey: 'danKinhConverted',
    targetKey: 'targetQDDanKinh',
    workerKey: 'workerCountDanKinh',
    hsqdSource: 'Tab HSQĐ Dán Kính 2 (Công đoạn riêng lẻ)',
  },
};

function DashboardView({ inventory, parts, transactions, labels, refreshData, setDefectModal, dclrNorms, dclrNormsMap, glazingHSQD }: any) {
  const [selectedStageDetail, setSelectedStageDetail] = useState<StageId | null>(null);
  const [expandedInventoryKeys, setExpandedInventoryKeys] = useState<Set<string>>(new Set());
  const [searchTerm, setSearchTerm] = useState('');
  const [showExportModal, setShowExportModal] = useState(false);
  const [showExportPreparationModal, setShowExportPreparationModal] = useState(false);
  const [showExportHourlyModal, setShowExportHourlyModal] = useState(false);
  const [activeChartTab, setActiveChartTab] = useState<'MECHANICAL' | 'LASER' | 'BENDING' | 'WELDING' | 'PAINTING' | 'GLAZING'>('MECHANICAL');
  const [activeChartShift, setActiveChartShift] = useState<'DAY' | 'NIGHT'>('DAY');
  const [chartDimensions, setChartDimensions] = useState<{ width: number; height: number } | null>(null);
  const chartWrapperRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    let cancelled = false;
    let resizeObserver: ResizeObserver | null = null;

    const measure = () => {
      if (cancelled || !chartWrapperRef.current) return;
      const el = chartWrapperRef.current;
      const rect = el.getBoundingClientRect();
      const w = Math.floor(rect.width || el.clientWidth || 0);
      const h = Math.floor(rect.height || el.clientHeight || 0);

      if (w > 0 && h > 0) {
        setChartDimensions({ width: w, height: h });
      }
    };

    // Delay initial measure slightly to allow CSS layout / animations to settle
    const initialTimer = setTimeout(() => {
      measure();
    }, 50);

    if (typeof ResizeObserver !== 'undefined' && chartWrapperRef.current) {
      resizeObserver = new ResizeObserver((entries) => {
        for (const entry of entries) {
          const w = Math.floor(entry.contentRect.width);
          const h = Math.floor(entry.contentRect.height);
          if (w > 0 && h > 0 && !cancelled) {
            setChartDimensions({ width: w, height: h });
          }
        }
      });
      resizeObserver.observe(chartWrapperRef.current);
    } else {
      const interval = setInterval(measure, 150);
      return () => {
        cancelled = true;
        clearTimeout(initialTimer);
        clearInterval(interval);
      };
    }

    return () => {
      cancelled = true;
      clearTimeout(initialTimer);
      if (resizeObserver) resizeObserver.disconnect();
    };
  }, [activeChartTab]);

  const getPastDateStr = (days: number) => {
    const d = new Date();
    d.setDate(d.getDate() - days);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  };

  const getPastDateTimeStr = (days: number) => {
    const d = new Date();
    d.setDate(d.getDate() - days);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    const hours = String(d.getHours()).padStart(2, '0');
    const minutes = String(d.getMinutes()).padStart(2, '0');
    return `${y}-${m}-${day}T${hours}:${minutes}`;
  };

  const [chartDate, setChartDate] = useState(() => getPastDateStr(0));
  const [hourlyPeopleLaser, setHourlyPeopleLaser] = useState<Record<string, number>>(() => storageService.getHourlyPeopleLaser(getPastDateStr(0)));
  const [hourlyPeoplePainting, setHourlyPeoplePainting] = useState<Record<string, number>>(() => storageService.getHourlyPeoplePainting(getPastDateStr(0)));
  const [hourlyPeopleGlazing, setHourlyPeopleGlazing] = useState<Record<string, number>>(() => storageService.getHourlyPeopleGlazing(getPastDateStr(0)));
  const [hourlyPeopleBending, setHourlyPeopleBending] = useState<Record<string, number>>(() => storageService.getHourlyPeopleBending(getPastDateStr(0)));
  const [hourlyPeopleWelding, setHourlyPeopleWelding] = useState<Record<string, number>>(() => storageService.getHourlyPeopleWelding(getPastDateStr(0)));
  const [stagePlannedWorkers, setStagePlannedWorkers] = useState<StagePlannedWorkers>(() => storageService.getStagePlannedWorkers(getPastDateStr(0)));

  useEffect(() => {
    setHourlyPeopleLaser(storageService.getHourlyPeopleLaser(chartDate));
    setHourlyPeoplePainting(storageService.getHourlyPeoplePainting(chartDate));
    setHourlyPeopleGlazing(storageService.getHourlyPeopleGlazing(chartDate));
    setHourlyPeopleBending(storageService.getHourlyPeopleBending(chartDate));
    setHourlyPeopleWelding(storageService.getHourlyPeopleWelding(chartDate));
    setStagePlannedWorkers(storageService.getStagePlannedWorkers(chartDate));
  }, [chartDate]);

  const handleStagePlannedWorkersChange = (stageKey: keyof StagePlannedWorkers, val: number | undefined) => {
    const updated = { ...stagePlannedWorkers };
    if (val === undefined || isNaN(val)) {
      delete updated[stageKey];
    } else {
      updated[stageKey] = val;
    }
    setStagePlannedWorkers(updated);
    storageService.saveStagePlannedWorkers(updated, chartDate);
  };

  const [showManualAddModal, setShowManualAddModal] = useState<{stageId: StageId, location: 'IN' | 'OUT'} | null>(null);
  const [manualAddPart, setManualAddPart] = useState('');
  const [manualAddPartSearch, setManualAddPartSearch] = useState('');
  const [manualAddQty, setManualAddQty] = useState('');

  const [exportStartDate, setExportStartDate] = useState(() => getPastDateStr(35));
  const [exportEndDate, setExportEndDate] = useState(() => getPastDateStr(0));
  const [peoplePerDay, setPeoplePerDay] = useState<Record<string, number>>(() => storageService.getPeoplePerDay());
  const [mandaysPerDay, setMandaysPerDay] = useState<Record<string, number>>(() => storageService.getMandaysPerDay());
  const [exportTotalMandays, setExportTotalMandays] = useState("");
  const [exportWeekName, setExportWeekName] = useState(() => storageService.getExportWeekName());
  const [selectedMasterPoIdsForExport, setSelectedMasterPoIdsForExport] = useState<string[]>([]);
  const [exportModelSearchTerm, setExportModelSearchTerm] = useState('');
  const [preparationExpectedDate, setPreparationExpectedDate] = useState(() => getPastDateTimeStr(-2)); // Tương lai 2 ngày (có thể điều chỉnh)

  const exportDateRange = useMemo(() => {
    if (!exportStartDate || !exportEndDate) return [];
    const dates = [];
    const [sY, sM, sD] = exportStartDate.split('-').map(Number);
    const curr = new Date(sY, sM - 1, sD, 0, 0, 0, 0);
    const [eY, eM, eD] = exportEndDate.split('-').map(Number);
    const end = new Date(eY, eM - 1, eD, 23, 59, 59, 999);
    
    let count = 0;
    while (curr <= end && count < 100) {
      const y = curr.getFullYear();
      const m = String(curr.getMonth() + 1).padStart(2, '0');
      const d = String(curr.getDate()).padStart(2, '0');
      dates.push(`${y}-${m}-${d}`);
      curr.setDate(curr.getDate() + 1);
      count++;
    }
    return dates;
  }, [exportStartDate, exportEndDate]);

  const calculatedTotalMandays = useMemo(() => {
    const total = exportDateRange.reduce((acc, d) => {
      const autoM = storageService.getDailyTotalMandays(d);
      const val = mandaysPerDay[d] !== undefined ? mandaysPerDay[d] : autoM;
      return acc + (val > 0 ? val : 0);
    }, 0);
    return Math.round(total * 10) / 10;
  }, [exportDateRange, mandaysPerDay]);

  useEffect(() => {
    setExportTotalMandays(String(calculatedTotalMandays));
  }, [calculatedTotalMandays]);

  const calculateAllStagesAverageForDate = (
    lRecord: Record<string, number>,
    bRecord: Record<string, number>,
    wRecord: Record<string, number>,
    pRecord: Record<string, number>,
    gRecord: Record<string, number>
  ) => {
    const timeSlots = ['6h-8h', '8h-10h', '10h-12h', '13h-15h', '15h-17h', '17h-19h', '19h-20h', '20h-22h', '22h-0h', '0h-2h', '2h-4h', '4h-6h'];
    let totalPeople = 0;
    let validSlots = 0;
    timeSlots.forEach(slot => {
      const l = lRecord[slot] !== undefined ? lRecord[slot] : 0;
      const b = bRecord[slot] !== undefined ? bRecord[slot] : 0;
      const w = wRecord[slot] !== undefined ? wRecord[slot] : 0;
      const p = pRecord[slot] !== undefined ? pRecord[slot] : 0;
      const g = gRecord[slot] !== undefined ? gRecord[slot] : 0;
      const slotTotal = l + b + w + p + g;
      if (slotTotal > 0) {
        totalPeople += slotTotal;
        validSlots++;
      }
    });
    return validSlots > 0 ? Math.round(totalPeople / validSlots) : null;
  };

  const syncDailyPeopleAverage = (
    l: Record<string, number>,
    b: Record<string, number>,
    w: Record<string, number>,
    p: Record<string, number>,
    g: Record<string, number>
  ) => {
    const autoPeople = calculateAllStagesAverageForDate(l, b, w, p, g);
    setPeoplePerDay(prev => {
      const updated = { ...prev };
      if (autoPeople !== null) {
        updated[chartDate] = autoPeople;
      } else {
        delete updated[chartDate];
      }
      storageService.savePeoplePerDay(updated);
      return updated;
    });

    // Tự động tính số công của ngày theo các khung giờ có nhân sự
    // Các khung giờ nào không có nhân sự nào hết ở tất cả công đoạn sẽ không tính công khung giờ đó
    const autoMandays = storageService.getDailyTotalMandays(chartDate);
    setMandaysPerDay(prev => {
      const updated = { ...prev };
      if (autoMandays > 0) {
        updated[chartDate] = Math.round(autoMandays * 10) / 10;
      } else {
        delete updated[chartDate];
      }
      storageService.saveMandaysPerDay(updated);
      return updated;
    });
  };

  const handleHourlyLaserChange = (slotLabel: string, val: number | undefined) => {
    const newLaser = { ...hourlyPeopleLaser };
    if (val === undefined || isNaN(val)) {
      delete newLaser[slotLabel];
    } else {
      newLaser[slotLabel] = val;
    }
    setHourlyPeopleLaser(newLaser);
    storageService.saveHourlyPeopleLaser(newLaser, chartDate);
    syncDailyPeopleAverage(newLaser, hourlyPeopleBending, hourlyPeopleWelding, hourlyPeoplePainting, hourlyPeopleGlazing);
  };

  const handleHourlyBendingChange = (slotLabel: string, val: number | undefined) => {
    const newBending = { ...hourlyPeopleBending };
    if (val === undefined || isNaN(val)) {
      delete newBending[slotLabel];
    } else {
      newBending[slotLabel] = val;
    }
    setHourlyPeopleBending(newBending);
    storageService.saveHourlyPeopleBending(newBending, chartDate);
    syncDailyPeopleAverage(hourlyPeopleLaser, newBending, hourlyPeopleWelding, hourlyPeoplePainting, hourlyPeopleGlazing);
  };

  const handleHourlyWeldingChange = (slotLabel: string, val: number | undefined) => {
    const newWelding = { ...hourlyPeopleWelding };
    if (val === undefined || isNaN(val)) {
      delete newWelding[slotLabel];
    } else {
      newWelding[slotLabel] = val;
    }
    setHourlyPeopleWelding(newWelding);
    storageService.saveHourlyPeopleWelding(newWelding, chartDate);
    syncDailyPeopleAverage(hourlyPeopleLaser, hourlyPeopleBending, newWelding, hourlyPeoplePainting, hourlyPeopleGlazing);
  };

  const handleHourlyPaintingChange = (slotLabel: string, val: number | undefined) => {
    const newPainting = { ...hourlyPeoplePainting };
    if (val === undefined || isNaN(val)) {
      delete newPainting[slotLabel];
    } else {
      newPainting[slotLabel] = val;
    }
    setHourlyPeoplePainting(newPainting);
    storageService.saveHourlyPeoplePainting(newPainting, chartDate);
    syncDailyPeopleAverage(hourlyPeopleLaser, hourlyPeopleBending, hourlyPeopleWelding, newPainting, hourlyPeopleGlazing);
  };

  const handleHourlyGlazingChange = (slotLabel: string, val: number | undefined) => {
    const newGlazing = { ...hourlyPeopleGlazing };
    if (val === undefined || isNaN(val)) {
      delete newGlazing[slotLabel];
    } else {
      newGlazing[slotLabel] = val;
    }
    setHourlyPeopleGlazing(newGlazing);
    storageService.saveHourlyPeopleGlazing(newGlazing, chartDate);
    syncDailyPeopleAverage(hourlyPeopleLaser, hourlyPeopleBending, hourlyPeopleWelding, hourlyPeoplePainting, newGlazing);
  };

  const productionChartData = useMemo(() => {
    const dayTimeSlots = [
      { label: '6h-8h', start: 6, end: 8, factor: 0.25 },
      { label: '8h-10h', start: 8, end: 10, factor: 0.25 },
      { label: '10h-12h', start: 10, end: 12, factor: 0.25 },
      { label: '13h-15h', start: 13, end: 15, factor: 0.25 },
      { label: '15h-17h', start: 15, end: 17, factor: 0.25 },
      { label: '17h-19h', start: 17, end: 19, factor: 0.25 },
    ];

    const nightTimeSlots = [
      { label: '19h-21h', start: 19, end: 21, factor: 0.25 },
      { label: '21h-23h', start: 21, end: 23, factor: 0.25 },
      { label: '23h-0h', start: 23, end: 24, factor: 0.125 },
      { label: '1h-3h', start: 1, end: 3, factor: 0.25 },
      { label: '3h-5h', start: 3, end: 5, factor: 0.25 },
      { label: '5h-6h', start: 5, end: 6, factor: 0.125 },
    ];

    const timeSlots = activeChartShift === 'DAY' ? dayTimeSlots : nightTimeSlots;

    const targetStages = ['LASER', 'BENDING', 'WELDING', 'PAINTING', 'GLAZING'];
    
    
    // Filter transactions by date and type STAGE_OUT (location IN -> processing done)
    const txFilterFn = (t: any) => {
      // "Năng lực sản xuất" logic: When something leaves 'IN' at a stage (moves to OUT),
      // OR when something is manually inbound directly to 'OUT' (like Laser Cutting)
      const isProducedAtStage = 
        (t.type === 'STAGE_OUT' && t.location === 'IN') ||
        (t.type === 'STAGE_IN' && t.location === 'OUT');

      if (!isProducedAtStage) return false;
      if (!targetStages.includes(t.stageId)) return false;
      
      const txDateObj = new Date(t.timestamp);
      const offset = txDateObj.getTimezoneOffset();
      const localDate = new Date(txDateObj.getTime() - (offset*60*1000));
      const txDate = localDate.toISOString().split('T')[0];
      
      return txDate === chartDate;
    };

    const productionTransactions = transactions.filter(txFilterFn);
    
    // Specifically for GLAZING, we might need to also count STAGE_OUT labels
    // similar to how export reports work.
    const glazingLabels = labels.filter(l => {
       if (l.type !== 'STAGE_OUT' || l.stageId !== 'GLAZING' || (l as any).kpiRecorded) return false;
       const txDateObj = new Date(l.timestamp);
       const offset = txDateObj.getTimezoneOffset();
       const localDate = new Date(txDateObj.getTime() - (offset*60*1000));
       return localDate.toISOString().split('T')[0] === chartDate;
    });

    const laserHSQDArray = storageService.getLaserHSQD();
    const laserHSQDMap = new Map(laserHSQDArray.map(n => [n.partId.toUpperCase(), n.hsqd]));
    const getLaserHSQD = (partId: string) => {
      const upper = (partId || '').trim().toUpperCase();
      if (laserHSQDMap.has(upper)) return laserHSQDMap.get(upper)!;
      const cleanId = upper.replace(/^LSR-OUT-/i, '').replace(/\s*-\s*CẮT LASER$/i, '').trim();
      if (cleanId && laserHSQDMap.has(cleanId)) return laserHSQDMap.get(cleanId)!;
      const matchedPart = parts.find((p: any) => p.id.toUpperCase() === upper || (cleanId && (p.name.toUpperCase().trim() === cleanId || p.id.toUpperCase() === cleanId)));
      if (matchedPart && laserHSQDMap.has(matchedPart.id.toUpperCase())) return laserHSQDMap.get(matchedPart.id.toUpperCase())!;
      return 0;
    };

    const bendingWeldingHSQDArray = storageService.getBendingWeldingHSQD();
    const bendingWeldingHSQDMap = new Map(bendingWeldingHSQDArray.map(n => [n.partId.toUpperCase(), n.hsqd]));
    const getBendingWeldingHSQD = (partId: string) => {
      const upper = (partId || '').trim().toUpperCase();
      if (bendingWeldingHSQDMap.has(upper)) return bendingWeldingHSQDMap.get(upper)!;
      const cleanId = upper.replace(/^(BND|WLD)-OUT-/i, '').replace(/\s*-\s*(CHẤN\/DẬP|HÀN|CHẤN|DẬP)$/i, '').trim();
      if (cleanId && bendingWeldingHSQDMap.has(cleanId)) return bendingWeldingHSQDMap.get(cleanId)!;
      const matchedPart = parts.find((p: any) => p.id.toUpperCase() === upper || (cleanId && (p.name.toUpperCase().trim() === cleanId || p.id.toUpperCase() === cleanId)));
      if (matchedPart && bendingWeldingHSQDMap.has(matchedPart.id.toUpperCase())) return bendingWeldingHSQDMap.get(matchedPart.id.toUpperCase())!;
      return 0;
    };

    const glazingStageHSQDArray = storageService.getGlazingStageHSQD();
    const glazingStageHSQDMap = new Map(glazingStageHSQDArray.map(n => [n.partId.toUpperCase(), n.hsqd]));
    const getGlazingStageHSQD = (partId: string) => {
      const upper = (partId || '').trim().toUpperCase();
      if (glazingStageHSQDMap.has(upper)) return glazingStageHSQDMap.get(upper)!;
      const cleanId = upper.replace(/^GLZ-OUT-/i, '').replace(/\s*-\s*DÁN KÍNH$/i, '').trim();
      if (cleanId && glazingStageHSQDMap.has(cleanId)) return glazingStageHSQDMap.get(cleanId)!;
      const matchedPart = parts.find((p: any) => p.id.toUpperCase() === upper || (cleanId && (p.name.toUpperCase().trim() === cleanId || p.id.toUpperCase() === cleanId)));
      if (matchedPart && glazingStageHSQDMap.has(matchedPart.id.toUpperCase())) return glazingStageHSQDMap.get(matchedPart.id.toUpperCase())!;
      // Fallback: nếu chưa cài đặt riêng trong tab HSQĐ Dán kính 2, dùng HSQĐ của tab Dán kính 1 (xưởng)
      return glazingHSQD(partId);
    };

    const paintingHSQDArray = storageService.getPaintingHSQD();
    const paintingHSQDMap = new Map(paintingHSQDArray.map(n => [n.partId.toUpperCase(), n.hsqd]));
    const getPaintingHSQD = (partId: string) => {
      const upper = (partId || '').trim().toUpperCase();
      if (paintingHSQDMap.has(upper)) return paintingHSQDMap.get(upper)!;
      const cleanId = upper.replace(/^PNT-OUT-/i, '').replace(/\s*-\s*SƠN$/i, '').trim();
      if (cleanId && paintingHSQDMap.has(cleanId)) return paintingHSQDMap.get(cleanId)!;
      const matchedPart = parts.find((p: any) => p.id.toUpperCase() === upper || (cleanId && (p.name.toUpperCase().trim() === cleanId || p.id.toUpperCase() === cleanId)));
      if (matchedPart && paintingHSQDMap.has(matchedPart.id.toUpperCase())) return paintingHSQDMap.get(matchedPart.id.toUpperCase())!;
      return dclrNormsMap.get(partId) || 0;
    };

    return timeSlots.map(slot => {
      const slotData: any = { name: slot.label, details: {} };
      targetStages.forEach(stageId => {
        const stageName = STAGES.find(s => s.id === stageId)?.name || stageId;
        slotData[stageName] = 0;
        slotData.details[stageName] = [];
      });

      const slotTxs = productionTransactions.filter(t => {
        const hour = new Date(t.timestamp).getHours();
        if (slot.label === '10h-12h') return hour >= 10 && hour <= 12; // Nghỉ trưa gộp vào 10h-12h
        if (slot.label === '23h-0h') return hour === 23 || hour === 0; // Nghỉ đêm gộp vào 23h-0h
        return hour >= slot.start && hour < slot.end;
      });
      
      const slotGlazingLabels = glazingLabels.filter(t => {
        const hour = new Date(t.timestamp).getHours();
        if (slot.label === '10h-12h') return hour >= 10 && hour <= 12;
        if (slot.label === '23h-0h') return hour === 23 || hour === 0;
        return hour >= slot.start && hour < slot.end;
      });

      let laserConverted = 0;
      let sonConverted = 0; // Tính theo HSQĐ Sơn riêng (Tab HSQĐ Sơn) cho tab riêng Sơn
      let sonConvertedDclr = 0; // Tính theo HSQĐ DCLR cho Biểu đồ Năng suất Xưởng Cơ Khí
      let danKinhConvertedMechanical = 0; // Tính theo HSQĐ Dán Kính 1 (Xưởng Cơ Khí)
      let danKinhConvertedStage = 0;      // Tính theo HSQĐ Dán Kính 2 (Công đoạn riêng lẻ)
      let chanDapConverted = 0;
      let hanConverted = 0;

      slotTxs.forEach(t => {
        const stageName = STAGES.find(s => s.id === t.stageId)?.name || t.stageId;
        slotData[stageName] += (t.quantity || 0);
        
        const part = parts.find((p: any) => p.id === t.partId);
        const partName = part ? part.name : (t.partName || t.partId);
        const unit = part?.unit || 'cái';
        
        let itemHsqd = 0;
        let itemHsqdMech = 0;
        if (t.stageId === 'LASER') {
           if (!isLaserScrapItem(t.partId, t.partName)) {
             itemHsqd = getLaserHSQD(t.partId);
             itemHsqdMech = itemHsqd;
             laserConverted += ((t.quantity || 0) * itemHsqd);
           }
        } else if (t.stageId === 'PAINTING') {
           itemHsqd = getPaintingHSQD(t.partId); // Tab HSQĐ Sơn (cho tab riêng Sơn)
           itemHsqdMech = dclrNormsMap.get(t.partId) || 0; // Tab HSQĐ DCLR (cho Năng suất Xưởng Cơ Khí)
           sonConverted += ((t.quantity || 0) * itemHsqd);
           sonConvertedDclr += ((t.quantity || 0) * itemHsqdMech);
        } else if (t.stageId === 'GLAZING') {
           itemHsqdMech = glazingHSQD(t.partId); // Tab HSQĐ Dán Kính 1 (Xưởng CK)
           itemHsqd = getGlazingStageHSQD(t.partId); // Tab HSQĐ Dán Kính 2 (Công đoạn)
           danKinhConvertedMechanical += ((t.quantity || 0) * itemHsqdMech);
           danKinhConvertedStage += ((t.quantity || 0) * itemHsqd);
        } else if (t.stageId === 'BENDING') {
           itemHsqd = getBendingWeldingHSQD(t.partId);
           itemHsqdMech = itemHsqd;
           chanDapConverted += ((t.quantity || 0) * itemHsqd);
        } else if (t.stageId === 'WELDING') {
           itemHsqd = getBendingWeldingHSQD(t.partId);
           itemHsqdMech = itemHsqd;
           hanConverted += ((t.quantity || 0) * itemHsqd);
        }

        const existingDetail = slotData.details[stageName].find((d: any) => d.partName === partName);
        if (existingDetail) {
          existingDetail.quantity += (t.quantity || 0);
          if ((!existingDetail.hsqd || existingDetail.hsqd === 0) && itemHsqd > 0) {
            existingDetail.hsqd = itemHsqd;
          }
          if ((!existingDetail.hsqdMech || existingDetail.hsqdMech === 0) && itemHsqdMech > 0) {
            existingDetail.hsqdMech = itemHsqdMech;
          }
        } else {
          slotData.details[stageName].push({ 
            partName, 
            quantity: (t.quantity || 0), 
            unit,
            hsqd: itemHsqd,
            hsqdMech: itemHsqdMech,
            partId: t.partId 
          });
        }
      });
      
      // Handle GLAZING using labels to match export exactly
      slotGlazingLabels.forEach(t => {
        const stageName = 'Dán Kính';
        slotData[stageName] += (t.quantity || 0);
        
        const part = parts.find((p: any) => p.id === t.partId);
        const partName = part ? part.name : (t.partName || t.partId);
        const unit = part?.unit || 'cái';
        
        const hsqdMech = glazingHSQD(t.partId); // Tab HSQĐ Dán Kính 1 (Xưởng CK)
        const itemHsqd = getGlazingStageHSQD(t.partId); // Tab HSQĐ Dán Kính 2 (Công đoạn)
        danKinhConvertedMechanical += ((t.quantity || 0) * hsqdMech);
        danKinhConvertedStage += ((t.quantity || 0) * itemHsqd);

        const existingDetail = slotData.details[stageName].find((d: any) => d.partName === partName);
        if (existingDetail) {
          existingDetail.quantity += (t.quantity || 0);
          if ((!existingDetail.hsqd || existingDetail.hsqd === 0) && itemHsqd > 0) {
            existingDetail.hsqd = itemHsqd;
          }
          if ((!existingDetail.hsqdMech || existingDetail.hsqdMech === 0) && hsqdMech > 0) {
            existingDetail.hsqdMech = hsqdMech;
          }
        } else {
          slotData.details[stageName].push({ 
            partName, 
            quantity: (t.quantity || 0), 
            unit,
            hsqd: itemHsqd,
            hsqdMech: hsqdMech,
            partId: t.partId
          });
        }
      });
      
      // Calculate productivity percentage based on user's exact shift standards:
      // Cắt Laser: 250 SPQĐ / người / ca 8h
      // Chấn / Dập: 132.7 SPQĐ / người / ca 8h
      // Hàn: 77.2 SPQĐ / người / ca 8h
      // Sơn: 127.5 SPQĐ / người / ca 8h
      // Dán Kính: 242 SPQĐ / người / ca 8h
      const hourFactor = slot.factor !== undefined ? slot.factor : ((slot.label === '23h-0h' || slot.label === '5h-6h') ? 0.125 : 0.25);

      const workerCountLaser = hourlyPeopleLaser[slot.label] !== undefined ? hourlyPeopleLaser[slot.label] : 0;
      const targetQDLaser = workerCountLaser * 250 * hourFactor;

      const workerCountSon = hourlyPeoplePainting[slot.label] !== undefined ? hourlyPeoplePainting[slot.label] : 0;
      const targetQDSon = workerCountSon * 127.5 * hourFactor;
      
      const workerCountDanKinh = hourlyPeopleGlazing[slot.label] !== undefined ? hourlyPeopleGlazing[slot.label] : 0;
      const targetQDDanKinh = workerCountDanKinh * 242 * hourFactor;
      
      const workerCountChanDap = hourlyPeopleBending[slot.label] !== undefined ? hourlyPeopleBending[slot.label] : 0;
      const targetQDChanDap = workerCountChanDap * 132.7 * hourFactor;
      
      const workerCountHan = hourlyPeopleWelding[slot.label] !== undefined ? hourlyPeopleWelding[slot.label] : 0;
      const targetQDHan = workerCountHan * 77.2 * hourFactor;

      slotData['NSLD_Laser'] = targetQDLaser > 0 ? (laserConverted / targetQDLaser) : 0;
      slotData['NSLD_Son'] = targetQDSon > 0 ? (sonConverted / targetQDSon) : 0;
      slotData['NSLD_DanKinh'] = targetQDDanKinh > 0 ? (danKinhConvertedStage / targetQDDanKinh) : 0;
      slotData['NSLD_ChanDap'] = targetQDChanDap > 0 ? (chanDapConverted / targetQDChanDap) : 0;
      slotData['NSLD_Han'] = targetQDHan > 0 ? (hanConverted / targetQDHan) : 0;

      // Năng suất lao động phân xưởng cơ khí từng khung giờ:
      // NSLĐ = Tổng sản phẩm quy đổi của khung giờ (Sơn DCLR + Dán kính tính theo HSQĐ Xưởng) / (Tổng số người của khung giờ đó * hourFactor) / 28.3
      const totalConvertedMechanical = sonConvertedDclr + danKinhConvertedMechanical;
      const totalWorkersMechanical = workerCountLaser + workerCountChanDap + workerCountHan + workerCountSon + workerCountDanKinh;
      const slotMandays = totalWorkersMechanical * hourFactor;
      
      const nsldMechanical = (slotMandays > 0 && totalConvertedMechanical > 0)
        ? (totalConvertedMechanical / slotMandays / 28.3)
        : 0;
      
      slotData['NSLD_CoKhi'] = nsldMechanical;
      slotData['totalConvertedMechanical'] = totalConvertedMechanical;
      slotData['workerCountLaser'] = workerCountLaser;
      slotData['workerCountChanDap'] = workerCountChanDap;
      slotData['workerCountHan'] = workerCountHan;
      slotData['workerCountSon'] = workerCountSon;
      slotData['workerCountDanKinh'] = workerCountDanKinh;
      slotData['totalWorkersMechanical'] = totalWorkersMechanical;
      slotData['hourFactor'] = hourFactor;
      slotData['slotMandays'] = slotMandays;

      slotData['laserConverted'] = laserConverted;
      slotData['targetQDLaser'] = targetQDLaser;
      slotData['sonConverted'] = sonConverted;
      slotData['targetQDSon'] = targetQDSon;
      slotData['danKinhConverted'] = danKinhConvertedStage;
      slotData['targetQDDanKinh'] = targetQDDanKinh;
      slotData['danKinhConvertedMechanical'] = danKinhConvertedMechanical;
      slotData['chanDapConverted'] = chanDapConverted;
      slotData['targetQDChanDap'] = targetQDChanDap;
      slotData['hanConverted'] = hanConverted;
      slotData['targetQDHan'] = targetQDHan;

      return slotData;
    });
  }, [transactions, labels, chartDate, activeChartShift, hourlyPeopleLaser, hourlyPeoplePainting, hourlyPeopleGlazing, hourlyPeopleBending, hourlyPeopleWelding]);

  const laserScrapItems = useMemo(() => {
    return inventory.filter(item => 
      item.stageId === 'LASER' && 
      item.quantity > 0 && 
      isLaserScrapItem(item.partId, parts.find(p => p.id === item.partId || p.name === item.partId)?.name)
    );
  }, [inventory, parts]);

  const [laserYieldPoFilter, setLaserYieldPoFilter] = useState<string>('ALL');
  const [showLaserYieldDetails, setShowLaserYieldDetails] = useState<boolean>(false);

  // Tính toán Tỷ lệ cắt tôn & Tỷ lệ phế liệu (Yield & Scrap Calculation) - Lịch sử 35 ngày qua
  const laserYieldReport = useMemo(() => {
    const bom = storageService.getBOM();
    const productionOrders = storageService.getProductionOrders().filter(po => po.stageId === 'LASER');
    
    // Khoảng thời gian thống kê: 35 ngày qua (tương đồng lịch sử xuất nhập)
    const thirtyFiveDaysAgoMs = Date.now() - (35 * 24 * 60 * 60 * 1000);
    const thirtyFiveDaysAgoStr = getPastDateStr(35);
    const todayStr = getPastDateStr(0);
    const dateRangeStr = `${thirtyFiveDaysAgoStr} đến ${todayStr} (35 ngày)`;

    // Lọc giao dịch trong 35 ngày qua
    const recent35DaysTxs = transactions.filter(t => t.timestamp >= thirtyFiveDaysAgoMs);

    // Helper tìm BOM theo mã linh kiện
    const findBOMForPart = (pId: string) => {
      const clean = (pId || '')
        .trim()
        .toUpperCase()
        .replace(/^LSR-OUT-/i, '')
        .replace(/\s*-\s*CẮT LASER$/i, '')
        .split(' - ')[0]
        .trim();
      const partObj = parts.find(p => p.id === pId || p.name === pId);
      const cleanObjId = partObj ? partObj.id.trim().toUpperCase() : '';
      const cleanObjName = partObj ? partObj.name.trim().toUpperCase() : '';

      return bom.find(b => {
        const child = (b.childPartId || '').trim().toUpperCase();
        return child === clean || child === pId.trim().toUpperCase() || child === cleanObjId || child === cleanObjName;
      });
    };

    // Helper kiểm tra mã phế liệu
    const isScrap = (pId: string) => {
      const pObj = parts.find(p => p.id === pId || p.name === pId);
      return isLaserScrapItem(pId, pObj?.name);
    };

    // 1. TẬP HỢP TẤT CẢ CÁC LINH KIỆN ĐÃ TỪNG XUẤT HIỆN TẠI KHO OUT CẮT LASER TRONG 35 NGÀY QUA:
    // - Giao dịch hoàn thành sản xuất vào kho OUT: (type STAGE_OUT && location IN) hoặc (type STAGE_IN && location OUT)
    // - Giao dịch xuất kho từ OUT chuyển công đoạn: (stageId LASER && type STAGE_OUT && location OUT) hoặc (sourceStageId LASER)
    // - Tồn kho thực tế hiện tại tại kho OUT Cắt Laser
    // - Lệnh sản xuất Cắt Laser đã có sản lượng đạt (producedQuantity / completedQuantity) trong 35 ngày
    const partStatsMap = new Map<string, {
      producedTxsQty: number;
      exportedTxsQty: number;
      currentStockQty: number;
      poProducedQty: number;
      originalPartId?: string;
    }>();

    const getOrInitStats = (pId: string, origId?: string) => {
      if (!partStatsMap.has(pId)) {
        partStatsMap.set(pId, {
          producedTxsQty: 0,
          exportedTxsQty: 0,
          currentStockQty: 0,
          poProducedQty: 0,
          originalPartId: origId
        });
      }
      const s = partStatsMap.get(pId)!;
      if (origId && !s.originalPartId) s.originalPartId = origId;
      return s;
    };

    // Quét giao dịch 35 ngày qua
    recent35DaysTxs.forEach(t => {
      if (isScrap(t.partId)) return;

      // Hàng được cắt và đưa vào kho OUT tại Cắt Laser
      const isProducedAtLaser = t.stageId === 'LASER' && (
        (t.type === 'STAGE_OUT' && t.location === 'IN') ||
        (t.type === 'STAGE_IN' && t.location === 'OUT')
      );

      if (isProducedAtLaser) {
        const stats = getOrInitStats(t.partId, t.originalPartId);
        stats.producedTxsQty += (t.quantity || 0);
      }

      // Hàng từ kho OUT Cắt Laser được chuyển/xuất sang công đoạn tiếp theo
      const isExportedFromLaserOut = (
        (t.stageId === 'LASER' && t.type === 'STAGE_OUT' && t.location === 'OUT') ||
        (t.sourceStageId === 'LASER' && t.type === 'STAGE_IN')
      );

      if (isExportedFromLaserOut) {
        const stats = getOrInitStats(t.partId, t.originalPartId);
        stats.exportedTxsQty += (t.quantity || 0);
      }
    });

    // Quét tồn kho hiện tại tại kho OUT Cắt Laser
    inventory.forEach(item => {
      if (item.stageId === 'LASER' && item.location === 'OUT' && item.quantity > 0 && !isScrap(item.partId)) {
        const stats = getOrInitStats(item.partId, item.originalPartId);
        stats.currentStockQty += item.quantity;
      }
    });

    // Quét các Lệnh sản xuất Cắt Laser
    productionOrders.forEach(po => {
      if (po.partId && !isScrap(po.partId)) {
        const poProduced = po.producedQuantity || 0;
        if (poProduced > 0) {
          const stats = getOrInitStats(po.partId);
          stats.poProducedQty = Math.max(stats.poProducedQty, poProduced);
        }
      }
    });

    // Bóc tách chi tiết từng linh kiện đã từng xuất hiện ở kho OUT Cắt Laser
    const outDetails = Array.from(partStatsMap.entries()).map(([partId, stats]) => {
      const partObj = parts.find(p => p.id === partId || p.name === partId);
      const bomDef = findBOMForPart(partId);
      const compWeightPerUnit = bomDef?.componentWeight || 0;
      const scrapWeightPerUnit = bomDef?.scrapWeight || 0;
      const parentSheetId = bomDef?.parentPartId || '';
      const parentPart = parts.find(p => p.id === parentSheetId);

      // Số lượng tích lũy đạt trong 35 ngày:
      // Kết hợp số lượng ghi nhận sản xuất, số lượng đã xuất kho + còn tồn, và số lượng PO
      const historicalThroughput = Math.max(
        stats.producedTxsQty,
        stats.exportedTxsQty + stats.currentStockQty,
        stats.poProducedQty,
        stats.currentStockQty
      );

      const effectiveQty = historicalThroughput;
      const itemGoodWeight = effectiveQty * compWeightPerUnit;
      const itemBomScrapWeight = effectiveQty * scrapWeightPerUnit;
      const itemTheoreticalTotal = itemGoodWeight + itemBomScrapWeight;
      const itemBomYield = itemTheoreticalTotal > 0 ? (itemGoodWeight / itemTheoreticalTotal) * 100 : 0;

      return {
        partId,
        partName: partObj?.name || partId,
        originalPartId: stats.originalPartId,
        quantity: effectiveQty, // Tổng SL đã từng xuất hiện / sản xuất trong 35 ngày
        producedTxsQty: stats.producedTxsQty,
        exportedTxsQty: stats.exportedTxsQty,
        currentStockQty: stats.currentStockQty,
        unit: partObj?.unit || 'Bộ',
        parentSheetId,
        parentSheetName: parentPart?.name || parentSheetId,
        compWeightPerUnit,
        scrapWeightPerUnit,
        itemGoodWeight,
        itemBomScrapWeight,
        itemTheoreticalTotal,
        itemBomYield,
        hasBOM: !!bomDef,
      };
    }).filter(d => d.quantity > 0);

    // 2. TẬP HỢP HÀNG LỖI / PHẾ NG CẮT LASER (Đồng bộ chuẩn xác với KHO_NG Cắt Laser):
    // Ưu tiên trực tiếp dữ liệu từ KHO_NG (tồn kho DEFECT hiện tại tại LASER) để khớp 100% với giao diện KHO_NG (Ảnh 3)
    const currentLaserDefectItems = inventory.filter(item => item.stageId === 'LASER' && item.location === 'DEFECT' && item.quantity > 0);
    const defectMap = new Map<string, { quantity: number; originalPartId?: string }>();

    if (currentLaserDefectItems.length > 0) {
      currentLaserDefectItems.forEach(item => {
        defectMap.set(item.partId, { quantity: item.quantity, originalPartId: item.originalPartId });
      });
    } else {
      // Nếu KHO_NG tạm thời trống, quét từ lịch sử báo lỗi trong 35 ngày qua
      recent35DaysTxs.forEach(t => {
        if (t.stageId === 'LASER' && t.type === 'DEFECT') {
          const cur = defectMap.get(t.partId) || { quantity: 0, originalPartId: t.originalPartId };
          cur.quantity += (t.quantity || 0);
          defectMap.set(t.partId, cur);
        }
      });
    }

    const defectDetails = Array.from(defectMap.entries()).map(([partId, info]) => {
      const partObj = parts.find(p => p.id === partId || p.name === partId);
      const isRawSheet = (partObj?.level === 3) || 
                         (partId.startsWith('NVL-TT')) || 
                         (partObj?.unit?.toUpperCase() === 'KG') ||
                         (partObj?.name?.toLowerCase().includes('tôn tấm'));
      
      let itemWeight = 0;
      let bomDef: BOMDefinition | undefined = undefined;

      if (isRawSheet) {
        itemWeight = info.quantity;
      } else {
        bomDef = findBOMForPart(partId);
        if (bomDef) {
          itemWeight = info.quantity * (bomDef.componentWeight + bomDef.scrapWeight);
        } else {
          itemWeight = info.quantity;
        }
      }

      return {
        partId,
        partName: partObj?.name || partId,
        quantity: info.quantity,
        unit: partObj?.unit || (isRawSheet ? 'KG' : 'Bộ'),
        isRawSheet,
        weightKg: itemWeight,
        originalPartId: info.originalPartId,
      };
    }).filter(d => d.quantity > 0);

    // Tổng phế liệu tồn kho tại Cắt Laser (PL-TON-SX) từ Khu vực thu hồi phế liệu tôn
    const scrapInStock = laserScrapItems.reduce((sum, item) => sum + item.quantity, 0);

    // Lọc theo Lệnh sản xuất nếu người dùng chọn
    let filteredOutDetails = outDetails;
    let filteredDefectDetails = defectDetails;

    if (laserYieldPoFilter !== 'ALL') {
      const targetPO = productionOrders.find(po => po.id === laserYieldPoFilter);
      if (targetPO) {
        const targetClean = (targetPO.partId || '').trim().toUpperCase();
        filteredOutDetails = outDetails.filter(d => {
          const dClean = (d.partId || '').trim().toUpperCase();
          return dClean.includes(targetClean) || targetClean.includes(dClean);
        });

        filteredDefectDetails = defectDetails.filter(d => {
          const dClean = (d.partId || '').trim().toUpperCase();
          return dClean.includes(targetClean) || targetClean.includes(dClean);
        });
      }
    }

    // 1. Tổng khối lượng linh kiện tốt (W_good) = SL sản phẩm đạt * KL Linh kiện (Kg)
    const W_good = filteredOutDetails.reduce((sum, d) => sum + d.itemGoodWeight, 0);

    // 2. Tổng khối lượng phế định mức (W_bom_scrap)
    // ĐỒNG BỘ 100% với Khu vực thu hồi phế liệu tôn (SX) PL-TON-SX (Ảnh 1) khi xem toàn bộ kho
    const calculatedBomScrap = filteredOutDetails.reduce((sum, d) => sum + d.itemBomScrapWeight, 0);
    const W_bom_scrap = (laserYieldPoFilter === 'ALL')
      ? (scrapInStock > 0 ? scrapInStock : calculatedBomScrap)
      : calculatedBomScrap;

    // 3. Khối lượng phế NG / lỗi thực tế (W_ng) lấy từ dữ liệu KHO_NG Cắt Laser (Đồng bộ 100% với Ảnh 3)
    const W_ng = filteredDefectDetails.reduce((sum, d) => sum + d.weightKg, 0);

    // 4. Tổng tôn cấp phát đầu vào (W_total_input) = W_good + W_bom_scrap + W_ng
    const W_total_input = W_good + W_bom_scrap + W_ng;

    // 5. Các chỉ số KPI
    // Tỷ lệ cắt tôn chuẩn (BOM Yield %) = [W_good / (W_good + W_bom_scrap)] * 100%
    const theoreticalSum = W_good + W_bom_scrap;
    const bomYield = theoreticalSum > 0 ? (W_good / theoreticalSum) * 100 : 0;

    // Tỷ lệ cắt tôn thực tế (Actual Yield %) = [W_good / W_total_input] * 100%
    const actualYield = W_total_input > 0 ? (W_good / W_total_input) * 100 : 0;

    // Tỷ lệ hao hụt lỗi (NG Rate %) = [W_ng / W_total_input] * 100%
    const ngRate = W_total_input > 0 ? (W_ng / W_total_input) * 100 : 0;

    // Tỷ lệ phế liệu định mức (% so với tổng cấp phát)
    const bomScrapRate = W_total_input > 0 ? (W_bom_scrap / W_total_input) * 100 : 0;

    return {
      W_good,
      W_bom_scrap,
      W_ng,
      W_total_input,
      bomYield,
      actualYield,
      ngRate,
      bomScrapRate,
      scrapInStock,
      dateRangeStr,
      thirtyFiveDaysAgoStr,
      todayStr,
      outDetails: filteredOutDetails,
      defectDetails: filteredDefectDetails,
      productionOrders,
      totalOutQty: filteredOutDetails.reduce((sum, d) => sum + d.quantity, 0),
      totalDefectQty: filteredDefectDetails.reduce((sum, d) => sum + d.quantity, 0),
    };
  }, [inventory, parts, laserScrapItems, laserYieldPoFilter, transactions]);

  const exportLaserYieldReportExcel = () => {
    const wb = XLSX.utils.book_new();
    const title = `BÁO CÁO HIỆU SUẤT CẮT TÔN & TỶ LỆ PHẾ LIỆU (YIELD & SCRAP) - 35 NGÀY`;
    const dateStrFormatted = format(new Date(), 'dd/MM/yyyy HH:mm');
    
    const summaryRows = [
      [title],
      [`Thời gian xuất: ${dateStrFormatted}`],
      [`Khoảng thời gian thống kê: ${laserYieldReport.dateRangeStr}`],
      [`Cơ chế tính toán: Dựa vào toàn bộ lịch sử linh kiện đã từng xuất hiện / xuất nhập tại kho OUT Cắt Laser trong 35 ngày qua`],
      [`Phạm vi lọc: ${laserYieldPoFilter === 'ALL' ? 'Tất cả (Tổng công đoạn Cắt Laser)' : `Lệnh sản xuất ${laserYieldPoFilter}`}`],
      [''],
      ['CHỈ SỐ KPI CHÍNH', 'GIÁ TRỊ', 'ĐƠN VỊ', 'CÔNG THỨC / Ý NGHĨA'],
      ['1. Tổng khối lượng linh kiện tốt (W_good)', laserYieldReport.W_good.toFixed(4), 'KG', 'Tổng KL sản phẩm đạt theo định mức linh kiện (Kg) trong 35 ngày'],
      ['2. Tổng khối lượng phế định mức (W_bom_scrap)', laserYieldReport.W_bom_scrap.toFixed(4), 'KG', 'Tổng KL phế liệu sinh ra theo định mức (PL-TON-SX) trong 35 ngày'],
      ['3. Khối lượng phế NG / lỗi thực tế (W_ng)', laserYieldReport.W_ng.toFixed(4), 'KG', 'Tổng KL lỗi ghi nhận từ KHO_NG và Báo lỗi nhanh trong 35 ngày'],
      ['4. Tổng tôn cấp phát đầu vào (W_total_input)', laserYieldReport.W_total_input.toFixed(4), 'KG', 'W_good + W_bom_scrap + W_ng'],
      ['5. Tỷ lệ cắt tôn chuẩn (BOM Yield %)', `${laserYieldReport.bomYield.toFixed(2)}%`, '%', '[W_good / (W_good + W_bom_scrap)] * 100% (Thu hồi lý thuyết định mức)'],
      ['6. Tỷ lệ cắt tôn thực tế (Actual Yield %)', `${laserYieldReport.actualYield.toFixed(2)}%`, '%', '[W_good / W_total_input] * 100% (Thu hồi thực tế sau lỗi)'],
      ['7. Tỷ lệ hao hụt lỗi (NG Rate %)', `${laserYieldReport.ngRate.toFixed(2)}%`, '%', '[W_ng / W_total_input] * 100% (Tỷ lệ hàng lỗi phế phẩm)'],
      ['8. Tỷ lệ phế liệu định mức (% BOM Scrap)', `${laserYieldReport.bomScrapRate.toFixed(2)}%`, '%', '[W_bom_scrap / W_total_input] * 100%'],
      [''],
      ['CHI TIẾT LINH KIỆN ĐÃ SẢN XUẤT TẠI KHO OUT CẮT LASER (35 NGÀY QUA)'],
      ['Mã linh kiện', 'Tên linh kiện', 'Mã tôn tấm (BOM)', 'SL Tích lũy (35 ngày)', 'Đã xuất kho', 'Tồn OUT hiện tại', 'Đơn vị', 'KL Linh kiện (kg/cái)', 'KL Phế (kg/cái)', 'W_good (kg)', 'W_bom_scrap (kg)', 'Tỷ lệ thu hồi BOM (%)'],
      ...laserYieldReport.outDetails.map(d => [
        d.partId,
        d.partName,
        d.parentSheetId || '(Chưa có BOM)',
        d.quantity,
        d.exportedTxsQty || 0,
        d.currentStockQty || 0,
        d.unit,
        d.compWeightPerUnit.toFixed(4),
        d.scrapWeightPerUnit.toFixed(4),
        d.itemGoodWeight.toFixed(4),
        d.itemBomScrapWeight.toFixed(4),
        `${d.itemBomYield.toFixed(2)}%`
      ]),
      [''],
      ['CHI TIẾT HÀNG LỖI KHO_NG CẮT LASER (35 NGÀY QUA)'],
      ['Mã linh kiện lỗi', 'Tên linh kiện / Tôn tấm', 'Số lượng lỗi', 'Đơn vị', 'Khối lượng lỗi W_ng (KG)'],
      ...laserYieldReport.defectDetails.map(d => [
        d.partId,
        d.partName,
        d.quantity,
        d.unit,
        d.weightKg.toFixed(4)
      ])
    ];

    const ws = XLSX.utils.aoa_to_sheet(summaryRows);
    XLSX.utils.book_append_sheet(wb, ws, "HieuSuat_CatTon_35Ngay");
    XLSX.writeFile(wb, `BaoCao_Yield_CatTon_35Ngay_${format(new Date(), 'yyyyMMdd_HHmm')}.xlsx`);
  };

  const stageSummaries = useMemo(() => {
    return STAGES.map(stage => {
      const stageItems = inventory.filter(item => item.stageId === stage.id);
      
      const formatQty = (val: number) => parseFloat(val.toFixed(4));
      
      const isScrap = (partId: string) => {
        if (stage.id !== 'LASER') return false;
        const partObj = parts.find(p => p.id === partId || p.name === partId);
        return isLaserScrapItem(partId, partObj?.name);
      };

      const inQty = stageItems.filter(i => i.location === 'IN').reduce((sum, i) => sum + i.quantity, 0);
      const outQty = stageItems.filter(i => i.location === 'OUT' && !isScrap(i.partId)).reduce((sum, i) => sum + i.quantity, 0);
      const defectQty = stageItems.filter(i => i.location === 'DEFECT').reduce((sum, i) => sum + i.quantity, 0);
      const scrapQty = stageItems.filter(i => isScrap(i.partId)).reduce((sum, i) => sum + i.quantity, 0);

      return {
        id: stage.id,
        name: stage.name,
        total: formatQty(inQty + outQty + defectQty),
        in: formatQty(inQty),
        out: formatQty(outQty),
        defect: formatQty(defectQty),
        scrap: formatQty(scrapQty),
      };
    });
  }, [inventory, parts]);

  const activeStageConfig = activeChartTab !== 'MECHANICAL' ? STAGE_CHART_CONFIGS[activeChartTab] : null;

  const stageDaySummary = useMemo(() => {
    if (activeChartTab === 'MECHANICAL') {
      const totalConv = productionChartData.reduce((acc, d) => acc + (d.totalConvertedMechanical || 0), 0);
      const totalMandays = productionChartData.reduce((acc, d) => acc + (d.slotMandays || 0), 0);
      const totalTarget = totalMandays * 28.3;
      
      // Trung bình cộng NSLĐ của các khung giờ có ghi nhận người (dù sản phẩm = 0 vẫn tính 0% vào trung bình)
      const validSlotPercentages: number[] = [];
      productionChartData.forEach(d => {
        if ((d.totalWorkersMechanical || 0) > 0) {
          const slotTarget = (d.slotMandays || 0) * 28.3;
          const slotNsld = slotTarget > 0 ? ((d.totalConvertedMechanical || 0) / slotTarget) : 0;
          validSlotPercentages.push(slotNsld);
        }
      });
      const avgNsld = validSlotPercentages.length > 0 
        ? (validSlotPercentages.reduce((a, b) => a + b, 0) / validSlotPercentages.length)
        : (totalTarget > 0 ? (totalConv / totalTarget) : 0);

      return {
        totalConv: totalConv.toFixed(1),
        totalTarget: totalTarget.toFixed(1),
        avgNsld: Math.round(avgNsld * 100),
        mandays: totalMandays.toFixed(2),
      };
    }
    if (!activeStageConfig) return null;
    const cfg = activeStageConfig;
    const totalActual = productionChartData.reduce((acc, d) => acc + (d[cfg.barKey] || 0), 0);
    const totalConv = productionChartData.reduce((acc, d) => acc + (d[cfg.convertedKey] || 0), 0);
    const totalTarget = productionChartData.reduce((acc, d) => acc + (d[cfg.targetKey] || 0), 0);
    
    // Trung bình cộng NSLĐ của các khung giờ có ghi nhận người (dù sản phẩm = 0 vẫn tính 0% vào trung bình)
    const validSlotPercentages: number[] = [];
    productionChartData.forEach(d => {
      if ((d[cfg.targetKey] || 0) > 0) {
        const slotTarget = d[cfg.targetKey] || 0;
        const slotConv = d[cfg.convertedKey] || 0;
        const slotNsld = slotTarget > 0 ? (slotConv / slotTarget) : 0;
        validSlotPercentages.push(slotNsld);
      }
    });
    const avgNsld = validSlotPercentages.length > 0 
      ? (validSlotPercentages.reduce((a, b) => a + b, 0) / validSlotPercentages.length)
      : (totalTarget > 0 ? (totalConv / totalTarget) : 0);

    return {
      totalActual: Math.round(totalActual),
      totalConv: Math.round(totalConv * 10) / 10,
      totalTarget: Math.round(totalTarget * 10) / 10,
      avgNsld: Math.round(avgNsld * 100),
    };
  }, [activeChartTab, activeStageConfig, productionChartData]);

  return (
    <motion.div 
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, y: -20 }}
      className="space-y-8"
    >
      <div className="flex flex-col sm:flex-row justify-between items-stretch sm:items-center gap-3 mb-4 sm:mb-6">
        <div className="flex items-center gap-2">
          <div className="w-2.5 h-6 bg-[#F27D26] rounded-full hidden sm:block" />
          <h3 className="text-base sm:text-xl font-black uppercase text-gray-800 tracking-tight">Tổng quan Sản xuất</h3>
        </div>
        <div className="grid grid-cols-2 sm:flex sm:items-center gap-2 sm:gap-3 w-full sm:w-auto">
          <button
            onClick={() => {
              const savedPeople = storageService.getPeoplePerDay();
              const updatedPeople = { ...savedPeople };
              exportDateRange.forEach(d => {
                if (updatedPeople[d] === undefined) {
                  const autoAvg = storageService.getDailyTotalPeopleAverage(d);
                  if (autoAvg !== null) {
                    updatedPeople[d] = autoAvg;
                  }
                }
              });
              setPeoplePerDay(updatedPeople);
              setShowExportModal(true);
            }}
            className="flex items-center justify-center gap-1.5 px-3 sm:px-4 py-2 sm:py-2.5 bg-green-600 text-white rounded-xl shadow-sm hover:bg-green-700 active:scale-95 transition-all font-bold uppercase tracking-wider text-xs sm:text-sm cursor-pointer"
          >
            <FileSpreadsheet size={16} />
            <span className="hidden sm:inline">Xuất báo cáo SX</span>
            <span className="sm:hidden">Báo cáo SX</span>
          </button>
          <button
            onClick={() => setShowExportPreparationModal(true)}
            className="flex items-center justify-center gap-1.5 px-3 sm:px-4 py-2 sm:py-2.5 bg-green-700 text-white rounded-xl shadow-sm hover:bg-green-800 active:scale-95 transition-all font-bold uppercase tracking-wider text-xs sm:text-sm cursor-pointer"
          >
            <FileSpreadsheet size={16} />
            <span className="hidden sm:inline">Xuất Báo Cáo SX 2</span>
            <span className="sm:hidden">Báo cáo SX 2</span>
          </button>
        </div>
      </div>

      {/* Summary Cards (2 cols on mobile for clean glance, 4 cols on desktop) */}
      <div className="grid grid-cols-2 sm:grid-cols-2 lg:grid-cols-4 gap-2.5 sm:gap-4 lg:gap-6">
        {stageSummaries.filter(s => s.id !== 'GLAZING').map((summary, idx) => {
          const isActive = selectedStageDetail === summary.id;

          return (
            <button 
              key={summary.id} 
              onClick={() => setSelectedStageDetail(isActive ? null : summary.id)}
              className={cn(
                "bg-white p-3 sm:p-5 rounded-xl sm:rounded-2xl border transition-all text-left group relative cursor-pointer active:scale-[0.98]",
                isActive 
                  ? "border-[#F27D26] ring-2 sm:ring-4 ring-[#F27D26]/10 shadow-md" 
                  : "border-gray-200 shadow-2xs hover:shadow-md hover:border-gray-300"
              )}
            >
              <div className="flex justify-between items-start mb-2 sm:mb-4">
                <div className={cn(
                  "p-1.5 sm:p-2 rounded-lg transition-colors shrink-0",
                  isActive ? "bg-blue-600 text-white" : "bg-gray-100 text-gray-600"
                )}>
                  <Package size={16} className="sm:w-5 sm:h-5" />
                </div>
                <div className="flex flex-col items-end">
                  <span className="text-[10px] sm:text-xs font-mono font-bold text-gray-400 uppercase tracking-wider">S{idx + 1}</span>
                  {isActive && <div className="w-1.5 h-1.5 sm:w-2 sm:h-2 rounded-full bg-[#F27D26] mt-0.5 animate-pulse" />}
                </div>
              </div>
              <h3 className="font-bold text-xs sm:text-lg md:text-xl text-gray-900 mb-0.5 sm:mb-1 truncate">{summary.name}</h3>
              <p className="text-xl sm:text-3xl md:text-4xl font-mono font-black tracking-tight text-gray-900">{summary.total}</p>
              
              <div className="mt-2 sm:mt-3 grid grid-cols-3 gap-0.5 sm:gap-1 p-1 sm:p-2 bg-gray-50/90 rounded-lg text-center border border-gray-100">
                <div className="flex flex-col min-w-0">
                  <span className="text-[8px] sm:text-[10px] font-mono font-bold text-gray-400 uppercase leading-none mb-0.5 truncate">IN</span>
                  <span className="text-[11px] sm:text-sm font-mono font-bold text-gray-800 truncate">{summary.in}</span>
                </div>
                <div className="flex flex-col min-w-0 border-x border-gray-200">
                  <span className="text-[8px] sm:text-[10px] font-mono font-bold text-orange-600 uppercase leading-none mb-0.5 truncate">OUT</span>
                  <span className={cn("text-[11px] sm:text-sm font-mono font-bold truncate", isActive ? "text-[#F27D26]" : "text-[#F27D26]/80")}>
                    {summary.out}
                  </span>
                </div>
                <div className="flex flex-col min-w-0">
                  <span className="text-[8px] sm:text-[10px] font-mono font-bold text-red-500 uppercase leading-none mb-0.5 truncate">LỖI</span>
                  <span className="text-[11px] sm:text-sm font-mono font-bold text-red-600 truncate">
                    {summary.defect}
                  </span>
                </div>
              </div>
              {summary.id === 'LASER' && (summary.scrap || 0) > 0 && (
                <div className="mt-1.5 px-2 py-1 bg-amber-50 border border-amber-200/80 rounded-lg flex items-center justify-between text-[9px] sm:text-[10px] font-mono">
                  <span className="font-bold text-amber-700 flex items-center gap-1 truncate">
                    <Layers size={11} className="text-amber-600 shrink-0" /> Phế liệu:
                  </span>
                  <span className="font-black text-amber-800 shrink-0">{summary.scrap?.toFixed(1)} kg</span>
                </div>
              )}
              <div className="mt-2 sm:mt-3 pt-1.5 sm:pt-2.5 border-t border-gray-100 flex items-center justify-between text-[10px] sm:text-xs font-bold uppercase tracking-wider text-gray-400 group-hover:text-[#F27D26] transition-colors">
                <span className="truncate">{isActive ? 'Đang xem' : 'Chi tiết'}</span>
                <ChevronRight size={13} className={cn("transition-transform shrink-0", isActive && "rotate-90 text-[#F27D26]")} />
              </div>
            </button>
          );
        })}
      </div>

      <AnimatePresence mode="wait">
        {selectedStageDetail ? (
          <motion.div
            key="stage-detail"
            initial={{ opacity: 0, height: 0 }}
            animate={{ opacity: 1, height: 'auto' }}
            exit={{ opacity: 0, height: 0 }}
            className="overflow-hidden"
          >
            <div className="bg-white text-gray-900 rounded-xl sm:rounded-2xl p-3.5 sm:p-6 md:p-8 shadow-lg border border-gray-200">
              <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 sm:gap-6 mb-4 sm:mb-8">
                <div className="flex items-center gap-3">
                  <div className="bg-[#F27D26] p-2 sm:p-3 rounded-xl shrink-0">
                    <Package size={20} className="text-white sm:w-6 sm:h-6" />
                  </div>
                  <div className="min-w-0">
                    <h2 className="text-base sm:text-2xl md:text-3xl font-bold tracking-tight truncate">Chi tiết tồn kho: {STAGES.find(s => s.id === selectedStageDetail)?.name}</h2>
                    <p className="text-[11px] sm:text-sm font-mono opacity-80 uppercase tracking-wider text-gray-500">Phân rã theo từng mã linh kiện</p>
                  </div>
                </div>
                
                <div className="flex items-center gap-2 sm:gap-4 w-full md:w-auto">
                  <div className="relative flex-1 md:w-64">
                    <input 
                      type="text"
                      placeholder="Tìm linh kiện..."
                      value={searchTerm}
                      onChange={(e) => setSearchTerm(e.target.value)}
                      className="w-full pl-9 sm:pl-11 pr-3 py-2 sm:py-2.5 rounded-xl border border-gray-200 focus:border-[#F27D26] outline-none text-xs sm:text-sm"
                    />
                    <Edit2 size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                  </div>
                  <button 
                    onClick={() => {
                      if (selectedStageDetail) {
                        setDefectModal({ partId: '', stageId: selectedStageDetail });
                      }
                    }}
                    className="flex items-center justify-center gap-1.5 px-3 sm:px-5 py-2 sm:py-2.5 bg-red-50 text-red-600 rounded-xl border border-red-200 font-bold uppercase text-[11px] sm:text-xs tracking-wider hover:bg-red-100 transition-all shadow-2xs active:scale-95 shrink-0 cursor-pointer"
                  >
                    <AlertCircle size={15} />
                    <span className="hidden sm:inline">Báo lỗi nhanh (NG)</span>
                    <span className="sm:hidden">Báo lỗi</span>
                  </button>
                  <button 
                    onClick={() => { setSelectedStageDetail(null); setSearchTerm(''); }}
                    className="p-1.5 sm:p-2 hover:bg-gray-100 rounded-full transition-colors text-gray-400 shrink-0 cursor-pointer"
                  >
                    <X size={20} />
                  </button>
                </div>
              </div>

              {selectedStageDetail === 'DCLR' && (
                <div className="mb-4 sm:mb-8 p-3.5 sm:p-6 bg-red-50 border border-red-200 rounded-xl sm:rounded-2xl flex flex-col md:flex-row justify-between items-start md:items-center gap-3 sm:gap-6 shadow-2xs">
                  <div className="flex items-center gap-3">
                    <div className="bg-red-600 p-2.5 sm:p-4 rounded-xl text-white shadow-sm shrink-0">
                      <Trash2 size={22} className="sm:w-7 sm:h-7" />
                    </div>
                    <div>
                      <h4 className="text-sm sm:text-xl font-black text-red-700 uppercase italic">Xóa tất cả tồn kho DCLR OUT</h4>
                      <p className="text-xs sm:text-sm text-red-600 font-medium opacity-80">Hành động này sẽ xóa vĩnh viễn toàn bộ số lượng đang có tại kho OUT của DCLR.</p>
                    </div>
                  </div>
                  <button 
                    onClick={async () => {
                      const pwd = await customPrompt('Nhập mật khẩu ADMIN để xóa hết tồn kho DCLR OUT:', '', true);
                      if (pwd === 'admin123') {
                        if (await customConfirm('BẠN CÓ CHẮC CHẮN MUỐN XÓA HẾT TỒN KHO DCLR OUT? Hành động này không thể hoàn tác!')) {
                          storageService.clearStageInventory('DCLR', 'OUT');
                          refreshData();
                          await customAlert('Đã xóa sạch tồn kho DCLR OUT thành công!');
                        }
                      } else if (pwd !== null) {
                        await customAlert('Mật khẩu không chính xác!');
                      }
                    }}
                    className="w-full sm:w-auto px-4 sm:px-8 py-2.5 sm:py-3.5 bg-red-600 text-white rounded-xl font-black uppercase tracking-wider text-xs sm:text-sm hover:bg-black transition-all flex items-center justify-center gap-2 shadow-sm active:scale-95 group cursor-pointer"
                  >
                    <Trash2 size={16} className="group-hover:rotate-12 transition-transform" />
                    <span>Xác nhận xóa hết</span>
                  </button>
                </div>
              )}

              {selectedStageDetail === 'LASER' && (
                <>
                  {/* WIDGET: HIỆU SUẤT CẮT TÔN & TỶ LỆ PHẾ LIỆU (YIELD & SCRAP) */}
                  <div 
                    style={{ fontFamily: '"Times New Roman", Times, Georgia, serif' }}
                    className="mb-4 sm:mb-6 p-4 sm:p-6 bg-gradient-to-br from-white via-indigo-50/25 to-blue-50/35 border-2 border-indigo-300 rounded-xl sm:rounded-2xl shadow-sm text-gray-900"
                  >
                    {/* Header */}
                    <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-3 sm:gap-4 pb-4 border-b border-indigo-200">
                      <div className="flex items-center gap-3">
                        <div className="p-3 bg-gradient-to-br from-indigo-700 to-blue-700 text-white rounded-xl shadow-xs shrink-0">
                          <Scissors size={24} className="sm:w-6 sm:h-6" />
                        </div>
                        <div>
                          <div className="flex items-center gap-2.5 flex-wrap">
                            <h3 className="text-lg sm:text-2xl md:text-3xl font-black uppercase text-gray-950 tracking-normal">
                              Hiệu Suất Cắt Tôn & Tỷ Lệ Phế Liệu
                            </h3>
                            <span className="px-3 py-1 bg-indigo-100 text-indigo-900 text-xs sm:text-sm font-black rounded-lg uppercase tracking-wide border border-indigo-200">
                              Yield & Scrap KPI
                            </span>
                            <span className="px-3 py-1 bg-amber-100 text-amber-950 text-xs sm:text-sm font-black rounded-lg uppercase tracking-wide border border-amber-300 flex items-center gap-1">
                              📅 35 ngày: {laserYieldReport.thirtyFiveDaysAgoStr} ➔ {laserYieldReport.todayStr}
                            </span>
                          </div>
                          <p className="text-xs sm:text-sm text-gray-700 font-bold mt-1.5 leading-relaxed">
                            Thống kê dựa trên toàn bộ linh kiện đã từng xuất hiện / xuất nhập tại kho OUT Cắt Laser trong 35 ngày qua (W_good đạt tích lũy, W_bom_scrap phế định mức đồng bộ Kho phế, W_ng lỗi KHO_NG)
                          </p>
                        </div>
                      </div>

                      {/* Controls */}
                      <div className="flex items-center gap-2 sm:gap-3 w-full lg:w-auto flex-wrap justify-between lg:justify-end">
                        <div className="flex items-center gap-1.5 bg-white border-2 border-indigo-200 rounded-xl px-3 py-2 shadow-2xs">
                          <Filter size={15} className="text-indigo-700 shrink-0" />
                          <span className="text-xs sm:text-sm font-black text-gray-800 shrink-0">Lệnh SX:</span>
                          <select
                            value={laserYieldPoFilter}
                            onChange={(e) => setLaserYieldPoFilter(e.target.value)}
                            className="bg-transparent text-xs sm:text-sm font-black text-indigo-950 outline-none cursor-pointer pr-1 max-w-[170px] sm:max-w-[220px] truncate"
                          >
                            <option value="ALL">-- Tất cả (Tổng kho Cắt Laser 35 ngày) --</option>
                            {laserYieldReport.productionOrders.map(po => {
                              const pName = parts.find(p => p.id === po.partId || p.name === po.partId)?.name || po.partId;
                              return (
                                <option key={po.id} value={po.id}>
                                  {po.id} - {pName} ({po.producedQuantity || 0}/{po.targetQuantity} cái)
                                </option>
                              );
                            })}
                          </select>
                        </div>

                        <button
                          onClick={exportLaserYieldReportExcel}
                          className="flex items-center gap-1.5 px-3.5 py-2 bg-emerald-50 text-emerald-800 hover:bg-emerald-100 border-2 border-emerald-300 rounded-xl text-xs sm:text-sm font-black transition-all shadow-2xs active:scale-95 cursor-pointer shrink-0"
                          title="Xuất file Excel báo cáo chi tiết tỷ lệ cắt tôn & phế liệu (35 ngày qua)"
                        >
                          <FileSpreadsheet size={16} className="text-emerald-700" />
                          <span>Xuất Excel</span>
                        </button>

                        <button
                          onClick={() => setShowLaserYieldDetails(!showLaserYieldDetails)}
                          className="flex items-center gap-1 px-3 py-2 bg-white hover:bg-gray-100 border-2 border-gray-300 rounded-xl text-xs sm:text-sm font-black text-gray-800 transition-all shadow-2xs active:scale-95 cursor-pointer shrink-0"
                        >
                          <span>{showLaserYieldDetails ? 'Thu gọn' : 'Bóc tách chi tiết'}</span>
                          <ChevronDown size={16} className={cn("transition-transform", showLaserYieldDetails && "rotate-180")} />
                        </button>
                      </div>
                    </div>

                    {/* 4 KPI Cards */}
                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 sm:gap-4 mt-4">
                      {/* Card 1: BOM Yield */}
                      <div className="bg-white p-4 sm:p-5 rounded-xl border-2 border-blue-200 shadow-2xs flex flex-col justify-between">
                        <div>
                          <div className="flex items-center justify-between mb-2">
                            <span className="text-xs sm:text-sm font-black uppercase tracking-wider text-blue-900">1. Tỷ lệ cắt chuẩn (BOM Yield)</span>
                            <span className="px-2 py-0.5 bg-blue-100 text-blue-900 text-[11px] sm:text-xs font-black rounded border border-blue-300">Định mức</span>
                          </div>
                          <div className="flex items-baseline gap-2">
                            <span className="text-3xl sm:text-4xl font-black text-blue-700">
                              {laserYieldReport.bomYield.toFixed(2)}%
                            </span>
                            <span className="text-xs sm:text-sm text-gray-600 font-bold">lý thuyết</span>
                          </div>
                        </div>
                        <div className="mt-3.5">
                          <div className="w-full bg-blue-100 h-2.5 rounded-full overflow-hidden">
                            <div 
                              className="bg-blue-600 h-full rounded-full transition-all duration-500" 
                              style={{ width: `${Math.min(100, Math.max(0, laserYieldReport.bomYield))}%` }}
                            />
                          </div>
                          <div className="text-xs sm:text-sm text-gray-700 font-bold mt-2 flex justify-between">
                            <span>LK tốt (35N): <b className="text-blue-900 font-black">{laserYieldReport.W_good.toFixed(1)} kg</b></span>
                            <span>Phế: <b className="text-amber-800 font-black">{laserYieldReport.W_bom_scrap.toFixed(1)} kg</b></span>
                          </div>
                        </div>
                      </div>

                      {/* Card 2: Actual Yield */}
                      <div className="bg-white p-4 sm:p-5 rounded-xl border-2 border-emerald-200 shadow-2xs flex flex-col justify-between">
                        <div>
                          <div className="flex items-center justify-between mb-2">
                            <span className="text-xs sm:text-sm font-black uppercase tracking-wider text-emerald-900">2. Tỷ lệ cắt thực tế (Actual Yield)</span>
                            <span className="px-2 py-0.5 bg-emerald-100 text-emerald-900 text-[11px] sm:text-xs font-black rounded border border-emerald-300">Thực tế</span>
                          </div>
                          <div className="flex items-baseline gap-2">
                            <span className="text-3xl sm:text-4xl font-black text-emerald-700">
                              {laserYieldReport.actualYield.toFixed(2)}%
                            </span>
                            {laserYieldReport.bomYield > 0 && (
                              <span className={cn(
                                "text-xs font-black px-1.5 py-0.5 rounded",
                                laserYieldReport.actualYield >= laserYieldReport.bomYield ? "bg-emerald-100 text-emerald-900" : "bg-rose-100 text-rose-900"
                              )}>
                                {laserYieldReport.actualYield >= laserYieldReport.bomYield ? '+' : ''}{(laserYieldReport.actualYield - laserYieldReport.bomYield).toFixed(2)}%
                              </span>
                            )}
                          </div>
                        </div>
                        <div className="mt-3.5">
                          <div className="w-full bg-emerald-100 h-2.5 rounded-full overflow-hidden">
                            <div 
                              className="bg-emerald-500 h-full rounded-full transition-all duration-500" 
                              style={{ width: `${Math.min(100, Math.max(0, laserYieldReport.actualYield))}%` }}
                            />
                          </div>
                          <div className="text-xs sm:text-sm text-gray-700 font-bold mt-2 flex justify-between">
                            <span>W_good / W_total_input</span>
                            <span className="text-emerald-900 font-black">{laserYieldReport.W_good.toFixed(1)} kg</span>
                          </div>
                        </div>
                      </div>

                      {/* Card 3: NG Rate */}
                      <div className="bg-white p-4 sm:p-5 rounded-xl border-2 border-rose-200 shadow-2xs flex flex-col justify-between">
                        <div>
                          <div className="flex items-center justify-between mb-2">
                            <span className="text-xs sm:text-sm font-black uppercase tracking-wider text-rose-900">3. Tỷ lệ hao hụt lỗi (NG Rate)</span>
                            <span className="px-2 py-0.5 bg-rose-100 text-rose-900 text-[11px] sm:text-xs font-black rounded border border-rose-300">KHO_NG</span>
                          </div>
                          <div className="flex items-baseline gap-2">
                            <span className="text-3xl sm:text-4xl font-black text-rose-700">
                              {laserYieldReport.ngRate.toFixed(2)}%
                            </span>
                            <span className="text-xs sm:text-sm text-gray-600 font-bold">{laserYieldReport.defectDetails.length} mã lỗi</span>
                          </div>
                        </div>
                        <div className="mt-3.5">
                          <div className="w-full bg-rose-100 h-2.5 rounded-full overflow-hidden">
                            <div 
                              className="bg-rose-500 h-full rounded-full transition-all duration-500" 
                              style={{ width: `${Math.min(100, Math.max(0, laserYieldReport.ngRate))}%` }}
                            />
                          </div>
                          <div className="text-xs sm:text-sm text-gray-700 font-bold mt-2 flex justify-between">
                            <span>Khối lượng lỗi W_ng (KHO_NG)</span>
                            <span className="text-rose-900 font-black">{laserYieldReport.W_ng.toFixed(2)} kg</span>
                          </div>
                        </div>
                      </div>

                      {/* Card 4: Total Input */}
                      <div className="bg-white p-4 sm:p-5 rounded-xl border-2 border-indigo-200 shadow-2xs flex flex-col justify-between">
                        <div>
                          <div className="flex items-center justify-between mb-2">
                            <span className="text-xs sm:text-sm font-black uppercase tracking-wider text-indigo-900">4. Tổng tôn cấp phát (W_input)</span>
                            <span className="px-2 py-0.5 bg-indigo-100 text-indigo-900 text-[11px] sm:text-xs font-black rounded border border-indigo-300">Đầu vào</span>
                          </div>
                          <div className="flex items-baseline gap-2">
                            <span className="text-3xl sm:text-4xl font-black text-indigo-950">
                              {laserYieldReport.W_total_input.toFixed(2)}
                            </span>
                            <span className="text-sm font-black text-indigo-700 uppercase">KG</span>
                          </div>
                        </div>
                        <div className="mt-3.5">
                          <div className="text-xs sm:text-sm text-gray-700 font-bold flex items-center justify-between border-t border-gray-200 pt-2">
                            <span>LK: <b className="text-blue-900 font-black">{laserYieldReport.W_good.toFixed(1)}</b></span>
                            <span>Phế: <b className="text-amber-800 font-black">{laserYieldReport.W_bom_scrap.toFixed(1)}</b></span>
                            <span>NG: <b className="text-rose-800 font-black">{laserYieldReport.W_ng.toFixed(1)}</b></span>
                          </div>
                        </div>
                      </div>
                    </div>

                    {/* Stacked Progress Bar (100% Distribution) */}
                    <div className="mt-5 pt-3.5 border-t border-indigo-200">
                      <div className="flex items-center justify-between text-xs sm:text-sm font-black text-gray-800 mb-2">
                        <span className="flex items-center gap-1.5">
                          <Scale size={16} className="text-indigo-700" />
                          <span>Phân bổ 100% Khối lượng Tôn cấp phát đầu vào:</span>
                        </span>
                        <span className="text-indigo-950 font-black text-sm">
                          {laserYieldReport.W_total_input > 0 ? `${laserYieldReport.W_total_input.toFixed(2)} KG (100%)` : 'Chưa có phát sinh trong 35 ngày'}
                        </span>
                      </div>

                      <div className="w-full bg-gray-100 h-5 rounded-full overflow-hidden flex border-2 border-gray-300 p-0.5">
                        <div 
                          className="bg-emerald-500 h-full rounded-l-full transition-all duration-500 relative cursor-pointer" 
                          style={{ width: `${Math.min(100, Math.max(0, laserYieldReport.actualYield))}%` }}
                          title={`Linh kiện đạt: ${laserYieldReport.W_good.toFixed(2)} kg (${laserYieldReport.actualYield.toFixed(2)}%)`}
                        />
                        <div 
                          className="bg-amber-400 h-full transition-all duration-500 relative cursor-pointer" 
                          style={{ width: `${Math.min(100, Math.max(0, laserYieldReport.bomScrapRate))}%` }}
                          title={`Phế liệu định mức: ${laserYieldReport.W_bom_scrap.toFixed(2)} kg (${laserYieldReport.bomScrapRate.toFixed(2)}%)`}
                        />
                        <div 
                          className="bg-rose-500 h-full rounded-r-full transition-all duration-500 relative cursor-pointer" 
                          style={{ width: `${Math.min(100, Math.max(0, laserYieldReport.ngRate))}%` }}
                          title={`Hao hụt lỗi NG: ${laserYieldReport.W_ng.toFixed(2)} kg (${laserYieldReport.ngRate.toFixed(2)}%)`}
                        />
                      </div>

                      {/* Legend */}
                      <div className="flex flex-wrap items-center gap-4 sm:gap-8 mt-2.5 text-xs sm:text-sm font-bold text-gray-700">
                        <div className="flex items-center gap-2">
                          <div className="w-3.5 h-3.5 rounded-full bg-emerald-500 shrink-0" />
                          <span>Linh kiện đạt (W_good):</span>
                          <span className="font-black text-emerald-800">{laserYieldReport.W_good.toFixed(2)} kg ({laserYieldReport.actualYield.toFixed(2)}%)</span>
                        </div>
                        <div className="flex items-center gap-2">
                          <div className="w-3.5 h-3.5 rounded-full bg-amber-400 shrink-0" />
                          <span>Phế liệu thu hồi (W_bom_scrap):</span>
                          <span className="font-black text-amber-800">{laserYieldReport.W_bom_scrap.toFixed(2)} kg ({laserYieldReport.bomScrapRate.toFixed(2)}%)</span>
                        </div>
                        <div className="flex items-center gap-2">
                          <div className="w-3.5 h-3.5 rounded-full bg-rose-500 shrink-0" />
                          <span>Hàng lỗi / phế NG (W_ng):</span>
                          <span className="font-black text-rose-800">{laserYieldReport.W_ng.toFixed(2)} kg ({laserYieldReport.ngRate.toFixed(2)}%)</span>
                        </div>
                      </div>
                    </div>

                    {/* Expandable Breakdown Tables */}
                    {showLaserYieldDetails && (
                      <div className="mt-5 pt-4 border-t-2 border-indigo-200 space-y-5">
                        {/* Table 1: Sản phẩm đạt tại kho OUT */}
                        <div>
                          <div className="flex items-center justify-between mb-2.5 flex-wrap gap-2">
                            <h4 className="text-xs sm:text-sm font-black uppercase tracking-wider text-gray-900 flex items-center gap-1.5">
                              <CheckCircle2 size={16} className="text-emerald-700" />
                              <span>Chi tiết Linh kiện đã xuất hiện tại kho OUT Cắt Laser ({laserYieldReport.outDetails.length} linh kiện - Lịch sử 35 ngày)</span>
                            </h4>
                            <span className="text-xs sm:text-sm text-gray-700 font-bold">
                              Tổng SL đạt tích lũy: <b className="font-black text-gray-950">{laserYieldReport.totalOutQty}</b> cái | Tổng W_good: <b className="text-emerald-800 font-black">{laserYieldReport.W_good.toFixed(2)} kg</b>
                            </span>
                          </div>

                          <div className="bg-white rounded-xl border-2 border-gray-200 overflow-x-auto shadow-2xs">
                            <table className="w-full text-left text-xs sm:text-sm border-collapse">
                              <thead>
                                <tr className="bg-gray-100 border-b-2 border-gray-300 uppercase text-gray-800 text-[11px] sm:text-xs font-black">
                                  <th className="p-3 pl-3.5">Mã linh kiện</th>
                                  <th className="p-3">Tên linh kiện</th>
                                  <th className="p-3">Tôn tấm (BOM)</th>
                                  <th className="p-3 text-right">SL Tích lũy (35N)</th>
                                  <th className="p-3 text-right">Định mức LK</th>
                                  <th className="p-3 text-right">Định mức Phế</th>
                                  <th className="p-3 text-right text-emerald-800">W_good (kg)</th>
                                  <th className="p-3 text-right text-amber-800">W_phế (kg)</th>
                                  <th className="p-3 pr-3.5 text-right text-blue-800">BOM Yield</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-gray-200 font-bold">
                                {laserYieldReport.outDetails.length > 0 ? (
                                  laserYieldReport.outDetails.map((d, idx) => (
                                    <tr key={idx} className="hover:bg-blue-50/40 transition-colors">
                                      <td className="p-3 pl-3.5 font-black text-gray-950">{d.partId}</td>
                                      <td className="p-3 text-gray-800 max-w-[220px] truncate">{d.partName}</td>
                                      <td className="p-3 text-gray-600">
                                        {d.parentSheetId || <span className="text-amber-700 italic">Chưa có BOM</span>}
                                      </td>
                                      <td className="p-3 text-right">
                                        <span className="font-black text-gray-950">{d.quantity} {d.unit}</span>
                                        {(d.exportedTxsQty > 0 || d.currentStockQty > 0) && (
                                          <span className="block text-[10px] sm:text-xs text-gray-500 font-medium">
                                            (Đã xuất: {d.exportedTxsQty || 0} | Tồn OUT: {d.currentStockQty || 0})
                                          </span>
                                        )}
                                      </td>
                                      <td className="p-3 text-right text-gray-700">{d.compWeightPerUnit.toFixed(4)} kg</td>
                                      <td className="p-3 text-right text-gray-700">{d.scrapWeightPerUnit.toFixed(4)} kg</td>
                                      <td className="p-3 text-right font-black text-emerald-800">{d.itemGoodWeight.toFixed(4)}</td>
                                      <td className="p-3 text-right font-black text-amber-800">{d.itemBomScrapWeight.toFixed(4)}</td>
                                      <td className="p-3 pr-3.5 text-right font-black text-blue-800">
                                        {d.hasBOM ? `${d.itemBomYield.toFixed(2)}%` : '-'}
                                      </td>
                                    </tr>
                                  ))
                                ) : (
                                  <tr>
                                    <td colSpan={9} className="p-5 text-center text-gray-500 font-bold italic">
                                      Không có linh kiện nào phát sinh tại kho OUT Cắt Laser trong 35 ngày qua.
                                    </td>
                                  </tr>
                                )}
                              </tbody>
                            </table>
                          </div>
                        </div>

                        {/* Table 2: Chi tiết phế phẩm lỗi KHO_NG */}
                        <div>
                          <div className="flex items-center justify-between mb-2.5 flex-wrap gap-2">
                            <h4 className="text-xs sm:text-sm font-black uppercase tracking-wider text-gray-900 flex items-center gap-1.5">
                              <AlertCircle size={16} className="text-rose-700" />
                              <span>Chi tiết Hàng lỗi / Phế NG tại KHO_NG Cắt Laser ({laserYieldReport.defectDetails.length} mục)</span>
                            </h4>
                            <span className="text-xs sm:text-sm text-gray-700 font-bold">
                              Tổng khối lượng lỗi W_ng (KHO_NG): <b className="text-rose-800 font-black">{laserYieldReport.W_ng.toFixed(4)} kg</b>
                            </span>
                          </div>

                          <div className="bg-white rounded-xl border-2 border-gray-200 overflow-x-auto shadow-2xs">
                            <table className="w-full text-left text-xs sm:text-sm border-collapse">
                              <thead>
                                <tr className="bg-gray-100 border-b-2 border-gray-300 uppercase text-gray-800 text-[11px] sm:text-xs font-black">
                                  <th className="p-3 pl-3.5">Mã linh kiện lỗi</th>
                                  <th className="p-3">Tên linh kiện / Tôn tấm</th>
                                  <th className="p-3">Phân loại</th>
                                  <th className="p-3 text-right">Số lượng lỗi</th>
                                  <th className="p-3 pr-3.5 text-right text-rose-800">Khối lượng lỗi W_ng (KG)</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-gray-200 font-bold">
                                {laserYieldReport.defectDetails.length > 0 ? (
                                  laserYieldReport.defectDetails.map((d, idx) => (
                                    <tr key={idx} className="hover:bg-rose-50/40 transition-colors">
                                      <td className="p-3 pl-3.5 font-black text-gray-950">{d.partId}</td>
                                      <td className="p-3 text-gray-800">{d.partName}</td>
                                      <td className="p-3">
                                        <span className={cn(
                                          "px-2 py-0.5 rounded text-[11px] font-black",
                                          d.isRawSheet ? "bg-amber-100 text-amber-900 border border-amber-300" : "bg-blue-100 text-blue-900 border border-blue-300"
                                        )}>
                                          {d.isRawSheet ? 'Tôn tấm phôi' : 'Linh kiện dập/cắt'}
                                        </span>
                                      </td>
                                      <td className="p-3 text-right font-black text-gray-950">
                                        {d.quantity} {d.unit}
                                      </td>
                                      <td className="p-3 pr-3.5 text-right font-black text-rose-800">
                                        {d.weightKg.toFixed(4)}
                                      </td>
                                    </tr>
                                  ))
                                ) : (
                                  <tr>
                                    <td colSpan={5} className="p-5 text-center text-gray-500 font-bold italic">
                                      Không có hàng lỗi ghi nhận tại KHO_NG Cắt Laser.
                                    </td>
                                  </tr>
                                )}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      </div>
                    )}
                  </div>

                  <div className="mb-4 sm:mb-8 p-4 sm:p-6 bg-gradient-to-r from-amber-50/90 via-amber-50/40 to-orange-50/60 border-2 border-amber-300 rounded-xl sm:rounded-2xl shadow-sm">
                  <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
                    <div className="flex items-center gap-3.5">
                      <div className="p-3 bg-amber-500 text-white rounded-xl shadow-sm shrink-0">
                        <Layers size={22} className="sm:w-6 sm:h-6" />
                      </div>
                      <div>
                        <div className="flex items-center gap-2 flex-wrap">
                          <h3 className="text-sm sm:text-lg font-black uppercase text-amber-950 tracking-wide">
                            Khu vực thu hồi Phế liệu tôn (SX)
                          </h3>
                          <span className="px-2 py-0.5 bg-amber-200/80 text-amber-900 text-[10px] font-bold rounded-md uppercase tracking-wider">
                            Tách riêng khỏi kho OUT
                          </span>
                        </div>
                        <p className="text-xs text-amber-800/80 font-medium mt-0.5">
                          Tự động sinh ra từ định mức hao hụt khi cắt tôn tấm (kg phế liệu phôi cắt laser)
                        </p>
                      </div>
                    </div>

                    <div className="flex items-center gap-3 w-full sm:w-auto justify-between sm:justify-end">
                      <div className="text-right">
                        <div className="text-[10px] font-bold font-mono uppercase text-amber-700 tracking-wider">Tổng khối lượng phế liệu</div>
                        <div className="flex items-baseline gap-1.5 justify-end">
                          <span className="text-2xl sm:text-3xl font-black font-mono text-amber-600">
                            {laserScrapItems.reduce((sum, item) => sum + item.quantity, 0).toFixed(4)}
                          </span>
                          <span className="text-xs font-bold font-mono uppercase text-amber-800">KG</span>
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Danh sách các mã phế liệu */}
                  <div className="mt-4 pt-4 border-t border-amber-200/70 grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                    {laserScrapItems.length > 0 ? (
                      laserScrapItems.map((item, idx) => {
                        const part = parts.find(p => p.id === item.partId || p.name === item.partId);
                        const ageBatches = getInventoryBatches(item, transactions);
                        const ageInfo = ageBatches.length > 0 ? ageBatches[0].ageInfo : null;

                        return (
                          <div key={idx} className="bg-white p-3.5 rounded-xl border border-amber-200 shadow-2xs flex justify-between items-center gap-3 hover:border-amber-300 transition-colors">
                            <div className="min-w-0 flex-1">
                              <div className="font-bold text-sm text-gray-900 truncate">
                                {part?.name || item.partId}
                              </div>
                              <div className="text-xs font-mono text-amber-700/80 flex items-center gap-2">
                                <span>{item.partId}</span>
                                <span className="text-[10px] bg-amber-100 text-amber-800 px-1.5 py-0.2 rounded font-sans font-bold">Thu hồi</span>
                              </div>
                              {ageInfo && (
                                <div className={`text-[11px] font-bold mt-1 ${ageInfo.colorClass}`}>
                                  {ageInfo.text}
                                </div>
                              )}
                            </div>

                            <div className="flex items-center gap-1.5 shrink-0">
                              <div className="text-right mr-1">
                                <div className="text-lg font-black font-mono text-amber-600">
                                  {item.quantity.toFixed(4)}
                                </div>
                                <div className="text-[10px] font-mono text-gray-400 uppercase font-bold">KG</div>
                              </div>

                              <button
                                onClick={async () => {
                                  const pwd = await customPrompt('Nhập mật khẩu để sửa số lượng phế liệu:', '', true);
                                  if (pwd === 'admin123') {
                                    const newQty = await customPrompt(`Nhập khối lượng phế liệu mới (KG) cho ${item.partId}:`, String(item.quantity));
                                    if (newQty !== null) {
                                      storageService.setInventoryQuantity(item.partId, 'LASER', item.location, parseFloat(newQty) || 0, item.originalPartId);
                                      refreshData();
                                    }
                                  } else if (pwd !== null) {
                                    await customAlert('Mật khẩu không chính xác!');
                                  }
                                }}
                                className="p-2 text-blue-600 hover:bg-blue-50 rounded-lg transition-colors cursor-pointer"
                                title="Sửa số lượng phế liệu"
                              >
                                <Edit2 size={16} />
                              </button>

                              <button
                                onClick={async () => {
                                  const pwd = await customPrompt('Nhập mật khẩu để xóa tồn phế liệu:', '', true);
                                  if (pwd === 'admin123') {
                                    if (await customConfirm(`Bạn có chắc chắn muốn xóa tồn phế liệu của ${item.partId}?`)) {
                                      storageService.deleteInventoryItem(item.partId, 'LASER', item.location, item.originalPartId);
                                      refreshData();
                                    }
                                  } else if (pwd !== null) {
                                    await customAlert('Mật khẩu không chính xác!');
                                  }
                                }}
                                className="p-2 text-red-600 hover:bg-red-50 rounded-lg transition-colors cursor-pointer"
                                title="Xóa tồn phế liệu"
                              >
                                <Trash2 size={16} />
                              </button>
                            </div>
                          </div>
                        );
                      })
                    ) : (
                      <div className="col-span-full py-2 text-center text-xs font-mono italic text-amber-700/70">
                        Chưa có tồn phế liệu tôn phát sinh tại công đoạn Cắt Laser.
                      </div>
                    )}
                  </div>
                </div>
              </>
            )}

              <div className="grid grid-cols-1 lg:grid-cols-3 gap-4 sm:gap-8">
                {/* KHO_IN Detail */}
                <div className="space-y-4">
                  <div className="flex items-center justify-between px-2">
                    <div className="flex items-center gap-2">
                      <div className="w-3 h-3 rounded-full bg-blue-500" />
                      <h3 className="font-mono text-base font-bold uppercase tracking-widest opacity-80">KHO_IN (Chờ sản xuất)</h3>
                    </div>
                    {['LASER', 'BENDING', 'WELDING', 'PAINTING'].includes(selectedStageDetail!) && (
                      <button 
                        onClick={() => setShowManualAddModal({ stageId: selectedStageDetail as StageId, location: 'IN' })}
                        className="text-xs bg-blue-100 text-blue-600 hover:bg-blue-200 px-3 py-1.5 rounded-lg font-bold flex items-center gap-1 transition-colors cursor-pointer"
                      >
                        <Plus size={14} /> Thêm tồn
                      </button>
                    )}
                  </div>
<div className="bg-gray-50 rounded-2xl border border-gray-100 overflow-hidden">
  <table className="w-full text-left border-collapse">
    <thead>
      <tr className="border-b border-gray-200 bg-gray-100/50">
        <th className="p-4 text-sm font-mono uppercase opacity-60">Linh kiện</th>
        <th className="p-4 text-sm font-mono uppercase opacity-60 text-right">Số lượng</th>
      </tr>
    </thead>
                      <tbody className="divide-y divide-gray-100">
                        {inventory
                          .filter(item => item.stageId === selectedStageDetail && item.location === 'IN' && item.quantity > 0)
                          .filter(item => {
                            const partStr = item.partId;
                            const part = parts.find(p => p.id === partStr || p.name === partStr);
                            const originalPart = item.originalPartId ? parts.find(p => p.id === item.originalPartId || p.name === item.originalPartId) : null;
                            const searchLower = searchTerm.toLowerCase();
                            const matchesSearch = item.partId.toLowerCase().includes(searchLower) || 
                                                 (part && part.name.toLowerCase().includes(searchLower)) ||
                                                 (item.originalPartId && item.originalPartId.toLowerCase().includes(searchLower)) ||
                                                 (originalPart && originalPart.name.toLowerCase().includes(searchLower));
                            return matchesSearch;
                          })
                          .map((item, idx) => {
                            const partStr = item.partId;
                            const part = parts.find(p => p.id === partStr || p.name === partStr);
                            const originalPart = item.originalPartId ? parts.find(p => p.id === item.originalPartId || p.name === item.originalPartId) : null;
                            const qty = item.quantity;
                            const displayQty = (part?.level === 3 || originalPart?.level === 3) ? qty.toFixed(4) : qty;
                            const displayUnit = part?.unit || originalPart?.unit || 'Bộ';
                            const ageBatches = getInventoryBatches(item, transactions);
                            const ageInfo = ageBatches.length > 0 ? ageBatches[0].ageInfo : null;
                            const itemKey = `IN-${item.partId}-${item.originalPartId}`;
                            const isExpanded = expandedInventoryKeys.has(itemKey);

                            return (
                              <React.Fragment key={`${item.partId}-${item.originalPartId}-${idx}`}>
                                <tr 
                                  className="hover:bg-white transition-colors cursor-pointer"
                                  onClick={() => {
                                    const next = new Set(expandedInventoryKeys);
                                    if (next.has(itemKey)) next.delete(itemKey);
                                    else next.add(itemKey);
                                    setExpandedInventoryKeys(next);
                                  }}
                                >
                                  <td className="p-4">
                                    <div className="font-bold text-base">{getProcessValue(part?.name || item.partId, part, selectedStageDetail, 'IN')}</div>
                                    <div className="text-xs opacity-50 flex flex-wrap gap-2 items-center">
                                      {getProcessValue(part?.id || item.partId, part, selectedStageDetail, 'IN')}
                                      {item.originalPartId && (
                                        <span className="px-1.5 py-0.5 bg-blue-50 text-blue-600 rounded text-[10px] font-medium border border-blue-100">
                                          Nguồn: {originalPart?.name || item.originalPartId} ({originalPart?.id || item.originalPartId})
                                        </span>
                                      )}
                                    </div>
                                    {ageInfo && (
                                      <div className={`text-xs mt-1 font-bold ${ageInfo.colorClass} flex items-center gap-1`}>
                                        {ageInfo.text}
                                        {ageBatches.length > 1 && <span className="text-[10px] text-gray-500 bg-gray-100 px-1 rounded ml-1">+{ageBatches.length - 1} lần nhập</span>}
                                      </div>
                                    )}
                                  </td>
                                  <td className="p-4 text-right">
                                  <div className="flex items-center justify-end gap-3">
                                    <span className="font-mono font-bold text-xl">{displayQty}</span>
                                    <span className="text-xs font-mono opacity-40 uppercase mr-4">{displayUnit}</span>
                                    <button 
                                      onClick={() => setDefectModal({ partId: item.partId, stageId: selectedStageDetail, location: 'IN' })}
                                      className="p-2 hover:bg-red-50 text-red-600 rounded-lg transition-colors"
                                      title="Báo lỗi (NG)"
                                    >
                                      <AlertCircle size={16} />
                                    </button>
                                    <button 
                                      onClick={async () => {
                                        const pwd = await customPrompt('Nhập mật khẩu để sửa tồn kho:', '', true);
                                        if (pwd === 'admin123') {
                                          const newQty = await customPrompt(`Nhập số lượng tồn mới cho ${item.partId}:`, String(qty));
                                          if (newQty !== null) {
                                            storageService.setInventoryQuantity(item.partId, selectedStageDetail, 'IN', parseFloat(newQty) || 0, item.originalPartId);
                                            refreshData();
                                          }
                                        } else if (pwd !== null) {
                                          await customAlert('Mật khẩu không chính xác!');
                                        }
                                      }}
                                      className="p-2 hover:bg-blue-50 text-blue-600 rounded-lg transition-colors"
                                      title="Sửa số lượng"
                                    >
                                      <Edit2 size={16} />
                                    </button>
                                    <button 
                                      onClick={async () => {
                                        const pwd = await customPrompt('Nhập mật khẩu để xóa tồn kho:', '', true);
                                        if (pwd === 'admin123') {
                                          if (await customConfirm(`Bạn có chắc chắn muốn xóa tồn kho của ${item.partId} tại ${STAGES.find(s => s.id === selectedStageDetail)?.name} (Kho IN)?`)) {
                                            storageService.deleteInventoryItem(item.partId, selectedStageDetail, 'IN', item.originalPartId);
                                            refreshData();
                                          }
                                        } else if (pwd !== null) {
                                          await customAlert('Mật khẩu không chính xác!');
                                        }
                                      }}
                                      className="p-2 hover:bg-red-50 text-red-600 rounded-lg transition-colors"
                                      title="Xóa tồn kho"
                                    >
                                      <Trash2 size={16} />
                                    </button>
                                  </div>
                                </td>
                              </tr>
                              {isExpanded && ageBatches.length > 1 && (
                                <tr className="bg-gray-50/50">
                                  <td colSpan={2} className="px-4 py-3 border-t border-dashed border-gray-200">
                                    <div className="flex flex-col gap-1.5 pl-4 border-l-[3px] border-blue-300 ml-2">
                                      <div className="text-[10px] font-bold text-gray-500 uppercase tracking-wider mb-1">Chi tiết tuổi tồn kho</div>
                                      {ageBatches.map((batch, bIdx) => (
                                        <div key={bIdx} className="flex justify-between items-center text-xs">
                                          <span className="font-mono text-gray-700">SL: {(part?.level === 3 || originalPart?.level === 3) ? batch.quantity.toFixed(4) : batch.quantity} {displayUnit}</span>
                                          <span className={`font-bold ${batch.ageInfo.colorClass}`}>{batch.ageInfo.text}</span>
                                        </div>
                                      ))}
                                    </div>
                                  </td>
                                </tr>
                              )}
                            </React.Fragment>
                          );
                          })}
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* KHO_OUT Detail */}
                <div className="space-y-4">
                  <div className="flex items-center justify-between px-2">
                    <div className="flex items-center gap-2">
                      <div className="w-3 h-3 rounded-full bg-[#F27D26]" />
                      <h3 className="font-mono text-base font-bold uppercase tracking-widest opacity-80">
                        KHO_OUT (Đã hoàn thành)
                        {selectedStageDetail === 'LASER' && (
                          <span className="ml-2 text-xs font-normal normal-case text-gray-500 font-sans hidden sm:inline">
                            (Linh kiện thành phẩm)
                          </span>
                        )}
                      </h3>
                    </div>
                    {['LASER', 'BENDING', 'WELDING', 'PAINTING'].includes(selectedStageDetail!) && (
                      <button 
                        onClick={() => setShowManualAddModal({ stageId: selectedStageDetail as StageId, location: 'OUT' })}
                        className="text-xs bg-orange-100 text-orange-600 hover:bg-orange-200 px-3 py-1.5 rounded-lg font-bold flex items-center gap-1 transition-colors cursor-pointer"
                      >
                        <Plus size={14} /> Thêm tồn
                      </button>
                    )}
                  </div>
<div className="bg-gray-50 rounded-2xl border border-gray-100 overflow-hidden">
  <table className="w-full text-left border-collapse">
    <thead>
      <tr className="border-b border-gray-200 bg-gray-100/50">
        <th className="p-4 text-sm font-mono uppercase opacity-60">Linh kiện</th>
        <th className="p-4 text-sm font-mono uppercase opacity-60 text-right">Số lượng</th>
      </tr>
    </thead>
                      <tbody className="divide-y divide-gray-100">
                        {inventory
                          .filter(item => item.stageId === selectedStageDetail && item.location === 'OUT' && item.quantity > 0)
                          .filter(item => {
                            // Tách riêng phế liệu tôn khỏi kho OUT của Cắt Laser
                            if (selectedStageDetail === 'LASER') {
                              const partObj = parts.find(p => p.id === item.partId || p.name === item.partId);
                              if (isLaserScrapItem(item.partId, partObj?.name)) {
                                return false;
                              }
                            }
                            const partStr = item.partId;
                            const part = parts.find(p => p.id === partStr || p.name === partStr);
                            const originalPart = item.originalPartId ? parts.find(p => p.id === item.originalPartId || p.name === item.originalPartId) : null;
                            const searchLower = searchTerm.toLowerCase();
                            const matchesSearch = item.partId.toLowerCase().includes(searchLower) || 
                                                 (part && part.name.toLowerCase().includes(searchLower)) ||
                                                 (item.originalPartId && item.originalPartId.toLowerCase().includes(searchLower)) ||
                                                 (originalPart && originalPart.name.toLowerCase().includes(searchLower));
                            return matchesSearch;
                          })
                          .map((item, idx) => {
                            const partStr = item.partId;
                            const part = parts.find(p => p.id === partStr || p.name === partStr);
                            const originalPart = item.originalPartId ? parts.find(p => p.id === item.originalPartId || p.name === item.originalPartId) : null;
                            const qty = item.quantity;
                            const displayQty = (part?.level === 3 || originalPart?.level === 3) ? qty.toFixed(4) : qty;
                            const displayUnit = part?.unit || originalPart?.unit || 'Bộ';
                            const ageBatches = getInventoryBatches(item, transactions);
                            const ageInfo = ageBatches.length > 0 ? ageBatches[0].ageInfo : null;
                            const itemKey = `OUT-${item.partId}-${item.originalPartId}`;
                            const isExpanded = expandedInventoryKeys.has(itemKey);

                            return (
                              <React.Fragment key={`${item.partId}-${item.originalPartId}-${idx}`}>
                                <tr 
                                  className="hover:bg-white transition-colors cursor-pointer"
                                  onClick={() => {
                                    const next = new Set(expandedInventoryKeys);
                                    if (next.has(itemKey)) next.delete(itemKey);
                                    else next.add(itemKey);
                                    setExpandedInventoryKeys(next);
                                  }}
                                >
                                  <td className="p-4">
                                    <div className="font-bold text-base">{getProcessValue(part?.name || item.partId, part, selectedStageDetail, 'OUT')}</div>
                                    <div className="text-xs opacity-50 flex flex-wrap gap-2 items-center">
                                      {getProcessValue(part?.id || item.partId, part, selectedStageDetail, 'OUT')}
                                      {item.originalPartId && (
                                        <span className="px-1.5 py-0.5 bg-orange-50 text-orange-600 rounded text-[10px] font-medium border border-orange-100">
                                          Nguồn: {originalPart?.name || item.originalPartId} ({originalPart?.id || item.originalPartId})
                                        </span>
                                      )}
                                    </div>
                                    {ageInfo && (
                                      <div className={`text-xs mt-1 font-bold ${ageInfo.colorClass} flex items-center gap-1`}>
                                        {ageInfo.text}
                                        {ageBatches.length > 1 && <span className="text-[10px] text-gray-500 bg-gray-100 px-1 rounded ml-1">+{ageBatches.length - 1} lần nhập</span>}
                                      </div>
                                    )}
                                  </td>
                                <td className="p-4 text-right">
                                  <div className="flex items-center justify-end gap-3">
                                    <span className="font-mono font-bold text-xl text-[#F27D26]">{displayQty}</span>
                                    <span className="text-xs font-mono opacity-40 uppercase mr-4">{displayUnit}</span>
                                    <button 
                                      onClick={() => setDefectModal({ partId: item.partId, stageId: selectedStageDetail, location: 'OUT' })}
                                      className="p-2 hover:bg-red-50 text-red-600 rounded-lg transition-colors"
                                      title="Báo lỗi (NG)"
                                    >
                                      <AlertCircle size={16} />
                                    </button>
                                    <button 
                                      onClick={async () => {
                                        const pwd = await customPrompt('Nhập mật khẩu để sửa tồn kho:', '', true);
                                        if (pwd === 'admin123') {
                                          const newQty = await customPrompt(`Nhập số lượng tồn mới cho ${item.partId}:`, String(qty));
                                          if (newQty !== null) {
                                            storageService.setInventoryQuantity(item.partId, selectedStageDetail, 'OUT', parseFloat(newQty) || 0, item.originalPartId);
                                            refreshData();
                                          }
                                        } else if (pwd !== null) {
                                          await customAlert('Mật khẩu không chính xác!');
                                        }
                                      }}
                                      className="p-2 hover:bg-blue-50 text-blue-600 rounded-lg transition-colors"
                                      title="Sửa số lượng"
                                    >
                                      <Edit2 size={16} />
                                    </button>
                                    <button 
                                      onClick={async () => {
                                        const pwd = await customPrompt('Nhập mật khẩu để xóa tồn kho:', '', true);
                                        if (pwd === 'admin123') {
                                          if (await customConfirm(`Bạn có chắc chắn muốn xóa tồn kho của ${item.partId} tại ${STAGES.find(s => s.id === selectedStageDetail)?.name} (Kho OUT)?`)) {
                                            storageService.deleteInventoryItem(item.partId, selectedStageDetail, 'OUT', item.originalPartId);
                                            refreshData();
                                          }
                                        } else if (pwd !== null) {
                                          await customAlert('Mật khẩu không chính xác!');
                                        }
                                      }}
                                      className="p-2 hover:bg-red-50 text-red-600 rounded-lg transition-colors"
                                      title="Xóa tồn kho"
                                    >
                                      <Trash2 size={16} />
                                    </button>
                                  </div>
                                </td>
                              </tr>
                              {isExpanded && ageBatches.length > 1 && (
                                <tr className="bg-gray-50/50">
                                  <td colSpan={2} className="px-4 py-3 border-t border-dashed border-gray-200">
                                    <div className="flex flex-col gap-1.5 pl-4 border-l-[3px] border-[#F27D26] ml-2">
                                      <div className="text-[10px] font-bold text-gray-500 uppercase tracking-wider mb-1">Chi tiết tuổi tồn kho</div>
                                      {ageBatches.map((batch, bIdx) => (
                                        <div key={bIdx} className="flex justify-between items-center text-xs">
                                          <span className="font-mono text-gray-700">SL: {(part?.level === 3 || originalPart?.level === 3) ? batch.quantity.toFixed(4) : batch.quantity} {displayUnit}</span>
                                          <span className={`font-bold ${batch.ageInfo.colorClass}`}>{batch.ageInfo.text}</span>
                                        </div>
                                      ))}
                                    </div>
                                  </td>
                                </tr>
                              )}
                            </React.Fragment>
                          );
                          })}
                        {inventory
                          .filter(item => item.stageId === selectedStageDetail && item.location === 'OUT' && item.quantity > 0)
                          .filter(item => {
                            if (selectedStageDetail === 'LASER') {
                              const partObj = parts.find(p => p.id === item.partId || p.name === item.partId);
                              if (isLaserScrapItem(item.partId, partObj?.name)) return false;
                            }
                            const partStr = item.partId;
                            const part = parts.find(p => p.id === partStr || p.name === partStr);
                            const originalPart = item.originalPartId ? parts.find(p => p.id === item.originalPartId || p.name === item.originalPartId) : null;
                            const searchLower = searchTerm.toLowerCase();
                            return item.partId.toLowerCase().includes(searchLower) || 
                                   (part && part.name.toLowerCase().includes(searchLower)) ||
                                   (item.originalPartId && item.originalPartId.toLowerCase().includes(searchLower)) ||
                                   (originalPart && originalPart.name.toLowerCase().includes(searchLower));
                          }).length === 0 && (
                            <tr>
                              <td colSpan={2} className="p-8 text-center text-xs font-mono text-gray-400 italic">
                                Không có linh kiện tồn trong kho OUT
                              </td>
                            </tr>
                          )}
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* KHO_NG Detail */}
                <div className="space-y-4">
                  <div className="flex items-center gap-2 px-2">
                    <div className="w-3 h-3 rounded-full bg-red-500" />
                    <h3 className="font-mono text-base font-bold uppercase tracking-widest opacity-80 text-gray-900">KHO_NG (Hàng lỗi)</h3>
                  </div>
                  <div className="bg-gray-50 rounded-2xl border border-gray-100 overflow-hidden">
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="border-b border-gray-200 bg-gray-100/50">
                          <th className="p-4 text-sm font-mono uppercase opacity-60 text-gray-900">Linh kiện</th>
                          <th className="p-4 text-sm font-mono uppercase opacity-60 text-right text-gray-900">Số lượng</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-100">
                        {inventory
                          .filter(item => item.stageId === selectedStageDetail && item.location === 'DEFECT' && item.quantity > 0)
                          .filter(item => {
                            const partStr = item.partId;
                            const part = parts.find(p => p.id === partStr || p.name === partStr);
                            const originalPart = item.originalPartId ? parts.find(p => p.id === item.originalPartId || p.name === item.originalPartId) : null;
                            const searchLower = searchTerm.toLowerCase();
                            const matchesSearch = item.partId.toLowerCase().includes(searchLower) || 
                                                 (part && part.name.toLowerCase().includes(searchLower)) ||
                                                 (item.originalPartId && item.originalPartId.toLowerCase().includes(searchLower)) ||
                                                 (originalPart && originalPart.name.toLowerCase().includes(searchLower));
                            return matchesSearch;
                          })
                          .map((item, idx) => {
                            const partStr = item.partId;
                            const part = parts.find(p => p.id === partStr || p.name === partStr);
                            const originalPart = item.originalPartId ? parts.find(p => p.id === item.originalPartId || p.name === item.originalPartId) : null;
                            const qty = item.quantity;
                            const displayQty = (part?.level === 3 || originalPart?.level === 3) ? qty.toFixed(4) : qty;
                            const displayUnit = part?.unit || originalPart?.unit || 'Bộ';
                            const ageBatches = getInventoryBatches(item, transactions);
                            const ageInfo = ageBatches.length > 0 ? ageBatches[0].ageInfo : null;
                            const itemKey = `DEFECT-${item.partId}-${item.originalPartId}`;
                            const isExpanded = expandedInventoryKeys.has(itemKey);

                            return (
                              <React.Fragment key={`${item.partId}-${item.originalPartId}-${idx}`}>
                                <tr 
                                  className="hover:bg-white transition-colors cursor-pointer"
                                  onClick={() => {
                                    const next = new Set(expandedInventoryKeys);
                                    if (next.has(itemKey)) next.delete(itemKey);
                                    else next.add(itemKey);
                                    setExpandedInventoryKeys(next);
                                  }}
                                >
                                  <td className="p-4">
                                    <div className="font-bold text-base text-gray-900">{getProcessValue(part?.name || item.partId, part, selectedStageDetail || STAGES[0].id, 'OUT')}</div>
                                    <div className="text-xs opacity-50 text-gray-900 flex flex-wrap gap-2 items-center">
                                      {getProcessValue(part?.id || item.partId, part, selectedStageDetail || STAGES[0].id, 'OUT')}
                                      {item.originalPartId && (
                                        <span className="px-1.5 py-0.5 bg-red-100 text-red-600 rounded text-[10px] font-medium border border-red-200">
                                          Nguồn: {originalPart?.name || item.originalPartId} ({originalPart?.id || item.originalPartId})
                                        </span>
                                      )}
                                    </div>
                                    {ageInfo && (
                                      <div className={`text-xs mt-1 font-bold ${ageInfo.colorClass} flex items-center gap-1`}>
                                        {ageInfo.text}
                                        {ageBatches.length > 1 && <span className="text-[10px] text-gray-500 bg-gray-100 px-1 rounded ml-1">+{ageBatches.length - 1} lần nhập</span>}
                                      </div>
                                    )}
                                  </td>
                                <td className="p-4 text-right">
                                  <div className="flex items-center justify-end gap-3">
                                    <span className="font-mono font-bold text-xl text-red-600">{displayQty}</span>
                                    <span className="text-xs font-mono opacity-40 uppercase mr-4 text-gray-900">{displayUnit}</span>
                                    <button 
                                      onClick={async () => {
                                        const pwd = await customPrompt('Nhập mật khẩu để sửa tồn kho lỗi:', '', true);
                                        if (pwd === 'admin123') {
                                          const newQty = await customPrompt(`Nhập số lượng tồn mới cho ${item.partId} (Lỗi):`, String(qty));
                                          if (newQty !== null) {
                                            storageService.setInventoryQuantity(item.partId, selectedStageDetail || STAGES[0].id, 'DEFECT', parseFloat(newQty) || 0, item.originalPartId);
                                            refreshData();
                                          }
                                        } else if (pwd !== null) {
                                          await customAlert('Mật khẩu không chính xác!');
                                        }
                                      }}
                                      className="p-2 hover:bg-red-50 text-red-600 rounded-lg transition-colors"
                                      title="Sửa tồn lỗi"
                                    >
                                      <Edit2 size={16} />
                                    </button>
                                  </div>
                                </td>
                              </tr>
                              {isExpanded && ageBatches.length > 1 && (
                                <tr className="bg-gray-50/50">
                                  <td colSpan={2} className="px-4 py-3 border-t border-dashed border-gray-200">
                                    <div className="flex flex-col gap-1.5 pl-4 border-l-[3px] border-red-300 ml-2">
                                      <div className="text-[10px] font-bold text-gray-500 uppercase tracking-wider mb-1">Chi tiết tuổi tồn kho</div>
                                      {ageBatches.map((batch, bIdx) => (
                                        <div key={bIdx} className="flex justify-between items-center text-xs">
                                          <span className="font-mono text-gray-700">SL: {(part?.level === 3 || originalPart?.level === 3) ? batch.quantity.toFixed(4) : batch.quantity} {displayUnit}</span>
                                          <span className={`font-bold ${batch.ageInfo.colorClass}`}>{batch.ageInfo.text}</span>
                                        </div>
                                      ))}
                                    </div>
                                  </td>
                                </tr>
                              )}
                            </React.Fragment>
                          );
                          })}
                        {inventory.filter(item => item.stageId === selectedStageDetail && item.location === 'DEFECT' && item.quantity > 0).length === 0 && (
                          <tr>
                            <td colSpan={2} className="p-8 text-center text-gray-400 italic text-sm">Không có hàng lỗi tại công đoạn này.</td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            </div>
          </motion.div>
        ) : null}
      </AnimatePresence>

      <div className="bg-white p-3.5 sm:p-6 md:p-8 rounded-xl sm:rounded-2xl border border-gray-200 shadow-2xs w-full">
        {/* Chart Header & Tab Switcher */}
        <div className="flex flex-col lg:flex-row justify-between items-stretch lg:items-center mb-3 sm:mb-6 gap-3 sm:gap-4 pb-3 sm:pb-5 border-b border-gray-100">
          <div className="flex items-center gap-1.5 p-1 bg-gray-100/90 rounded-xl border border-gray-200 max-w-full overflow-x-auto no-scrollbar shrink-0">
            <button
              onClick={() => setActiveChartTab('MECHANICAL')}
              className={`px-2.5 sm:px-3.5 py-2 sm:py-2.5 rounded-lg text-xs md:text-sm font-bold uppercase tracking-wider transition-all whitespace-nowrap flex items-center gap-1.5 cursor-pointer active:scale-95 ${
                activeChartTab === 'MECHANICAL'
                  ? 'bg-white text-indigo-700 shadow-sm border border-gray-200 font-black'
                  : 'text-gray-500 hover:text-indigo-600'
              }`}
            >
              <span className={`w-2.5 h-2.5 rounded-full ${activeChartTab === 'MECHANICAL' ? 'bg-indigo-600' : 'bg-gray-300'}`} />
              <span className="hidden sm:inline">Năng suất Xưởng Cơ Khí</span>
              <span className="sm:hidden">NSLĐ Cơ Khí</span>
            </button>
            
            {([
              { id: 'LASER', name: 'Cắt Laser', color: '#3B82F6', activeText: 'text-blue-700' },
              { id: 'BENDING', name: 'Chấn / Dập', color: '#9333EA', activeText: 'text-purple-700' },
              { id: 'WELDING', name: 'Hàn', color: '#10B981', activeText: 'text-emerald-700' },
              { id: 'PAINTING', name: 'Sơn', color: '#EAB308', activeText: 'text-amber-700' },
              { id: 'GLAZING', name: 'Dán Kính', color: '#EF4444', activeText: 'text-red-700' },
            ] as const).map(tab => (
              <button
                key={tab.id}
                onClick={() => setActiveChartTab(tab.id)}
                className={`px-2.5 sm:px-3.5 py-2 sm:py-2.5 rounded-lg text-xs md:text-sm font-bold uppercase tracking-wider transition-all whitespace-nowrap flex items-center gap-1.5 cursor-pointer active:scale-95 ${
                  activeChartTab === tab.id
                    ? `bg-white ${tab.activeText} shadow-sm border border-gray-200 font-black`
                    : 'text-gray-500 hover:text-gray-900'
                }`}
              >
                <span className="w-2.5 h-2.5 rounded-full" style={{ backgroundColor: activeChartTab === tab.id ? tab.color : '#D1D5DB' }} />
                <span>{tab.name}</span>
              </button>
            ))}
          </div>

          <div className="flex items-center gap-2 sm:gap-4 flex-wrap justify-between sm:justify-end">
            <button
              onClick={() => setShowExportHourlyModal(true)}
              className="flex-1 sm:flex-none flex items-center justify-center gap-1.5 px-3 py-2 bg-emerald-50 text-emerald-700 hover:bg-emerald-100 rounded-xl transition-all text-xs font-bold uppercase border border-emerald-200 shadow-2xs cursor-pointer active:scale-95"
            >
              <FileSpreadsheet size={15} />
              <span>XUẤT BÁO CÁO GIỜ</span>
            </button>
            <div className="flex-1 sm:flex-none flex items-center justify-center gap-2 bg-gray-50 border border-gray-200 rounded-xl px-3 py-1.5 sm:py-2 relative group">
              <span className="text-[11px] sm:text-xs font-bold uppercase text-gray-500">Ngày:</span>
              <input 
                type="date"
                min={getPastDateStr(35)}
                max={getPastDateStr(0)}
                value={chartDate}
                onChange={(e) => setChartDate(e.target.value)}
                className="bg-transparent border-none outline-none font-bold text-indigo-600 text-xs cursor-pointer"
              />
              <div className="absolute -top-9 left-1/2 -translate-x-1/2 whitespace-nowrap bg-gray-800 text-white text-[10px] font-bold py-1 px-2.5 rounded-lg opacity-0 group-hover:opacity-100 transition-opacity pointer-events-none z-10">
                Chỉ xem báo cáo 35 ngày gần nhất
              </div>
            </div>
          </div>
        </div>

        {/* Sub-header: Shift Switcher (Ca Ngày / Ca Đêm) */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 mb-3 sm:mb-4 bg-gray-50/90 border border-gray-200 p-2 sm:p-2.5 rounded-xl">
          <div className="flex items-center gap-2 flex-wrap">
            <span className="text-xs font-bold uppercase text-gray-500 pl-1">Ca làm việc:</span>
            <div className="flex items-center p-1 bg-white rounded-lg border border-gray-200 shadow-2xs">
              <button
                onClick={() => setActiveChartShift('DAY')}
                className={`px-3 py-1.5 rounded-md text-xs font-black uppercase tracking-wider transition-all flex items-center gap-1.5 cursor-pointer active:scale-95 ${
                  activeChartShift === 'DAY'
                    ? 'bg-amber-500 text-white shadow-xs font-black'
                    : 'text-gray-600 hover:text-amber-800 font-bold'
                }`}
              >
                <span>☀️</span>
                <span>Ca Ngày (6h - 19h)</span>
              </button>
              <button
                onClick={() => setActiveChartShift('NIGHT')}
                className={`px-3 py-1.5 rounded-md text-xs font-black uppercase tracking-wider transition-all flex items-center gap-1.5 cursor-pointer active:scale-95 ${
                  activeChartShift === 'NIGHT'
                    ? 'bg-indigo-900 text-white shadow-xs font-black'
                    : 'text-gray-600 hover:text-indigo-950 font-bold'
                }`}
              >
                <span>🌙</span>
                <span>Ca Đêm (19h - 6h)</span>
              </button>
            </div>
          </div>
          <div className="text-[11px] text-gray-500 font-medium pr-1">
            {activeChartShift === 'DAY' 
              ? 'Khung giờ: 6h-8h • 8h-10h • 10h-12h • 13h-15h • 15h-17h • 17h-19h'
              : 'Khung giờ: 19h-21h • 21h-23h • 23h-0h • 1h-3h • 3h-5h • 5h-6h'}
          </div>
        </div>

        {/* Stage KPI Summary Banner */}
        {activeChartTab === 'MECHANICAL' ? (
          <div className="bg-indigo-50/70 border border-indigo-100 rounded-xl p-3 sm:p-4 mb-3 sm:mb-4">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-indigo-100/80 pb-2 mb-2.5">
              <div className="flex items-center gap-2 flex-wrap">
                <div className="w-3 h-3 rounded-full bg-[#4F46E5]" />
                <span className="font-extrabold uppercase text-xs sm:text-sm text-indigo-950">
                  Biểu đồ Năng suất Lao động Toàn Xưởng Cơ Khí - {activeChartShift === 'DAY' ? 'Ca Ngày' : 'Ca Đêm'}
                </span>
                <span className="text-[11px] font-bold text-indigo-700 bg-white px-2 py-0.5 rounded border border-indigo-200">Định mức chuẩn: 28.3 SPQĐ / công 8h</span>
              </div>
              <div className="text-[11px] text-gray-500">
                Hệ số quy đổi: <span className="font-bold text-gray-700">Sơn (DCLR) + Dán Kính 1 (Xưởng)</span>
              </div>
            </div>
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 sm:gap-3 text-xs">
              <div className="bg-white rounded-lg p-2.5 border border-indigo-100 shadow-2xs">
                <div className="text-gray-400 text-[10px] font-bold uppercase">Tổng SPQĐ thực tế</div>
                <div className="font-mono font-black text-sm sm:text-base text-indigo-900">{stageDaySummary?.totalConv} SPQĐ</div>
              </div>
              <div className="bg-white rounded-lg p-2.5 border border-indigo-100 shadow-2xs">
                <div className="text-gray-400 text-[10px] font-bold uppercase">Tổng số công làm việc</div>
                <div className="font-mono font-black text-sm sm:text-base text-gray-800">{stageDaySummary?.mandays} công</div>
              </div>
              <div className="bg-white rounded-lg p-2.5 border border-indigo-100 shadow-2xs">
                <div className="text-gray-400 text-[10px] font-bold uppercase">Mục tiêu quy đổi theo công</div>
                <div className="font-mono font-black text-sm sm:text-base text-gray-800">{stageDaySummary?.totalTarget} SPQĐ</div>
              </div>
              <div className="bg-white rounded-lg p-2.5 border border-indigo-100 shadow-2xs">
                <div className="text-gray-400 text-[10px] font-bold uppercase">NSLĐ trung bình {activeChartShift === 'DAY' ? 'ca ngày' : 'ca đêm'}</div>
                <div className={`font-mono font-black text-sm sm:text-base ${Number(stageDaySummary?.avgNsld || 0) >= 100 ? 'text-emerald-600' : Number(stageDaySummary?.avgNsld || 0) >= 80 ? 'text-amber-600' : 'text-red-600'}`}>
                  {stageDaySummary?.avgNsld}%
                </div>
              </div>
            </div>
          </div>
        ) : (
          activeStageConfig && (
            <div className="rounded-xl p-3 sm:p-4 mb-3 sm:mb-4 border bg-gray-50/70 border-gray-200">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-gray-200 pb-2 mb-2.5">
                <div className="flex items-center gap-2 flex-wrap">
                  <div className="w-3 h-3 rounded-full" style={{ backgroundColor: activeStageConfig.color }} />
                  <span className="font-extrabold uppercase text-xs sm:text-sm text-gray-900">
                    Năng suất Công đoạn {activeStageConfig.name} - {activeChartShift === 'DAY' ? 'Ca Ngày' : 'Ca Đêm'}
                  </span>
                  <span className="text-[11px] font-bold px-2 py-0.5 rounded bg-white border border-gray-200 text-gray-800">
                    Định mức chuẩn: <span className="font-black" style={{ color: activeStageConfig.lineColor }}>{activeStageConfig.normText}</span>
                  </span>
                </div>
                <div className="text-[11px] text-gray-500 font-medium">
                  Hệ số quy đổi: <span className="font-bold text-gray-700">{activeStageConfig.hsqdSource}</span>
                </div>
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 sm:gap-3 text-xs">
                <div className="bg-white rounded-lg p-2.5 border border-gray-200 shadow-2xs">
                  <div className="text-gray-400 text-[10px] font-bold uppercase">Sản lượng thực tế</div>
                  <div className="font-mono font-black text-sm sm:text-base text-gray-800">{stageDaySummary?.totalActual} cái</div>
                </div>
                <div className="bg-white rounded-lg p-2.5 border border-gray-200 shadow-2xs">
                  <div className="text-gray-400 text-[10px] font-bold uppercase">Sản phẩm quy đổi</div>
                  <div className="font-mono font-black text-sm sm:text-base" style={{ color: activeStageConfig.lineColor }}>{stageDaySummary?.totalConv} SPQĐ</div>
                </div>
                <div className="bg-white rounded-lg p-2.5 border border-gray-200 shadow-2xs">
                  <div className="text-gray-400 text-[10px] font-bold uppercase">Mục tiêu theo nhân sự</div>
                  <div className="font-mono font-black text-sm sm:text-base text-gray-800">{stageDaySummary?.totalTarget} SPQĐ</div>
                </div>
                <div className="bg-white rounded-lg p-2.5 border border-gray-200 shadow-2xs">
                  <div className="text-gray-400 text-[10px] font-bold uppercase">NSLĐ trung bình {activeChartShift === 'DAY' ? 'ca ngày' : 'ca đêm'}</div>
                  <div className={`font-mono font-black text-sm sm:text-base ${Number(stageDaySummary?.avgNsld || 0) >= 100 ? 'text-emerald-600' : Number(stageDaySummary?.avgNsld || 0) >= 80 ? 'text-amber-600' : 'text-red-600'}`}>
                    {stageDaySummary?.avgNsld}%
                  </div>
                </div>
              </div>
            </div>
          )
        )}

        {/* Legend area */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 sm:gap-4 mb-3 sm:mb-4 flex-wrap text-xs">
          {activeChartTab === 'MECHANICAL' ? (
            <div className="flex items-center gap-3 sm:gap-6 flex-wrap text-xs font-bold">
              <div className="flex items-center gap-3 flex-wrap">
                <span className="text-gray-400 uppercase text-[10px] sm:text-[11px]">Sản lượng thực tế:</span>
                <div className="flex items-center gap-1">
                  <div className="w-3 h-3 rounded bg-[#E5A84B]" />
                  <span className="text-gray-700 uppercase text-[11px] sm:text-xs">Sơn</span>
                </div>
                <div className="flex items-center gap-1">
                  <div className="w-3 h-3 rounded bg-[#E06A6A]" />
                  <span className="text-gray-700 uppercase text-[11px] sm:text-xs">Dán Kính</span>
                </div>
              </div>

              <div className="flex items-center gap-2 bg-indigo-50/80 border border-indigo-200 px-2.5 sm:px-3 py-1 rounded-xl">
                <div className="w-3.5 h-1.5 rounded bg-[#4F46E5]" />
                <span className="text-[11px] sm:text-xs font-black uppercase text-indigo-700 tracking-wider">NSLĐ XƯỞNG CƠ KHÍ</span>
              </div>
            </div>
          ) : (
            activeStageConfig && (
              <div className="flex items-center gap-4 sm:gap-8 flex-wrap text-xs font-bold">
                <div className="flex items-center gap-2">
                  <div className="w-3.5 h-3 rounded" style={{ backgroundColor: activeStageConfig.color }} />
                  <span className="text-gray-800 uppercase text-[11px] sm:text-xs">Sản lượng thực tế: {activeStageConfig.name}</span>
                </div>
                <div className="flex items-center gap-2 px-2.5 py-1 rounded-lg border bg-white" style={{ borderColor: activeStageConfig.color }}>
                  <div className="w-4 h-1.5 rounded" style={{ backgroundColor: activeStageConfig.lineColor }} />
                  <span className="text-[11px] sm:text-xs font-black uppercase tracking-wider" style={{ color: activeStageConfig.lineColor }}>
                    % NSLĐ
                  </span>
                </div>
                <div className="flex items-center gap-2 text-gray-500 font-medium text-[11px]">
                  <div className="w-4 h-0.5 border-t-2 border-dashed border-gray-600" />
                  <span>Đường chuẩn 100% Mục tiêu</span>
                </div>
              </div>
            )
          )}
        </div>

        {/* Chart View with horizontal scroll on mobile to view all 12 hours */}
        <div className="w-full">
          <div className="flex items-center justify-between sm:hidden text-[10px] text-gray-400 font-medium mb-1.5 px-0.5">
            <span>📊 Biểu đồ chi tiết theo giờ</span>
            <span>↔ Vuốt ngang xem đủ 12 khung giờ</span>
          </div>
          <div className="w-full overflow-x-auto no-scrollbar rounded-xl">
            <div 
              ref={chartWrapperRef}
              className="w-full h-72 min-h-[300px] sm:h-[450px] min-w-[580px] sm:min-w-full relative" 
              style={{ minHeight: '300px', width: '100%' }}
            >
              {chartDimensions && chartDimensions.width > 0 && chartDimensions.height > 0 ? (
                <ResponsiveContainer 
                  width="100%" 
                  height={chartDimensions.height > 250 ? chartDimensions.height : 420} 
                  minHeight={300} 
                  minWidth={0} 
                  debounce={50}
                >
                  <ComposedChart data={productionChartData} margin={{ top: 25, right: 25, left: -15, bottom: 15 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#E5E7EB" />
              <XAxis 
                dataKey="name" 
                axisLine={false} 
                tickLine={false} 
                tick={{ fontSize: 12, fontWeight: 700, fill: '#374151' }} 
                dy={10}
              />
              <YAxis 
                yAxisId="left"
                axisLine={false} 
                tickLine={false} 
                tick={{ fontSize: 12, fontFamily: 'monospace', fontWeight: 700, fill: '#374151' }} 
                dx={-10}
              />
              <YAxis 
                yAxisId="right"
                orientation="right"
                domain={[0, (dataMax: number) => Math.max(Math.ceil((dataMax || 0) * 1.25 * 10) / 10, 1.2)]}
                axisLine={false} 
                tickLine={false} 
                tick={{ fontSize: 12, fontFamily: 'monospace', fontWeight: 700, fill: '#374151' }} 
                tickFormatter={(value) => `${Math.round(value * 100)}%`}
                dx={10}
              />
              <Tooltip 
                cursor={{ fill: '#F9FAFB' }}
                content={({ active, payload, label }: any) => {
                  if (active && payload && payload.length) {
                    const data = payload[0].payload;
                    return (
                      <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xl text-sm text-gray-800 min-w-[280px] max-w-sm">
                        <h4 className="font-bold border-b border-gray-100 pb-2 mb-3 text-lg text-gray-900 flex justify-between items-center">
                          <span>{label}</span>
                          <span className="text-xs font-normal text-gray-400">{chartDate}</span>
                        </h4>

                        {/* Mechanical Workshop highlight in Mechanical Tab */}
                        {activeChartTab === 'MECHANICAL' && (
                          <div className="bg-indigo-50/70 border border-indigo-100 rounded-xl p-3 mb-3">
                            <div className="flex items-center justify-between mb-2">
                              <div className="flex items-center gap-2">
                                <div className="w-3.5 h-3.5 rounded bg-[#4F46E5]" />
                                <span className="font-extrabold uppercase text-xs text-indigo-950">NSLĐ Xưởng Cơ Khí</span>
                              </div>
                              <span className="font-mono font-black text-sm px-2 py-0.5 rounded bg-white border border-indigo-200 text-indigo-700">
                                {Math.round((data.NSLD_CoKhi || 0) * 100)}%
                              </span>
                            </div>
                            <div className="text-[11px] text-gray-700 space-y-1">
                              <div className="flex justify-between">
                                <span className="text-gray-500">• Tổng SPQĐ (Sơn + Dán kính):</span>
                                <span className="font-mono font-bold">{(data.totalConvertedMechanical || 0).toFixed(1)} SPQĐ</span>
                              </div>
                              <div className="flex justify-between">
                                <span className="text-gray-500">• Tổng nhân sự (5 công đoạn):</span>
                                <span className="font-mono font-bold">{data.totalWorkersMechanical || 0} người</span>
                              </div>
                              <div className="text-[10px] text-gray-500 pl-2">
                                (Laser: {data.workerCountLaser} | Chấn: {data.workerCountChanDap} | Hàn: {data.workerCountHan} | Sơn: {data.workerCountSon} | Dán: {data.workerCountDanKinh})
                              </div>
                              <div className="flex justify-between items-baseline gap-1">
                                <span className="text-gray-500 whitespace-nowrap">• Công thức NSLĐ:</span>
                                <span className="font-mono font-bold text-indigo-700 text-right">SPQĐ / ({data.totalWorkersMechanical || 0} × {data.hourFactor === 0.125 ? '0.125' : '0.25'}) / 28.3</span>
                              </div>
                            </div>
                          </div>
                        )}

                        {/* Stage Specific Highlight Box */}
                        {activeStageConfig && (
                          <div className="bg-gray-50 border border-gray-200 rounded-xl p-3 mb-3">
                            <div className="flex items-center justify-between mb-2">
                              <div className="flex items-center gap-2">
                                <div className="w-3.5 h-3.5 rounded" style={{ backgroundColor: activeStageConfig.color }} />
                                <span className="font-extrabold uppercase text-xs text-gray-900">NSLĐ {activeStageConfig.name}</span>
                              </div>
                              <span className="font-mono font-black text-sm px-2 py-0.5 rounded bg-white border border-gray-200" style={{ color: activeStageConfig.lineColor }}>
                                {Math.round((data[activeStageConfig.lineKey] || 0) * 100)}%
                              </span>
                            </div>
                            <div className="text-[11px] text-gray-700 space-y-1">
                              <div className="flex justify-between">
                                <span className="text-gray-500">• Sản phẩm quy đổi đạt:</span>
                                <span className="font-mono font-bold" style={{ color: activeStageConfig.lineColor }}>{(data[activeStageConfig.convertedKey] || 0).toFixed(1)} SPQĐ</span>
                              </div>
                              <div className="flex justify-between">
                                <span className="text-gray-500">• Mục tiêu khung giờ:</span>
                                <span className="font-mono font-bold text-gray-900">{(data[activeStageConfig.targetKey] || 0).toFixed(1)} SPQĐ</span>
                              </div>
                              <div className="flex justify-between">
                                <span className="text-gray-500">• Nhân sự khung giờ:</span>
                                <span className="font-mono font-bold">{data[activeStageConfig.workerKey] || 0} người</span>
                              </div>
                              <div className="flex justify-between items-baseline gap-1">
                                <span className="text-gray-500 whitespace-nowrap">• Công thức:</span>
                                <span className="font-mono font-bold text-gray-700 text-right">
                                  SPQĐ / ({data[activeStageConfig.workerKey] || 0} × {activeStageConfig.normPerShift} × {data.hourFactor === 0.125 ? '0.125' : '0.25'})
                                </span>
                              </div>
                            </div>
                          </div>
                        )}

                        {/* Stage Produced Items Details */}
                        {(() => {
                          if (activeChartTab === 'MECHANICAL') {
                            return ['Sơn', 'Dán Kính'].map((stName, sIdx) => {
                              const items = data.details[stName] || [];
                              if (items.length === 0) return null;
                              return (
                                <div key={sIdx} className="mb-3 last:mb-0">
                                  <div className="font-bold flex items-center justify-between mb-1.5 text-xs text-gray-800">
                                    <span className="uppercase tracking-widest">{stName}</span>
                                    <span className="font-mono bg-gray-100 px-2 py-0.5 rounded text-xs">{data[stName]} Cái</span>
                                  </div>
                                  <div className="pl-2 space-y-1.5 max-h-48 overflow-y-auto no-scrollbar">
                                    {items.map((detail: any, idx: number) => {
                                      const mechHsqd = detail.hsqdMech !== undefined ? detail.hsqdMech : detail.hsqd;
                                      const hasHsqd = mechHsqd !== undefined && mechHsqd > 0;
                                      const convertedVal = hasHsqd ? ((detail.quantity || 0) * mechHsqd) : 0;
                                      const hsqdLabel = stName === 'Sơn' ? 'DCLR' : 'Xưởng CK';
                                      return (
                                        <div key={idx} className="flex justify-between items-start text-xs text-gray-700 gap-3 border-b border-gray-50 pb-1 last:border-0 last:pb-0">
                                          <div className="flex-1 min-w-0">
                                            <span className="break-words leading-tight font-medium text-[11px] text-gray-800 block">{detail.partName}</span>
                                            <span className="text-[10px] text-gray-400 font-mono block mt-0.5">
                                              HSQĐ ({hsqdLabel}): {hasHsqd ? <span className="font-bold text-gray-700">{mechHsqd}</span> : <span className="italic text-amber-600 font-medium">Trống (tính 0)</span>}
                                              {hasHsqd && <span className="text-gray-500"> → <span className="font-bold text-indigo-600">{convertedVal.toFixed(1)} SPQĐ</span></span>}
                                            </span>
                                          </div>
                                          <span className="font-mono font-bold text-gray-900 text-[11px] whitespace-nowrap">{detail.quantity} {detail.unit}</span>
                                        </div>
                                      );
                                    })}
                                  </div>
                                </div>
                              );
                            });
                          }

                          if (activeStageConfig) {
                            const items = data.details[activeStageConfig.stageName] || [];
                            if (items.length === 0) return (
                              <div className="text-xs text-gray-400 italic py-1 text-center">Không có sản lượng trong khung giờ này</div>
                            );
                            return (
                              <div className="mb-2">
                                <div className="font-bold flex items-center justify-between mb-1.5 text-xs text-gray-800 border-b border-gray-100 pb-1">
                                  <span className="uppercase tracking-widest">Chi tiết sản lượng thực tế</span>
                                  <span className="font-mono bg-gray-100 px-2 py-0.5 rounded text-xs">{data[activeStageConfig.barKey] || 0} Cái</span>
                                </div>
                                <div className="pl-1 space-y-1.5 max-h-52 overflow-y-auto no-scrollbar">
                                  {items.map((detail: any, idx: number) => {
                                    const hasHsqd = detail.hsqd !== undefined && detail.hsqd > 0;
                                    const convertedVal = hasHsqd ? ((detail.quantity || 0) * detail.hsqd) : 0;
                                    return (
                                      <div key={idx} className="flex justify-between items-start text-xs text-gray-700 gap-3 border-b border-gray-50 pb-1 last:border-0 last:pb-0">
                                        <div className="flex-1 min-w-0">
                                          <span className="break-words leading-tight font-medium text-[11px] text-gray-800 block">{detail.partName}</span>
                                          <span className="text-[10px] text-gray-400 font-mono block mt-0.5">
                                            HSQĐ: {hasHsqd ? <span className="font-bold text-gray-700">{detail.hsqd}</span> : <span className="italic text-amber-600 font-medium">Trống (tính 0)</span>}
                                            {hasHsqd && <span className="text-gray-500"> → <span className="font-bold text-indigo-600">{convertedVal.toFixed(1)} SPQĐ</span></span>}
                                          </span>
                                        </div>
                                        <span className="font-mono font-bold text-gray-900 text-[11px] whitespace-nowrap">{detail.quantity} {detail.unit}</span>
                                      </div>
                                    );
                                  })}
                                </div>
                              </div>
                            );
                          }
                          return null;
                        })()}
                      </div>
                    );
                  }
                  return null;
                }}
              />

              {/* In MECHANICAL tab: Show ONLY Painting & Glazing bars with soothing soft colors */}
              {activeChartTab === 'MECHANICAL' && (
                <>
                  <Bar yAxisId="left" dataKey="Sơn" fill="#E5A84B" radius={[4, 4, 0, 0]} />
                  <Bar yAxisId="left" dataKey="Dán Kính" fill="#E06A6A" radius={[4, 4, 0, 0]} />
                </>
              )}

              {/* In Specific Stage tab: Show ONLY this stage's actual production bar */}
              {activeStageConfig && (
                <Bar 
                  yAxisId="left" 
                  dataKey={activeStageConfig.barKey} 
                  fill={activeStageConfig.color} 
                  radius={[4, 4, 0, 0]} 
                />
              )}
              
              <ReferenceLine y={1} yAxisId="right" stroke="#4B5563" strokeDasharray="3 3" strokeWidth={1.5} label={{ position: 'insideTopLeft', value: '100% Target', fill: '#4B5563', fontSize: 11, fontWeight: 'bold' }} />

              {/* Stage Specific Line */}
              {activeStageConfig && (
                <Line 
                  yAxisId="right" 
                  type="monotone" 
                  dataKey={activeStageConfig.lineKey} 
                  stroke={activeStageConfig.lineColor} 
                  strokeWidth={3.5} 
                  dot={{ r: 4.5, fill: activeStageConfig.lineColor, strokeWidth: 2, stroke: '#fff' }} 
                  activeDot={{ r: 6.5 }} 
                  label={(props: any) => {
                    const { x, y, value } = props;
                    if (!value || value <= 0) return null;
                    const labelY = y < 35 ? y + 18 : y - 14;
                    return (
                      <g>
                        <rect x={x - 22} y={labelY - 11} width={44} height={18} fill={activeStageConfig.lineColor} rx={5} />
                        <text x={x} y={labelY + 2.5} fill="#ffffff" fontSize={11} fontWeight="800" fontFamily="monospace" textAnchor="middle">
                          {Math.round(value * 100)}%
                        </text>
                      </g>
                    );
                  }} 
                />
              )}

              {/* Mechanical Workshop Productivity Single Line (Tab 1) - Soothing Indigo theme */}
              {activeChartTab === 'MECHANICAL' && (
                <Line 
                  yAxisId="right" 
                  type="monotone" 
                  dataKey="NSLD_CoKhi" 
                  name="NSLĐ Xưởng Cơ Khí"
                  stroke="#4F46E5" 
                  strokeWidth={3} 
                  dot={{ r: 4.5, fill: '#4F46E5', strokeWidth: 2, stroke: '#fff' }} 
                  activeDot={{ r: 6.5 }} 
                  label={(props: any) => {
                    const { x, y, value } = props;
                    // Do not render cluttering badges for 0%
                    if (!value || value <= 0) return null;
                    const labelY = y < 35 ? y + 18 : y - 14;
                    const textVal = `${Math.round(value * 100)}%`;
                    return (
                      <g>
                        <rect 
                          x={x - 21} 
                          y={labelY - 11} 
                          width={42} 
                          height={18} 
                          fill="#4F46E5" 
                          rx={5} 
                        />
                        <text 
                          x={x} 
                          y={labelY + 2.5} 
                          fill="#ffffff" 
                          fontSize={11} 
                          fontWeight="800" 
                          fontFamily="monospace"
                          textAnchor="middle"
                        >
                          {textVal}
                        </text>
                      </g>
                    );
                  }} 
                />
              )}
            </ComposedChart>
          </ResponsiveContainer>
        ) : (
          <div className="w-full h-72 min-h-[300px] sm:h-[450px] flex items-center justify-center bg-gray-50/50 rounded-xl border border-dashed border-gray-200" style={{ minHeight: '300px', width: '100%' }}>
            <div className="flex flex-col items-center gap-2 text-gray-400 font-bold text-xs">
              <span className="w-4 h-4 rounded-full border-2 border-indigo-600 border-t-transparent animate-spin" />
              <span>Đang tính toán kích thước biểu đồ...</span>
            </div>
          </div>
        )}
        </div>
      </div>
    </div>
  </div>

      <AnimatePresence>
        {showManualAddModal && (
          <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center p-3 sm:p-4 z-[9999]">
            <motion.div 
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-2xl sm:rounded-3xl p-4 sm:p-8 max-w-md w-full shadow-2xl relative max-h-[90vh] overflow-y-auto no-scrollbar"
            >
              <button 
                className="absolute top-4 right-4 text-gray-500 hover:text-gray-800 cursor-pointer p-2 hover:bg-gray-100 rounded-full transition-colors" 
                onClick={() => {
                  setShowManualAddModal(null);
                  setManualAddPart('');
                  setManualAddPartSearch('');
                  setManualAddQty('');
                }}
              >
                <X size={20} />
              </button>
              <h2 className="text-2xl font-bold uppercase text-gray-900 mb-6 flex items-center gap-3">
                <PackagePlus className={showManualAddModal.location === 'IN' ? "text-blue-600" : "text-orange-600"} />
                Thêm tồn kho {showManualAddModal.location}
              </h2>
              <div className="space-y-6">
                <div>
                  <label className="block text-sm font-bold uppercase text-gray-600 mb-2">Linh kiện</label>
                  <input
                    list="manual-add-parts"
                    value={manualAddPartSearch}
                    onChange={(e) => {
                      setManualAddPartSearch(e.target.value);
                      const match = parts.find(p => `${p.id} - ${p.name}` === e.target.value || p.id === e.target.value);
                      if (match) setManualAddPart(match.id);
                      else setManualAddPart('');
                    }}
                    className="w-full bg-gray-50 border border-gray-200 p-4 rounded-xl font-bold text-gray-800 outline-none focus:border-blue-500 transition-all font-mono"
                    placeholder="Nhập mã hoặc tên linh kiện..."
                  />
                  <datalist id="manual-add-parts">
                    {parts.map(p => (
                      <option key={p.id} value={`${p.id} - ${p.name}`} />
                    ))}
                  </datalist>
                </div>
                <div>
                  <label className="block text-sm font-bold uppercase text-gray-600 mb-2">Số lượng</label>
                  <input 
                    type="number"
                    value={manualAddQty}
                    onChange={(e) => setManualAddQty(e.target.value)}
                    className="w-full bg-gray-50 border border-gray-200 p-4 rounded-xl font-bold text-gray-800 outline-none focus:border-blue-500 transition-all font-mono"
                    placeholder="Nhập số lượng..."
                  />
                </div>
                <div className="pt-2">
                  <button 
                    onClick={async () => {
                      if (!manualAddPart || !manualAddQty || Number(manualAddQty) <= 0) {
                        await customAlert('Vui lòng chọn linh kiện và nhập số lượng hợp lệ.');
                        return;
                      }
                      
                      const part = manualAddPart;
                      const qty = Number(manualAddQty);
                      const stageId = showManualAddModal.stageId;
                      const location = showManualAddModal.location;
                      
                      setShowManualAddModal(null);
                      setManualAddPart('');
                      setManualAddPartSearch('');
                      setManualAddQty('');

                      const pwd = await customPrompt('Nhập mật khẩu để cấu hình tồn kho:', '', true);
                      if (pwd !== 'admin123') {
                        if (pwd !== null) await customAlert('Mật khẩu không chính xác!');
                        return;
                      }
                      try {
                        storageService.recordManualInbound(part, stageId, location, qty);
                        await customAlert('Thêm tồn kho thành công!');
                        refreshData();
                      } catch (err: any) {
                        const errMsg = err.message || '';
                        if (errMsg.startsWith('OVER_PO:')) {
                          await customAlert(errMsg.replace('OVER_PO:', ''));
                        } else {
                          await customAlert(errMsg || 'Lỗi thêm tồn kho');
                        }
                      }
                    }}
                    className={cn("w-full py-4 text-white rounded-xl font-bold text-lg uppercase transition-all shadow-lg flex items-center justify-center gap-2", showManualAddModal.location === 'IN' ? 'bg-blue-600 hover:bg-blue-700' : 'bg-orange-600 hover:bg-orange-700')}
                  >
                    <PackagePlus size={24} />
                    Xác nhận thêm
                  </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {showExportModal && (
          <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center p-3 sm:p-4 z-[9999]">
            <motion.div 
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-2xl sm:rounded-3xl p-4 sm:p-8 max-w-xl w-full shadow-2xl relative max-h-[90vh] overflow-y-auto no-scrollbar"
            >
              <button className="absolute top-4 right-4 text-gray-500 hover:text-gray-800 cursor-pointer p-2 hover:bg-gray-100 rounded-full transition-colors" onClick={() => setShowExportModal(false)}>
                <X size={20} />
              </button>
              <h2 className="text-2xl font-bold uppercase text-gray-900 mb-6 flex items-center gap-3">
                <FileSpreadsheet className="text-green-600" />
                Chọn Quãng Ngày
              </h2>
              <div className="space-y-6">
                <div>
                  <label className="block text-sm font-bold uppercase text-gray-600 mb-2">Từ Ngày <span className="text-[10px] text-gray-400 normal-case">(Tối đa 35 ngày gần nhất)</span></label>
                  <input 
                    type="date"
                    min={getPastDateStr(35)}
                    max={getPastDateStr(0)}
                    value={exportStartDate}
                    onChange={(e) => setExportStartDate(e.target.value)}
                    className="w-full bg-gray-50 border border-gray-200 p-4 rounded-xl font-bold text-gray-800 outline-none focus:border-green-500 focus:bg-green-50 transition-all font-mono cursor-pointer"
                  />
                </div>
                <div>
                  <label className="block text-sm font-bold uppercase text-gray-600 mb-2">Đến Ngày <span className="text-[10px] text-gray-400 normal-case">(Đến hiện tại)</span></label>
                  <input 
                    type="date"
                    min={getPastDateStr(35)}
                    max={getPastDateStr(0)}
                    value={exportEndDate}
                    onChange={(e) => setExportEndDate(e.target.value)}
                    className="w-full bg-gray-50 border border-gray-200 p-4 rounded-xl font-bold text-gray-800 outline-none focus:border-green-500 focus:bg-green-50 transition-all font-mono cursor-pointer"
                  />
                </div>
                
                <div>
                  <label className="block text-sm font-bold uppercase text-gray-600 mb-2">Tuần <span className="text-[10px] text-gray-400 normal-case">(Hiển thị trên tiêu đề)</span></label>
                  <input 
                    type="text"
                    placeholder="VD: 24,25"
                    value={exportWeekName}
                    onChange={(e) => {
                      setExportWeekName(e.target.value);
                      storageService.saveExportWeekName(e.target.value);
                    }}
                    className="w-full bg-gray-50 border border-gray-200 p-3.5 rounded-xl font-bold text-gray-800 outline-none focus:border-green-500 focus:bg-green-50 transition-all"
                  />
                </div>
                
                <div>
                  <div className="flex items-center justify-between mb-2">
                    <label className="block text-sm font-bold uppercase text-gray-600">
                      Tổng số công đi làm <span className="text-[10px] text-gray-400 normal-case">(Tính % thực hiện tuần)</span>
                    </label>
                    <span className="text-[11px] font-bold text-green-700 bg-green-50 border border-green-200 px-2.5 py-0.5 rounded-full font-mono">
                      Tự động tính: {calculatedTotalMandays} công
                    </span>
                  </div>
                  <input 
                    type="number"
                    min="1"
                    placeholder="VD: 50"
                    value={exportTotalMandays}
                    onChange={(e) => setExportTotalMandays(e.target.value)}
                    className="w-full bg-gray-50 border border-gray-200 p-4 rounded-xl font-bold text-gray-800 outline-none focus:border-green-500 focus:bg-green-50 transition-all font-mono"
                  />
                </div>

                <div className="bg-gray-50 border border-gray-100 rounded-xl p-4 max-h-60 overflow-y-auto custom-scrollbar">
                  <label className="block text-sm font-bold uppercase text-gray-600 mb-4 sticky top-0 bg-gray-50 pb-2 border-b border-gray-200 z-10">Dữ liệu mỗi ngày</label>
                  <div className="grid grid-cols-2 md:grid-cols-3 gap-3">
                    {exportDateRange.map(dateStr => {
                      const [y, m, d] = dateStr.split('-');
                      const dayAutoAvg = storageService.getDailyTotalPeopleAverage(dateStr);
                      const dayAutoMandays = storageService.getDailyTotalMandays(dateStr);
                      const currentPeopleVal = peoplePerDay[dateStr] !== undefined 
                        ? peoplePerDay[dateStr] 
                        : (dayAutoAvg !== null ? dayAutoAvg : '');
                      const currentMandaysVal = mandaysPerDay[dateStr] !== undefined 
                        ? mandaysPerDay[dateStr] 
                        : (dayAutoMandays > 0 ? Math.round(dayAutoMandays * 10) / 10 : '');

                      return (
                      <div key={dateStr} className="flex flex-col gap-2 bg-white p-2 border border-gray-200 rounded-lg shadow-sm">
                        <span className="text-[10px] font-bold text-gray-500 uppercase text-center border-b border-gray-100 pb-1">
                          {d}/{m}/{y}
                        </span>
                        
                        <div className="flex gap-2">
                           <div className="flex-1">
                             <label className="block text-[9px] font-bold text-gray-400 mb-1 leading-none uppercase" title="Tự động tính từ trung bình cộng số người tất cả công đoạn của ngày đó">Người</label>
                             <input
                               type="number"
                               min="0"
                               title={`Số người ngày ${dateStr} (Tự động tính từ trung bình cộng các khung giờ)`}
                               placeholder={dayAutoAvg !== null ? String(dayAutoAvg) : "0"}
                               value={currentPeopleVal}
                               onChange={(e) => {
                                 const raw = e.target.value;
                                 const updated = { ...peoplePerDay };
                                 if (raw === '') {
                                   delete updated[dateStr];
                                 } else {
                                   updated[dateStr] = parseInt(raw) || 0;
                                 }
                                 setPeoplePerDay(updated);
                                 storageService.savePeoplePerDay(updated);
                               }}
                               className="w-full bg-gray-50 outline-none text-sm font-bold p-1.5 rounded focus:border-blue-500 border border-transparent hover:border-gray-300 transition-colors text-blue-700 text-center font-mono"
                             />
                           </div>
                           <div className="flex-1">
                             <label className="block text-[9px] font-bold text-gray-400 mb-1 leading-none uppercase" title="Tự động tính từ tổng số công các khung giờ có nhân sự">Công</label>
                             <input
                               type="number"
                               step="0.1"
                               min="0"
                               title={`Số công ngày ${dateStr} (Tự động tính từ các khung giờ có nhân sự)`}
                               placeholder={dayAutoMandays > 0 ? String(Math.round(dayAutoMandays * 10) / 10) : "0"}
                               value={currentMandaysVal}
                               onChange={(e) => {
                                 const raw = e.target.value;
                                 const updated = { ...mandaysPerDay };
                                 if (raw === '') {
                                   delete updated[dateStr];
                                 } else {
                                   updated[dateStr] = parseFloat(raw) || 0;
                                 }
                                 setMandaysPerDay(updated);
                                 storageService.saveMandaysPerDay(updated);
                               }}
                               className="w-full bg-gray-50 outline-none text-sm font-bold p-1.5 rounded focus:border-orange-500 border border-transparent hover:border-gray-300 transition-colors text-orange-700 text-center font-mono"
                             />
                           </div>
                        </div>
                      </div>
                    )})}
                  </div>
                </div>

                <div className="pt-2">
                  <button 
                    onClick={() => {
                      exportProductionReportRange(
                        transactions, 
                        parts, 
                        exportStartDate, 
                        exportEndDate, 
                        peoplePerDay, 
                        mandaysPerDay, 
                        parseFloat(exportTotalMandays) || calculatedTotalMandays, 
                        exportWeekName
                      );
                      setShowExportModal(false);
                    }}
                    className="w-full py-4 bg-green-600 text-white rounded-xl font-bold text-lg uppercase hover:bg-green-700 active:scale-95 transition-all shadow-lg flex items-center justify-center gap-2"
                  >
                    <FileSpreadsheet size={24} />
                    Xuất File Excel
                  </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {showExportPreparationModal && (
          <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center p-3 sm:p-4 z-[9999]">
            <motion.div 
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-2xl sm:rounded-3xl p-4 sm:p-8 max-w-xl w-full shadow-2xl relative max-h-[90vh] overflow-y-auto no-scrollbar"
            >
              <button className="absolute top-4 right-4 text-gray-500 hover:text-gray-800 cursor-pointer p-2 hover:bg-gray-100 rounded-full transition-colors" onClick={() => setShowExportPreparationModal(false)}>
                <X size={20} />
              </button>
              <h2 className="text-2xl font-bold uppercase text-gray-900 mb-6 flex items-center gap-3">
                <FileSpreadsheet className="text-green-600" />
                Chọn Thông Tin Báo Cáo SX 2
              </h2>
              <div className="space-y-6">
                <div>
                  <label className="block text-sm font-bold uppercase text-gray-600 mb-2">Tuần <span className="text-[10px] text-gray-400 normal-case">(Hiển thị trên tiêu đề)</span></label>
                  <input 
                    type="text"
                    placeholder="VD: 24,25"
                    value={exportWeekName}
                    onChange={(e) => {
                      setExportWeekName(e.target.value);
                      storageService.saveExportWeekName(e.target.value);
                    }}
                    className="w-full bg-gray-50 border border-gray-200 p-3.5 rounded-xl font-bold text-gray-800 outline-none focus:border-green-500 focus:bg-green-50 transition-all"
                  />
                </div>

                <div>
                   <label className="block text-sm font-bold uppercase text-gray-600 mb-2">Chọn PO Tổng để báo cáo</label>
                   <input
                     type="text"
                     placeholder="Tìm kiếm model hoặc mã PO..."
                     value={exportModelSearchTerm}
                     onChange={(e) => setExportModelSearchTerm(e.target.value)}
                     className="w-full mb-2 bg-white border border-gray-200 p-2 rounded-lg font-mono text-sm outline-none focus:border-green-500"
                   />
                   <div className="max-h-[30vh] overflow-y-auto border border-gray-200 rounded-xl p-3 bg-gray-50 flex flex-col gap-1">
                     {(() => {
                       const thirtyFiveDaysAgo = Date.now() - 35 * 24 * 60 * 60 * 1000;
                       const allMasterPOs = storageService.getProductionOrders().filter(o => !o.masterPoId && !o.id.startsWith('REPAIR') && o.createdAt >= thirtyFiveDaysAgo);
                       const allGlazingPlans = storageService.getGlazingPlans().filter(o => o.createdAt >= thirtyFiveDaysAgo);
                       
                       const allPlans = [...allMasterPOs, ...allGlazingPlans];
                       
                       const grouped = allPlans.reduce((acc, po) => {
                         const modelId = 'partId' in po ? po.partId : po.modelId;
                         if (!acc[modelId]) acc[modelId] = [];
                         acc[modelId].push(po);
                         return acc;
                       }, {} as Record<string, any[]>);
                       
                       const filteredModels = Object.keys(grouped).filter(modelId => {
                         const modelName = parts.find(p => p.id === modelId)?.name || modelId;
                         const searchLower = exportModelSearchTerm.toLowerCase();
                         if (modelName.toLowerCase().includes(searchLower) || modelId.toLowerCase().includes(searchLower)) return true;
                         return grouped[modelId].some(po => po.id.toLowerCase().includes(searchLower));
                       });
                       
                       if (filteredModels.length === 0) {
                         return <div className="text-gray-500 text-sm italic p-2">Không tìm thấy kết quả phù hợp.</div>;
                       }
                       
                       return filteredModels.map(modelId => {
                         const modelName = parts.find((p: any) => p.id === modelId)?.name || modelId;
                         const pos = grouped[modelId];
                         
                         return (
                           <div key={modelId} className="border border-gray-200 rounded-lg p-2 bg-white shadow-sm mb-2">
                             <div className="font-bold text-gray-800 mb-2 border-b pb-1">
                               {modelName} <span className="text-gray-500 font-mono text-xs ml-1">({modelId})</span>
                             </div>
                             {pos.sort((a, b) => b.createdAt - a.createdAt).map(po => {
                                if (exportModelSearchTerm && !modelName.toLowerCase().includes(exportModelSearchTerm.toLowerCase()) && !modelId.toLowerCase().includes(exportModelSearchTerm.toLowerCase()) && !po.id.toLowerCase().includes(exportModelSearchTerm.toLowerCase())) return null;
                                
                                const isSelected = selectedMasterPoIdsForExport.includes(po.id);
                                const isGlazing = !('partId' in po);
                                const labelColor = isGlazing ? 'text-indigo-600' : 'text-blue-600';
                                
                                return (
                                   <label key={po.id} className="flex items-center gap-3 p-1.5 hover:bg-gray-50 rounded cursor-pointer ml-2">
                                      <input type="checkbox" checked={isSelected} onChange={(e) => {
                                         if (e.target.checked) setSelectedMasterPoIdsForExport(prev => [...prev, po.id]);
                                         else setSelectedMasterPoIdsForExport(prev => prev.filter(id => id !== po.id));
                                      }} className="w-4 h-4 text-green-600 rounded focus:ring-green-500 cursor-pointer" />
                                      <span className={`font-mono ${labelColor} font-bold text-sm`}>{po.id}</span>
                                      <span className="text-xs text-gray-500">(Kế hoạch: {po.targetQuantity}, Tạo: {format(po.createdAt, 'dd/MM HH:mm')})</span>
                                   </label>
                                );
                             })}
                           </div>
                         )
                       });
                     })()}
                   </div>
                </div>

                <div className="pt-2">
                  <button 
                    onClick={() => {
                      exportPreparationReport(
                        storageService.getProductionOrders(),
                        storageService.getGlazingPlans(),
                        parts,
                        storageService.getGlazingPlanNorms(),
                        selectedMasterPoIdsForExport,
                        exportWeekName
                      );
                      setShowExportPreparationModal(false);
                    }}
                    className="w-full py-4 bg-green-700 text-white rounded-xl font-bold text-lg uppercase hover:bg-green-800 active:scale-95 transition-all shadow-lg flex items-center justify-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed"
                    disabled={selectedMasterPoIdsForExport.length === 0}
                  >
                    <FileSpreadsheet size={24} />
                    Xuất File Excel
                  </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {showExportHourlyModal && (
          <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center p-3 sm:p-4 z-[9999]">
            <motion.div 
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-2xl sm:rounded-3xl p-4 sm:p-8 max-w-2xl w-full shadow-2xl relative max-h-[90vh] overflow-y-auto no-scrollbar"
            >
              <button className="absolute top-4 right-4 text-gray-500 hover:text-gray-800 cursor-pointer p-2 hover:bg-gray-100 rounded-full transition-colors" onClick={() => setShowExportHourlyModal(false)}>
                <X size={20} />
              </button>
              <h2 className="text-xl sm:text-2xl font-bold uppercase text-gray-900 mb-4 sm:mb-6 flex items-center gap-2 sm:gap-3">
                <FileSpreadsheet className="text-green-600" />
                <span>Dữ Liệu Giờ</span>
              </h2>
              <div className="space-y-4 sm:space-y-6">
                <div>
                  <label className="block text-xs sm:text-sm font-bold uppercase text-gray-600 mb-2">Ngày Xuất / Cập nhật số liệu</label>
                  <input 
                    type="date"
                    min={getPastDateStr(35)}
                    max={getPastDateStr(0)}
                    value={chartDate}
                    onChange={(e) => setChartDate(e.target.value)}
                    className="w-full bg-gray-50 border border-gray-200 p-3 sm:p-4 rounded-xl font-bold text-indigo-700 outline-none focus:border-indigo-500 font-mono cursor-pointer text-sm sm:text-base"
                  />
                  <p className="text-[11px] text-emerald-600 mt-1.5 font-medium flex items-center gap-1.5">
                    <span>✨</span> Tự động kế thừa số người từ ngày gần nhất khi sang ngày mới. Chỉ cần chỉnh sửa các ô có thay đổi.
                  </p>
                </div>

                <div className="bg-amber-50/70 border border-amber-200 rounded-xl sm:rounded-2xl p-3 sm:p-4">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-1 mb-2">
                    <label className="text-xs sm:text-sm font-bold uppercase text-amber-950 flex items-center gap-1.5">
                      <Users size={16} className="text-amber-700" />
                      <span>Số Người Định Biên Của Từng Công Đoạn</span>
                    </label>
                    <span className="text-[10px] sm:text-[11px] font-medium text-amber-700 bg-amber-100/70 px-2 py-0.5 rounded-full self-start sm:self-auto">
                      Dùng tính năng lực sản xuất 8h
                    </span>
                  </div>
                  <div className="grid grid-cols-2 sm:grid-cols-5 gap-2 pt-1">
                    <div className="bg-white border border-amber-200/80 rounded-xl p-2.5 text-center shadow-xs">
                      <label className="text-[10px] font-bold text-gray-500 uppercase leading-none block mb-1.5">Cắt Laser</label>
                      <input 
                        type="number" 
                        min="0"
                        placeholder=""
                        value={stagePlannedWorkers.LASER === undefined ? '' : stagePlannedWorkers.LASER}
                        onChange={(e) => {
                          const raw = e.target.value.trim();
                          const val = raw === '' ? undefined : parseInt(raw);
                          handleStagePlannedWorkersChange('LASER', val !== undefined && !isNaN(val) ? val : undefined);
                        }}
                        className="w-full border border-gray-200 rounded-lg p-1.5 text-base font-mono font-bold text-center text-amber-900 outline-none focus:border-amber-500 focus:bg-amber-50/30"
                      />
                    </div>
                    <div className="bg-white border border-amber-200/80 rounded-xl p-2.5 text-center shadow-xs">
                      <label className="text-[10px] font-bold text-gray-500 uppercase leading-none block mb-1.5">Chấn/Dập</label>
                      <input 
                        type="number" 
                        min="0"
                        placeholder=""
                        value={stagePlannedWorkers.BENDING === undefined ? '' : stagePlannedWorkers.BENDING}
                        onChange={(e) => {
                          const raw = e.target.value.trim();
                          const val = raw === '' ? undefined : parseInt(raw);
                          handleStagePlannedWorkersChange('BENDING', val !== undefined && !isNaN(val) ? val : undefined);
                        }}
                        className="w-full border border-gray-200 rounded-lg p-1.5 text-base font-mono font-bold text-center text-amber-900 outline-none focus:border-amber-500 focus:bg-amber-50/30"
                      />
                    </div>
                    <div className="bg-white border border-amber-200/80 rounded-xl p-2.5 text-center shadow-xs">
                      <label className="text-[10px] font-bold text-gray-500 uppercase leading-none block mb-1.5">Hàn</label>
                      <input 
                        type="number" 
                        min="0"
                        placeholder=""
                        value={stagePlannedWorkers.WELDING === undefined ? '' : stagePlannedWorkers.WELDING}
                        onChange={(e) => {
                          const raw = e.target.value.trim();
                          const val = raw === '' ? undefined : parseInt(raw);
                          handleStagePlannedWorkersChange('WELDING', val !== undefined && !isNaN(val) ? val : undefined);
                        }}
                        className="w-full border border-gray-200 rounded-lg p-1.5 text-base font-mono font-bold text-center text-amber-900 outline-none focus:border-amber-500 focus:bg-amber-50/30"
                      />
                    </div>
                    <div className="bg-white border border-amber-200/80 rounded-xl p-2.5 text-center shadow-xs">
                      <label className="text-[10px] font-bold text-gray-500 uppercase leading-none block mb-1.5">Sơn</label>
                      <input 
                        type="number" 
                        min="0"
                        placeholder=""
                        value={stagePlannedWorkers.PAINTING === undefined ? '' : stagePlannedWorkers.PAINTING}
                        onChange={(e) => {
                          const raw = e.target.value.trim();
                          const val = raw === '' ? undefined : parseInt(raw);
                          handleStagePlannedWorkersChange('PAINTING', val !== undefined && !isNaN(val) ? val : undefined);
                        }}
                        className="w-full border border-gray-200 rounded-lg p-1.5 text-base font-mono font-bold text-center text-amber-900 outline-none focus:border-amber-500 focus:bg-amber-50/30"
                      />
                    </div>
                    <div className="bg-white border border-amber-200/80 rounded-xl p-2.5 text-center shadow-xs">
                      <label className="text-[10px] font-bold text-gray-500 uppercase leading-none block mb-1.5">Dán Kính</label>
                      <input 
                        type="number" 
                        min="0"
                        placeholder=""
                        value={stagePlannedWorkers.GLAZING === undefined ? '' : stagePlannedWorkers.GLAZING}
                        onChange={(e) => {
                          const raw = e.target.value.trim();
                          const val = raw === '' ? undefined : parseInt(raw);
                          handleStagePlannedWorkersChange('GLAZING', val !== undefined && !isNaN(val) ? val : undefined);
                        }}
                        className="w-full border border-gray-200 rounded-lg p-1.5 text-base font-mono font-bold text-center text-amber-900 outline-none focus:border-amber-500 focus:bg-amber-50/30"
                      />
                    </div>
                  </div>
                </div>
                
                <div>
                  <label className="block text-sm font-bold uppercase text-gray-600 mb-2">Nhân sự mỗi giờ <span className="text-[10px] text-gray-400 normal-case">(CẮT LASER / CHẤN/DẬP / HÀN / SƠN / DÁN KÍNH - Ca ngày & Ca đêm)</span></label>
                  <div className="max-h-[400px] overflow-y-auto space-y-2 pr-2 custom-scrollbar">
                    {ALL_HOURLY_TIME_SLOTS.map((slotLabel, sIdx) => (
                      <React.Fragment key={slotLabel}>
                        {sIdx === 0 && (
                          <div className="px-2 py-1 bg-amber-100/80 rounded-md text-[11px] font-black uppercase text-amber-800 flex items-center gap-1.5 mt-1">
                            <span>☀️</span>
                            <span>Ca Ngày (06:00 - 19:00)</span>
                          </div>
                        )}
                        {sIdx === 6 && (
                          <div className="px-2 py-1 bg-indigo-100/80 rounded-md text-[11px] font-black uppercase text-indigo-900 flex items-center gap-1.5 mt-3">
                            <span>🌙</span>
                            <span>Ca Đêm (19:00 - 06:00)</span>
                          </div>
                        )}
                        <div className="flex gap-2 items-center bg-gray-50 border border-gray-200 p-2 rounded-lg">
                          <span className="text-xs font-bold w-16 text-gray-600 shrink-0">{slotLabel}</span>
                        <div className="flex-1">
                          <label className="text-[10px] font-bold text-gray-400 uppercase leading-none block mb-1">Cắt Laser</label>
                          <input 
                            type="number" 
                            min="0"
                            placeholder=""
                            value={hourlyPeopleLaser[slotLabel] === undefined ? '' : hourlyPeopleLaser[slotLabel]}
                            onChange={(e) => {
                               const raw = e.target.value.trim();
                               const val = raw === '' ? undefined : parseInt(raw);
                               handleHourlyLaserChange(slotLabel, val !== undefined && !isNaN(val) ? val : undefined);
                            }}
                            className="w-full border border-gray-200 rounded p-1.5 text-sm font-mono text-center outline-none focus:border-green-500"
                          />
                        </div>
                        <div className="flex-1">
                          <label className="text-[10px] font-bold text-gray-400 uppercase leading-none block mb-1">Chấn/Dập</label>
                          <input 
                            type="number" 
                            min="0"
                            placeholder=""
                            value={hourlyPeopleBending[slotLabel] === undefined ? '' : hourlyPeopleBending[slotLabel]}
                            onChange={(e) => {
                               const raw = e.target.value.trim();
                               const val = raw === '' ? undefined : parseInt(raw);
                               handleHourlyBendingChange(slotLabel, val !== undefined && !isNaN(val) ? val : undefined);
                            }}
                            className="w-full border border-gray-200 rounded p-1.5 text-sm font-mono text-center outline-none focus:border-green-500"
                          />
                        </div>
                        <div className="flex-1">
                          <label className="text-[10px] font-bold text-gray-400 uppercase leading-none block mb-1">Hàn</label>
                          <input 
                            type="number" 
                            min="0"
                            placeholder=""
                            value={hourlyPeopleWelding[slotLabel] === undefined ? '' : hourlyPeopleWelding[slotLabel]}
                            onChange={(e) => {
                               const raw = e.target.value.trim();
                               const val = raw === '' ? undefined : parseInt(raw);
                               handleHourlyWeldingChange(slotLabel, val !== undefined && !isNaN(val) ? val : undefined);
                            }}
                            className="w-full border border-gray-200 rounded p-1.5 text-sm font-mono text-center outline-none focus:border-green-500"
                          />
                        </div>
                        <div className="flex-1">
                          <label className="text-[10px] font-bold text-gray-400 uppercase leading-none block mb-1">Sơn</label>
                          <input 
                            type="number" 
                            min="0"
                            placeholder=""
                            value={hourlyPeoplePainting[slotLabel] === undefined ? '' : hourlyPeoplePainting[slotLabel]}
                            onChange={(e) => {
                               const raw = e.target.value.trim();
                               const val = raw === '' ? undefined : parseInt(raw);
                               handleHourlyPaintingChange(slotLabel, val !== undefined && !isNaN(val) ? val : undefined);
                            }}
                            className="w-full border border-gray-200 rounded p-1.5 text-sm font-mono text-center outline-none focus:border-green-500"
                          />
                        </div>
                        <div className="flex-1">
                          <label className="text-[10px] font-bold text-gray-400 uppercase leading-none block mb-1">Dán Kính</label>
                          <input 
                            type="number" 
                            min="0"
                            placeholder=""
                            value={hourlyPeopleGlazing[slotLabel] === undefined ? '' : hourlyPeopleGlazing[slotLabel]}
                            onChange={(e) => {
                               const raw = e.target.value.trim();
                               const val = raw === '' ? undefined : parseInt(raw);
                               handleHourlyGlazingChange(slotLabel, val !== undefined && !isNaN(val) ? val : undefined);
                            }}
                            className="w-full border border-gray-200 rounded p-1.5 text-sm font-mono text-center outline-none focus:border-green-500"
                          />
                        </div>
                      </div>
                    </React.Fragment>
                  ))}
                  </div>
                </div>

                <div className="pt-2">
                  <button 
                    onClick={() => {
                      exportHourlyProductionReport(transactions, parts, chartDate);
                      setShowExportHourlyModal(false);
                    }}
                    className="w-full py-4 bg-green-600 text-white rounded-xl font-bold text-lg uppercase hover:bg-green-700 active:scale-95 transition-all shadow-lg flex items-center justify-center gap-2"
                  >
                    <FileSpreadsheet size={24} />
                    Xuất Báo Cáo
                  </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </motion.div>
  );
}

function ProduceView({ 
  selectedStage, setSelectedStage, 
  selectedPart, setSelectedPart, 
  quantity, setQuantity, 
  handleProduce,
  lastTransaction, setLastTransaction,
  inventory,
  productionOrders,
  parts,
  onPrint,
  onCopy,
  setDefectModal,
  labelSettings
}: any) {
  const [sourceLocation, setSourceLocation] = useState<'IN' | 'OUT'>('IN');
  const [selectedPoId, setSelectedPoId] = useState<string>("");
  const [targetStageId, setTargetStageId] = useState<StageId>(
    (STAGES.find(s => s.id === selectedStage)?.nextStageId || STAGES[0].id) as StageId
  );

  const allAvailablePos = useMemo(() => {
    return productionOrders.filter(p => {
      if (p.stageId !== selectedStage || p.status === 'COMPLETED' || p.status === 'PENDING_APPROVAL' || p.status === 'REJECTED') return false;
      
      if (sourceLocation === 'IN') {
        return p.producedQuantity < p.targetQuantity;
      } else {
        // Allow exporting if there's stock available manually, so we just check against targetQuantity
        return (p.exportedQuantity || 0) < p.targetQuantity;
      }
    });
  }, [selectedStage, sourceLocation, productionOrders]);

  const partsMap = useMemo(() => new Map<string, any>(parts.map((p: any) => [p.id, p])), [parts]);

  // Filter parts based on stage, BOM level, and selected PO
  const filteredParts = React.useMemo(() => {
    let baseParts = parts;
    const glazingConfigs = storageService.getGlazingConfigs();
    const glazingConfigPartIds = new Set(glazingConfigs.map((c: any) => c.partId.toUpperCase()));

    // Inject pseudo parts if we are looking at Glazing OUT
    if (selectedStage === 'GLAZING' && sourceLocation === 'OUT') {
      // Show ONLY items that actually have inventory in GLAZING OUT
      const glzInv = inventory.filter(i => i.stageId === 'GLAZING' && i.location === 'OUT' && i.quantity > 0);
      return glzInv.map(i => {
        const isPseudo = i.partId.startsWith('GLZ-OUT-');
        const name = isPseudo ? i.partId.replace('GLZ-OUT-', '') : (partsMap.get(i.partId)?.name || i.partId);
        return {
          id: i.partId,
          name: name,
          level: 1,
          skipLaser: false, skipBending: false, skipWelding: false, skipPainting: false
        };
      });
    } else if (selectedStage === 'GLAZING') {
      // For GLAZING IN, just return empty because we don't 'produce' IN using this view normally
      return [];
    }

    const availablePoPartIds = new Set(allAvailablePos.map(po => po.partId));
    const selectedPo = selectedPoId ? allAvailablePos.find(po => po.id === selectedPoId) : null;
    const isPaintingExempt = selectedStage === 'PAINTING' && sourceLocation === 'IN';

    // Pre-build set of part IDs that have stock in selectedStage and sourceLocation
    const inStockPartIds = new Set<string>();
    for (const i of inventory) {
      if (i.stageId === selectedStage && i.location === sourceLocation && i.quantity > 0) {
        inStockPartIds.add(i.partId.toUpperCase());
        const effId = storageService.getEffectivePartId(i.partId, selectedStage, selectedPoId);
        if (effId) inStockPartIds.add(effId.toUpperCase());
      }
    }

    return baseParts.filter((p: any) => {
      const pUpper = p.id.toUpperCase();
      if (selectedPoId) {
        if (selectedPo && p.id !== selectedPo.partId) return false;
      } else {
        const hasAvailablePo = availablePoPartIds.has(p.id);
        const hasInventory = inStockPartIds.has(pUpper);
        if (!hasAvailablePo && selectedStage !== 'GLAZING' && !isPaintingExempt && !hasInventory) return false; 
      }

      if (selectedStage === 'LASER') {
        return (sourceLocation === 'IN' ? p.level === 3 : p.level === 2) && !p.skipLaser;
      }
      if (selectedStage === 'BENDING') {
        return p.level === 2 && !p.skipBending;
      }
      if (selectedStage === 'WELDING') {
        return (sourceLocation === 'IN' ? p.level === 2 : p.level === 1) && !p.skipWelding;
      }
      if (selectedStage === 'PAINTING') return (p.level === 1 || p.hasPaintingPO) && !p.skipPainting;
      
      return true;
    });
  }, [parts, inventory, selectedStage, sourceLocation, selectedPoId, allAvailablePos, partsMap]);

  const availablePoOptions = useMemo(() => {
    return allAvailablePos.map(po => {
      const poPart = partsMap.get(po.partId);
      const isCurrentPart = po.partId === selectedPart;
      const currentQty = sourceLocation === 'IN' ? po.producedQuantity : (po.exportedQuantity || 0);
      const rem = Math.max(0, po.targetQuantity - currentQty);
      return {
        id: po.id,
        label: `${isCurrentPart ? '★ ' : ''}${po.id.startsWith('REPAIR') ? '🛠️ [BÙ] ' : ''}${po.id} - ${getProcessValue(poPart?.name || po.partId, poPart, selectedStage, sourceLocation)} (${currentQty}/${po.targetQuantity} - Còn: ${rem})`
      };
    });
  }, [allAvailablePos, partsMap, selectedPart, selectedStage, sourceLocation]);

  const filteredPartOptions = useMemo(() => {
    return filteredParts.map((p: any) => ({ 
      id: p.id, 
      label: `${getProcessValue(p.id, p, selectedStage, sourceLocation)} - ${getProcessValue(p.name, p, selectedStage, sourceLocation)}` 
    }));
  }, [filteredParts, selectedStage, sourceLocation]);

  const availablePos = allAvailablePos.filter(p => p.partId === selectedPart);

  // Check if selected part is in Glazing incoming configuration FOR THIS CURRENT STAGE
  const matchedGlazingConfig = useMemo(() => {
    if (!selectedPart) return null;
    const cleanId = selectedPart.startsWith('GLZ-OUT-') 
      ? selectedPart.trim().toUpperCase() 
      : selectedPart.split(' - ')[0].trim().toUpperCase();
    const partObj = parts.find((p: any) => p.id.toUpperCase() === cleanId);
    const glazingConfigs = storageService.getGlazingConfigs();
    
    // Strict requirement: Only route to GLAZING if this component is specifically configured
    // to be pulled from THIS current stage's OUT warehouse!
    return glazingConfigs.find(c => 
      c.sourceStageId === selectedStage && (
        c.partId.trim().toUpperCase() === cleanId || 
        (partObj?.name && c.partName && c.partName.trim().toUpperCase() === partObj.name.trim().toUpperCase())
      )
    ) || null;
  }, [selectedPart, parts, selectedStage]);

  const isGlazingPart = Boolean(matchedGlazingConfig);

  // Handlers for selection
  const handlePartChange = (partId: string) => {
    setSelectedPart(partId);
    const pos = allAvailablePos.filter(p => p.partId === partId);
    setSelectedPoId(pos.length > 0 ? pos[0].id : "");

    const cleanId = partId.startsWith('GLZ-OUT-') ? partId.trim().toUpperCase() : partId.split(' - ')[0].trim().toUpperCase();
    const partObj = parts.find((p: any) => p.id.toUpperCase() === cleanId);
    const glazingConfigs = storageService.getGlazingConfigs();
    const isGlzThisStage = glazingConfigs.some(c => 
      c.sourceStageId === selectedStage && (
        c.partId.trim().toUpperCase() === cleanId || 
        (partObj?.name && c.partName && c.partName.trim().toUpperCase() === partObj.name.trim().toUpperCase())
      )
    );
    if (sourceLocation === 'OUT') {
      if (isGlzThisStage && selectedStage !== 'GLAZING') {
        setTargetStageId('GLAZING');
      } else {
        const standardNext = (STAGES.find(s => s.id === selectedStage)?.nextStageId || 'BENDING') as StageId;
        setTargetStageId(standardNext);
      }
    }
  };

  const handlePoChange = (poId: string) => {
    setSelectedPoId(poId);
    if (poId) {
      const po = allAvailablePos.find(p => p.id === poId);
      if (po && po.partId !== selectedPart) {
        setSelectedPart(po.partId);
        const cleanId = po.partId.split(' - ')[0].trim().toUpperCase();
        const partObj = parts.find((p: any) => p.id.toUpperCase() === cleanId);
        const glazingConfigs = storageService.getGlazingConfigs();
        const isGlzThisStage = glazingConfigs.some(c => 
          c.sourceStageId === selectedStage && (
            c.partId.trim().toUpperCase() === cleanId || 
            (partObj?.name && c.partName && c.partName.trim().toUpperCase() === partObj.name.trim().toUpperCase())
          )
        );
        if (sourceLocation === 'OUT') {
          if (isGlzThisStage && selectedStage !== 'GLAZING') {
            setTargetStageId('GLAZING');
          } else {
            const standardNext = (STAGES.find(s => s.id === selectedStage)?.nextStageId || 'BENDING') as StageId;
            setTargetStageId(standardNext);
          }
        }
      }
    }
  };

  // Auto-select first filtered part if current selection is not in list
  useEffect(() => {
    if (filteredParts.length > 0 && !filteredParts.find((p: any) => p.id === selectedPart)) {
      const firstPartId = filteredParts[0].id;
      setSelectedPart(firstPartId);
      const pos = allAvailablePos.filter(p => p.partId === firstPartId);
      setSelectedPoId(pos.length > 0 ? pos[0].id : "");
    }
  }, [selectedStage, filteredParts, selectedPart, allAvailablePos]);

  useEffect(() => {
    // If operator is exporting OUT and this component is configured to go to Glazing FROM THIS STAGE
    if (sourceLocation === 'OUT' && isGlazingPart && selectedStage !== 'GLAZING') {
      setTargetStageId('GLAZING');
      return;
    }

    const currentIdx = STAGES.findIndex(s => s.id === selectedStage);
    const part = parts.find((p: any) => p.id === selectedPart);
    
    // Find the next non-skipped stage
    const nextAvailableStage = STAGES.find((s, idx) => {
      if (idx <= currentIdx) return false;
      if (s.id === 'LASER' && part?.skipLaser) return false;
      if (s.id === 'BENDING' && part?.skipBending) return false;
      if (s.id === 'WELDING' && part?.skipWelding) return false;
      if (s.id === 'PAINTING' && part?.skipPainting) return false;
      return true;
    });

    if (nextAvailableStage) {
      setTargetStageId(nextAvailableStage.id);
    } else {
      setTargetStageId(STAGES[0].id as StageId); // Default if no next valid stage
    }
    
    // Auto-switch to OUT for stages with automatic BOM deduction
    if (selectedStage === 'LASER' || selectedStage === 'WELDING' || selectedStage === 'GLAZING') {
      setSourceLocation('OUT');
    }
  }, [selectedStage, selectedPart, sourceLocation, parts, isGlazingPart]);

  const currentStock = useMemo(() => {
    const cleanId = selectedPart.startsWith('GLZ-OUT-') ? selectedPart.trim().toUpperCase() : selectedPart.split(' - ')[0].trim().toUpperCase();
    const effectiveId = storageService.getEffectivePartId(cleanId, selectedStage, selectedPoId);
    
    return inventory.reduce((sum: number, item: any) => {
      if (item.stageId === selectedStage && item.location === sourceLocation) {
        const itPartId = item.partId.toUpperCase();
        const targetId = effectiveId.toUpperCase();
        const partInCatalog = parts.find((p: any) => p.id.toUpperCase() === targetId);
        const targetName = partInCatalog ? partInCatalog.name.toUpperCase() : '';

        if (itPartId === targetId || (targetName && itPartId === targetName)) {
          return sum + item.quantity;
        }
      }
      return sum;
    }, 0);
  }, [inventory, selectedPart, selectedStage, sourceLocation, selectedPoId]);

  const activePo = allAvailablePos.find(p => p.id === selectedPoId) || availablePos[0];
  const poProcessed = activePo ? (sourceLocation === 'IN' ? activePo.producedQuantity : (activePo.exportedQuantity || 0)) : 0;
  const poTarget = activePo?.targetQuantity || 0;
  const poRemaining = activePo ? Math.max(0, poTarget - poProcessed) : 0;
  const isOverPo = Boolean(activePo && quantity > 0 && ((poProcessed + quantity) > poTarget));
  const overAmount = isOverPo ? ((poProcessed + quantity) - poTarget) : 0;

  return (
    <motion.div 
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, y: -20 }}
      className="grid grid-cols-1 lg:grid-cols-2 gap-4 sm:gap-8"
    >
      <div className="bg-white p-3.5 sm:p-8 rounded-2xl border border-gray-200 shadow-sm space-y-4 sm:space-y-6">
        {activePo && (
          <div className={cn(
            "p-3.5 sm:p-5 rounded-xl sm:rounded-2xl border transition-all space-y-2 sm:space-y-3",
            isOverPo ? "bg-amber-50 border-amber-300" : "bg-blue-50 border-blue-200"
          )}>
            <div className="flex justify-between items-center">
              <div>
                <div className="text-[11px] sm:text-xs font-bold uppercase text-blue-600 tracking-wider">Lệnh PO đã chọn</div>
                <div className="font-mono font-bold text-blue-900 text-sm sm:text-lg flex items-center gap-1.5 sm:gap-2">
                  <span>{activePo.id}</span>
                  {activePo.id.startsWith('REPAIR') && (
                    <span className="text-[10px] bg-amber-200 text-amber-900 px-1.5 py-0.5 rounded font-sans font-bold">LỆNH BÙ</span>
                  )}
                </div>
              </div>
              <div className="text-right">
                <div className="text-[11px] sm:text-xs font-bold uppercase text-blue-600 tracking-wider">
                  {sourceLocation === 'IN' ? 'Tiến độ SX (IN)' : 'Tiến độ Xuất (OUT)'}
                </div>
                <div className="font-mono font-bold text-blue-900 text-sm sm:text-lg">
                  {poProcessed} / {poTarget}
                </div>
              </div>
            </div>

            <div className="flex flex-wrap items-center justify-between gap-1.5 pt-2 border-t border-blue-200/60 text-xs">
              <div className="font-medium text-gray-700 text-xs">
                Còn lại theo PO: <span className="font-mono font-bold text-blue-900 text-sm">{poRemaining}</span>
              </div>
              {isOverPo ? (
                <div className="flex items-center gap-1.5 text-red-800 font-bold bg-red-100 border border-red-300 px-2 py-0.5 rounded-lg text-[11px]">
                  <AlertCircle size={13} className="text-red-600 flex-shrink-0" />
                  <span>VƯỢT QUÁ PO (+{overAmount}) — KHÔNG ĐƯỢC XUẤT!</span>
                </div>
              ) : quantity > 0 ? (
                <div className="flex items-center gap-1 text-emerald-800 font-bold bg-emerald-100/90 border border-emerald-300 px-2 py-0.5 rounded-lg text-[11px]">
                  <Check size={13} className="text-emerald-600 flex-shrink-0" />
                  <span>Trong hạn mức PO (Còn lại: {poRemaining - quantity})</span>
                </div>
              ) : null}
            </div>
          </div>
        )}
        <div>
          <h2 className="text-xl sm:text-3xl font-bold tracking-tight mb-1 sm:mb-2 text-gray-900">Xuất kho (Finish & Label)</h2>
          <p className="text-xs sm:text-sm text-gray-500 leading-relaxed">Ghi nhận hoàn thành công đoạn hoặc xuất linh kiện từ kho thành phẩm để in nhãn QR.</p>
        </div>
        
        <form onSubmit={async (e) => {
          e.preventDefault();
          if (quantity <= 0) {
            await customAlert('Vui lòng nhập số lượng lớn hơn 0');
            return;
          }
          if (quantity > currentStock + 0.0001) {
            await customAlert(`Số lượng xuất (${quantity}) vượt quá tồn kho hiện tại (${currentStock})`);
            return;
          }
          if (selectedPoId) {
            const po = productionOrders.find((p: any) => p.id === selectedPoId);
            if (po) {
              const currentProgress = sourceLocation === 'IN' ? po.producedQuantity : (po.exportedQuantity || 0);
              const nextTotal = currentProgress + quantity;
              if (nextTotal > po.targetQuantity) {
                const maxRem = Math.max(0, po.targetQuantity - currentProgress);
                await customAlert(`Số lượng xuất kho (${quantity}) sẽ vượt quá số lượng mục tiêu của PO [${po.id}] (${po.targetQuantity}).\nHiện đã xuất/làm: ${currentProgress}, tối đa chỉ được làm thêm: ${maxRem}.\n\nHệ thống KHÔNG cho phép xuất vượt quá PO!`);
                return;
              }
            }
          }
          handleProduce(e, sourceLocation, sourceLocation === 'OUT' ? targetStageId : undefined, selectedPoId);
        }} className="space-y-4 sm:space-y-6">
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 sm:gap-6">
            <div className="space-y-1.5 sm:space-y-3">
              <label className="text-xs sm:text-sm font-bold uppercase tracking-wider text-gray-700">1. Công đoạn hiện tại</label>
              <select 
                value={selectedStage}
                onChange={(e) => setSelectedStage(e.target.value)}
                className="w-full p-2.5 sm:p-4 rounded-xl border border-gray-200 font-bold text-xs sm:text-base focus:border-blue-600 outline-none bg-white cursor-pointer"
              >
                {STAGES.map(stage => (
                  <option key={stage.id} value={stage.id}>{stage.name}</option>
                ))}
              </select>
            </div>
            <div className="space-y-1.5 sm:space-y-3">
              <label className="text-xs sm:text-sm font-bold uppercase tracking-wider text-gray-700">2. Vị trí xuất:</label>
              <div className="flex bg-gray-100 p-1 rounded-xl h-[42px] sm:h-[54px]">
                {(selectedStage !== 'LASER' && selectedStage !== 'WELDING') && (
                  <button 
                    type="button"
                    onClick={() => setSourceLocation('IN')}
                    className={cn(
                      "flex-1 rounded-lg font-bold text-xs sm:text-sm uppercase transition-all cursor-pointer",
                      sourceLocation === 'IN' ? "bg-gray-900 text-white shadow-sm" : "text-gray-500 hover:text-gray-900"
                    )}
                  >
                    Kho IN
                  </button>
                )}
                <button 
                  type="button"
                  onClick={() => setSourceLocation('OUT')}
                  className={cn(
                    "flex-1 rounded-lg font-bold text-xs sm:text-sm uppercase transition-all cursor-pointer",
                    sourceLocation === 'OUT' ? (selectedStage === 'LASER' || selectedStage === 'WELDING' ? "bg-blue-600 text-white shadow-sm" : "bg-[#F27D26] text-white shadow-sm") : "text-gray-500 hover:text-gray-900"
                  )}
                >
                  Kho OUT
                </button>
              </div>
            </div>
          </div>

          {allAvailablePos.length > 0 && (
            <div className="space-y-1.5 sm:space-y-3">
              <div className="flex flex-col sm:flex-row sm:justify-between sm:items-center gap-1">
                <label className="text-xs sm:text-sm font-bold uppercase tracking-wider text-gray-700">3. Chọn Lệnh PO (Ràng buộc tiến độ)</label>
                {availablePos.length > 1 && (
                  <span className="text-[10px] sm:text-xs bg-purple-100 text-purple-700 font-bold px-2 py-0.5 rounded-md self-start sm:self-auto">
                    Linh kiện nằm trong {availablePos.length} PO khác nhau
                  </span>
                )}
              </div>
              <SearchableSelect 
                options={availablePoOptions}
                value={selectedPoId}
                onChange={handlePoChange}
                placeholder="Tìm và chọn đúng Lệnh PO..."
              />
            </div>
          )}

          <div className="space-y-1.5 sm:space-y-3">
            <label className="text-xs sm:text-sm font-bold uppercase tracking-wider text-gray-700">4. Chọn mã linh kiện</label>
            <SearchableSelect 
              options={filteredPartOptions}
              value={selectedPart}
              onChange={handlePartChange}
              placeholder="Tìm mã linh kiện..."
            />
          </div>

          {sourceLocation === 'OUT' && (
            <div className="space-y-1.5 sm:space-y-3">
              <div className="flex flex-wrap justify-between items-center gap-1">
                <label className="text-xs sm:text-sm font-bold uppercase tracking-wider text-gray-700">4. Công đoạn đích (Nhập kho IN)</label>
                {isGlazingPart && targetStageId === 'GLAZING' && (
                  <span className="text-[10px] font-bold text-emerald-800 bg-emerald-100 border border-emerald-300 px-2.5 py-0.5 rounded-full flex items-center gap-1 shadow-2xs">
                    <Sparkles size={11} className="text-emerald-600" />
                    Tự động chọn: Dán Kính (Linh kiện kho IN Dán Kính)
                  </span>
                )}
              </div>
              <select 
                value={targetStageId}
                onChange={(e) => setTargetStageId(e.target.value as StageId)}
                className={cn(
                  "w-full p-2.5 sm:p-4 rounded-xl border font-bold text-xs sm:text-base outline-none bg-white cursor-pointer transition-all",
                  isGlazingPart && targetStageId === 'GLAZING'
                    ? "border-emerald-500 bg-emerald-50/30 text-emerald-950 ring-2 ring-emerald-200"
                    : "border-gray-200 focus:border-blue-600"
                )}
              >
                {STAGES.map(s => (
                  <option key={s.id} value={s.id}>{s.name}</option>
                ))}
              </select>
            </div>
          )}

          <div className="space-y-1.5 sm:space-y-3">
            <div className="flex justify-between items-center">
              <label className="text-xs sm:text-sm font-bold uppercase tracking-wider text-gray-700">{sourceLocation === 'OUT' ? '5' : '4'}. Nhập số lượng xuất</label>
              <div className={cn(
                "flex items-center gap-2 px-2.5 py-1 rounded-lg border",
                sourceLocation === 'IN' ? "bg-blue-50 border-blue-100" : "bg-orange-50 border-orange-100"
              )}>
                <span className={cn(
                  "text-[10px] sm:text-xs font-mono font-bold uppercase",
                  sourceLocation === 'IN' ? "text-blue-600" : "text-orange-600"
                )}>Tồn {sourceLocation}:</span>
                <span className={cn(
                  "text-xs sm:text-sm font-mono font-bold",
                  sourceLocation === 'IN' ? "text-blue-700" : "text-orange-700"
                )}>{currentStock}</span>
              </div>
            </div>
            <div className="flex items-center gap-2.5 sm:gap-4">
              <button 
                type="button"
                onClick={() => setQuantity(Math.max(0, quantity - 1))}
                className="w-12 h-12 sm:w-16 sm:h-16 rounded-xl border-2 border-gray-200 flex items-center justify-center font-bold text-xl sm:text-3xl hover:bg-gray-50 active:scale-95 transition-all cursor-pointer shrink-0"
              >
                -
              </button>
              <input 
                type="number" 
                step="any"
                value={quantity || ''} 
                onChange={(e) => setQuantity(parseFloat(e.target.value) || 0)}
                className="flex-1 h-12 sm:h-16 text-center border-2 border-gray-900 rounded-xl font-mono font-bold text-xl sm:text-3xl outline-none focus:ring-4 ring-black/5"
                placeholder="0"
              />
              <button 
                type="button"
                onClick={() => setQuantity(quantity + 1)}
                className="w-12 h-12 sm:w-16 sm:h-16 rounded-xl border-2 border-gray-200 flex items-center justify-center font-bold text-xl sm:text-3xl hover:bg-gray-50 active:scale-95 transition-all cursor-pointer shrink-0"
              >
                +
              </button>
            </div>
          </div>

          {availablePos.length === 0 && selectedStage !== 'GLAZING' && !(selectedStage === 'PAINTING' && sourceLocation === 'IN') ? (
            <div className="w-full bg-red-50 border border-red-200 p-4 sm:p-6 rounded-xl flex flex-col items-center gap-2.5 text-center">
              <AlertCircle size={26} className="text-red-600" />
              <div>
                <p className="font-bold text-red-800 uppercase text-sm sm:text-base">Không có Lệnh sản xuất (PO)</p>
                <p className="text-xs sm:text-sm text-red-600">Linh kiện này chưa được tạo PO cho công đoạn {STAGES.find(s => s.id === selectedStage)?.name}. Vui lòng vào mục "Lệnh sản xuất" để tạo lệnh trước.</p>
              </div>
            </div>
          ) : (
            <div className="flex gap-2.5 sm:gap-4 pt-1">
              <button 
                type="submit"
                disabled={(!selectedPoId && selectedStage !== 'GLAZING' && !(selectedStage === 'PAINTING' && sourceLocation === 'IN')) || isOverPo}
                className={cn(
                  "flex-1 py-3.5 sm:py-5 rounded-xl font-bold uppercase text-xs sm:text-base tracking-wider flex items-center justify-center gap-2 active:scale-[0.98] transition-all shadow-md cursor-pointer",
                  ((!selectedPoId && selectedStage !== 'GLAZING' && !(selectedStage === 'PAINTING' && sourceLocation === 'IN')) || isOverPo)
                    ? "bg-gray-300 text-gray-500 cursor-not-allowed" 
                    : (selectedStage === 'LASER' || selectedStage === 'WELDING' || sourceLocation === 'IN' ? "bg-blue-600 text-white hover:bg-blue-700 shadow-blue-200" : "bg-[#F27D26] text-white hover:bg-[#F27D26]/90 shadow-[#F27D26]/20")
                )}
              >
                {sourceLocation === 'IN' ? <CheckCircle2 size={18} /> : <Printer size={18} />}
                {sourceLocation === 'IN' ? 'Xác nhận hoàn thành' : 'In nhãn QR'}
              </button>
              {sourceLocation === 'IN' || selectedStage === 'LASER' ? (
                <button 
                  type="button"
                  onClick={() => setDefectModal({ partId: selectedPart, stageId: selectedStage, poId: selectedPoId, location: sourceLocation })}
                  className="w-24 sm:w-32 py-3.5 sm:py-5 rounded-xl border border-red-200 bg-red-50 text-red-600 font-bold uppercase text-[10px] sm:text-xs flex flex-col items-center justify-center gap-1 hover:bg-red-100 transition-all active:scale-[0.98] shadow-sm cursor-pointer shrink-0"
                >
                  <AlertCircle size={16} />
                  Báo lỗi (NG)
                </button>
              ) : null}
            </div>
          )}
        </form>
      </div>

      {/* QR Label Preview */}
      <div className="flex flex-col">
        <AnimatePresence mode="wait">
          {lastTransaction && lastTransaction.qrData ? (
            <motion.div 
              key="label"
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white p-8 rounded-2xl border-2 border-dashed border-[#141414] flex flex-col items-center space-y-6 h-full"
            >
              <div className="flex justify-between w-full items-center no-print">
                <div className="flex items-center gap-2">
                  <Printer size={20} className="text-[#F27D26]" />
                  <h3 className="font-mono text-xs uppercase font-bold tracking-widest">Xem trước nhãn In</h3>
                </div>
                <button onClick={() => setLastTransaction(null)} className="text-gray-400 hover:text-black transition-colors">
                  <X size={24} />
                </button>
              </div>
              
              <A7QRLabelCard label={lastTransaction} parts={parts} labelSettings={labelSettings} isPrintArea={false} />

              <div className="w-full flex gap-4 no-print mt-auto">
                <button 
                  onClick={() => onPrint(lastTransaction)}
                  className="flex-1 bg-black text-white py-5 rounded-xl flex items-center justify-center gap-3 font-bold text-lg uppercase hover:bg-black/90 active:scale-[0.98] transition-all"
                >
                  <Printer size={24} />
                  In nhãn QR
                </button>
                <button 
                  onClick={() => onCopy(lastTransaction.qrData || '')}
                  className="px-6 bg-gray-100 text-gray-600 py-5 rounded-xl flex items-center justify-center gap-3 font-bold text-lg uppercase hover:bg-gray-200 transition-all"
                  title="Sao chép mã QR"
                >
                  <QrCode size={24} />
                </button>
              </div>
            </motion.div>
          ) : (
            <motion.div 
              key="empty"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className="bg-gray-100 rounded-2xl border-2 border-dashed border-gray-300 flex flex-col items-center justify-center h-full p-12 text-center text-gray-400"
            >
              <Printer size={64} strokeWidth={1} />
              <p className="mt-6 font-medium">Sau khi xác nhận sản xuất,<br />nhãn QR sẽ hiển thị tại đây để in.</p>
            </motion.div>
          )}
        </AnimatePresence>
      </div>
    </motion.div>
  );
}

function WeldingInboundView({ parts, onManualInbound, setDefectModal, productionOrders }: any) {
  const [manualPart, setManualPart] = useState('');
  const [manualQty, setManualQty] = useState(0);
  const [selectedPoId, setSelectedPoId] = useState<string>("");
  const [lastScanned, setLastScanned] = useState<any>(null);

  const selectedStage = 'WELDING';
  const targetLocation = 'OUT';

  const allAvailablePos = useMemo(() => {
    return (productionOrders || []).filter(
      (p: any) => p.stageId === selectedStage && p.status !== 'COMPLETED' && p.status !== 'PENDING_APPROVAL' && p.status !== 'REJECTED'
    );
  }, [selectedStage, productionOrders]);

  const filteredParts = parts.filter((p: any) => {
    if (selectedPoId) {
      const selectedPo = allAvailablePos.find(po => po.id === selectedPoId);
      if (selectedPo && p.id !== selectedPo.partId) return false;
    } else {
      const hasAvailablePo = allAvailablePos.some(po => po.partId === p.id);
      if (!hasAvailablePo) return false;
    }
    return p.level === 1 && !p.skipWelding;
  });

  const availablePos = allAvailablePos.filter(p => p.partId === manualPart);

  const handlePartChange = (partId: string) => {
    setManualPart(partId);
    const pos = allAvailablePos.filter(p => p.partId === partId);
    setSelectedPoId(pos.length > 0 ? pos[0].id : "");
  };

  const handlePoChange = (poId: string) => {
    setSelectedPoId(poId);
    if (poId) {
      const po = allAvailablePos.find(p => p.id === poId);
      if (po && po.partId !== manualPart) {
        setManualPart(po.partId);
      }
    }
  };

  useEffect(() => {
    if (filteredParts.length > 0 && !filteredParts.find((p: any) => p.id === manualPart)) {
      const firstPartId = filteredParts[0].id;
      setManualPart(firstPartId);
      const pos = allAvailablePos.filter(p => p.partId === firstPartId);
      setSelectedPoId(pos.length > 0 ? pos[0].id : "");
    }
  }, [filteredParts, manualPart, allAvailablePos]);

  const selectedPo = allAvailablePos.find((p: any) => p.id === selectedPoId);
  const isOverPo = Boolean(selectedPo && manualQty > 0 && (selectedPo.producedQuantity + manualQty > selectedPo.targetQuantity));
  const overAmount = isOverPo && selectedPo ? (selectedPo.producedQuantity + manualQty - selectedPo.targetQuantity) : 0;

  const handleManualSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (isOverPo) {
      alert(`Số lượng nhập kho (${manualQty}) sẽ vượt quá số lượng mục tiêu của PO [${selectedPo?.id}] (${selectedPo?.targetQuantity}).\nHiện đã làm: ${selectedPo?.producedQuantity}.\nHệ thống KHÔNG cho phép nhập vượt quá PO!`);
      return;
    }
    if (manualPart && manualQty > 0) {
      try {
        const part = parts.find((p: any) => p.id === manualPart);
        onManualInbound(manualPart, selectedStage, targetLocation, manualQty, selectedPoId);
        setLastScanned({
          partId: getProcessValue(manualPart, part, selectedStage, targetLocation),
          quantity: manualQty,
          partName: getProcessValue(part?.name, part, selectedStage, targetLocation),
          status: 'success',
          isManual: true,
          poId: selectedPoId
        });
        setManualQty(0);
      } catch (err) {
        const part = parts.find((p: any) => p.id === manualPart);
        setLastScanned({
          partId: getProcessValue(manualPart, part, selectedStage, targetLocation),
          quantity: manualQty,
          partName: getProcessValue(part?.name, part, selectedStage, targetLocation),
          status: 'error',
          isManual: true,
          errorMsg: err instanceof Error ? err.message : 'Lỗi không xác định'
        });
      }
    }
  };

  return (
    <motion.div 
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, y: -20 }}
      className="grid grid-cols-1 lg:grid-cols-2 gap-8"
    >
      <div className="bg-white p-10 rounded-3xl border border-gray-200 shadow-xl space-y-8">
        <div className="flex items-center gap-4 mb-2">
          <div className="bg-blue-600 p-3 rounded-xl text-white">
            <Flame size={24} />
          </div>
          <div>
            <h3 className="text-xl font-bold uppercase tracking-tight">Nhập kho Hàn</h3>
            <p className="text-sm text-gray-500">Nhập linh kiện (IN) hoặc kết quả hàn (OUT) thủ công</p>
          </div>
        </div>

        <form onSubmit={handleManualSubmit} className="space-y-8 text-left">
          <div className="grid grid-cols-2 gap-8">
            <div className="space-y-4">
              <label className="text-sm font-bold uppercase tracking-widest opacity-50">Công đoạn:</label>
              <div className="w-full p-5 rounded-xl border-2 border-gray-100 font-bold text-lg bg-gray-50 text-gray-500">
                HÀN
              </div>
            </div>
            <div className="space-y-4">
              <label className="text-sm font-bold uppercase tracking-widest opacity-50">Vị trí nhập:</label>
              <div className="flex bg-gray-100 p-1 rounded-xl h-[68px]">
                <div className="flex-1 rounded-lg font-bold text-sm uppercase bg-[#F27D26] text-white shadow-sm flex items-center justify-center">
                  Kho OUT (Thành phẩm)
                </div>
              </div>
            </div>
          </div>

          {allAvailablePos.length > 0 && (
            <div className="space-y-4">
              <label className="text-sm font-bold uppercase tracking-widest opacity-50">Chọn Lệnh PO (Tiến độ)</label>
              <SearchableSelect 
                options={allAvailablePos.map(po => ({
                  id: po.id,
                  label: `${po.id.startsWith('REPAIR') ? '🛠️ [BÙ] ' : ''}${po.id} - ${getProcessValue(parts.find((p: any) => p.id === po.partId)?.name || po.partId, parts.find((p: any) => p.id === po.partId), selectedStage, targetLocation)} (${po.producedQuantity}/${po.targetQuantity})`
                }))}
                value={selectedPoId}
                onChange={handlePoChange}
                placeholder="Tìm Lệnh PO..."
              />
            </div>
          )}

          <div className="space-y-4">
            <label className="text-sm font-bold uppercase tracking-widest opacity-50">Chọn mã thành phẩm:</label>
            <SearchableSelect 
              options={filteredParts.map((p: any) => ({ 
                id: p.id, 
                label: `${getProcessValue(p.id, p, selectedStage, targetLocation)} - ${getProcessValue(p.name, p, selectedStage, targetLocation)}` 
              }))}
              value={manualPart}
              onChange={handlePartChange}
              placeholder="Tìm kiếm..."
            />
          </div>

          <div className="space-y-4">
            <div className="flex justify-between items-center">
              <label className="text-sm font-bold uppercase tracking-widest opacity-50">Số lượng nhập:</label>
              {manualPart && (
                <span className="text-xs font-mono font-bold text-gray-400 uppercase">
                  Đơn vị: {parts.find((p: any) => p.id === manualPart)?.unit}
                </span>
              )}
            </div>
            <input 
              type="number"
              step="any"
              value={manualQty || ''}
              onChange={(e) => setManualQty(parseFloat(e.target.value) || 0)}
              className="w-full p-5 rounded-xl border-2 border-gray-100 font-mono text-2xl font-bold focus:border-blue-600 outline-none"
              placeholder="Nhập số lượng..."
            />
          </div>

          {isOverPo && (
            <div className="w-full bg-red-50 border-2 border-red-200 p-4 rounded-xl flex items-center gap-3 text-red-700 font-bold">
              <AlertCircle size={24} className="text-red-600 flex-shrink-0" />
              <span>VƯỢT QUÁ PO (+{overAmount}) — KHÔNG ĐƯỢC PHÉP NHẬP!</span>
            </div>
          )}

          {availablePos.length === 0 ? (
            <div className="w-full bg-red-50 border-2 border-red-200 p-6 rounded-xl flex flex-col items-center gap-3 text-center">
              <AlertCircle size={32} className="text-red-600" />
              <div>
                <p className="font-bold text-red-800 uppercase">Yêu cầu Lệnh sản xuất (PO)</p>
                <p className="text-xs text-red-600">Bạn phải chọn PO cụ thể để ghi nhận sản xuất cho công đoạn Hàn.</p>
              </div>
            </div>
          ) : (
            <button 
              type="submit"
              disabled={!selectedPoId || isOverPo}
              className={cn(
                "w-full py-6 rounded-xl font-bold uppercase tracking-widest flex items-center justify-center gap-4 active:scale-[0.98] transition-all shadow-lg text-xl",
                (!selectedPoId || isOverPo) 
                  ? "bg-gray-300 text-gray-500 cursor-not-allowed" 
                  : "bg-blue-600 text-white hover:bg-blue-700 shadow-blue-200"
              )}
            >
              <CheckCircle2 size={28} />
              Xác nhận nhập kho Hàn
            </button>
          )}
        </form>
      </div>

      <div className="flex flex-col">
        <AnimatePresence mode="wait">
          {lastScanned ? (
            <motion.div
              key="scan-result"
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              className={cn(
                "p-10 rounded-3xl border-2 shadow-2xl h-full flex flex-col items-center justify-center space-y-6 bg-white",
                lastScanned.status === 'success' ? "border-green-500" : "border-red-500"
              )}
            >
              <div className={cn(
                "w-16 h-16 rounded-full flex items-center justify-center",
                lastScanned.status === 'success' ? "bg-green-100 text-green-600" : "bg-red-100 text-red-600"
              )}>
                {lastScanned.status === 'success' ? <CheckCircle2 size={32} /> : <AlertCircle size={32} />}
              </div>
              <div className="text-center space-y-2">
                <h3 className="text-xl font-bold">
                  {lastScanned.status === 'success' ? 'Nhập kho thành công!' : 'Lỗi nhập kho!'}
                </h3>
                <p className="text-gray-500 text-sm">
                  {lastScanned.status === 'success' 
                    ? 'Thông tin linh kiện vừa nhập tay:' 
                    : lastScanned.errorMsg}
                </p>
              </div>
              
              {lastScanned.partId && (
                <div className="w-full bg-gray-50 p-8 rounded-2xl space-y-6">
                  <div className="flex justify-between items-center border-b border-gray-200 pb-4">
                    <span className="text-sm font-mono uppercase opacity-50">Mã LK</span>
                    <span className="font-bold text-xl">{lastScanned.partId}</span>
                  </div>
                  <div className="flex justify-between items-center border-b border-gray-200 pb-4">
                    <span className="text-sm font-mono uppercase opacity-50">Linh kiện</span>
                    <span className="font-bold text-xl">{lastScanned.partName}</span>
                  </div>
                  <div className="flex justify-between items-center border-b border-gray-200 pb-4">
                    <span className="text-sm font-mono uppercase opacity-50">Số lượng</span>
                    <span className="font-bold text-2xl">{lastScanned.quantity}</span>
                  </div>
                  <div className="flex justify-between items-center">
                    <span className="text-sm font-mono uppercase opacity-50">Đích</span>
                    <div className="flex items-center gap-3 font-bold text-[#F27D26] text-xl">
                      <span>HÀN</span>
                      <ArrowRight size={20} />
                      <span className="px-3 py-1 rounded text-xs bg-orange-100 text-orange-700">
                        KHO_OUT
                      </span>
                    </div>
                  </div>
                  {lastScanned.poId && (
                    <div className="flex justify-between items-center pt-4 border-t border-gray-200">
                      <span className="text-xs font-mono uppercase opacity-50">Lệnh PO</span>
                      <span className="text-sm font-bold text-blue-600">{lastScanned.poId}</span>
                    </div>
                  )}
                </div>
              )}

              <div className="flex gap-4 w-full">
                <button 
                  onClick={() => setLastScanned(null)}
                  className="flex-1 py-4 text-sm font-bold uppercase tracking-widest text-gray-400 hover:text-[#141414] transition-colors border-2 border-gray-100 rounded-xl"
                >
                  Đóng
                </button>
                {lastScanned.status === 'success' && (
                  <button 
                    onClick={() => setDefectModal({ partId: lastScanned.partId, stageId: 'WELDING' })}
                    className="flex-1 py-4 text-sm font-bold uppercase tracking-widest text-red-600 hover:bg-red-50 transition-colors border-2 border-red-100 rounded-xl flex items-center justify-center gap-2"
                  >
                    <AlertCircle size={18} />
                    Báo lỗi (NG)
                  </button>
                )}
              </div>
            </motion.div>
          ) : (
            <div className="bg-gray-100 rounded-3xl border-2 border-dashed border-gray-300 h-full flex flex-col items-center justify-center p-12 text-center text-gray-400">
              <Package size={64} strokeWidth={1} />
              <p className="mt-6 font-medium">Chi tiết nhập kho Hàn<br />sẽ hiển thị tại đây.</p>
            </div>
          )}
        </AnimatePresence>
      </div>
    </motion.div>
  );
}

function LaserInboundView({ parts, onManualInbound, setDefectModal, productionOrders }: any) {
  const [targetLocation, setTargetLocation] = useState<'IN' | 'OUT'>('IN');
  const [manualPart, setManualPart] = useState('');
  const [manualQty, setManualQty] = useState(0);
  const [selectedPoId, setSelectedPoId] = useState<string>("");
  const [lastScanned, setLastScanned] = useState<any>(null);

  const selectedStage = 'LASER';

  const allAvailablePos = useMemo(() => {
    return (productionOrders || []).filter(
      (p: any) => p.stageId === selectedStage && p.status !== 'COMPLETED' && p.status !== 'PENDING_APPROVAL' && p.status !== 'REJECTED'
    );
  }, [selectedStage, productionOrders]);

  const filteredParts = parts.filter((p: any) => {
    if (selectedPoId && targetLocation === 'OUT') {
      const selectedPo = allAvailablePos.find(po => po.id === selectedPoId);
      if (selectedPo && p.id !== selectedPo.partId) return false;
    } else if (targetLocation === 'OUT') {
      const hasAvailablePo = allAvailablePos.some(po => po.partId === p.id);
      if (!hasAvailablePo) return false;
    }
    
    if (targetLocation === 'IN') return p.level === 3;
    return p.level === 2;
  });

  const availablePos = allAvailablePos.filter(p => p.partId === manualPart);

  const handlePartChange = (partId: string) => {
    setManualPart(partId);
    const pos = allAvailablePos.filter(p => p.partId === partId);
    setSelectedPoId(pos.length > 0 ? pos[0].id : "");
  };

  const handlePoChange = (poId: string) => {
    setSelectedPoId(poId);
    if (poId && targetLocation === 'OUT') {
      const po = allAvailablePos.find(p => p.id === poId);
      if (po && po.partId !== manualPart) {
        setManualPart(po.partId);
      }
    }
  };

  useEffect(() => {
    if (filteredParts.length > 0 && !filteredParts.find((p: any) => p.id === manualPart)) {
      const firstPartId = filteredParts[0].id;
      setManualPart(firstPartId);
      const pos = allAvailablePos.filter(p => p.partId === firstPartId);
      setSelectedPoId(pos.length > 0 ? pos[0].id : "");
    }
  }, [filteredParts, manualPart, allAvailablePos]);

  const selectedPo = allAvailablePos.find(p => p.id === selectedPoId);
  const isOverPo = Boolean(targetLocation === 'OUT' && selectedPo && manualQty > 0 && (selectedPo.producedQuantity + manualQty > selectedPo.targetQuantity));
  const overAmount = isOverPo && selectedPo ? (selectedPo.producedQuantity + manualQty - selectedPo.targetQuantity) : 0;

  const handleManualSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (isOverPo) {
      alert(`Số lượng nhập kho Laser (${manualQty}) sẽ vượt quá số lượng mục tiêu của PO [${selectedPo?.id}] (${selectedPo?.targetQuantity}).\nHiện đã làm: ${selectedPo?.producedQuantity}.\nHệ thống KHÔNG cho phép nhập vượt quá PO!`);
      return;
    }
    if (manualPart && manualQty > 0) {
      try {
        const part = parts.find((p: any) => p.id === manualPart);
        onManualInbound(manualPart, selectedStage, targetLocation, manualQty, selectedPoId);
        setLastScanned({
          partId: getProcessValue(manualPart, part, selectedStage, targetLocation),
          quantity: manualQty,
          partName: getProcessValue(part?.name, part, selectedStage, targetLocation),
          status: 'success',
          isManual: true,
          poId: selectedPoId
        });
        setManualQty(0);
      } catch (err) {
        const part = parts.find((p: any) => p.id === manualPart);
        setLastScanned({
          partId: getProcessValue(manualPart, part, selectedStage, targetLocation),
          quantity: manualQty,
          partName: getProcessValue(part?.name, part, selectedStage, targetLocation),
          status: 'error',
          isManual: true,
          errorMsg: err instanceof Error ? err.message : 'Lỗi không xác định'
        });
      }
    }
  };

  return (
    <motion.div 
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, y: -20 }}
      className="grid grid-cols-1 lg:grid-cols-2 gap-8"
    >
      <div className="bg-white p-10 rounded-3xl border border-gray-200 shadow-xl space-y-8">
        <div className="flex items-center gap-4 mb-2">
          <div className="bg-blue-600 p-3 rounded-xl text-white">
            <Layers size={24} />
          </div>
          <div>
            <h3 className="text-xl font-bold uppercase tracking-tight">Nhập kho Laser</h3>
            <p className="text-sm text-gray-500">Nhập tôn (IN) hoặc kết quả cắt (OUT) thủ công</p>
          </div>
        </div>

        <form onSubmit={handleManualSubmit} className="space-y-8 text-left">
          <div className="grid grid-cols-2 gap-8">
            <div className="space-y-4">
              <label className="text-sm font-bold uppercase tracking-widest opacity-50">Công đoạn:</label>
              <div className="w-full p-5 rounded-xl border-2 border-gray-100 font-bold text-lg bg-gray-50 text-gray-500">
                CẮT LASER
              </div>
            </div>
            <div className="space-y-4">
              <label className="text-sm font-bold uppercase tracking-widest opacity-50">Vị trí nhập:</label>
              <div className="flex bg-gray-100 p-1 rounded-xl h-[68px]">
                <button 
                  type="button"
                  onClick={() => setTargetLocation('IN')}
                  className={cn(
                    "flex-1 rounded-lg font-bold text-sm uppercase transition-all",
                    targetLocation === 'IN' ? "bg-gray-900 text-white shadow-sm" : "text-gray-400"
                  )}
                >
                  Kho IN (Tôn)
                </button>
                <button 
                  type="button"
                  onClick={() => setTargetLocation('OUT')}
                  className={cn(
                    "flex-1 rounded-lg font-bold text-sm uppercase transition-all",
                    targetLocation === 'OUT' ? "bg-[#F27D26] text-white shadow-sm" : "text-gray-400"
                  )}
                >
                  Kho OUT (LK)
                </button>
              </div>
            </div>
          </div>

          {targetLocation === 'OUT' && allAvailablePos.length > 0 && (
            <div className="space-y-4">
              <label className="text-sm font-bold uppercase tracking-widest opacity-50">Chọn Lệnh PO (Tiến độ)</label>
              <SearchableSelect 
                options={allAvailablePos.map(po => ({
                  id: po.id,
                  label: `${po.id.startsWith('REPAIR') ? '🛠️ [BÙ] ' : ''}${po.id} - ${getProcessValue(parts.find((p: any) => p.id === po.partId)?.name || po.partId, parts.find((p: any) => p.id === po.partId), selectedStage, targetLocation)} (${po.producedQuantity}/${po.targetQuantity})`
                }))}
                value={selectedPoId}
                onChange={handlePoChange}
                placeholder="Tìm Lệnh PO..."
              />
            </div>
          )}

          <div className="space-y-4">
            <label className="text-sm font-bold uppercase tracking-widest opacity-50">Chọn {targetLocation === 'IN' ? 'mã tôn' : 'mã linh kiện'}:</label>
            <SearchableSelect 
              options={filteredParts.map((p: any) => ({ 
                id: p.id, 
                label: `${getProcessValue(p.id, p, selectedStage, targetLocation)} - ${getProcessValue(p.name, p, selectedStage, targetLocation)}` 
              }))}
              value={manualPart}
              onChange={handlePartChange}
              placeholder="Tìm kiếm..."
            />
          </div>

          <div className="space-y-4">
            <div className="flex justify-between items-center">
              <label className="text-sm font-bold uppercase tracking-widest opacity-50">Số lượng nhập:</label>
              {manualPart && (
                <span className="text-xs font-mono font-bold text-gray-400 uppercase">
                  Đơn vị: {parts.find((p: any) => p.id === manualPart)?.unit}
                </span>
              )}
            </div>
            <input 
              type="number"
              step="any"
              value={manualQty || ''}
              onChange={(e) => setManualQty(parseFloat(e.target.value) || 0)}
              className="w-full p-5 rounded-xl border-2 border-gray-100 font-mono text-2xl font-bold focus:border-blue-600 outline-none"
              placeholder="Nhập số lượng..."
            />
          </div>

          {isOverPo && (
            <div className="w-full bg-red-50 border-2 border-red-200 p-4 rounded-xl flex items-center gap-3 text-red-700 font-bold">
              <AlertCircle size={24} className="text-red-600 flex-shrink-0" />
              <span>VƯỢT QUÁ PO (+{overAmount}) — KHÔNG ĐƯỢC PHÉP NHẬP!</span>
            </div>
          )}

          {availablePos.length === 0 && targetLocation === 'OUT' ? (
            <div className="w-full bg-red-50 border-2 border-red-200 p-6 rounded-xl flex flex-col items-center gap-3 text-center">
              <AlertCircle size={32} className="text-red-600" />
              <div>
                <p className="font-bold text-red-800 uppercase">Yêu cầu Lệnh sản xuất (PO)</p>
                <p className="text-xs text-red-600">Linh kiện cắt Laser (OUT) phải gắn với một PO đang chạy.</p>
              </div>
            </div>
          ) : (
            <button 
              type="submit"
              disabled={(targetLocation === 'OUT' && !selectedPoId) || isOverPo}
              className={cn(
                "w-full py-6 rounded-xl font-bold uppercase tracking-widest flex items-center justify-center gap-4 active:scale-[0.98] transition-all shadow-lg text-xl",
                ((targetLocation === 'OUT' && !selectedPoId) || isOverPo)
                  ? "bg-gray-300 text-gray-500 cursor-not-allowed"
                  : "bg-blue-600 text-white hover:bg-blue-700 shadow-blue-200"
              )}
            >
              <CheckCircle2 size={28} />
              Xác nhận nhập kho Laser
            </button>
          )}
        </form>
      </div>

      <div className="flex flex-col">
        <AnimatePresence mode="wait">
          {lastScanned ? (
            <motion.div
              key="scan-result"
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              className={cn(
                "p-10 rounded-3xl border-2 shadow-2xl h-full flex flex-col items-center justify-center space-y-6 bg-white",
                lastScanned.status === 'success' ? "border-green-500" : "border-red-500"
              )}
            >
              <div className={cn(
                "w-16 h-16 rounded-full flex items-center justify-center",
                lastScanned.status === 'success' ? "bg-green-100 text-green-600" : "bg-red-100 text-red-600"
              )}>
                {lastScanned.status === 'success' ? <CheckCircle2 size={32} /> : <AlertCircle size={32} />}
              </div>
              <div className="text-center space-y-2">
                <h3 className="text-xl font-bold">
                  {lastScanned.status === 'success' ? 'Nhập kho thành công!' : 'Lỗi nhập kho!'}
                </h3>
                <p className="text-gray-500 text-sm">
                  {lastScanned.status === 'success' 
                    ? 'Thông tin linh kiện vừa nhập tay:' 
                    : lastScanned.errorMsg}
                </p>
              </div>
              
              {lastScanned.partId && (
                <div className="w-full bg-gray-50 p-8 rounded-2xl space-y-6">
                  <div className="flex justify-between items-center border-b border-gray-200 pb-4">
                    <span className="text-sm font-mono uppercase opacity-50">Mã LK</span>
                    <span className="font-bold text-xl">{lastScanned.partId}</span>
                  </div>
                  <div className="flex justify-between items-center border-b border-gray-200 pb-4">
                    <span className="text-sm font-mono uppercase opacity-50">Linh kiện</span>
                    <span className="font-bold text-xl">{lastScanned.partName}</span>
                  </div>
                  <div className="flex justify-between items-center border-b border-gray-200 pb-4">
                    <span className="text-sm font-mono uppercase opacity-50">Số lượng</span>
                    <span className="font-bold text-2xl">{lastScanned.quantity}</span>
                  </div>
                  <div className="flex justify-between items-center">
                    <span className="text-sm font-mono uppercase opacity-50">Đích</span>
                    <div className="flex items-center gap-3 font-bold text-[#F27D26] text-xl">
                      <span>CẮT LASER</span>
                      <ArrowRight size={20} />
                      <span className={cn(
                        "px-3 py-1 rounded text-xs",
                        targetLocation === 'IN' ? "bg-blue-100 text-blue-700" : "bg-orange-100 text-orange-700"
                      )}>
                        KHO_{targetLocation}
                      </span>
                    </div>
                  </div>
                  {lastScanned.poId && (
                    <div className="flex justify-between items-center pt-4 border-t border-gray-200">
                      <span className="text-xs font-mono uppercase opacity-50">Lệnh PO</span>
                      <span className="text-sm font-bold text-blue-600">{lastScanned.poId}</span>
                    </div>
                  )}
                </div>
              )}

              <div className="flex gap-4 w-full">
                <button 
                  onClick={() => setLastScanned(null)}
                  className="flex-1 py-4 text-sm font-bold uppercase tracking-widest text-gray-400 hover:text-[#141414] transition-colors border-2 border-gray-100 rounded-xl"
                >
                  Đóng
                </button>
                {lastScanned.status === 'success' && (
                  <button 
                    onClick={() => setDefectModal({ partId: lastScanned.partId, stageId: selectedStage, poId: lastScanned.poId })}
                    className="flex-1 py-4 text-sm font-bold uppercase tracking-widest text-red-600 hover:bg-red-50 transition-colors border-2 border-red-100 rounded-xl flex items-center justify-center gap-2"
                  >
                    <AlertCircle size={18} />
                    Báo lỗi (NG)
                  </button>
                )}
              </div>
            </motion.div>
          ) : (
            <div className="bg-gray-100 rounded-3xl border-2 border-dashed border-gray-300 h-full flex flex-col items-center justify-center p-12 text-center text-gray-400">
              <Package size={64} strokeWidth={1} />
              <p className="mt-6 font-medium">Chi tiết nhập kho Laser<br />sẽ hiển thị tại đây.</p>
            </div>
          )}
        </AnimatePresence>
      </div>
    </motion.div>
  );
}

function OddProductionOrderView({ 
  parts, 
  onUpdate, 
  productionOrders, 
  currentUser 
}: { 
  parts: Part[]; 
  onUpdate: () => void; 
  productionOrders: ProductionOrder[]; 
  currentUser?: UserAccount | null; 
  key?: string; 
}) {
  // Chỉ chuyên biệt cho công đoạn Cắt Laser
  const selectedStage: StageId = 'LASER';
  const [selectedPartId, setSelectedPartId] = useState<string>('');
  const [quantity, setQuantity] = useState<number>(0);
  const [isCreating, setIsCreating] = useState(false);

  // Quyền duyệt: Admin hoặc tài khoản được Admin phân quyền duyệt PO dư lẻ
  const canApprove = currentUser?.role === 'ADMIN' || currentUser?.canApproveOddPO === true;

  // Filters for PO list: Đặt 'PENDING_APPROVAL' (Chờ duyệt) lên hàng đầu theo yêu cầu
  const [statusFilter, setStatusFilter] = useState<'PENDING_APPROVAL' | 'PENDING' | 'IN_PROGRESS' | 'COMPLETED' | 'ALL'>('PENDING_APPROVAL');
  const [searchTerm, setSearchTerm] = useState('');

  // Modals for editing/deleting PO
  const [showDeleteModal, setShowDeleteModal] = useState<string | null>(null);
  const [showEditModal, setShowEditModal] = useState<{ id: string, qty: number } | null>(null);
  const [password, setPassword] = useState("");

  // Chỉ hiển thị linh kiện Bậc 2 gia công Cắt Laser theo đúng yêu cầu
  const availableLaserParts = useMemo(() => {
    return parts.filter(p => {
      const lvl = p.level ?? 2;
      return lvl === 2 && !p.skipLaser;
    });
  }, [parts]);

  const selectedPart = useMemo(() => {
    return parts.find(p => p.id === selectedPartId);
  }, [parts, selectedPartId]);

  // Đếm số lượng PO đang chờ duyệt
  const pendingApprovalCount = useMemo(() => {
    return (productionOrders || []).filter(o => 
      (o.isOddPO || o.id.startsWith('PO-LE') || (o.note && (o.note.includes('Dư Lẻ') || o.note.includes('Tồn Lẻ')))) && 
      o.status === 'PENDING_APPROVAL'
    ).length;
  }, [productionOrders]);

  // Tạo PO Dư Lẻ cho công đoạn Cắt Laser
  const handleCreateOddPO = (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedPartId) {
      alert("Vui lòng chọn linh kiện Bậc 2 cần tạo PO Cắt Laser!");
      return;
    }
    if (!quantity || quantity <= 0) {
      alert("Vui lòng nhập số lượng cần thực hiện lớn hơn 0!");
      return;
    }

    setIsCreating(true);
    try {
      const newPo = storageService.createSingleOddPO(selectedPartId, 'LASER', quantity, {
        username: currentUser?.username || 'khach',
        fullName: currentUser?.fullName || 'Người dùng',
        role: currentUser?.role,
        canApprove
      });
      onUpdate();
      const partName = selectedPart?.name || selectedPartId;
      if (canApprove) {
        alert(`Đã tạo và phê duyệt thành công PO Dư Lẻ [${newPo.id}] cho linh kiện "${partName}" với số lượng ${quantity} ${selectedPart?.unit || 'cái'}!\n\nLệnh đã sẵn sàng để công đoạn Cắt Laser thực hiện.`);
      } else {
        alert(`Đã gửi yêu cầu tạo PO Dư Lẻ [${newPo.id}] cho linh kiện "${partName}" với số lượng ${quantity} ${selectedPart?.unit || 'cái'}!\n\nLưu ý: Lệnh đang ở trạng thái [Chờ Admin duyệt]. Vui lòng báo Quản trị viên (Admin) phê duyệt trước khi bắt đầu sản xuất.`);
      }
      setQuantity(0);
      setSelectedPartId('');
    } catch (err: any) {
      alert(err?.message || "Có lỗi xảy ra khi tạo PO Dư Lẻ!");
    } finally {
      setIsCreating(false);
    }
  };

  const handleApproveOddPO = (poId: string) => {
    if (!canApprove) {
      alert('Chỉ Quản trị viên hoặc tài khoản được Admin phân quyền mới có thể phê duyệt lệnh này!');
      return;
    }
    const res = storageService.approveOddPO(poId, {
      username: currentUser?.username || 'admin',
      fullName: currentUser?.fullName || 'Quản trị viên'
    });
    if (res.success) {
      alert(`Đã phê duyệt lệnh [${poId}] thành công! Lệnh đã sẵn sàng để công đoạn Cắt Laser gia công.`);
      onUpdate();
    } else {
      alert(res.error || 'Có lỗi xảy ra');
    }
  };

  const handleRejectOddPO = (poId: string) => {
    if (!canApprove) {
      alert('Chỉ Quản trị viên hoặc tài khoản được Admin phân quyền mới có thể từ chối lệnh này!');
      return;
    }
    if (confirm(`Bạn có chắc muốn từ chối yêu cầu PO Dư Lẻ [${poId}] này không?`)) {
      storageService.rejectOddPO(poId);
      onUpdate();
    }
  };

  const handleDeletePO = () => {
    if (!canApprove) {
      alert('Chỉ Quản trị viên (Admin) hoặc tài khoản được phân quyền mới có thể xóa PO Dư Lẻ!');
      return;
    }
    const isAuthorized = currentUser?.role === 'ADMIN' || password === 'admin123' || (currentUser?.password && password === currentUser.password);
    if (isAuthorized) {
      if (showDeleteModal) {
        storageService.deletePO(showDeleteModal);
        onUpdate();
        setShowDeleteModal(null);
        setPassword("");
      }
    } else {
      alert('Mật khẩu xác nhận không chính xác!');
    }
  };

  const handleUpdatePoQty = () => {
    if (!canApprove) {
      alert('Chỉ Quản trị viên (Admin) hoặc tài khoản được phân quyền mới có thể chỉnh sửa số lượng PO Dư Lẻ!');
      return;
    }
    const isAuthorized = currentUser?.role === 'ADMIN' || password === 'admin123' || (currentUser?.password && password === currentUser.password);
    if (isAuthorized) {
      if (showEditModal) {
        storageService.updateSubPoQty(showEditModal.id, showEditModal.qty);
        onUpdate();
        setShowEditModal(null);
        setPassword("");
      }
    } else {
      alert('Mật khẩu xác nhận không chính xác!');
    }
  };

  const handleCompletePO = (poId: string) => {
    storageService.forceCompleteMasterPO(poId);
    onUpdate();
  };

  // Danh sách PO Dư Lẻ đã tạo
  const oddOrdersList = useMemo(() => {
    const searchLower = searchTerm.toLowerCase();
    const partsMap = new Map(parts.map(p => [p.id, p]));

    return (productionOrders || [])
      .filter(o => o.isOddPO || o.id.startsWith('PO-LE') || (o.note && (o.note.includes('Dư Lẻ') || o.note.includes('Tồn Lẻ'))))
      .filter(po => {
        if (statusFilter !== 'ALL' && po.status !== statusFilter) return false;
        
        const part = partsMap.get(po.partId);
        const partName = part?.name || "";
        return (
          po.id.toLowerCase().includes(searchLower) ||
          po.partId.toLowerCase().includes(searchLower) ||
          partName.toLowerCase().includes(searchLower)
        );
      });
  }, [productionOrders, parts, statusFilter, searchTerm]);

  return (
    <div className="space-y-8">
      {/* Create Odd PO Form Card - Chuyên biệt Cắt Laser */}
      <div className="bg-white p-8 rounded-3xl border border-gray-200 shadow-sm space-y-8">
        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 pb-6 border-b border-gray-100">
          <div>
            <h2 className="text-2xl font-bold flex items-center gap-3 text-amber-700">
              <Layers className="text-amber-600" size={28} />
              Tạo PO Dư Lẻ - Cắt Laser
            </h2>
            <p className="text-sm text-gray-500 mt-1">
              Chuyên biệt cho công đoạn Cắt Laser. Chỉ xuất hiện linh kiện Bậc 2. Chỉ tạo PO duy nhất tại công đoạn Cắt Laser. Sau khi cắt và xuất kho OUT, linh kiện có thể quét nhập vào công đoạn tiếp theo hoặc dùng cho các PO khác cùng mã.
            </p>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            {canApprove ? (
              <div className="bg-emerald-50 border border-emerald-200 px-3.5 py-2 rounded-2xl flex items-center gap-2 text-xs font-bold text-emerald-800 shadow-xs">
                <ShieldCheck size={16} className="text-emerald-600" />
                <span>Quyền: Duyệt lệnh trực tiếp (Admin)</span>
              </div>
            ) : (
              <div className="bg-amber-50 border border-amber-200 px-3.5 py-2 rounded-2xl flex items-center gap-2 text-xs font-bold text-amber-900 shadow-xs">
                <Clock size={16} className="text-amber-600" />
                <span>Cần Admin phê duyệt sau khi tạo</span>
              </div>
            )}
            <div className="bg-amber-100/70 border border-amber-300 px-3.5 py-2 rounded-2xl flex items-center gap-1.5 text-xs font-bold text-amber-900 shadow-xs">
              <span>⚡</span>
              <span>Công đoạn: Cắt Laser (Bậc 2)</span>
            </div>
          </div>
        </div>

        <form onSubmit={handleCreateOddPO} className="space-y-6">
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 items-start">
            {/* BƯỚC 1: CHỌN LINH KIỆN CẮT LASER (BẬC 2) */}
            <div className="lg:col-span-2 space-y-3">
              <div className="flex justify-between items-center">
                <label className="text-sm font-bold uppercase tracking-wider text-gray-700 flex items-center gap-2">
                  <span className="w-6 h-6 rounded-full bg-amber-600 text-white flex items-center justify-center text-xs">1</span>
                  <span>Chọn Linh Kiện Cắt Laser (Bậc 2)</span>
                </label>
                <span className="text-xs text-blue-600 font-mono font-semibold">
                  ({availableLaserParts.length} linh kiện Bậc 2)
                </span>
              </div>

              <SearchableSelect 
                options={availableLaserParts.map(p => ({ 
                  id: p.id, 
                  label: `${p.id} - ${p.name} (Bậc ${p.level ?? 2}, ĐVT: ${p.unit})` 
                }))}
                value={selectedPartId}
                onChange={(val) => setSelectedPartId(val)}
                placeholder="Tìm chọn mã linh kiện bậc 2 Cắt Laser..."
              />

              {/* Thông tin linh kiện đã chọn */}
              {selectedPart && (
                <div className="p-4 bg-gray-50 rounded-2xl border border-gray-200 grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
                  <div>
                    <span className="text-gray-400 block font-medium">Mã LK:</span>
                    <strong className="font-mono text-gray-900 text-sm">{selectedPart.id}</strong>
                  </div>
                  <div>
                    <span className="text-gray-400 block font-medium">Tên linh kiện:</span>
                    <strong className="text-gray-900">{selectedPart.name}</strong>
                  </div>
                  <div>
                    <span className="text-gray-400 block font-medium">Bậc linh kiện:</span>
                    <span className="inline-block bg-blue-100 text-blue-800 px-2 py-0.5 rounded font-bold mt-0.5">
                      Bậc {selectedPart.level ?? 2} (Cắt Laser)
                    </span>
                  </div>
                  <div>
                    <span className="text-gray-400 block font-medium">Đơn vị tính:</span>
                    <strong className="text-gray-800">{selectedPart.unit}</strong>
                  </div>
                </div>
              )}
            </div>

            {/* BƯỚC 2: SỐ LƯỢNG & NÚT TẠO */}
            <div className="space-y-3">
              <label className="text-sm font-bold uppercase tracking-wider text-gray-700 flex items-center gap-2">
                <span className="w-6 h-6 rounded-full bg-amber-600 text-white flex items-center justify-center text-xs">2</span>
                <span>Số Lượng Cần Thực Hiện</span>
              </label>

              <div className="space-y-3">
                <div className="relative">
                  <input 
                    type="number" 
                    min="1"
                    step="1"
                    value={quantity || ''} 
                    onChange={(e) => setQuantity(parseInt(e.target.value) || 0)}
                    placeholder="Nhập số lượng..."
                    className="w-full p-4 rounded-2xl border-2 border-gray-200 font-mono font-bold text-xl focus:border-amber-500 outline-none transition-all pr-16"
                    required
                  />
                  {selectedPart && (
                    <span className="absolute right-4 top-1/2 -translate-y-1/2 text-xs font-bold text-gray-400 uppercase">
                      {selectedPart.unit}
                    </span>
                  )}
                </div>

                <button
                  type="submit"
                  disabled={!selectedPartId || quantity <= 0 || isCreating}
                  className="w-full bg-amber-600 hover:bg-amber-700 disabled:opacity-50 text-white p-4 rounded-2xl font-bold uppercase tracking-wider flex items-center justify-center gap-2 cursor-pointer transition-all shadow-lg shadow-amber-200 active:scale-95 text-sm"
                >
                  <Plus size={18} />
                  <span>
                    {isCreating 
                      ? 'Đang xử lý...' 
                      : canApprove 
                        ? 'Tạo & Phê Duyệt PO Dư Lẻ (Cắt Laser)' 
                        : 'Tạo Yêu Cầu PO Dư Lẻ (Chờ Admin Duyệt)'}
                  </span>
                </button>
                {canApprove ? (
                  <p className="text-[11px] text-emerald-700 font-medium text-center">
                    ✓ Bạn có quyền quản trị/duyệt lệnh. Lệnh tạo ra sẽ được phê duyệt ngay và sẵn sàng sản xuất.
                  </p>
                ) : (
                  <p className="text-[11px] text-amber-800 font-medium text-center">
                    ⚠️ Tài khoản thông thường chỉ được phép tạo đề xuất. Lệnh sẽ ở trạng thái <strong>Chờ Admin duyệt</strong>.
                  </p>
                )}
              </div>
            </div>
          </div>
        </form>
      </div>

      {/* DANH SÁCH CÁC PO DƯ LẺ ĐÃ TẠO */}
      <div className="bg-white rounded-3xl border border-gray-200 shadow-sm overflow-hidden">
        <div className="p-8 border-b border-gray-200 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
          <div>
            <h3 className="text-xl font-bold flex items-center gap-3">
              <ClipboardList className="text-amber-600" />
              <span>Danh Sách PO Dư Lẻ Cắt Laser ({oddOrdersList.length})</span>
            </h3>
            <p className="text-xs text-gray-400 mt-1">
              Theo dõi tiến độ gia công Cắt Laser và xuất kho OUT các PO dư lẻ
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-3 w-full md:w-auto">
            {/* Status Filter: Đưa Chờ duyệt lên trước, Tất cả ở cuối cùng */}
            <div className="flex bg-gray-100 p-1 rounded-xl">
              <button
                onClick={() => setStatusFilter('PENDING_APPROVAL')}
                className={cn(
                  "px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5",
                  statusFilter === 'PENDING_APPROVAL' ? "bg-amber-500 text-white shadow-sm" : "text-amber-700 hover:text-amber-900"
                )}
              >
                <span>Chờ duyệt</span>
                {pendingApprovalCount > 0 && (
                  <span className="px-1.5 py-0.5 rounded-full text-[10px] bg-white text-amber-600 font-black animate-pulse">
                    {pendingApprovalCount}
                  </span>
                )}
              </button>
              <button
                onClick={() => setStatusFilter('PENDING')}
                className={cn(
                  "px-3 py-1.5 rounded-lg text-xs font-bold transition-all",
                  statusFilter === 'PENDING' ? "bg-white text-gray-900 shadow-sm" : "text-gray-500 hover:text-gray-700"
                )}
              >
                Chờ SX
              </button>
              <button
                onClick={() => setStatusFilter('IN_PROGRESS')}
                className={cn(
                  "px-3 py-1.5 rounded-lg text-xs font-bold transition-all",
                  statusFilter === 'IN_PROGRESS' ? "bg-white text-gray-900 shadow-sm" : "text-gray-500 hover:text-gray-700"
                )}
              >
                Đang làm
              </button>
              <button
                onClick={() => setStatusFilter('COMPLETED')}
                className={cn(
                  "px-3 py-1.5 rounded-lg text-xs font-bold transition-all",
                  statusFilter === 'COMPLETED' ? "bg-white text-gray-900 shadow-sm" : "text-gray-500 hover:text-gray-700"
                )}
              >
                Hoàn thành
              </button>
              <button
                onClick={() => setStatusFilter('ALL')}
                className={cn(
                  "px-3 py-1.5 rounded-lg text-xs font-bold transition-all",
                  statusFilter === 'ALL' ? "bg-white text-gray-900 shadow-sm" : "text-gray-500 hover:text-gray-700"
                )}
              >
                Tất cả
              </button>
            </div>

            {/* Search */}
            <div className="relative flex-1 md:w-56">
              <input 
                type="text" 
                placeholder="Tìm mã PO, linh kiện..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full p-2.5 pl-9 rounded-xl border border-gray-200 text-xs focus:border-amber-500 outline-none"
              />
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={14} />
            </div>
          </div>
        </div>

        {oddOrdersList.length === 0 ? (
          <div className="p-16 text-center text-gray-400 flex flex-col items-center justify-center">
            <Layers size={48} strokeWidth={1} className="mb-4 text-gray-300" />
            <p className="font-medium text-gray-600">Chưa có PO Dư Lẻ nào.</p>
            <p className="text-xs text-gray-400 mt-1">Sử dụng biểu mẫu phía trên để tạo PO dư lẻ cho công đoạn cần làm.</p>
          </div>
        ) : (
          <div className="divide-y divide-gray-100">
            {oddOrdersList.map(po => {
              const partObj = parts.find(p => p.id === po.partId);
              const stageName = STAGES.find(s => s.id === po.stageId)?.name || (po.stageId ? po.stageId : 'Cắt Laser');
              const stageIcon = po.stageId === 'LASER' ? '⚡' : '⚙️';
              const produced = po.producedQuantity || 0;
              const exported = po.exportedQuantity || 0;
              const target = po.targetQuantity || 1;
              const pct = Math.min(100, Math.round((produced / target) * 100));

              return (
                <div key={po.id} className={cn(
                  "p-6 transition-all flex flex-col lg:flex-row justify-between items-start lg:items-center gap-5 hover:shadow-sm",
                  po.status === 'PENDING_APPROVAL' ? "bg-amber-50/50 hover:bg-amber-50/80 border-l-4 border-amber-500" : "hover:bg-gray-50/60"
                )}>
                  {/* Left: Info & Prominent Target Quantity Badge */}
                  <div className="space-y-2 flex-1">
                    <div className="flex items-center gap-3 flex-wrap">
                      <span className="font-mono font-black text-amber-900 text-lg tracking-tight">
                        {po.id}
                      </span>
                      <span className="inline-flex items-center gap-1.5 bg-amber-100/90 text-amber-950 px-3 py-1 rounded-full text-xs font-bold border border-amber-200">
                        <span>{stageIcon}</span>
                        <span>{stageName}</span>
                      </span>
                      {partObj?.level && (
                        <span className="bg-blue-100 text-blue-900 px-2.5 py-1 rounded-md text-xs font-bold">
                          Bậc {partObj.level}
                        </span>
                      )}
                      <span className={cn(
                        "px-3 py-1 rounded-full text-xs font-bold",
                        po.status === 'COMPLETED' ? "bg-green-100 text-green-800" :
                        po.status === 'IN_PROGRESS' ? "bg-blue-100 text-blue-800" :
                        po.status === 'PENDING_APPROVAL' ? "bg-amber-100 text-amber-900 border border-amber-300 animate-pulse font-black" :
                        po.status === 'REJECTED' ? "bg-rose-100 text-rose-800 border border-rose-200" :
                        "bg-gray-100 text-gray-700"
                      )}>
                        {po.status === 'COMPLETED' ? 'Hoàn thành' :
                         po.status === 'IN_PROGRESS' ? 'Đang làm' :
                         po.status === 'PENDING_APPROVAL' ? '⏳ Chờ Admin duyệt' :
                         po.status === 'REJECTED' ? 'Đã từ chối' : 'Chờ sản xuất'}
                      </span>
                    </div>

                    {/* Prominent Product Info & Quantity Display */}
                    <div className="flex items-center gap-3 flex-wrap pt-1">
                      {/* Big prominent quantity badge */}
                      <div className="inline-flex items-baseline gap-1.5 bg-amber-100/80 border-2 border-amber-300/80 px-3 py-1 rounded-xl shadow-xs">
                        <span className="text-xs uppercase font-extrabold text-amber-900">Mục tiêu:</span>
                        <span className="font-mono font-black text-2xl text-amber-950 leading-none">{target.toLocaleString('vi-VN')}</span>
                        <span className="text-xs font-bold text-amber-900 uppercase">{partObj?.unit || 'cái'}</span>
                      </div>

                      <div className="text-base text-gray-800 flex items-center gap-2 flex-wrap">
                        <span className="font-mono font-black text-gray-950 bg-gray-100 px-2.5 py-1 rounded-lg border border-gray-200">
                          {po.partId}
                        </span>
                        <span className="font-bold text-gray-900 text-base">
                          {partObj?.name || po.partId}
                        </span>
                      </div>
                    </div>

                    <div className="text-xs text-gray-500 flex items-center gap-3 pt-1 font-mono flex-wrap">
                      <span>Tạo lúc: {format(po.createdAt || Date.now(), 'dd/MM/yyyy HH:mm')}</span>
                      {po.createdBy && (
                        <span className="text-amber-950 font-sans font-bold bg-amber-100/90 px-2.5 py-0.5 rounded border border-amber-200">
                          Tạo bởi: @{po.createdBy} {po.createdByName ? `(${po.createdByName})` : ''}
                        </span>
                      )}
                      {po.approvedBy && (
                        <span className="text-emerald-950 font-sans font-bold bg-emerald-100 px-2.5 py-0.5 rounded border border-emerald-300">
                          Duyệt bởi: @{po.approvedBy}
                        </span>
                      )}
                      {po.completedAt && (
                        <span className="text-green-700 font-bold">Hoàn tất: {format(po.completedAt, 'dd/MM/yyyy HH:mm')}</span>
                      )}
                    </div>
                  </div>

                  {/* Right: Progress & Actions */}
                  <div className="flex items-center gap-6 w-full lg:w-auto justify-between lg:justify-end">
                    <div className="text-right">
                      <div className="text-xs text-gray-500 uppercase font-bold tracking-wide">Tiến độ SX / Xuất OUT</div>
                      <div className="font-mono font-black text-base text-gray-900 mt-0.5">
                        <span className="text-emerald-700">SX: {produced}/{target}</span>
                        <span className="text-gray-300 mx-1.5">•</span>
                        <span className="text-blue-700">Xuất: {exported}/{target}</span>
                        <span className="ml-1.5 text-xs font-bold text-amber-700">({pct}%)</span>
                      </div>
                      <div className="w-44 bg-gray-200 h-2.5 rounded-full overflow-hidden mt-1.5 shadow-inner">
                        <div 
                          className={cn(
                            "h-full rounded-full transition-all",
                            po.status === 'COMPLETED' ? "bg-green-600" : 
                            po.status === 'PENDING_APPROVAL' ? "bg-amber-400" : "bg-amber-600"
                          )}
                          style={{ width: `${pct}%` }}
                        />
                      </div>
                    </div>

                    <div className="flex items-center gap-2">
                      {/* Approval buttons when PENDING_APPROVAL */}
                      {po.status === 'PENDING_APPROVAL' ? (
                        canApprove ? (
                          <div className="flex items-center gap-2">
                            <button
                              onClick={() => handleApproveOddPO(po.id)}
                              className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-black transition-all shadow-md shadow-emerald-600/20 active:scale-95 flex items-center gap-1.5 cursor-pointer"
                              title="Phê duyệt lệnh PO Dư Lẻ này"
                            >
                              <Check size={15} strokeWidth={3} />
                              <span>Duyệt lệnh</span>
                            </button>
                            <button
                              onClick={() => handleRejectOddPO(po.id)}
                              className="px-3.5 py-2 bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 rounded-xl text-xs font-bold transition-all cursor-pointer"
                              title="Từ chối yêu cầu PO Dư Lẻ này"
                            >
                              <X size={15} />
                              <span>Từ chối</span>
                            </button>
                          </div>
                        ) : (
                          <span className="text-xs font-bold text-amber-800 bg-amber-50 px-3 py-1.5 rounded-xl border border-amber-200">
                            ⏳ Chờ Admin duyệt
                          </span>
                        )
                      ) : (
                        <>
                          <button
                            onClick={() => setShowEditModal({ id: po.id, qty: po.targetQuantity })}
                            className="p-2 text-gray-400 hover:text-blue-600 hover:bg-blue-50 rounded-xl transition-colors"
                            title="Sửa số lượng PO"
                          >
                            <Edit2 size={16} />
                          </button>

                          {po.status !== 'COMPLETED' && po.status !== 'REJECTED' && (
                            <button
                              onClick={() => handleCompletePO(po.id)}
                              className="px-3 py-1.5 bg-green-50 hover:bg-green-100 text-green-700 rounded-xl text-xs font-bold transition-colors"
                              title="Đánh dấu hoàn thành PO này"
                            >
                              Hoàn tất
                            </button>
                          )}
                        </>
                      )}

                      <button
                        onClick={() => setShowDeleteModal(po.id)}
                        className="p-2 text-red-400 hover:text-red-600 hover:bg-red-50 rounded-xl transition-colors"
                        title="Xóa PO Dư Lẻ"
                      >
                        <Trash2 size={16} />
                      </button>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {/* Footer guidance note */}
        <div className="p-4 bg-amber-50/60 border-t border-amber-100 flex items-center gap-2 text-xs text-amber-900">
          <AlertCircle size={16} className="text-amber-600 flex-shrink-0" />
          <span>
            <strong>Gợi ý luân chuyển:</strong> Sau khi công đoạn gia công xong, vào mục <strong>Xuất kho & In nhãn QR</strong> để ghi nhận xuất kho OUT. Nhãn QR in ra có thể quét nhập vào kho IN của công đoạn kế tiếp, hoặc các PO khác cùng mã linh kiện có thể lấy sử dụng trực tiếp từ tồn kho OUT này.
          </span>
        </div>
      </div>

      {/* Delete Confirmation Modal */}
      {showDeleteModal && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white p-6 rounded-2xl max-w-md w-full space-y-4">
            <h3 className="font-bold text-lg text-red-600 flex items-center gap-2">
              <AlertCircle size={20} />
              Xác nhận xóa PO Dư Lẻ
            </h3>
            <p className="text-sm text-gray-600">
              Bạn có chắc chắn muốn xóa PO <strong>{showDeleteModal}</strong>?
            </p>
            <input 
              type="password"
              placeholder="Nhập mật khẩu Admin (admin123)..."
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              className="w-full p-3 border rounded-xl outline-none text-sm"
            />
            <div className="flex gap-3 justify-end">
              <button 
                onClick={() => { setShowDeleteModal(null); setPassword(""); }}
                className="px-4 py-2 border rounded-xl text-sm font-bold text-gray-600"
              >
                Hủy
              </button>
              <button 
                onClick={handleDeletePO}
                className="px-4 py-2 bg-red-600 text-white rounded-xl text-sm font-bold"
              >
                Xác nhận xóa
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Edit Sub-PO Quantity Modal */}
      {showEditModal && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white p-6 rounded-2xl max-w-md w-full space-y-4">
            <h3 className="font-bold text-lg">Chỉnh sửa số lượng PO Dư Lẻ</h3>
            <p className="text-xs text-gray-500 font-mono">{showEditModal.id}</p>
            <input 
              type="number"
              value={showEditModal.qty}
              onChange={(e) => setShowEditModal({ ...showEditModal, qty: parseFloat(e.target.value) || 0 })}
              className="w-full p-3 border rounded-xl outline-none font-mono font-bold text-lg"
            />
            <input 
              type="password"
              placeholder="Mật khẩu Admin (admin123)..."
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              className="w-full p-3 border rounded-xl outline-none text-sm"
            />
            <div className="flex gap-3 justify-end">
              <button 
                onClick={() => { setShowEditModal(null); setPassword(""); }}
                className="px-4 py-2 border rounded-xl text-sm font-bold text-gray-600"
              >
                Hủy
              </button>
              <button 
                onClick={handleUpdatePoQty}
                className="px-4 py-2 bg-blue-600 text-white rounded-xl text-sm font-bold"
              >
                Lưu
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function GlazingComponentPrintCard({ norm, plan, idx, onPrint, onRecordProduction, onReportDefect }: any) {
  const produced = (plan.producedQuantities || {})[norm.id] || 0;
  const printed = (plan.printedQuantities || {})[norm.id] || 0;
  const total = plan.targetQuantity;
  const remainingProduce = Math.max(0, total - produced);
  const remainingPrint = Math.max(0, total - printed);

  // Bắt buộc để trống, không được đề xuất số sẵn (người dùng phải bấm vào tự điền)
  const [producedInputQty, setProducedInputQty] = useState<number | ''>('');
  const [printInputQty, setPrintInputQty] = useState<number | ''>('');
  const [timeSlot, setTimeSlot] = useState<string>('');

  // Luôn reset về trống khi thay đổi kế hoạch hoặc linh kiện
  useEffect(() => {
    setProducedInputQty('');
    setPrintInputQty('');
  }, [plan.id, norm.id]);

  const timeSlots = [
    { label: '6h-8h', start: 6, end: 8 },
    { label: '8h-10h', start: 8, end: 10 },
    { label: '10h-12h', start: 10, end: 12 },
    { label: '13h-15h', start: 13, end: 15 },
    { label: '15h-17h', start: 15, end: 17 },
    { label: '17h-19h', start: 17, end: 19 },
    { label: '19h-20h', start: 19, end: 20 },
  ];

  const progress = Math.min(100, (produced / total) * 100);

  const numProduced = typeof producedInputQty === 'number' ? producedInputQty : 0;
  const isProduceDisabled = produced >= total || producedInputQty === '' || numProduced <= 0 || numProduced > remainingProduce;

  const numPrint = typeof printInputQty === 'number' ? printInputQty : 0;
  const isPrintDisabled = printed >= total || printInputQty === '' || numPrint <= 0 || numPrint > remainingPrint;

  return (
    <div key={`${plan.id}-${norm.id}-${idx}`} className="bg-white p-3.5 sm:p-6 rounded-2xl sm:rounded-3xl shadow-sm sm:shadow-xl border border-gray-100 flex flex-col gap-3.5 sm:gap-6 group hover:shadow-md transition-all">
      <div className="flex justify-between items-start gap-2">
        <div className="space-y-0.5 sm:space-y-1 min-w-0">
          <span className="text-[10px] font-bold text-blue-500 uppercase tracking-wider">{norm.id}</span>
          <h5 className="font-bold text-gray-900 leading-tight text-sm sm:text-base md:text-lg truncate">{norm.partName}</h5>
        </div>
        <div className="flex flex-col items-end gap-1.5 shrink-0">
          <div className="bg-blue-50 px-2 sm:px-3 py-0.5 sm:py-1 rounded-full text-[10px] font-bold text-blue-600 uppercase">
            Norm: {norm.norm}
          </div>
          <button
            type="button"
            onClick={() => onReportDefect && onReportDefect()}
            className="text-[10px] font-bold text-red-500 hover:text-red-700 bg-red-50 px-2 py-0.5 sm:py-1 border border-red-100 flex items-center gap-1 transition-all rounded-lg shadow-2xs cursor-pointer"
          >
            <AlertCircle size={11} /> BÁO LỖI
          </button>
        </div>
      </div>

      <div className="space-y-1.5 sm:space-y-2">
        <div className="flex justify-between text-xs font-bold uppercase">
          <span className="text-gray-400">Tiến độ Kế hoạch:</span>
          <span className={cn(progress >= 100 ? "text-green-600" : "text-blue-600")}>
            {produced.toLocaleString()} / {total.toLocaleString()}
          </span>
        </div>
        <div className="h-2 sm:h-3 bg-gray-100 rounded-full overflow-hidden p-0.5 border border-gray-50">
          <motion.div 
            initial={{ width: 0 }}
            animate={{ width: `${progress}%` }}
            className={cn(
              "h-full rounded-full shadow-inner",
              progress >= 100 ? "bg-green-500" : "bg-blue-500"
            )}
          />
        </div>
      </div>

      <div className="space-y-3 sm:space-y-4 bg-gray-50/80 p-3 sm:p-4 rounded-xl sm:rounded-2xl border border-gray-100">
          <div className="flex flex-col gap-2.5 sm:gap-3">
              <div className="grid grid-cols-2 gap-2 sm:gap-4">
                  <div>
                    <span className="text-[10px] font-bold text-gray-400 uppercase block mb-1">Khung giờ</span>
                    <select
                      value={timeSlot}
                      onChange={(e) => setTimeSlot(e.target.value)}
                      className="w-full bg-white border border-gray-200 rounded-xl px-2 py-1.5 sm:py-2 font-bold text-gray-700 outline-none transition-all text-xs cursor-pointer"
                    >
                      <option value="">-- Hiện tại --</option>
                      {timeSlots.map(s => <option key={s.label} value={s.label}>{s.label}</option>)}
                    </select>
                  </div>
                  <div>
                    <span className="text-[10px] font-bold text-gray-700 uppercase block mb-1">
                      SL Đã Làm <span className="text-red-500">*</span>
                    </span>
                    <input 
                      type="number"
                      min="1"
                      max={remainingProduce}
                      value={producedInputQty === '' ? '' : producedInputQty}
                      onChange={(e) => {
                          const raw = e.target.value;
                          if (raw === '') {
                            setProducedInputQty('');
                          } else {
                            const val = parseInt(raw);
                            setProducedInputQty(isNaN(val) ? '' : val);
                          }
                      }}
                      placeholder="Bắt buộc điền..."
                      autoComplete="off"
                      className="w-full bg-white border border-gray-300 focus:border-blue-500 rounded-xl px-2.5 py-1.5 sm:py-2 font-black text-blue-700 outline-none transition-all text-xs placeholder:font-normal placeholder:text-gray-400 placeholder:italic"
                    />
                  </div>
              </div>
              <button 
                disabled={isProduceDisabled}
                onClick={() => {
                  if (isProduceDisabled) return;
                  let timestamp = Date.now();
                  if (timeSlot) {
                      const slot = timeSlots.find(s => s.label === timeSlot);
                      if (slot) {
                          const d = new Date();
                          d.setHours(slot.start + 1, 0, 0, 0);
                          timestamp = d.getTime();
                      }
                  }
                  onRecordProduction(numProduced, timestamp);
                  // Không để số sẵn sau khi lưu: reset về trống
                  setProducedInputQty('');
                }}
                className={cn(
                  "h-9 sm:h-10 px-4 rounded-xl flex items-center justify-center gap-1.5 transition-all font-bold uppercase text-xs tracking-tight shadow-sm active:scale-95 cursor-pointer",
                  isProduceDisabled
                    ? "bg-gray-200 text-gray-400 cursor-not-allowed"
                    : "bg-blue-600 text-white hover:bg-blue-700 shadow-blue-500/20"
                )}
              >
               <PackagePlus size={15} /> LƯU SẢN LƯỢNG
              </button>
          </div>
          
          <div className="flex flex-col gap-2.5 sm:gap-3 pt-3 sm:pt-4 border-t border-gray-200 mt-1">
              <div className="flex justify-between items-center text-xs font-bold uppercase">
                <span className="text-gray-500 font-bold">Đã in nhãn:</span>
                <span className={cn(printed >= total ? "text-emerald-600 font-black" : "text-blue-700 font-black")}>
                  {printed.toLocaleString()} / {total.toLocaleString()} {printed >= total ? '✓ (Đủ kế hoạch)' : ''}
                </span>
              </div>
              <div className="flex flex-col sm:flex-row items-stretch sm:items-end gap-2 sm:gap-4">
                  <div className="flex-1">
                    <span className="text-[10px] font-bold text-gray-700 uppercase block mb-1">
                      Số lượng IN <span className="text-red-500">*</span>
                      {remainingPrint > 0 && (
                        <span className="text-gray-400 font-normal lowercase ml-1">(tối đa: {remainingPrint})</span>
                      )}
                    </span>
                    <input 
                      type="number"
                      min="1"
                      max={remainingPrint}
                      value={printInputQty === '' ? '' : printInputQty}
                      onChange={(e) => {
                          const raw = e.target.value;
                          if (raw === '') {
                            setPrintInputQty('');
                          } else {
                            const val = parseInt(raw);
                            setPrintInputQty(isNaN(val) ? '' : val);
                          }
                      }}
                      placeholder={printed >= total ? "Đã in đủ kế hoạch" : "Bắt buộc điền..."}
                      disabled={printed >= total}
                      autoComplete="off"
                      className="w-full bg-white border border-gray-300 focus:border-green-500 rounded-xl px-2.5 py-1.5 sm:py-2 font-black text-green-700 outline-none transition-all text-xs sm:text-sm placeholder:font-normal placeholder:text-gray-400 placeholder:italic disabled:bg-gray-100 disabled:cursor-not-allowed"
                    />
                  </div>
                  <button 
                    disabled={isPrintDisabled}
                    onClick={() => {
                      if (isPrintDisabled) return;
                      const txId = `GLZ-${Date.now()}-${norm.id}`;
                      const formatQR = `${norm.id}|${numPrint}|GLAZING|${Date.now()}|${txId}|DCLR|${norm.id}|||`;
                      
                      const labelTx: import('./types').Transaction = {
                        id: txId,
                        partId: norm.id,
                        partName: norm.partName,
                        quantity: numPrint,
                        type: 'STAGE_OUT',
                        stageId: 'GLAZING',
                        location: 'IN',
                        targetStageId: 'DCLR',
                        timestamp: Date.now(),
                        qrData: formatQR,
                        planId: plan.id,
                        printed: true,
                        kpiRecorded: true
                      };
                      
                      onPrint(labelTx, numPrint);
                      // Không để số sẵn sau khi in: reset về trống
                      setPrintInputQty('');
                    }}
                    className={cn(
                      "h-9 sm:h-11 px-4 sm:px-6 rounded-xl flex items-center justify-center gap-1.5 transition-all font-bold uppercase text-xs tracking-tight shadow-md active:scale-95 cursor-pointer",
                      isPrintDisabled
                        ? "bg-gray-200 text-gray-400 cursor-not-allowed"
                        : "bg-green-600 text-white hover:bg-green-700 shadow-green-500/20"
                    )}
                  >
                   <Printer size={15} /> IN NHÃN
                  </button>
              </div>
          </div>
      </div>
    </div>
  );
}

function GlazingView({ parts, inventory: globalInventory, onManualInbound, setDefectModal, refreshData, labels, onPrint, aggregatedGlazingOut }: any) {
  const inventory = globalInventory || [];
  const [activeTab, setActiveTab] = useState<'INVENTORY' | 'INBOUND' | 'OUTBOUND' | 'QUICK_PRINT' | 'PLANNING' | 'CONFIG'>('INVENTORY');
  const [showEditInventoryModal, setShowEditInventoryModal] = useState<{ partId: string, quantity: number, stageId: any, location: any } | null>(null);
  const [editPassword, setEditPassword] = useState("");
  const [configs, setConfigs] = useState<import('./types').GlazingConfig[]>([]);
  const [outConfigs, setOutConfigs] = useState<import('./types').GlazingOutConfig[]>([]);
  const [quickPrintParts, setQuickPrintParts] = useState<{id: string, name: string, quantity: number}[]>(() => storageService.getQuickPrintParts());
  
  const handleDeleteGlazingInventoryItem = async (item: import('./types').InventoryItem) => {
    const pName = parts.find((p: any) => p.id === item.partId)?.name || item.partId;
    if (await customConfirm(`Bạn có chắc muốn xóa linh kiện "${pName}" (Mã: ${item.partId}, SL: ${item.quantity}) khỏi Tồn kho IN Dán Kính? Thao tác này sẽ đồng bộ xóa lên Supabase database.`)) {
      storageService.deleteInventoryItem(item.partId, 'GLAZING', 'IN', item.originalPartId);
      refreshData();
    }
  };

  const handleEditInventory = () => {
    if (editPassword === 'admin123') {
      if (showEditInventoryModal) {
        const { partId, stageId, location, quantity } = showEditInventoryModal;
        storageService.setAbsoluteInventory(partId, stageId, location, quantity);
        refreshData();
        setShowEditInventoryModal(null);
        setEditPassword("");
      }
    } else {
      alert('Mật khẩu không chính xác!');
    }
  };

  // Planning state
  const [selectedModel, setSelectedModel] = useState("");
  const [planQuantity, setPlanQuantity] = useState(0);
  const [targetCompletion, setTargetCompletion] = useState(format(new Date(), "yyyy-MM-dd'T'HH:mm"));
  const [glazingPlans, setGlazingPlans] = useState<import('./types').GlazingPlan[]>(() => storageService.getGlazingPlans());
  const [glazingPlanNorms, setGlazingPlanNorms] = useState<import('./types').GlazingPlanNorm[]>(() => storageService.getGlazingPlanNorms());
  const [planSubTab, setPlanSubTab] = useState<'IN_PROGRESS' | 'COMPLETED'>('IN_PROGRESS');
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedPlanIdForPrint, setSelectedPlanIdForPrint] = useState<string | "">("");
  const [estimatedGlazingStart, setEstimatedGlazingStart] = useState<number | null>(null);
  const [estimatedGlazingEnd, setEstimatedGlazingEnd] = useState<number | null>(null);

  useEffect(() => {
    if (selectedModel && planQuantity > 0 && targetCompletion) {
      const endTs = new Date(targetCompletion).getTime();
      const schedule = storageService.getGlazingSchedule(selectedModel, planQuantity, endTs);
      setEstimatedGlazingStart(schedule.start);
      setEstimatedGlazingEnd(schedule.end);
    } else {
      setEstimatedGlazingStart(null);
      setEstimatedGlazingEnd(null);
    }
  }, [selectedModel, planQuantity, targetCompletion, glazingPlanNorms]);

  useEffect(() => {
    setGlazingPlans(storageService.getGlazingPlans());
  }, [labels]);

  const pendingGlazingLabels = useMemo(() => {
    return (labels || []).filter((l: any) => l.stageId === 'GLAZING' && l.type === 'STAGE_OUT' && l.printed === false);
  }, [labels]);
  

  useEffect(() => {
    const loadedConfigs = storageService.getGlazingConfigs();
    const loadedOutConfigs = storageService.getGlazingOutConfigs();
    
    // Auto deduplicate any legacy duplicates in localStorage
    const uniqueConfigs: import('./types').GlazingConfig[] = [];
    const seen = new Set<string>();
    loadedConfigs.forEach(c => {
      if (!seen.has(c.partId)) {
        seen.add(c.partId);
        uniqueConfigs.push(c);
      }
    });

    if (uniqueConfigs.length !== loadedConfigs.length) {
      storageService.saveGlazingConfigs(uniqueConfigs);
    }

    setConfigs(uniqueConfigs);
    setOutConfigs(loadedOutConfigs.filter(c => !c.finalPartName.toUpperCase().includes('DCLR')));
  }, []);

  useEffect(() => {
    storageService.saveQuickPrintParts(quickPrintParts);
  }, [quickPrintParts]);

  const handleGlazingExport = (item: import('./types').InventoryItem) => {
    const qtyInput = glazingOutQty[item.partId];
    const qty = parseFloat(qtyInput);
    
    if (!qty || qty <= 0) return alert('Vui lòng nhập số lượng xuất');
    if (qty > item.quantity) return alert('Số lượng xuất vượt quá tồn kho');

    try {
      // 1. Record Stage Out (From GLAZING OUT to DCLR)
      const tx = storageService.recordStageOut(item.partId, 'GLAZING', qty, 'OUT', 'DCLR');
      
      // 2. Clear input
      setGlazingOutQty(prev => ({ ...prev, [item.partId]: '' }));
      
      // 3. Refresh
      refreshData();
      
      // 4. Trigger print modal immediately
      const isPseudo = item.partId.startsWith('GLZ-OUT-');
      const displayName = isPseudo ? item.partId.replace('GLZ-OUT-', '') : (parts.find((p: any) => p.id === item.partId)?.name || item.partId);
      onPrint({ ...tx, partName: displayName });
    } catch (err) {
      alert(err instanceof Error ? err.message : 'Lỗi xuất kho');
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
        
        if (data.length < 2) return alert('File Excel không có dữ liệu');

        const headers = data[0].map(h => String(h || '').normalize('NFC').toLowerCase().trim());
        const nameIdx = headers.findIndex(h => h.includes('tên'));
        const codeIdx = headers.findIndex(h => h.includes('mã') || h.includes('code') || h.includes('part'));
        const sourceIdx = headers.findIndex(h => 
          h.includes('tiếp nhận') || 
          h.includes('từ kho') || 
          h.includes('công đoạn') || 
          h.includes('nguồn') || 
          h.includes('source')
        );

        if (codeIdx === -1 || sourceIdx === -1) {
          return alert('Không tìm thấy cột "Mã Linh kiện" hoặc "Công đoạn tiếp nhận từ kho OUT của" trong file Excel.');
        }

        const newConfigs: import('./types').GlazingConfig[] = [];
        const seenPartKeys = new Set<string>();

        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          if (!row || !row[codeIdx]) continue;
          
          const partIdStr = String(row[codeIdx]).trim();
          if (!partIdStr) continue;

          const dedupeKey = partIdStr.toUpperCase();
          if (seenPartKeys.has(dedupeKey)) continue;
          seenPartKeys.add(dedupeKey);

          let partNameStr = nameIdx !== -1 ? String(row[nameIdx] || '').trim() : '';
          if (!partNameStr) {
            const foundP = parts.find((p: any) => p.id.toUpperCase() === partIdStr.toUpperCase());
            if (foundP) partNameStr = foundP.name;
          }

          const sourceRaw = String(row[sourceIdx] || '').normalize('NFC').toUpperCase().trim();
          let sId: StageId = 'BENDING';
          if (sourceRaw.includes('SƠN') || sourceRaw.includes('SON') || sourceRaw.includes('PAINT')) {
            sId = 'PAINTING';
          } else if (sourceRaw.includes('CHẤN') || sourceRaw.includes('DẬP') || sourceRaw.includes('CHAN') || sourceRaw.includes('DAP') || sourceRaw.includes('BEND')) {
            sId = 'BENDING';
          } else if (sourceRaw.includes('HÀN') || sourceRaw.includes('HAN') || sourceRaw.includes('WELD')) {
            sId = 'WELDING';
          } else if (sourceRaw.includes('LASER') || sourceRaw.includes('CẮT')) {
            sId = 'LASER';
          }

          newConfigs.push({
            partName: partNameStr || partIdStr,
            partId: partIdStr,
            sourceStageId: sId
          });
        }

        storageService.saveGlazingConfigs(newConfigs);
        setConfigs(newConfigs);
        alert(`Đã nạp thành công ${newConfigs.length} cấu hình tiếp nhận linh kiện Dán Kính từ kho OUT các công đoạn.`);
      } catch (err) {
        alert('Lỗi đọc file: ' + err);
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleOutFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
        if (data.length < 2) return alert('File Excel không có dữ liệu');

        const headers = data[0].map(h => String(h).toLowerCase().trim());
        const finalNameIdx = headers.findIndex(h => h.includes('thành phẩm') && !h.includes('con'));
        const subNameIdx = headers.findIndex(h => h.includes('tên') && h.includes('con'));
        const subCodeIdx = headers.findIndex(h => h.includes('mã') && h.includes('con'));

        if (finalNameIdx === -1 || subNameIdx === -1 || subCodeIdx === -1) {
          return alert('Không tìm thấy cột Tên thành phẩm, Tên thành phẩm con hoặc Mã thành phẩm con.');
        }

        const outConfigMap = new Map<string, { partId: string, partName: string }[]>();

        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const finalName = String(row[finalNameIdx] || '').trim();
          const subName = String(row[subNameIdx] || '').trim();
          const subCode = String(row[subCodeIdx] || '').trim();

          if (!finalName || !subCode) continue;

          const existingList = outConfigMap.get(finalName) || [];
          existingList.push({ partId: subCode, partName: subName });
          outConfigMap.set(finalName, existingList);
        }

        const newOutConfigs: import('./types').GlazingOutConfig[] = [];
        outConfigMap.forEach((subParts, finalPartName) => {
          // Hide DCLR specific final part names as per request
          if (!finalPartName.toUpperCase().includes('DCLR')) {
            newOutConfigs.push({ finalPartName, subParts });
          }
        });

        storageService.saveGlazingOutConfigs(newOutConfigs);
        setOutConfigs(newOutConfigs);
        alert(`Đã nạp ${newOutConfigs.length} cấu hình gói Dán Kính OUT.`);
      } catch (err) {
        alert('Lỗi đọc file OUT: ' + err);
      }
    };
    reader.readAsBinaryString(file);
  };

  const [inboundQty, setInboundQty] = useState<Record<string, string>>({});
  const [glazingOutQty, setGlazingOutQty] = useState<Record<string, string>>({});

  const handleReceive = (config: import('./types').GlazingConfig) => {
    const qty = parseFloat(inboundQty[config.partId]);
    if (!qty || qty <= 0) return alert('Vui lòng nhập số lượng hợp lệ');
    
    const norm = (s: string) => (s || '').normalize('NFC').trim().toUpperCase().replace(/\s*-\s*(H|CD|S|C|GLZ|OUT|IN)$/i, '');
    const cleanTargetId = norm(config.partId);
    const cleanTargetName = config.partName ? norm(config.partName) : '';

    // Find all matching stocks in source stage OUT
    const matchingStocks = globalInventory.filter((i: any) => {
      if (i.stageId !== config.sourceStageId || i.location !== 'OUT' || (i.quantity || 0) <= 0) return false;
      const cleanInvId = norm(i.partId);
      if (cleanInvId === cleanTargetId) return true;
      if (i.originalPartId && norm(i.originalPartId) === cleanTargetId) return true;
      if (cleanTargetName) {
        const catalogPart = parts.find((p: any) => norm(p.id) === cleanInvId);
        if (catalogPart && norm(catalogPart.name) === cleanTargetName) return true;
      }
      return false;
    });

    const totalAvailable = matchingStocks.reduce((sum: number, i: any) => sum + (i.quantity || 0), 0);
    const sourceStageName = STAGES.find(s => s.id === config.sourceStageId)?.name || config.sourceStageId;

    if (totalAvailable < qty) {
      return alert(`Không đủ tồn kho OUT tại ${sourceStageName}. Kho hiện có: ${totalAvailable}, Yêu cầu nhập: ${qty}`);
    }

    try {
      // 1. Deduct from source stage OUT across matching stocks
      let remainingToDeduct = qty;
      for (const stock of matchingStocks) {
        if (remainingToDeduct <= 0) break;
        const toTake = Math.min(stock.quantity, remainingToDeduct);
        storageService.updateInventory(stock.partId, config.sourceStageId, 'OUT', -toTake, stock.originalPartId);
        remainingToDeduct -= toTake;
      }

      // 2. Add to GLAZING IN
      storageService.updateInventory(config.partId, 'GLAZING', 'IN', qty);

      // 3. Record transaction
      const ts = Date.now();
      const txId = `GLZ-RCV-${ts}-${Math.random().toString(36).substring(2, 7).toUpperCase()}`;
      const receiveTx: import('./types').Transaction = {
        id: txId,
        type: 'STAGE_IN',
        partId: config.partId,
        partName: config.partName || (parts.find((p: any) => p.id === config.partId)?.name || config.partId),
        quantity: qty,
        stageId: 'GLAZING',
        location: 'IN',
        sourceStageId: config.sourceStageId,
        timestamp: ts
      };
      storageService.saveTransactions([receiveTx, ...storageService.getTransactions()]);

      setInboundQty(prev => ({ ...prev, [config.partId]: '' }));
      refreshData();
      alert(`Đã tiếp nhận ${qty} ${config.partName || config.partId} từ kho OUT ${sourceStageName} vào kho IN Dán Kính thành công!`);
    } catch(err) {
      alert(err instanceof Error ? err.message : 'Lỗi nhập kho');
    }
  };

  const [outboundQty, setOutboundQty] = useState<Record<string, string>>({});
  const handleTransferOut = (finalPartName: string) => {
    const qty = parseFloat(outboundQty[finalPartName]);
    if (!qty || qty <= 0) return alert('Vui lòng nhập số lượng hợp lệ');
    
    const config = outConfigs.find(c => c.finalPartName === finalPartName);
    if (!config) return alert('Không tìm thấy cấu hình cho ' + finalPartName);
    
    // Aggregate subparts in case there are duplicates in the config
    const requiredQtys: Record<string, number> = {};
    for (const sp of config.subParts) {
      if (!requiredQtys[sp.partId]) requiredQtys[sp.partId] = 0;
      requiredQtys[sp.partId] += 1;
    }

    // Validate all subparts have enough qty
    for (const partId of Object.keys(requiredQtys)) {
      const required = requiredQtys[partId] * qty;
      const inv = globalInventory.find(i => i.partId.toUpperCase() === partId.toUpperCase() && i.location === 'IN' && i.stageId === 'GLAZING');
      if (!inv || inv.quantity < required) {
        return alert(`Không đủ tồn kho IN cho linh kiện mã ${partId}. Cần: ${required}, Có: ${inv?.quantity || 0}`);
      }
    }

    try {
      // Deduct subparts from IN
      for (const partId of Object.keys(requiredQtys)) {
        const required = requiredQtys[partId] * qty;
        storageService.updateInventory(partId, 'GLAZING', 'IN', -required);
      }

      // We use finalPartName as the "partId" for the final compiled object in inventory OUT
      const cleanFinalName = finalPartName.trim();
      const pseudoPartId = `GLZ-OUT-${cleanFinalName}`;
      storageService.updateInventory(pseudoPartId, 'GLAZING', 'OUT', qty);
      
      const ts = Date.now();
      const txId = `GLZ-${ts}-${pseudoPartId}`;
      
      // STAGE_IN transaction into OUT for KPIs
      const transaction: import('./types').Transaction = {
         id: txId,
         type: 'STAGE_IN',
         partId: pseudoPartId,
         quantity: qty,
         stageId: 'GLAZING',
         location: 'OUT',
         timestamp: ts,
      };
      
      storageService.saveTransactions([transaction, ...storageService.getTransactions()]);
      
      // Label awaiting printing
      const formatQR = `${pseudoPartId}|${qty}|GLAZING|${ts}|${txId}|DCLR|${pseudoPartId}|||`;
      const labelTx: import('./types').Transaction = {
         id: txId,
         partId: pseudoPartId,
         partName: cleanFinalName,
         quantity: qty,
         type: 'STAGE_OUT',
         stageId: 'GLAZING',
         targetStageId: 'DCLR',
         timestamp: ts,
         qrData: formatQR,
         printed: false,
         kpiRecorded: true
      };
      storageService.saveLabel(labelTx);
      
      setOutboundQty({...outboundQty, [finalPartName]: ''});
      refreshData();
      alert('Đóng gói & Chuyển sang OUT thành công!');
    } catch(err) {
      alert('Lỗi xuất kho: ' + (err instanceof Error ? err.message : String(err)));
    }
  };

  const handleQuickPrintImport = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const dataBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(dataBuffer);
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<any[]>(firstSheet, { header: 1 });

      if (rows.length < 2) {
        alert('File Excel không có dữ liệu');
        return;
      }

      const headers = rows[0].map((h: any) => String(h || '').toLowerCase().trim());
      const nameIdx = headers.findIndex((h: any) => h.includes('tên'));
      const idIdx = headers.findIndex((h: any) => h.includes('mã'));

      if (nameIdx === -1 || idIdx === -1) {
        alert('Không tìm thấy cột "Tên linh kiện" hoặc "Mã linh kiện" trong file Excel.');
        return;
      }

      const imported = rows.slice(1).map(row => ({
        id: String(row[idIdx] || '').trim(),
        name: String(row[nameIdx] || '').trim(),
        quantity: 1
      })).filter(p => p.id && p.name);

      // Lọc trùng mã (ID)
      const uniqueMap = new Map();
      imported.forEach(item => {
        if (!uniqueMap.has(item.id)) {
          uniqueMap.set(item.id, item);
        }
      });

      setQuickPrintParts(Array.from(uniqueMap.values()));
    } catch (err) {
      alert('Lỗi đọc file: ' + (err instanceof Error ? err.message : String(err)));
    }
  };

  const handleGlazingPlanImport = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const dataBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(dataBuffer);
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<any[]>(firstSheet, { header: 1 });

      if (rows.length < 2) {
        alert('File Excel không có dữ liệu');
        return;
      }

      const headers = rows[0].map((h: any) => String(h || '').toLowerCase().trim());
      const nameIdx = headers.findIndex((h: any) => h.includes('tên lk') || h.includes('tên linh kiện'));
      const idIdx = headers.findIndex((h: any) => h.includes('mã lk') || h.includes('mã linh kiện') || h.includes('part id'));
      const normIdx = headers.findIndex((h: any) => h.includes('định mức') || h.includes('đm'));
      const modelIdx = headers.findIndex((h: any) => h.includes('tên model') || h.includes('mã model'));
      const appliedIdx = headers.findIndex((h: any) => h.includes('model áp dụng') || h.includes('model sx') || h.includes('mã model'));

      if (nameIdx === -1 || normIdx === -1) {
        alert('Không tìm thấy các cột tối thiểu: Tên LK (hoặc Tên linh kiện), Định mức.');
        return;
      }

      if (modelIdx === -1 && appliedIdx === -1) {
        alert('Cần có ít nhất một trong các cột: Tên model hoặc Model áp dụng.');
        return;
      }

      const imported: import('./types').GlazingPlanNorm[] = rows.slice(1).map((row, idx) => {
        const partName = String(row[nameIdx] || '').trim();
        const norm = parseFloat(String(row[normIdx]).replace(',', '.')) || 0;
        
        // Cần đảm bảo có modelName. Nếu thiếu modelIdx thì lấy appliedIdx
        const modelVal = modelIdx > -1 ? String(row[modelIdx] || '').trim() : '';
        const appliedVal = appliedIdx > -1 ? String(row[appliedIdx] || '').trim() : '';
        
        const modelName = modelVal || appliedVal;
        const appliedModel = appliedVal || modelVal;
        
        let partId = '';
        if (idIdx > -1 && row[idIdx]) {
          partId = String(row[idIdx]).trim();
        } else {
          const foundPart = parts.find((p: any) => p.name.toLowerCase() === partName.toLowerCase());
          if (foundPart) partId = foundPart.id;
          else partId = `GLZ-PLAN-${idx}-${Date.now()}`;
        }

        return {
          id: partId,
          partName,
          norm,
          modelName,
          appliedModel
        };
      }).filter(n => n.partName);

      storageService.saveGlazingPlanNorms(imported);
      setGlazingPlanNorms(imported);
      alert(`Đã nhập thành công ${imported.length} định mức kế hoạch dán kính.`);
    } catch (err) {
      alert('Lỗi đọc file: ' + (err instanceof Error ? err.message : String(err)));
    }
  };

  const [excelGlazingImportList, setExcelGlazingImportList] = useState<import('./types').GlazingPlan[] | null>(null);
  const [missingNormList, setMissingNormList] = useState<MissingNormModelInfo[]>([]);
  const [excelGlazingStats, setExcelGlazingStats] = useState<{ grandTotal: number; validTotal: number; missingTotal: number } | null>(null);
  const [glazingModalTab, setGlazingModalTab] = useState<'VALID' | 'MISSING_NORM'>('VALID');
  const [isImportingGlazingExcel, setIsImportingGlazingExcel] = useState(false);
  const [isBatchCreatingGlazing, setIsBatchCreatingGlazing] = useState(false);

  const handleExcelGlazingUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsImportingGlazingExcel(true);
    setGlazingModalTab('VALID');
    try {
      const currentNorms = storageService.getGlazingPlanNorms();
      const result = await importExcelToGlazingPlan(file, currentNorms);
      
      setMissingNormList(result.missingNormModels || []);
      setExcelGlazingImportList(result.plans);
      setExcelGlazingStats({
        grandTotal: result.grandTotalFileQty,
        validTotal: result.validTotalQty,
        missingTotal: result.missingTotalQty
      });

      if (result.plans.length === 0 && result.missingNormModels && result.missingNormModels.length > 0) {
        setGlazingModalTab('MISSING_NORM');
      } else if (result.plans.length === 0 && (!result.missingNormModels || result.missingNormModels.length === 0)) {
        alert('Không tìm thấy dữ liệu số lượng kế hoạch sản xuất (> 0) trong các cột ngày của file Excel!');
      }
    } catch (err: any) {
      alert(err?.message || 'Lỗi khi đọc file Excel dán kính!');
    } finally {
      setIsImportingGlazingExcel(false);
      e.target.value = '';
    }
  };

  const handleConfirmBatchCreateGlazingPlans = () => {
    if (!excelGlazingImportList || excelGlazingImportList.length === 0) return;
    setIsBatchCreatingGlazing(true);
    setTimeout(() => {
      try {
        const batchItems = excelGlazingImportList.map(p => ({
          modelId: p.modelId,
          quantity: p.targetQuantity,
          targetCompletion: p.targetCompletionTime,
          note: p.note,
          excelRequiredDate: p.excelRequiredDate,
          id: p.id
        }));
        const created = storageService.createBatchGlazingPlans(batchItems);
        setGlazingPlans(storageService.getGlazingPlans());
        refreshData();
        setExcelGlazingImportList(null);
        alert(`Đã nhập và tạo thành công ${created.length} Kế Hoạch Sản Xuất Dán Kính từ file Excel!`);
      } catch (err: any) {
        alert(err?.message || 'Có lỗi xảy ra khi tạo các Kế Hoạch Dán Kính!');
      } finally {
        setIsBatchCreatingGlazing(false);
      }
    }, 50);
  };

  const handleCreatePlan = (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedModel || planQuantity <= 0) return;
    
    try {
      const endTs = targetCompletion ? new Date(targetCompletion).getTime() : Date.now();
      storageService.createGlazingPlan(selectedModel, planQuantity, endTs);
      setGlazingPlans(storageService.getGlazingPlans());
      setSelectedModel("");
      setPlanQuantity(0);
      refreshData();
      alert('Đã lập kế hoạch dán kính thành công!');
    } catch (err) {
      alert(err instanceof Error ? err.message : 'Lỗi khi tạo kế hoạch');
    }
  };

  const currentGlazingPlans = useMemo(() => {
    return glazingPlans.filter(p => p.status !== 'COMPLETED');
  }, [glazingPlans]);

  const completedGlazingPlans = useMemo(() => {
    return glazingPlans.filter(p => p.status === 'COMPLETED');
  }, [glazingPlans]);

  const filteredQuickPrintParts = useMemo(() => {
    if (!searchQuery.trim()) return quickPrintParts;
    const q = searchQuery.toLowerCase().trim();
    return quickPrintParts.filter(p => 
      p.id.toLowerCase().includes(q) || 
      p.name.toLowerCase().includes(q)
    );
  }, [quickPrintParts, searchQuery]);

  return (
    <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} className="space-y-6">
      {/* Glazing Tabs (Touch-friendly & Scrollable for mobile portrait) */}
      <div className="flex items-center gap-1.5 sm:gap-2 border-b border-gray-200 pb-2 overflow-x-auto no-scrollbar scroll-smooth whitespace-nowrap">
        {([
          { id: 'INVENTORY', label: 'Tồn kho IN', fullLabel: 'Tồn kho IN', icon: <Boxes size={15} /> },
          { id: 'PLANNING', label: 'Kế hoạch SX', fullLabel: 'Kế hoạch sản xuất', icon: <ClipboardList size={15} /> },
          { id: 'INBOUND', label: 'Nhập IN', fullLabel: 'Nhập IN (Từ OUT)', icon: <ArrowDownLeft size={15} /> },
          { id: 'QUICK_PRINT', label: 'In nhanh', fullLabel: 'In Nhãn Nhanh', icon: <Printer size={15} /> },
          { id: 'OUTBOUND', label: 'Đóng gói OUT', fullLabel: 'Đóng gói OUT', icon: <Package size={15} /> },
          { id: 'CONFIG', label: 'Cấu hình', fullLabel: 'Cấu hình & Định mức', icon: <Settings size={15} /> },
        ] as const).map(item => (
          <button
            key={item.id}
            onClick={() => setActiveTab(item.id as any)}
            className={cn(
              "flex items-center gap-1.5 px-3 sm:px-4 py-2 sm:py-2.5 rounded-xl text-xs sm:text-sm font-bold transition-all shrink-0 cursor-pointer active:scale-95",
              activeTab === item.id 
                ? "bg-blue-600 text-white shadow-sm shadow-blue-500/30" 
                : "bg-gray-100 hover:bg-gray-200 text-gray-600 hover:text-gray-900"
            )}
          >
            {item.icon}
            <span className="sm:hidden">{item.label}</span>
            <span className="hidden sm:inline">{item.fullLabel || item.label}</span>
          </button>
        ))}
      </div>

      <div className="bg-white p-2.5 sm:p-6 rounded-xl sm:rounded-2xl shadow-sm sm:shadow-xl border border-gray-100 min-h-[400px]">
        {activeTab === 'PLANNING' && (
          <div className="space-y-8">
            <div className="bg-blue-50/70 p-3 sm:p-6 rounded-xl sm:rounded-2xl border border-blue-100">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4 sm:mb-6">
                <div className="flex items-center gap-2.5">
                  <div className="bg-blue-600 p-1.5 sm:p-2 rounded-lg text-white">
                    <ClipboardList size={20} />
                  </div>
                  <h3 className="text-sm sm:text-lg md:text-xl font-bold uppercase text-blue-900 tracking-tight">Lập Kế Hoạch SX Dán Kính</h3>
                </div>

                <div className="flex items-center gap-2 flex-wrap">
                  <button
                    type="button"
                    onClick={() => {
                      const systemModels = Array.from(new Set(glazingPlanNorms.map(n => n.appliedModel || n.modelName).filter(Boolean))) as string[];
                      downloadSamplePOExcelTemplate(systemModels);
                    }}
                    className="flex-1 sm:flex-none flex items-center justify-center gap-1.5 bg-white hover:bg-blue-50 text-blue-700 border border-blue-200 px-3 sm:px-4 py-2 sm:py-2.5 rounded-xl font-bold text-xs sm:text-sm cursor-pointer transition-all active:scale-95 shadow-2xs"
                    title="Tải file Excel mẫu đúng chuẩn cấu trúc để nhập liệu"
                  >
                    <Download size={15} />
                    <span>File Mẫu Excel</span>
                  </button>

                  <label className="flex-1 sm:flex-none flex items-center justify-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white px-3 sm:px-4 py-2 sm:py-2.5 rounded-xl font-bold text-xs sm:text-sm cursor-pointer shadow-sm shadow-emerald-500/20 transition-all active:scale-95">
                    <FileSpreadsheet size={15} />
                    <span>{isImportingGlazingExcel ? 'Đang đọc...' : 'Upload Excel Kế Hoạch'}</span>
                    <input 
                      type="file" 
                      accept=".xlsx, .xls" 
                      className="hidden" 
                      onChange={handleExcelGlazingUpload}
                      disabled={isImportingGlazingExcel}
                    />
                  </label>
                </div>
              </div>

              <form onSubmit={handleCreatePlan} className="space-y-4 sm:space-y-6">
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3 sm:gap-6 items-end">
                  <div className="space-y-1 sm:space-y-2">
                    <label className="text-[10px] sm:text-xs font-bold uppercase text-blue-900/60 tracking-wider pl-1">Chọn Model</label>
                    <SearchableSelect 
                      options={Array.from(new Set(storageService.getModelBOM().map(b => b.modelId))).map(id => ({ 
                        id, 
                        label: parts.find((p: any) => p.id === id)?.name || id 
                      }))}
                      value={selectedModel}
                      onChange={setSelectedModel}
                      placeholder="Chọn model sản xuất..."
                    />
                  </div>
                  <div className="space-y-1 sm:space-y-2">
                    <label className="text-[10px] sm:text-xs font-bold uppercase text-blue-900/60 tracking-wider pl-1">Số lượng (Máy)</label>
                    <input 
                      type="number"
                      value={planQuantity || ''}
                      onChange={e => setPlanQuantity(parseFloat(e.target.value) || 0)}
                      className="w-full p-2.5 sm:p-3.5 rounded-xl border border-blue-200 font-bold text-xs sm:text-base focus:border-blue-600 outline-none bg-white"
                      placeholder="Nhập số lượng..."
                    />
                  </div>
                  <div className="space-y-1 sm:space-y-2">
                    <label className="text-[10px] sm:text-xs font-bold uppercase text-blue-900/60 tracking-wider pl-1">Ngày mong muốn hoàn thành</label>
                    <input 
                      type="datetime-local"
                      value={targetCompletion}
                      onChange={e => setTargetCompletion(e.target.value)}
                      className="w-full p-2.5 sm:p-3.5 rounded-xl border border-blue-200 font-bold text-xs sm:text-base focus:border-blue-600 outline-none bg-white"
                    />
                  </div>
                </div>

                <div className="flex flex-col sm:flex-row justify-between items-stretch sm:items-center bg-white/70 p-3 sm:p-4 rounded-xl border border-blue-100 gap-3">
                  <div className="grid grid-cols-2 gap-2 bg-white px-3 sm:px-5 py-2.5 rounded-xl border border-blue-100 shadow-2xs">
                    <div className="flex flex-col">
                      <span className="text-[9px] font-bold text-blue-900/50 uppercase tracking-wider leading-none mb-1 flex items-center gap-1">
                        <Clock size={10} className="text-blue-400" /> Bắt đầu dự kiến
                      </span>
                      <span className="text-xs sm:text-sm font-black text-blue-800">
                        {estimatedGlazingStart ? format(new Date(estimatedGlazingStart), 'HH:mm - dd/MM') : '--:--'}
                      </span>
                    </div>
                    <div className="flex flex-col border-l border-blue-100 pl-2">
                      <span className="text-[9px] font-bold text-orange-900/50 uppercase tracking-wider leading-none mb-1 flex items-center gap-1">
                        <CheckCircle2 size={10} className="text-orange-400" /> Dự kiến hoàn thành
                      </span>
                      <span className="text-xs sm:text-sm font-black text-orange-600">
                        {estimatedGlazingEnd ? format(new Date(estimatedGlazingEnd), 'HH:mm - dd/MM') : '--:--'}
                      </span>
                    </div>
                  </div>
                  <button 
                    type="submit"
                    disabled={!selectedModel || planQuantity <= 0}
                    className="w-full sm:w-auto bg-blue-600 text-white px-6 sm:px-10 py-2.5 sm:py-3.5 rounded-xl font-bold uppercase text-xs sm:text-sm tracking-wider hover:bg-blue-700 transition-all shadow-md shadow-blue-500/20 active:scale-95 disabled:opacity-50 cursor-pointer"
                  >
                    Tạo Kế Hoạch
                  </button>
                </div>
              </form>
            </div>

            {selectedModel && (
              <div className="bg-white rounded-2xl border border-blue-100 p-6 flex flex-col gap-4">
                <div className="flex justify-between items-center">
                  <h4 className="font-black text-gray-900 uppercase tracking-tight flex items-center gap-2">
                    <div className="w-1.5 h-4 bg-blue-600 rounded-full" />
                    Thành phần Model (Gói Cánh / Gói Đỉnh)
                  </h4>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                  {glazingPlanNorms.filter(n => n.appliedModel === selectedModel).map((c, i) => (
                    <div key={i} className="bg-blue-50/50 p-4 rounded-xl border border-blue-50 flex flex-col gap-1">
                      <span className="text-[10px] font-black text-blue-400 uppercase tracking-widest">Linh kiện</span>
                      <span className="font-bold text-gray-800">{c.partName}</span>
                      <div className="flex border-t border-blue-100 mt-2 pt-2 gap-4">
                        <div className="flex-1">
                          <span className="text-[9px] font-bold text-gray-400 uppercase block">Số lượng:</span>
                          <span className="text-base font-black text-blue-700">{planQuantity.toLocaleString()} Tấm</span>
                        </div>
                        <div className="flex-1">
                          <span className="text-[9px] font-bold text-gray-400 uppercase block">Tổng thời gian:</span>
                          <span className="text-base font-black text-orange-600">{(c.norm * planQuantity).toLocaleString()} s</span>
                        </div>
                      </div>
                    </div>
                  ))}
                  {glazingPlanNorms.filter(n => n.appliedModel === selectedModel).length === 0 && (
                    <div className="col-span-full py-6 text-center text-gray-400 italic text-sm">
                      Model này chưa được cài đặt định mức dán kính
                    </div>
                  )}
                </div>
              </div>
            )}

            <div className="space-y-4">
              <div className="flex border-b border-gray-200 gap-3 sm:gap-6 px-1 sm:px-4 overflow-x-auto no-scrollbar">
                <button
                  onClick={() => setPlanSubTab('IN_PROGRESS')}
                  className={cn(
                    "pb-2.5 sm:pb-3 text-xs sm:text-sm font-bold uppercase tracking-wider transition-all relative shrink-0 cursor-pointer",
                    planSubTab === 'IN_PROGRESS' ? "text-blue-600" : "text-gray-400 hover:text-gray-600"
                  )}
                >
                  <div className="flex items-center gap-1.5">
                    <span>Đang thực hiện</span>
                    <span className="bg-blue-100 text-blue-700 px-2 py-0.5 rounded-full text-[10px] font-black">{currentGlazingPlans.length}</span>
                  </div>
                  {planSubTab === 'IN_PROGRESS' && (
                    <div className="absolute bottom-0 left-0 right-0 h-1 bg-blue-600 rounded-t-full" />
                  )}
                </button>
                <button
                  onClick={() => setPlanSubTab('COMPLETED')}
                  className={cn(
                    "pb-2.5 sm:pb-3 text-xs sm:text-sm font-bold uppercase tracking-wider transition-all relative shrink-0 cursor-pointer",
                    planSubTab === 'COMPLETED' ? "text-green-600" : "text-gray-400 hover:text-gray-600"
                  )}
                >
                  <div className="flex items-center gap-1.5">
                    <span>Đã hoàn thành</span>
                    <span className="bg-green-100 text-green-700 px-2 py-0.5 rounded-full text-[10px] font-black">{completedGlazingPlans.length}</span>
                  </div>
                  {planSubTab === 'COMPLETED' && (
                    <div className="absolute bottom-0 left-0 right-0 h-1 bg-green-600 rounded-t-full" />
                  )}
                </button>
              </div>

              {planSubTab === 'IN_PROGRESS' && (
                <>
                  {currentGlazingPlans.length > 0 ? (
                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3 sm:gap-4">
                      {currentGlazingPlans.map((plan) => {
                        const modelName = parts.find((p: any) => p.id === plan.modelId)?.name || plan.modelId;
                        const components = glazingPlanNorms.filter(n => n.appliedModel === plan.modelId);
                        
                        return (
                          <div key={plan.id} className="bg-white border border-gray-200 rounded-xl sm:rounded-2xl p-3.5 sm:p-5 hover:border-blue-200 transition-all shadow-2xs">
                            <div className="flex justify-between items-start mb-3 gap-2">
                              <div className="min-w-0">
                                <div className="text-[10px] font-bold text-blue-600 uppercase tracking-wider mb-0.5">{plan.id}</div>
                                <h5 className="font-bold text-gray-900 leading-tight text-xs sm:text-base truncate">{modelName}</h5>
                              </div>
                              <div className="px-2.5 py-0.5 rounded-full text-[10px] bg-blue-100 text-blue-700 font-bold uppercase shrink-0">
                                {plan.targetQuantity} MÁY
                              </div>
                            </div>

                            <div className="space-y-2 mb-3">
                              <div className="text-[10px] font-bold text-gray-400 uppercase tracking-wider pl-0.5">Thành phần & Tiến độ in</div>
                              <div className="space-y-1">
                                {components.map((c, i) => {
                                  const printed = (plan.printedQuantities || {})[c.id] || 0;
                                  const target = plan.targetQuantity;
                                  const compProgress = Math.min(100, (printed / target) * 100);
                                  
                                  return (
                                    <div key={i} className="flex flex-col bg-gray-50 px-2.5 py-1.5 rounded-lg gap-1">
                                      <div className="flex justify-between items-center text-[11px] sm:text-xs">
                                        <span className="font-medium text-gray-700 truncate mr-2">{c.partName}</span>
                                        <span className="font-bold text-blue-700 whitespace-nowrap">{printed}/{target}</span>
                                      </div>
                                      <div className="h-1 bg-gray-200 rounded-full overflow-hidden">
                                        <div 
                                          className={cn("h-full transition-all duration-500", compProgress >= 100 ? "bg-green-500" : "bg-blue-500")}
                                          style={{ width: `${compProgress}%` }}
                                        />
                                      </div>
                                    </div>
                                  );
                                })}
                                {components.length === 0 && (
                                  <div className="text-[10px] text-gray-400 italic">Chưa cài đặt định mức cho model này</div>
                                )}
                              </div>
                            </div>

                            <div className="pt-2.5 border-t border-gray-100 flex flex-col gap-2.5">
                              <div className="flex justify-between items-center text-xs">
                                <div className="flex flex-col">
                                  <span className="text-[9px] font-bold text-gray-400 uppercase tracking-wider">Bắt đầu</span>
                                  <span className="text-[11px] sm:text-xs font-bold text-blue-700">
                                    {plan.plannedStartTime ? format(new Date(plan.plannedStartTime), 'HH:mm - dd/MM') : '--:--'}
                                  </span>
                                </div>
                                <div className="flex flex-col items-end">
                                  <span className="text-[9px] font-bold text-orange-400 uppercase tracking-wider">Dự kiến xong</span>
                                  <span className="text-[11px] sm:text-xs font-bold text-orange-700">
                                    {plan.expectedCompletionTime ? format(new Date(plan.expectedCompletionTime), 'HH:mm - dd/MM') : '--:--'}
                                  </span>
                                </div>
                              </div>
                              
                              <div className="flex justify-between items-center border-t border-gray-50 pt-2">
                                <div className="flex flex-col">
                                  <span className="text-[9px] font-bold text-gray-400 uppercase tracking-wider">Target:</span>
                                  <span className="text-[10px] font-bold text-gray-600">
                                    {format(new Date(plan.targetCompletionTime), 'HH:mm - dd/MM')}
                                  </span>
                                </div>
                                <div className="flex items-center gap-1">
                                  <button 
                                    onClick={() => {
                                      if (confirm('Đánh dấu kế hoạch này đã hoàn thành? Kế hoạch sẽ được chuyển sang mục Đã hoàn thành.')) {
                                        storageService.completeGlazingPlan(plan.id);
                                        setGlazingPlans(storageService.getGlazingPlans());
                                        refreshData();
                                      }
                                    }}
                                    className="px-2.5 py-1 text-blue-600 hover:text-white hover:bg-blue-600 rounded-lg transition-all font-bold text-[10px] uppercase tracking-wider flex items-center gap-1 border border-blue-200 cursor-pointer"
                                  >
                                    <CheckCircle2 size={13} /> 
                                    Xong
                                  </button>
                                  <button 
                                    onClick={() => {
                                      if (confirm('Xóa kế hoạch dán kính này?')) {
                                        storageService.deleteGlazingPlan(plan.id);
                                        setGlazingPlans(storageService.getGlazingPlans());
                                        refreshData();
                                      }
                                    }}
                                    className="p-1.5 text-gray-400 hover:text-red-500 hover:bg-red-50 rounded-lg transition-all cursor-pointer"
                                  >
                                    <Trash2 size={14} />
                                  </button>
                                </div>
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  ) : (
                    <div className="py-12 flex flex-col items-center justify-center text-gray-300 bg-gray-50 rounded-2xl border-2 border-dashed border-gray-200">
                      <ClipboardList size={48} strokeWidth={1} />
                      <p className="text-sm sm:text-base font-bold mt-3 uppercase tracking-wider opacity-60">Chưa có kế hoạch dán kính</p>
                      <p className="text-xs opacity-50 mt-1">Sử dụng form bên trên để tạo kế hoạch mới</p>
                    </div>
                  )}
                </>
              )}

              {planSubTab === 'COMPLETED' && (
                <div className="space-y-4">
                  {completedGlazingPlans.length > 0 ? (
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                    {completedGlazingPlans.map((plan) => {
                       const modelName = parts.find((p: any) => p.id === plan.modelId)?.name || plan.modelId;
                       return (
                         <div key={plan.id} className="bg-green-50 border-2 border-green-100 rounded-2xl p-5 hover:border-green-200 transition-all group flex flex-col justify-between opacity-80 hover:opacity-100">
                           <div className="flex justify-between items-start mb-4">
                             <div>
                               <div className="text-[10px] font-black text-green-600 uppercase tracking-widest mb-1">{plan.id}</div>
                               <h5 className="font-bold text-gray-900 leading-tight">{modelName}</h5>
                               <div className="mt-1 text-xs text-green-700 font-medium">Đã in đủ: {plan.targetQuantity} MÁY</div>
                             </div>
                             <div className="px-3 py-1 rounded-full text-[10px] bg-green-100 text-green-700 font-black uppercase tracking-widest flex items-center gap-1">
                               <CheckCircle2 size={12}/> XONG
                             </div>
                           </div>
                           <div className="flex justify-between items-center border-t border-green-200 pt-3 mt-4">
                             <div className="flex flex-col">
                               <span className="text-[9px] font-black text-gray-500 uppercase tracking-widest pl-1">Target</span>
                               <span className="text-[10px] font-bold text-gray-600 italic">
                                 {format(new Date(plan.targetCompletionTime), 'HH:mm - dd/MM/yyyy')}
                               </span>
                             </div>
                             <div className="flex gap-1">
                              <button 
                                onClick={() => {
                                  // Hack to trigger status change safely
                                  const plans = storageService.getGlazingPlans();
                                  const planIndex = plans.findIndex(p => p.id === plan.id);
                                  if(planIndex !== -1) {
                                    plans[planIndex].status = 'IN_PROGRESS';
                                    storageService.saveGlazingPlans(plans);
                                    setGlazingPlans(storageService.getGlazingPlans());
                                    refreshData();
                                  }
                                }}
                                className="p-2 text-gray-400 hover:text-blue-500 hover:bg-blue-50 rounded-lg transition-all"
                                title="Khôi phục"
                              >
                                <RotateCcw size={16} />
                              </button>
                               <button 
                                 onClick={() => {
                                   if (confirm('Xóa vĩnh viễn kế hoạch này?')) {
                                     storageService.deleteGlazingPlan(plan.id);
                                     setGlazingPlans(storageService.getGlazingPlans());
                                     refreshData();
                                   }
                                 }}
                                 className="p-2 text-gray-400 hover:text-red-500 hover:bg-red-50 rounded-lg transition-all"
                               >
                                 <Trash2 size={16} />
                               </button>
                             </div>
                           </div>
                         </div>
                       );
                    })}
                  </div>
                  ) : (
                    <div className="py-20 flex flex-col items-center justify-center text-gray-300 bg-gray-50 rounded-3xl border-2 border-dashed border-gray-200">
                      <CheckCircle2 size={64} strokeWidth={1} />
                      <p className="text-lg font-bold mt-4 uppercase tracking-widest opacity-50">Chưa có kế hoạch nào hoàn thành</p>
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>
        )}
        {activeTab === 'QUICK_PRINT' && (
          <div className="space-y-6 sm:space-y-8">
            {/* Step 1: Select Plan */}
            <div className="bg-white p-3.5 sm:p-8 rounded-xl sm:rounded-3xl border border-blue-100 shadow-sm sm:shadow-xl space-y-4 sm:space-y-6">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 sm:w-12 sm:h-12 bg-blue-600 rounded-xl sm:rounded-2xl flex items-center justify-center text-white shadow-sm shadow-blue-500/20 shrink-0">
                  <ClipboardList size={20} />
                </div>
                <div>
                  <h3 className="text-sm sm:text-lg md:text-xl font-bold text-gray-900 uppercase tracking-tight">Chọn Kế Hoạch Dán Kính</h3>
                  <p className="text-xs text-gray-500 font-bold italic">* Bạn cần chọn một kế hoạch đang thực hiện để in nhãn</p>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-2.5 sm:gap-4">
                {glazingPlans.filter(p => p.status !== 'COMPLETED').map(plan => (
                  <button
                    key={plan.id}
                    onClick={() => setSelectedPlanIdForPrint(plan.id)}
                    className={cn(
                      "p-3.5 sm:p-5 rounded-xl sm:rounded-2xl border-2 transition-all text-left flex flex-col gap-1.5 sm:gap-2 relative group cursor-pointer active:scale-98",
                      selectedPlanIdForPrint === plan.id 
                        ? "border-blue-600 bg-blue-50 shadow-sm shadow-blue-500/10" 
                        : "border-gray-100 hover:border-blue-300 hover:bg-gray-50"
                    )}
                  >
                    {selectedPlanIdForPrint === plan.id && (
                      <div className="absolute -top-2.5 -right-2.5 w-6 h-6 sm:w-8 sm:h-8 bg-blue-600 text-white rounded-full flex items-center justify-center shadow-sm border-2 border-white">
                        <Check size={14} strokeWidth={3} />
                      </div>
                    )}
                    <span className="text-[10px] font-bold text-gray-400 uppercase tracking-wider">{plan.modelId}</span>
                    <span className="font-bold text-gray-800 text-sm sm:text-lg">Hạn: {format(new Date(plan.targetCompletionTime), 'dd/MM')}</span>
                    <div className="flex justify-between items-center mt-1">
                      <span className="text-xs font-bold text-blue-600 bg-blue-100/50 px-2 py-0.5 rounded-full">
                        SL: {plan.targetQuantity}
                      </span>
                      <span className={cn(
                        "text-[9px] font-bold uppercase px-2 py-0.5 rounded-full",
                        plan.status === 'IN_PROGRESS' ? "text-orange-600 bg-orange-100" : "text-blue-600 bg-blue-100"
                      )}>
                        {plan.status === 'IN_PROGRESS' ? 'Đang chạy' : 'Chờ'}
                      </span>
                    </div>
                  </button>
                ))}
                {glazingPlans.filter(p => p.status !== 'COMPLETED').length === 0 && (
                  <div className="col-span-full py-8 text-center bg-gray-50 rounded-xl border-2 border-dashed border-gray-200">
                    <p className="text-gray-400 font-bold uppercase tracking-wider text-xs sm:text-sm">Chưa có kế hoạch nào đang hoạt động</p>
                  </div>
                )}
              </div>

            </div>

            {/* Step 2: Show Components if Plan selected */}
            {selectedPlanIdForPrint && (
              <div className="space-y-4 sm:space-y-6">
                <div className="flex justify-between items-center px-1 sm:px-4">
                  <h4 className="font-bold text-gray-900 uppercase tracking-tight flex items-center gap-2 text-xs sm:text-base">
                    <div className="w-1.5 h-4 bg-blue-600 rounded-full" />
                    Linh kiện trong kế hoạch
                  </h4>
                  <div className="text-[11px] sm:text-xs font-bold text-gray-400 uppercase tracking-wider">
                    Model: <span className="text-blue-600">{glazingPlans.find(p => p.id === selectedPlanIdForPrint)?.modelId}</span>
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3 sm:gap-6">
                  {glazingPlanNorms.filter(n => n.appliedModel === glazingPlans.find(p => p.id === selectedPlanIdForPrint)?.modelId).map((norm, idx) => {
                    const plan = glazingPlans.find(p => p.id === selectedPlanIdForPrint)!;
                    return (
                      <GlazingComponentPrintCard
                        key={`${plan.id}-${norm.id}-${idx}`}
                        norm={norm}
                        plan={plan}
                        idx={idx}
                        onRecordProduction={(producedQty: number, timestamp: number) => {
                          storageService.recordStageOut(norm.id, 'GLAZING', producedQty, 'IN', 'DCLR', undefined, true, timestamp);
                          storageService.updateGlazingPlanProgress(plan.id, norm.id, producedQty);
                          setGlazingPlans(storageService.getGlazingPlans());
                          refreshData();
                        }}
                        onPrint={(labelTx: any, printQty: number) => {
                          storageService.saveLabel(labelTx);
                          // Lưu vào cả transactions (lưu 35 ngày) để Báo cáo sản lượng & Biểu đồ năng suất xem được trọn vẹn 35 ngày
                          const txList = storageService.getTransactions();
                          storageService.saveTransactions([labelTx, ...txList]);
                          storageService.updateGlazingPlanPrinted(plan.id, norm.id, printQty);
                          setGlazingPlans(storageService.getGlazingPlans());
                          onPrint(labelTx);
                          refreshData();
                        }}
                        onReportDefect={() => setDefectModal({ partId: norm.id, stageId: 'GLAZING', location: 'IN', poId: plan.id })}
                      />
                    );
                  })}
                </div>

                <div className="mt-8 space-y-3">
                  <h4 className="font-bold text-gray-900 uppercase tracking-tight flex items-center gap-2 px-1 sm:px-4 text-xs sm:text-base">
                    <History size={16} className="text-gray-400" />
                    Lịch sử in gần đây
                  </h4>
                  <div className="bg-white rounded-xl sm:rounded-2xl border border-gray-100 shadow-2xs overflow-hidden">
                    <div className="overflow-x-auto no-scrollbar">
                      <table className="w-full text-left text-xs sm:text-sm">
                        <thead className="bg-gray-50 border-b border-gray-100 text-gray-500 font-bold uppercase text-[10px] tracking-wider">
                          <tr>
                            <th className="px-3 sm:px-6 py-2.5 sm:py-3.5">Thời gian</th>
                            <th className="px-3 sm:px-6 py-2.5 sm:py-3.5">Tên linh kiện / Gói</th>
                            <th className="px-3 sm:px-6 py-2.5 sm:py-3.5 text-center">Số lượng</th>
                            <th className="px-3 sm:px-6 py-2.5 sm:py-3.5">Thao tác</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-50">
                          {labels && labels.filter((l: any) => l.planId === selectedPlanIdForPrint).sort((a: any, b: any) => b.timestamp - a.timestamp).map((label: any) => (
                            <tr key={label.id} className="hover:bg-gray-50/50">
                              <td className="px-3 sm:px-6 py-2.5 sm:py-3.5 font-mono text-[10px] sm:text-[11px] font-bold text-gray-500 whitespace-nowrap">
                                {format(label.timestamp, 'HH:mm dd/MM')}
                              </td>
                              <td className="px-3 sm:px-6 py-2.5 sm:py-3.5">
                                <div className="font-bold text-gray-900 text-xs sm:text-sm leading-tight">{label.partName || label.partId}</div>
                                <div className="text-[10px] font-mono text-gray-400 uppercase">{label.partId}</div>
                              </td>
                              <td className="px-3 sm:px-6 py-2.5 sm:py-3.5 text-center">
                                <span className="font-bold text-blue-600 bg-blue-50 px-2.5 py-0.5 rounded-full text-xs">{label.quantity}</span>
                              </td>
                              <td className="px-3 sm:px-6 py-2.5 sm:py-3.5 whitespace-nowrap">
                                <button 
                                  onClick={() => onPrint(label)}
                                  className="flex items-center gap-1.5 text-blue-600 hover:text-blue-800 font-bold text-xs uppercase bg-white border border-blue-200 px-2.5 sm:px-3 py-1 sm:py-1.5 rounded-lg hover:border-blue-400 transition-all shadow-2xs active:scale-95 cursor-pointer"
                                >
                                  <Printer size={13} />
                                  <span>In lại</span>
                                </button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              </div>
            )}
            
            {!selectedPlanIdForPrint && (
              <div className="py-12 flex flex-col items-center justify-center text-gray-300 bg-gray-50 rounded-2xl border-2 border-dashed border-gray-200">
                <Printer size={48} strokeWidth={1} />
                <p className="text-sm sm:text-base font-bold mt-3 uppercase tracking-wider opacity-60">Chọn kế hoạch để bắt đầu in</p>
              </div>
            )}
          </div>
        )}

        {activeTab === 'CONFIG' && (
          <div className="space-y-12">
            <div className="space-y-6 bg-blue-50/30 p-8 rounded-3xl border border-blue-100">
              <hgroup>
                <h3 className="text-xl font-black uppercase text-blue-800 tracking-tight">Định mức kế hoạch dán kính</h3>
                <p className="text-xs text-blue-600 font-bold uppercase tracking-widest mt-1 italic">Dành cho lập kế hoạch gói cánh / gói đỉnh</p>
              </hgroup>
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                <div className="p-8 border-2 border-dashed border-blue-200 rounded-2xl bg-white text-center hover:bg-blue-50 transition-all group overflow-hidden">
                  <input type="file" accept=".xlsx,.xls" onChange={handleGlazingPlanImport} className="hidden" id="glazing-plan-norm-upload" />
                  <label htmlFor="glazing-plan-norm-upload" className="cursor-pointer flex flex-col items-center justify-center gap-4">
                    <FileSpreadsheet size={40} className="text-blue-500 group-hover:scale-110 transition-transform" />
                    <div>
                      <div className="font-black text-blue-600 uppercase tracking-tight">Tải lên định mức dán kính</div>
                      <div className="text-[10px] text-gray-400 font-bold uppercase tracking-widest mt-2 leading-relaxed">
                        Tên linh kiện | Định mức | Mã linh kiện | Model áp dụng
                      </div>
                    </div>
                  </label>
                </div>

                <div className="bg-white rounded-2xl border border-blue-100 shadow-sm flex flex-col max-h-[300px]">
                  <div className="p-4 border-b border-blue-50 flex justify-between items-center bg-blue-50/50">
                    <span className="text-[10px] font-black uppercase text-blue-600 tracking-widest">Định mức đã tải ({glazingPlanNorms.length})</span>
                    {glazingPlanNorms.length > 0 && (
                      <button 
                        onClick={() => { if(confirm('Xóa sạch định mức?')) { storageService.saveGlazingPlanNorms([]); setGlazingPlanNorms([]); } }}
                        className="text-red-400 hover:text-red-600 transition-colors"
                      >
                        <Trash2 size={16} />
                      </button>
                    )}
                  </div>
                  <div className="overflow-y-auto scrollbar-hide flex-1">
                    {glazingPlanNorms.length > 0 ? (
                      <table className="w-full text-left text-[11px]">
                        <thead className="bg-gray-50 font-black uppercase text-gray-400">
                          <tr>
                            <th className="px-3 py-2">Linh kiện</th>
                            <th className="px-3 py-2 text-center">T/G (Giây)</th>
                            <th className="px-3 py-2">Model áp dụng</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-50">
                          {glazingPlanNorms.map((n, i) => (
                            <tr key={i} className="hover:bg-blue-50/20">
                              <td className="px-3 py-2 font-bold text-gray-700">
                                <div>{n.partName}</div>
                              </td>
                              <td className="px-3 py-2 text-center font-black text-blue-600">{n.norm}</td>
                              <td className="px-3 py-2 font-bold text-blue-600">{n.appliedModel || n.modelName}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    ) : (
                      <div className="h-full flex items-center justify-center text-gray-400 italic text-[10px] uppercase font-bold tracking-widest">
                        Chưa có dữ liệu định mức
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 sm:gap-8">
              <div className="space-y-4 sm:space-y-6">
                <h3 className="text-base sm:text-xl font-bold uppercase text-blue-800">Cấu hình linh kiện (Kho IN)</h3>
                <div className="p-6 border-2 border-dashed border-gray-300 rounded-xl bg-gray-50 text-center">
                  <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} className="hidden" id="glazing-upload" />
                  <label htmlFor="glazing-upload" className="cursor-pointer flex flex-col items-center justify-center gap-4">
                    <FileSpreadsheet size={48} className="text-blue-500 opacity-50" />
                    <div>
                      <div className="font-bold text-blue-600 text-lg">Click tải lên Excel (IN)</div>
                      <div className="text-sm text-gray-500 mt-2">Cột yêu cầu: Tên linh kiện | Mã linh kiện | Công đoạn tiếp nhận</div>
                    </div>
                  </label>
                </div>

                {configs.length > 0 && (
                  <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-2xs max-h-[350px] overflow-y-auto">
                    <div className="p-3 bg-gray-50 border-b border-gray-200 flex justify-between items-center">
                      <span className="text-xs font-bold uppercase text-gray-700">Đã nạp {configs.length} linh kiện tiếp nhận IN</span>
                      <button 
                        onClick={() => {
                          if (confirm('Xóa toàn bộ cấu hình tiếp nhận linh kiện kho IN?')) {
                            storageService.saveGlazingConfigs([]);
                            setConfigs([]);
                          }
                        }}
                        className="text-red-500 hover:text-red-700 text-xs font-bold flex items-center gap-1 cursor-pointer"
                      >
                        <Trash2 size={13} /> Xóa tất cả
                      </button>
                    </div>
                    <table className="w-full text-left bg-white text-xs">
                      <thead>
                        <tr className="bg-gray-100/70 text-gray-500 uppercase text-[10px] tracking-wider border-b">
                          <th className="p-2.5">Tên linh kiện</th>
                          <th className="p-2.5">Mã linh kiện</th>
                          <th className="p-2.5">Kho OUT nguồn</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-100">
                        {configs.map((c) => (
                          <tr key={c.partId} className="hover:bg-blue-50/30">
                            <td className="p-2.5 font-bold text-gray-800">{c.partName || '-'}</td>
                            <td className="p-2.5 font-mono text-gray-600">{c.partId}</td>
                            <td className="p-2.5">
                              <span className={cn(
                                "px-2 py-0.5 rounded text-[10px] font-black uppercase",
                                c.sourceStageId === 'BENDING' ? "bg-amber-100 text-amber-800" :
                                c.sourceStageId === 'PAINTING' ? "bg-purple-100 text-purple-800" :
                                "bg-blue-100 text-blue-800"
                              )}>
                                {STAGES.find(s => s.id === c.sourceStageId)?.name || c.sourceStageId}
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>

              <div className="space-y-4 sm:space-y-6">
                <h3 className="text-base sm:text-xl font-bold uppercase text-orange-800">Cấu hình Gói Thành Phẩm (Kho OUT)</h3>
                <div className="p-6 border-2 border-dashed border-gray-300 rounded-xl bg-gray-50 text-center">
                  <input type="file" accept=".xlsx,.xls" onChange={handleOutFileUpload} className="hidden" id="glazing-upload-out" />
                  <label htmlFor="glazing-upload-out" className="cursor-pointer flex flex-col items-center justify-center gap-4">
                    <FileSpreadsheet size={48} className="text-orange-500 opacity-50" />
                    <div>
                      <div className="font-bold text-orange-600 text-lg">Click tải lên Excel (OUT)</div>
                      <div className="text-sm text-gray-500 mt-2">Tên linh kiện thành phẩm | Tên linh kiện thành phẩm con | Mã...</div>
                    </div>
                  </label>
                </div>

                {outConfigs.length > 0 && (
                  <div className="space-y-4">
                    {outConfigs.map(c => (
                      <div key={c.finalPartName} className="border border-gray-200 rounded-lg overflow-hidden">
                        <div className="bg-orange-50 font-bold p-3 border-b border-gray-200">
                          {c.finalPartName}
                        </div>
                        <div className="p-3 bg-white space-y-2">
                          {c.subParts.map((sp, idx) => (
                            <div key={`${sp.partId}-${idx}`} className="flex justify-between text-sm border-b border-gray-100 last:border-0 pb-2 last:pb-0">
                              <span className="text-gray-600">{sp.partName}</span>
                              <span className="font-mono text-xs opacity-70">{sp.partId}</span>
                            </div>
                          ))}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {activeTab === 'INVENTORY' && (
          <div>
            <h3 className="text-sm sm:text-base md:text-lg font-bold uppercase mb-3 sm:mb-4 flex items-center gap-2">
              <Boxes size={18} className="text-blue-600" />
              <span>Tồn kho IN Dán Kính</span>
            </h3>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2.5 sm:gap-4">
              {inventory.filter(i => i.location === 'IN' && i.stageId === 'GLAZING' && i.quantity > 0).map((i, idx) => (
                <div key={`${i.partId}-${idx}`} className="bg-white p-3 sm:p-4 rounded-xl border border-gray-200 shadow-2xs flex justify-between items-center hover:border-blue-300 transition-all">
                  <div className="min-w-0 mr-2">
                    <div className="font-bold text-gray-900 leading-tight mb-0.5 text-xs sm:text-sm truncate">{parts.find((p: any) => p.id === i.partId)?.name || i.partId}</div>
                    <div className="font-mono text-[10px] sm:text-xs text-gray-400 bg-gray-100 px-1.5 py-0.5 rounded inline-block">{i.partId}</div>
                  </div>
                  <div className="flex items-center gap-2 sm:gap-4 shrink-0">
                    <div className="font-black text-lg sm:text-2xl text-blue-600">{i.quantity}</div>
                    <button
                      onClick={() => setShowEditInventoryModal({ partId: i.partId, quantity: i.quantity, stageId: 'GLAZING', location: 'IN' })}
                      className="p-1.5 text-blue-600 hover:bg-blue-50 rounded-lg transition-colors cursor-pointer"
                      title="Chỉnh sửa tồn kho"
                    >
                      <Edit2 size={15} />
                    </button>
                    <button
                      onClick={() => handleDeleteGlazingInventoryItem(i)}
                      className="p-1.5 text-red-500 hover:bg-red-50 rounded-lg transition-colors cursor-pointer"
                      title="Xóa tồn kho"
                    >
                      <Trash2 size={15} />
                    </button>
                  </div>
                </div>
              ))}
              {inventory.filter(i => i.location === 'IN' && i.stageId === 'GLAZING' && i.quantity > 0).length === 0 && (
                <div className="col-span-full py-8 text-center text-gray-400 border-2 border-dashed border-gray-100 rounded-xl text-xs sm:text-sm italic">
                  Không có linh kiện nào trong kho IN.
                </div>
              )}
            </div>
          </div>
        )}

        {activeTab === 'INBOUND' && (() => {
          const norm = (s: string) => (s || '').normalize('NFC').trim().toUpperCase().replace(/\s*-\s*(H|CD|S|C|GLZ|OUT|IN)$/i, '');
          const configsWithStock = [...configs].map(c => {
            const cleanTargetId = norm(c.partId);
            const cleanTargetName = c.partName ? norm(c.partName) : '';

            const matchingStocks = globalInventory.filter((i: any) => {
              if (i.stageId !== c.sourceStageId || i.location !== 'OUT' || (i.quantity || 0) <= 0) return false;
              const cleanInvId = norm(i.partId);
              if (cleanInvId === cleanTargetId) return true;
              if (i.originalPartId && norm(i.originalPartId) === cleanTargetId) return true;
              if (cleanTargetName) {
                const catalogPart = parts.find((p: any) => norm(p.id) === cleanInvId);
                if (catalogPart && norm(catalogPart.name) === cleanTargetName) return true;
              }
              return false;
            });

            const sourceQty = matchingStocks.reduce((sum: number, i: any) => sum + (i.quantity || 0), 0);
            return { ...c, sourceQty };
          }).sort((a, b) => {
            if (a.sourceQty > 0 && b.sourceQty <= 0) return -1;
            if (a.sourceQty <= 0 && b.sourceQty > 0) return 1;
            return 0;
          });

          const availableCount = configsWithStock.filter(c => c.sourceQty > 0).length;

          return (
            <div className="space-y-4">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-gradient-to-r from-blue-50 via-sky-50 to-indigo-50 p-3.5 sm:p-4 rounded-xl border border-blue-200 shadow-2xs">
                <div className="flex items-center gap-2.5">
                  <div className="w-9 h-9 rounded-xl bg-blue-600 text-white flex items-center justify-center shadow-xs shrink-0">
                    <ArrowDownLeft size={20} />
                  </div>
                  <div>
                    <h3 className="text-xs sm:text-base font-bold uppercase text-blue-900 leading-tight">
                      Tiếp nhận linh kiện từ công đoạn khác vào IN
                    </h3>
                    <p className="text-[11px] text-gray-500 font-medium mt-0.5">
                      Linh kiện cấu hình sẽ <span className="text-emerald-700 font-bold">tự động sáng lên</span> khi công đoạn nguồn có hàng tại kho OUT
                    </p>
                  </div>
                </div>
                {availableCount > 0 ? (
                  <div className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-emerald-100 text-emerald-800 border border-emerald-300 rounded-xl text-xs font-black shadow-xs self-start sm:self-auto animate-pulse">
                    <Sparkles size={14} className="text-emerald-600 shrink-0" />
                    <span>{availableCount} / {configs.length} linh kiện có sẵn tại OUT nguồn!</span>
                  </div>
                ) : (
                  <div className="text-[11px] text-gray-400 font-bold self-start sm:self-auto px-2.5 py-1 bg-white/80 rounded-lg border border-gray-200">
                    Tất cả OUT nguồn hiện đang bằng 0
                  </div>
                )}
              </div>

              {/* Mobile View: Clean Touch-Friendly Cards */}
              <div className="sm:hidden space-y-2.5">
                {configsWithStock.map(c => {
                  const hasStock = c.sourceQty > 0;
                  return (
                    <div 
                      key={c.partId} 
                      className={cn(
                        "p-3.5 rounded-xl border transition-all space-y-2 relative overflow-hidden",
                        hasStock 
                          ? "bg-gradient-to-br from-emerald-50 via-teal-50/60 to-white border-2 border-emerald-500 shadow-md ring-2 ring-emerald-300/60" 
                          : "bg-white border-gray-200 shadow-2xs opacity-75"
                      )}
                    >
                      {hasStock && (
                        <div className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-emerald-600 text-white text-[10px] font-black uppercase tracking-wider shadow-xs mb-1">
                          <Sparkles size={11} className="text-yellow-300 shrink-0" />
                          <span>Có hàng sẵn sàng: {c.sourceQty}</span>
                        </div>
                      )}
                      <div className="flex justify-between items-start gap-2">
                        <div className="min-w-0">
                          <div className={cn("font-bold text-xs leading-tight", hasStock ? "text-emerald-950 font-black text-sm" : "text-gray-900")}>
                            {c.partName}
                          </div>
                          <div className="font-mono text-[10px] text-gray-400">{c.partId}</div>
                          <span className={cn(
                            "inline-block mt-1 px-1.5 py-0.5 rounded text-[10px] font-bold",
                            hasStock ? "bg-emerald-100 text-emerald-800 border border-emerald-200" : "bg-gray-100 text-gray-600"
                          )}>
                            Từ: {STAGES.find(s => s.id === c.sourceStageId)?.name || c.sourceStageId}
                          </span>
                        </div>
                        <div className="text-right shrink-0">
                          <div className={cn("text-[10px] font-bold uppercase", hasStock ? "text-emerald-700" : "text-gray-400")}>
                            OUT Nguồn
                          </div>
                          <div className={cn("font-black text-base font-mono", hasStock ? "text-emerald-700 text-lg" : "text-gray-400")}>
                            {c.sourceQty}
                          </div>
                        </div>
                      </div>
                      <div className="flex gap-2 items-center pt-2 border-t border-gray-100">
                        <input 
                          type="number"
                          value={inboundQty[c.partId] || ''}
                          onChange={e => setInboundQty({...inboundQty, [c.partId]: e.target.value})}
                          placeholder="SL"
                          min="1"
                          max={c.sourceQty}
                          className={cn(
                            "w-20 px-2.5 py-1.5 border rounded-lg text-center font-bold text-xs outline-none transition-all",
                            hasStock 
                              ? "bg-white border-emerald-400 focus:border-emerald-600 text-emerald-900 font-black" 
                              : "bg-gray-50 border-gray-200 text-gray-400"
                          )}
                        />
                        <button 
                          onClick={() => handleReceive(c)}
                          disabled={!hasStock}
                          className={cn(
                            "flex-1 py-1.5 px-3 rounded-lg font-bold text-xs uppercase text-white transition-all active:scale-95 cursor-pointer shadow-xs",
                            hasStock 
                              ? "bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 shadow-emerald-500/30 font-black" 
                              : "bg-gray-300 text-gray-500 cursor-not-allowed"
                          )}
                        >
                          Nhập IN
                        </button>
                      </div>
                    </div>
                  );
                })}
                {configs.length === 0 && (
                  <div className="text-center text-gray-400 py-8 text-xs italic bg-gray-50 rounded-xl border border-dashed border-gray-200">
                    Chưa có cấu hình tiếp nhận linh kiện. Vui lòng tải file ở tab Cấu hình.
                  </div>
                )}
              </div>

              {/* Desktop View: Standard Table */}
              <div className="hidden sm:block overflow-x-auto no-scrollbar">
                <table className="w-full text-left text-sm">
                  <thead>
                    <tr className="bg-gray-50 text-gray-500 uppercase text-[10px] tracking-wider border-b border-gray-100">
                      <th className="p-3">Linh kiện</th>
                      <th className="p-3 w-40 text-center">Slg tại OUT nguồn</th>
                      <th className="p-3 w-56">Thao tác</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-50">
                    {configsWithStock.map(c => {
                      const hasStock = c.sourceQty > 0;
                      return (
                        <tr 
                          key={c.partId} 
                          className={cn(
                            "transition-all",
                            hasStock 
                              ? "bg-gradient-to-r from-emerald-50/90 via-teal-50/40 to-white hover:bg-emerald-100/50 border-l-4 border-l-emerald-500 shadow-2xs font-medium" 
                              : "hover:bg-gray-50/50 opacity-75"
                          )}
                        >
                          <td className="p-3">
                            <div className="flex items-center gap-2">
                              <div className={cn("font-bold text-sm", hasStock ? "text-emerald-950 font-black" : "text-gray-900")}>
                                {c.partName}
                              </div>
                              {hasStock && (
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-800 text-[10px] font-black uppercase tracking-wider border border-emerald-300 animate-pulse">
                                  <Sparkles size={10} className="text-emerald-600" /> Sẵn sàng nhập
                                </span>
                              )}
                            </div>
                            <div className="font-mono text-[10px] text-gray-400">{c.partId}</div>
                            <div className={cn(
                              "inline-block mt-1 px-1.5 py-0.5 rounded text-[10px] font-bold",
                              hasStock ? "bg-emerald-100/80 text-emerald-800 border border-emerald-200" : "bg-gray-100 text-gray-600"
                            )}>
                              Từ: {STAGES.find(s => s.id === c.sourceStageId)?.name || c.sourceStageId}
                            </div>
                          </td>
                          <td className="p-3 text-center">
                            <div className={cn("font-black text-xl font-mono", hasStock ? "text-emerald-700 drop-shadow-xs" : "text-gray-400")}>
                              {c.sourceQty}
                            </div>
                            {hasStock && (
                              <div className="text-[9px] font-black uppercase text-emerald-600">Có tồn OUT</div>
                            )}
                          </td>
                          <td className="p-3">
                            <div className={cn(
                              "flex rounded-xl overflow-hidden border p-0.5 max-w-[200px] transition-all",
                              hasStock ? "bg-white border-emerald-400 shadow-sm" : "bg-gray-50 border-gray-200"
                            )}>
                              <input 
                                type="number"
                                value={inboundQty[c.partId] || ''}
                                onChange={e => setInboundQty({...inboundQty, [c.partId]: e.target.value})}
                                placeholder="SL"
                                min="1"
                                max={c.sourceQty}
                                className={cn(
                                  "bg-transparent border-none p-2 w-16 outline-none text-center font-bold text-xs",
                                  hasStock ? "text-emerald-900 font-black" : "text-gray-500"
                                )}
                              />
                              <button 
                                onClick={() => handleReceive(c)}
                                disabled={!hasStock}
                                className={cn(
                                  "flex-1 py-1.5 px-3 font-bold uppercase text-xs text-white transition-all rounded-lg active:scale-95 cursor-pointer shadow-xs", 
                                  hasStock 
                                    ? "bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 shadow-emerald-500/20 font-black" 
                                    : "bg-gray-300 text-gray-500 cursor-not-allowed"
                                )}
                              >
                                Nhập
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          );
        })()}

        {activeTab === 'OUTBOUND' && (
          <div className="space-y-8 sm:space-y-12">
            <div>
              <h3 className="text-xs sm:text-base font-bold uppercase mb-3 sm:mb-4 bg-orange-50 text-orange-800 p-3 sm:p-4 rounded-xl border border-orange-100 flex items-center gap-2">
                <Package size={18} />
                <span>Đóng gói thành phẩm (Chuyển từ IN sang OUT)</span>
              </h3>
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-3 sm:gap-6 mt-3">
                {outConfigs.map(c => {
                  const requiredQtys: Record<string, number> = {};
                  for (const sp of c.subParts) {
                    if (!requiredQtys[sp.partId]) requiredQtys[sp.partId] = 0;
                    requiredQtys[sp.partId] += 1;
                  }

                  let maxPossible = Infinity;
                  for (const partId of Object.keys(requiredQtys)) {
                    const reqPerItem = requiredQtys[partId];
                    const invQt = inventory.find(i => i.partId === partId && i.location === 'IN' && i.stageId === 'GLAZING')?.quantity || 0;
                    const possible = Math.floor(invQt / reqPerItem);
                    if (possible < maxPossible) maxPossible = possible;
                  }
                  if (maxPossible === Infinity) maxPossible = 0;

                  return (
                    <div key={c.finalPartName} className="border border-gray-200 rounded-xl overflow-hidden shadow-2xs bg-white">
                      <div className="bg-gray-50 border-b border-gray-200 p-3 sm:p-4">
                        <div className="font-bold text-sm sm:text-base text-gray-900">{c.finalPartName}</div>
                        <div className="text-xs font-bold text-orange-600 mt-0.5">Có thể đóng gói: {maxPossible}</div>
                      </div>
                      <div className="p-3 sm:p-4 space-y-2.5">
                        <div className="text-[10px] font-bold uppercase text-gray-400">Linh kiện yêu cầu (từ kho IN Dán Kính)</div>
                        {c.subParts.map((sp, idx) => {
                          const invQt = inventory.find(i => i.partId === sp.partId && i.location === 'IN' && i.stageId === 'GLAZING')?.quantity || 0;
                          return (
                            <div key={`${sp.partId}-${idx}`} className="flex justify-between items-center text-xs sm:text-sm">
                              <div className="min-w-0 mr-2">
                                <div className="font-medium text-gray-700 truncate">{sp.partName}</div>
                                <div className="font-mono text-[10px] text-gray-400">{sp.partId}</div>
                              </div>
                              <div className={cn("font-bold shrink-0", invQt > 0 ? "text-green-600" : "text-red-500")}>
                                Kho IN: {invQt}
                              </div>
                            </div>
                          );
                        })}
                        <div className="pt-3 border-t border-gray-100 flex gap-2">
                          <input 
                            type="number"
                            value={outboundQty[c.finalPartName] || ''}
                            onChange={e => setOutboundQty({...outboundQty, [c.finalPartName]: e.target.value})}
                            placeholder="SL"
                            min="1"
                            max={maxPossible || 1}
                            disabled={maxPossible === 0}
                            className="w-16 sm:w-20 px-2.5 py-1.5 sm:py-2 bg-gray-50 border border-gray-200 rounded-lg font-bold text-xs sm:text-sm outline-none focus:border-blue-500 disabled:opacity-50 text-center"
                          />
                          <button 
                            onClick={() => handleTransferOut(c.finalPartName)}
                            disabled={maxPossible === 0}
                            className="flex-1 bg-blue-600 text-white font-bold py-1.5 sm:py-2 px-3 rounded-lg hover:bg-blue-700 transition-colors uppercase text-xs disabled:bg-gray-400 disabled:cursor-not-allowed active:scale-95 cursor-pointer"
                          >
                            Đóng gói & Nhập OUT
                          </button>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
              {outConfigs.length === 0 && (
                <div className="text-center text-gray-400 py-8 text-xs italic bg-gray-50 rounded-xl border border-dashed border-gray-200">
                  Chưa có cấu hình Gói Thành Phẩm OUT. Vui lòng tải lên ở tab Cấu hình.
                </div>
              )}
            </div>

            {/* Warehouse OUT and Export to DCLR section */}
            <div className="pt-6 sm:pt-8 border-t-2 border-dashed border-gray-200">
              <h3 className="text-xs sm:text-base font-bold uppercase mb-4 bg-gray-900 text-white p-3 sm:p-4 rounded-xl border border-gray-800 flex items-center gap-2">
                <ArrowUpRight size={18} />
                <span>Linh kiện/Gói thành phẩm sẵn sàng xuất sang DCLR</span>
              </h3>
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3 sm:gap-6">
                {aggregatedGlazingOut.map(i => {
                  const isPseudo = i.partId.startsWith('GLZ-OUT-');
                  const displayName = isPseudo ? i.partId.replace('GLZ-OUT-', '') : (parts.find((p: any) => p.id === i.partId)?.name || i.partId);
                  
                  return (
                    <div key={i.partId} className="bg-white p-3.5 sm:p-5 rounded-xl sm:rounded-2xl shadow-2xs border border-orange-100 flex flex-col sm:flex-row justify-between items-stretch sm:items-center gap-3 group hover:border-orange-300 transition-all">
                      <div className="flex-1 min-w-0">
                        <div className="font-bold text-gray-900 text-sm sm:text-base leading-tight truncate">{displayName}</div>
                        {isPseudo ? (
                          <div className="font-mono text-[9px] uppercase font-bold text-orange-600 mt-0.5 bg-orange-100 px-1.5 py-0.5 rounded inline-block">Gói thành phẩm</div>
                        ) : (
                          <div className="font-mono text-[10px] text-gray-400 mt-0.5">{i.partId}</div>
                        )}
                        <div className="mt-1 text-2xl sm:text-3xl font-black text-[#F27D26]">{i.quantity}</div>
                      </div>
                      
                      <div className="flex items-center gap-2 w-full sm:w-auto shrink-0">
                        <div className="flex bg-white rounded-lg border-2 border-orange-200 overflow-hidden shadow-2xs h-10 sm:h-11 w-full sm:w-auto">
                          <input 
                            type="number"
                            value={glazingOutQty[i.partId] || ''}
                            onChange={e => setGlazingOutQty({...glazingOutQty, [i.partId]: e.target.value})}
                            placeholder="SL"
                            className="w-16 sm:w-20 px-2 text-center font-black outline-none border-r border-orange-100 text-sm sm:text-base text-orange-600"
                          />
                          <button 
                            onClick={() => handleGlazingExport(i)}
                            className="flex-1 sm:flex-none px-3 sm:px-5 bg-[#F27D26] text-white font-bold text-xs uppercase hover:bg-orange-700 transition-all active:scale-95 flex items-center justify-center gap-1.5 whitespace-nowrap cursor-pointer"
                          >
                            <ArrowUpRight size={14} />
                            <span>XUẤT DCLR</span>
                          </button>
                        </div>
                      </div>
                    </div>
                  );
                })}
                {aggregatedGlazingOut.length === 0 && (
                  <div className="col-span-full text-center py-8 text-gray-400 text-xs italic bg-gray-50 rounded-xl border border-dashed border-gray-200">
                    Chưa có linh kiện nào trong kho OUT Dán Kính...
                  </div>
                )}
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Edit Inventory Modal */}
      <AnimatePresence>
        {showEditInventoryModal && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="bg-white rounded-2xl p-8 max-w-md w-full shadow-2xl space-y-6"
            >
              <div className="flex items-center gap-4 text-blue-600">
                <Edit2 size={32} />
                <h3 className="text-xl font-bold">Chỉnh sửa Tồn Kho</h3>
              </div>
              <div className="space-y-4">
                <div className="space-y-2">
                  <label className="text-xs font-bold uppercase opacity-50">Số lượng tồn kho (Đặt 0 để xóa)</label>
                  <input 
                    type="number"
                    min="0"
                    value={showEditInventoryModal.quantity}
                    onChange={e => setShowEditInventoryModal({...showEditInventoryModal, quantity: parseInt(e.target.value) || 0})}
                    className="w-full p-4 rounded-lg border border-gray-200 focus:border-blue-500 outline-none text-lg font-bold"
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-bold uppercase opacity-50">Mật khẩu xác nhận (Admin)</label>
                  <input 
                    type="password"
                    autoComplete="off"
                    value={editPassword}
                    onChange={e => setEditPassword(e.target.value)}
                    className="w-full p-4 rounded-lg border border-gray-200 focus:border-blue-500 outline-none text-lg"
                    placeholder="Nhập mật khẩu..."
                    onKeyDown={(e) => e.key === 'Enter' && handleEditInventory()}
                  />
                </div>
              </div>
              <div className="flex gap-3 pt-2">
                <button 
                  onClick={() => { setShowEditInventoryModal(null); setEditPassword(''); }}
                  className="flex-1 py-4 bg-gray-100 rounded-lg font-bold text-sm uppercase hover:bg-gray-200 transition-all"
                >
                  Hủy
                </button>
                <button 
                  onClick={handleEditInventory}
                  disabled={showEditInventoryModal.quantity < 0 || !editPassword}
                  className="flex-1 py-4 bg-blue-600 text-white rounded-lg font-bold text-sm uppercase hover:bg-blue-700 transition-all shadow-lg disabled:opacity-50"
                >
                  Cập nhật
                </button>
              </div>
            </motion.div>
          </div>
        )}

        {/* Modal Import Excel Kế Hoạch Dán Kính */}
        {excelGlazingImportList && (
          <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
            <motion.div 
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white rounded-3xl p-6 md:p-8 max-w-5xl w-full shadow-2xl border border-gray-100 flex flex-col max-h-[90vh]"
            >
              {/* Header */}
              <div className="flex items-center justify-between pb-4 border-b border-gray-100">
                <div className="flex items-center gap-3">
                  <div className="w-12 h-12 rounded-2xl bg-emerald-100 text-emerald-700 flex items-center justify-center">
                    <FileSpreadsheet size={24} />
                  </div>
                  <div>
                    <h3 className="text-xl font-bold text-gray-900">Xem Trước Kế Hoạch Sản Xuất Dán Kính Từ Excel</h3>
                    <p className="text-xs text-gray-500 font-medium">Tự động phân tách kế hoạch theo từng ngày, lùi 48 giờ thời gian mong muốn hoàn thành và áp dụng định mức dán kính</p>
                  </div>
                </div>
                <button 
                  onClick={() => setExcelGlazingImportList(null)}
                  className="p-2 hover:bg-gray-100 rounded-xl text-gray-400 hover:text-gray-700 transition-colors"
                >
                  <X size={20} />
                </button>
              </div>

              {/* Summary Stats */}
              <div className="grid grid-cols-2 md:grid-cols-4 gap-4 my-4">
                <div className="bg-emerald-50/80 border border-emerald-100 p-4 rounded-2xl">
                  <span className="text-xs font-bold uppercase text-emerald-600 block mb-1">Số Kế Hoạch Tạo Ra</span>
                  <span className="text-2xl font-black text-emerald-800">{excelGlazingImportList.length}</span>
                  <span className="text-xs text-emerald-600 ml-1">kế hoạch ({new Set(excelGlazingImportList.map(o => o.modelId)).size} model)</span>
                </div>

                <div className="bg-indigo-50/80 border border-indigo-100 p-4 rounded-2xl">
                  <span className="text-xs font-bold uppercase text-indigo-600 block mb-1">Tổng SL Trong File</span>
                  <span className="text-2xl font-black text-indigo-800">
                    {(excelGlazingStats?.grandTotal ?? excelGlazingImportList.reduce((sum, o) => sum + o.targetQuantity, 0)).toLocaleString('vi-VN')}
                  </span>
                  <span className="text-xs text-indigo-600 ml-1">Máy</span>
                </div>

                <div className="bg-blue-50/80 border border-blue-100 p-4 rounded-2xl">
                  <span className="text-xs font-bold uppercase text-blue-600 block mb-1">SL Hợp Lệ Sẽ Tạo Kế Hoạch</span>
                  <span className="text-2xl font-black text-blue-800">
                    {(excelGlazingStats?.validTotal ?? excelGlazingImportList.reduce((sum, o) => sum + o.targetQuantity, 0)).toLocaleString('vi-VN')}
                  </span>
                  <span className="text-xs text-blue-600 ml-1">Máy</span>
                </div>

                <div className={cn(
                  "p-4 rounded-2xl border",
                  missingNormList.length > 0 
                    ? "bg-rose-50/80 border-rose-200" 
                    : "bg-emerald-50/80 border-emerald-100"
                )}>
                  <span className={cn(
                    "text-xs font-bold uppercase block mb-1",
                    missingNormList.length > 0 ? "text-rose-600" : "text-emerald-600"
                  )}>
                    {missingNormList.length > 0 ? 'Model Thiếu Định Mức Dán Kính' : 'Định Mức Dán Kính'}
                  </span>
                  <span className={cn(
                    "text-2xl font-black",
                    missingNormList.length > 0 ? "text-rose-700" : "text-emerald-800"
                  )}>
                    {missingNormList.length > 0 ? `${missingNormList.length} model` : '100% Đầy đủ'}
                  </span>
                  <span className={cn(
                    "text-xs ml-1 font-bold",
                    missingNormList.length > 0 ? "text-rose-600" : "text-emerald-600"
                  )}>
                    {missingNormList.length > 0 ? `(${excelGlazingStats?.missingTotal.toLocaleString('vi-VN')} Máy)` : 'hợp lệ'}
                  </span>
                </div>
              </div>

              {/* View Mode Switch (Tabs) */}
              <div className="flex items-center justify-between mb-3">
                <div className="flex bg-gray-100 p-1 rounded-xl gap-1">
                  <button
                    onClick={() => setGlazingModalTab('VALID')}
                    className={cn(
                      "px-4 py-2 rounded-lg font-bold text-xs uppercase transition-all flex items-center gap-2 cursor-pointer",
                      glazingModalTab === 'VALID' ? "bg-white text-emerald-700 shadow-sm" : "text-gray-500 hover:text-gray-700"
                    )}
                  >
                    <CheckCircle2 size={15} className="text-emerald-600" />
                    <span>Kế Hoạch Dán Kính Hợp Lệ ({excelGlazingImportList.length})</span>
                  </button>

                  <button
                    onClick={() => setGlazingModalTab('MISSING_NORM')}
                    className={cn(
                      "px-4 py-2 rounded-lg font-bold text-xs uppercase transition-all flex items-center gap-2 cursor-pointer",
                      glazingModalTab === 'MISSING_NORM' 
                        ? (missingNormList.length > 0 ? "bg-rose-600 text-white shadow-sm" : "bg-white text-blue-700 shadow-sm") 
                        : (missingNormList.length > 0 ? "text-rose-600 hover:text-rose-800 bg-rose-50" : "text-gray-500 hover:text-gray-700")
                    )}
                  >
                    <AlertCircle size={15} className={missingNormList.length > 0 ? (glazingModalTab === 'MISSING_NORM' ? "text-white" : "text-rose-600") : "text-gray-400"} />
                    <span>Model Chưa Cài Đặt Định Mức Dán Kính ({missingNormList.length})</span>
                  </button>
                </div>
              </div>

              {/* Main Content Area */}
              <div className="flex-1 overflow-y-auto border border-gray-100 rounded-2xl min-h-[220px]">
                {glazingModalTab === 'MISSING_NORM' ? (
                  <div className="p-4 space-y-4">
                    {missingNormList.length === 0 ? (
                      <div className="bg-emerald-50 border border-emerald-200 rounded-2xl p-6 text-center">
                        <CheckCircle2 size={40} className="text-emerald-600 mx-auto mb-2" />
                        <h4 className="text-base font-bold text-emerald-800">Tất cả Model đều hợp lệ!</h4>
                        <p className="text-xs text-emerald-600 mt-1">
                          100% Model trong file Excel đều đã được cài đặt Định mức kế hoạch dán kính (Gói cánh / Gói đỉnh) đầy đủ trong tab Cấu hình.
                        </p>
                      </div>
                    ) : (
                      <>
                        <div className="bg-rose-50 border border-rose-200 rounded-2xl p-4 flex items-start gap-3">
                          <AlertCircle size={24} className="text-rose-600 shrink-0 mt-0.5" />
                          <div>
                            <h4 className="text-sm font-bold text-rose-800">
                              Phát hiện {missingNormList.length} Model trong file Excel chưa có Định mức dán kính
                            </h4>
                            <p className="text-xs text-rose-600 mt-1">
                              Các model này sẽ <strong>tạm thời không được tạo Kế hoạch dán kính</strong> để đảm bảo có đầy đủ thành phần linh kiện và định mức thời gian sản xuất. Vui lòng vào tab <strong>Cấu hình &gt; Định Mức Kế Hoạch Dán Kính</strong> để tải lên file định mức bổ sung.
                            </p>
                          </div>
                        </div>

                        <table className="w-full text-left text-sm">
                          <thead className="bg-rose-50/50 text-xs font-bold uppercase text-rose-700 sticky top-0 border-b border-rose-200">
                            <tr>
                              <th className="p-3.5 text-center w-12">STT</th>
                              <th className="p-3.5">Mã Model (Cột B)</th>
                              <th className="p-3.5">Tên Sản Phẩm</th>
                              <th className="p-3.5 text-right">SL Trong File (Máy)</th>
                              <th className="p-3.5">Nguyên Nhân Lỗi / Khắc Phục</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-gray-100 font-medium">
                            {missingNormList.map((m, idx) => (
                              <tr key={m.modelId || idx} className="hover:bg-rose-50/30 transition-colors">
                                <td className="p-3.5 text-center text-gray-400 font-mono">{idx + 1}</td>
                                <td className="p-3.5 font-mono font-bold text-rose-700">{m.modelId}</td>
                                <td className="p-3.5 text-gray-700">{m.name || '-'}</td>
                                <td className="p-3.5 text-right font-bold text-gray-900 font-mono">
                                  {m.totalQty > 0 ? m.totalQty.toLocaleString('vi-VN') : '-'}
                                </td>
                                <td className="p-3.5 text-xs text-rose-600">
                                  <span className="bg-rose-100 text-rose-800 px-2.5 py-1 rounded-md font-semibold">
                                    {m.issueReason}
                                  </span>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </>
                    )}
                  </div>
                ) : (
                  <table className="w-full text-left text-sm">
                    <thead className="bg-gray-50 text-xs font-bold uppercase text-gray-500 sticky top-0 border-b border-gray-200">
                      <tr>
                        <th className="p-3.5 text-center w-12">STT</th>
                        <th className="p-3.5">Mã Kế Hoạch</th>
                        <th className="p-3.5">Model Sản Xuất</th>
                        <th className="p-3.5 text-right">Số Lượng (Máy)</th>
                        <th className="p-3.5">Thành Phần Dán Kính</th>
                        <th className="p-3.5">Bắt Đầu Dự Kiến</th>
                        <th className="p-3.5">Dự Kiến Hoàn Thành</th>
                        <th className="p-3.5">Ghi Chú Đợt Chạy</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100 font-medium">
                      {excelGlazingImportList.map((plan, index) => {
                        const modelNorms = glazingPlanNorms.filter(n => n.appliedModel === plan.modelId || n.modelName === plan.modelId);
                        return (
                          <tr key={plan.id || index} className="hover:bg-blue-50/40 transition-colors">
                            <td className="p-3.5 text-center text-gray-400 font-mono">{index + 1}</td>
                            <td className="p-3.5 font-mono font-bold text-indigo-700">{plan.id}</td>
                            <td className="p-3.5">
                              <span className="font-bold text-gray-900">{plan.modelId}</span>
                              {parts.find(p => p.id === plan.modelId)?.name && (
                                <span className="text-xs text-gray-500 block">
                                  {parts.find(p => p.id === plan.modelId)?.name}
                                </span>
                              )}
                            </td>
                            <td className="p-3.5 text-right font-black text-emerald-700 text-base">
                              {plan.targetQuantity.toLocaleString('vi-VN')}
                            </td>
                            <td className="p-3.5 text-xs text-gray-700">
                              {modelNorms.length > 0 ? (
                                <div className="flex flex-col gap-1">
                                  {modelNorms.map((n, ni) => (
                                    <span key={ni} className="bg-blue-50 text-blue-800 px-2 py-0.5 rounded text-[11px] font-semibold border border-blue-100">
                                      {n.partName} ({n.norm}s)
                                    </span>
                                  ))}
                                </div>
                              ) : (
                                <span className="text-gray-400 italic">Theo định mức chung</span>
                              )}
                            </td>
                            <td className="p-3.5 text-xs font-mono text-gray-600">
                              {plan.plannedStartTime ? format(new Date(plan.plannedStartTime), 'dd/MM/yyyy HH:mm') : '-'}
                            </td>
                            <td className="p-3.5 text-xs font-mono text-gray-800 font-bold">
                              {plan.targetCompletionTime ? format(new Date(plan.targetCompletionTime), 'dd/MM/yyyy HH:mm') : '-'}
                            </td>
                            <td className="p-3.5 text-xs text-gray-600">
                              <span className="bg-gray-100 text-gray-700 px-2.5 py-1 rounded-md font-medium">
                                {plan.note || '-'}
                              </span>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                )}
              </div>

              {/* Action Buttons */}
              <div className="flex gap-3 pt-4 border-t border-gray-100 mt-4">
                <button
                  onClick={() => setExcelGlazingImportList(null)}
                  className="flex-1 py-3.5 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-xl font-bold text-sm uppercase transition-all cursor-pointer"
                >
                  Đóng
                </button>
                <button
                  onClick={handleConfirmBatchCreateGlazingPlans}
                  disabled={isBatchCreatingGlazing || excelGlazingImportList.length === 0}
                  className="flex-2 py-3.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-bold text-sm uppercase transition-all shadow-lg shadow-emerald-200 flex items-center justify-center gap-2 active:scale-95 disabled:opacity-50 cursor-pointer"
                >
                  <CheckCircle2 size={18} />
                  <span>{isBatchCreatingGlazing ? 'Đang tạo kế hoạch...' : `Xác Nhận Tạo Hàng Loạt (${excelGlazingImportList.length} Kế Hoạch)`}</span>
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </motion.div>
  );
}

function InboundView({ selectedStage, setSelectedStage, onScanSuccess, parts, setDefectModal }: any) {
  const [targetLocation, setTargetLocation] = useState<'IN' | 'OUT'>('IN');
  const [scanInput, setScanInput] = useState('');
  const [lastScanned, setLastScanned] = useState<any>(null);
  const [isCameraOpen, setIsCameraOpen] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    inputRef.current?.focus();
  }, []);

  const handleScanSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (scanInput.trim()) {
      const input = scanInput.trim();
      processQrCodeInput(input);
      setScanInput('');
    }
  };

  const processQrCodeInput = (input: string) => {
    const qrParts = input.split('|');
    
    try {
      onScanSuccess(input, targetLocation);
      
      if (qrParts.length >= 5) {
        const pId = qrParts[0];
        const sId = qrParts[2];
        setLastScanned({
          partId: pId,
          quantity: qrParts[1],
          sourceStageId: sId,
          partName: parts.find((p: any) => p.id === pId)?.name || pId,
          sourceName: STAGES.find(s => s.id === sId)?.name || sId,
          status: 'success'
        });
      }
    } catch (err) {
      if (qrParts.length >= 5) {
        const pId = qrParts[0];
        const sId = qrParts[2];
        setLastScanned({
          partId: pId,
          quantity: qrParts[1],
          sourceStageId: sId,
          partName: parts.find((p: any) => p.id === pId)?.name || pId,
          sourceName: STAGES.find(s => s.id === sId)?.name || sId,
          status: 'error',
          errorMsg: err instanceof Error ? err.message : 'Lỗi không xác định'
        });
      } else {
        setLastScanned({
          status: 'error',
          errorMsg: 'Định dạng mã QR không hợp lệ'
        });
      }
    }
  };

  return (
    <motion.div 
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, y: -20 }}
      className="grid grid-cols-1 lg:grid-cols-2 gap-6 sm:gap-8"
    >
      <div className="bg-white p-4 sm:p-10 rounded-2xl sm:rounded-3xl border border-gray-200 shadow-xl space-y-6 sm:space-y-8">
        <div className="flex items-center gap-3 sm:gap-4 mb-2">
          <div className="bg-blue-600 p-2.5 sm:p-3 rounded-xl text-white">
            <QrCode size={22} />
          </div>
          <div>
            <h3 className="text-base sm:text-xl font-bold uppercase tracking-tight">Nhập kho bằng mã QR</h3>
            <p className="text-xs sm:text-sm text-gray-500">Dùng súng quét hoặc Camera điện thoại để nhập kho</p>
          </div>
        </div>

        {/* Camera QR Scanner Trigger Button for Mobile */}
        <button
          type="button"
          onClick={() => setIsCameraOpen(true)}
          className="w-full py-3.5 sm:py-4 px-4 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 text-white rounded-2xl font-bold text-sm sm:text-base uppercase tracking-wider transition-all flex items-center justify-center gap-2.5 shadow-lg active:scale-95 cursor-pointer"
        >
          <Camera size={22} />
          <span>Quét mã bằng Camera điện thoại</span>
        </button>

        <div className="space-y-6 sm:space-y-8">
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 sm:gap-8">
            <div className="space-y-2 sm:space-y-4 text-left">
              <label className="text-xs sm:text-sm font-bold uppercase tracking-widest opacity-50">Công đoạn tiếp nhận:</label>
              <select 
                value={selectedStage}
                onChange={(e) => setSelectedStage(e.target.value as StageId)}
                className="w-full p-3.5 sm:p-5 rounded-xl border-2 border-gray-100 font-bold text-sm sm:text-lg focus:border-blue-600 outline-none bg-white cursor-pointer"
              >
                {STAGES.map(stage => (
                  <option key={stage.id} value={stage.id}>{stage.name}</option>
                ))}
              </select>
            </div>
            <div className="space-y-2 sm:space-y-4 text-left">
              <label className="text-xs sm:text-sm font-bold uppercase tracking-widest opacity-50">Vị trí nhập:</label>
              <div className="flex bg-gray-100 p-1 rounded-xl h-[52px] sm:h-[68px]">
                <div className="flex-1 rounded-lg font-bold text-xs sm:text-sm uppercase bg-gray-900 text-white shadow-xs flex items-center justify-center">
                  Kho IN
                </div>
              </div>
            </div>
          </div>

          <form onSubmit={handleScanSubmit} className="space-y-6 sm:space-y-8">
            <div className="relative">
              <input 
                ref={inputRef}
                type="text"
                value={scanInput}
                onChange={(e) => setScanInput(e.target.value)}
                placeholder="Nhập mã hoặc quét QR..."
                className="w-full p-4 sm:p-10 rounded-2xl sm:rounded-3xl border-2 sm:border-4 border-gray-900 bg-gray-50 text-center font-mono text-base sm:text-3xl font-bold focus:ring-4 sm:ring-8 ring-black/5 outline-none placeholder:text-gray-300"
                autoComplete="off"
              />
              <div className="hidden sm:flex absolute right-8 top-1/2 -translate-y-1/2 items-center gap-3 text-xs font-mono font-bold text-blue-600 bg-blue-50 px-4 py-2 rounded-full border border-blue-100">
                <div className="w-2 h-2 rounded-full bg-blue-500 animate-pulse" />
                READY
              </div>
            </div>
            <div className="flex gap-4">
              <button 
                type="submit"
                className="flex-1 h-14 sm:h-20 bg-gray-900 text-white rounded-2xl sm:rounded-3xl font-bold text-sm sm:text-base uppercase tracking-widest hover:bg-black transition-all flex items-center justify-center gap-3 shadow-xl active:scale-95 cursor-pointer"
              >
                <CheckCircle2 size={22} />
                Xác nhận
              </button>
            </div>
          </form>
        </div>

        <div className="p-4 sm:p-8 bg-blue-50 border border-blue-100 rounded-2xl sm:rounded-3xl text-left space-y-2 sm:space-y-3">
          <div className="flex items-center gap-2 sm:gap-3 text-blue-600">
            <AlertCircle size={20} />
            <span className="font-bold text-xs sm:text-sm uppercase">Hướng dẫn</span>
          </div>
          <p className="text-xs sm:text-sm text-blue-800 leading-relaxed">
            Bạn có thể bấm nút <b>"Quét bằng Camera"</b> để quét trực tiếp từ điện thoại di động hoặc dùng súng quét mã barcode.
          </p>
        </div>

        {/* Mobile Camera Modal */}
        <MobileQRScannerModal
          isOpen={isCameraOpen}
          onClose={() => setIsCameraOpen(false)}
          onScan={(qr) => {
            processQrCodeInput(qr);
          }}
          title={`Quét mã Nhập kho (${STAGES.find(s => s.id === selectedStage)?.name || selectedStage})`}
        />
      </div>

      <div className="flex flex-col">
        <AnimatePresence mode="wait">
          {lastScanned ? (
            <motion.div
              key="scan-result"
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              className={cn(
                "p-10 rounded-3xl border-2 shadow-2xl h-full flex flex-col items-center justify-center space-y-6 bg-white",
                lastScanned.status === 'success' ? "border-green-500" : "border-red-500"
              )}
            >
              <div className={cn(
                "w-16 h-16 rounded-full flex items-center justify-center",
                lastScanned.status === 'success' ? "bg-green-100 text-green-600" : "bg-red-100 text-red-600"
              )}>
                {lastScanned.status === 'success' ? <CheckCircle2 size={32} /> : <AlertCircle size={32} />}
              </div>
              <div className="text-center space-y-2">
                <h3 className="text-xl font-bold">
                  {lastScanned.status === 'success' ? 'Đã nhập kho thành công!' : 'Lỗi nhập kho!'}
                </h3>
                <p className="text-gray-500 text-sm">
                  {lastScanned.status === 'success' 
                    ? (lastScanned.isManual ? 'Thông tin linh kiện vừa nhập tay:' : 'Thông tin linh kiện vừa quét:') 
                    : lastScanned.errorMsg}
                </p>
              </div>
              
              {lastScanned.partId && (
                <div className="w-full bg-gray-50 p-8 rounded-2xl space-y-6">
                  <div className="flex justify-between items-center border-b border-gray-200 pb-4">
                    <span className="text-sm font-mono uppercase opacity-50">Mã LK</span>
                    <span className="font-bold text-xl">{lastScanned.partId}</span>
                  </div>
                  <div className="flex justify-between items-center border-b border-gray-200 pb-4">
                    <span className="text-sm font-mono uppercase opacity-50">Linh kiện</span>
                    <span className="font-bold text-xl">{lastScanned.partName}</span>
                  </div>
                  <div className="flex justify-between items-center border-b border-gray-200 pb-4">
                    <span className="text-sm font-mono uppercase opacity-50">Số lượng</span>
                    <span className="font-bold text-2xl">{lastScanned.quantity}</span>
                  </div>
                  <div className="flex justify-between items-center">
                    <span className="text-sm font-mono uppercase opacity-50">Đích</span>
                    <div className="flex items-center gap-3 font-bold text-[#F27D26] text-xl">
                      <span>{STAGES.find(s => s.id === selectedStage)?.name}</span>
                      <ArrowRight size={20} />
                      <span className={cn(
                        "px-3 py-1 rounded text-xs",
                        targetLocation === 'IN' ? "bg-blue-100 text-blue-700" : "bg-orange-100 text-orange-700"
                      )}>
                        KHO_{targetLocation}
                      </span>
                    </div>
                  </div>
                  {!lastScanned.isManual && lastScanned.sourceName && (
                    <div className="flex justify-between items-center pt-4 border-t border-gray-200">
                      <span className="text-xs font-mono uppercase opacity-50">Nguồn từ</span>
                      <span className="text-sm font-bold opacity-60">{lastScanned.sourceName} (OUT)</span>
                    </div>
                  )}
                </div>
              )}

              <div className="flex gap-4 w-full">
                <button 
                  onClick={() => setLastScanned(null)}
                  className="flex-1 py-4 text-sm font-bold uppercase tracking-widest text-gray-400 hover:text-[#141414] transition-colors border-2 border-gray-100 rounded-xl"
                >
                  Quét tiếp
                </button>
                {lastScanned.status === 'success' && (
                  <button 
                    onClick={() => setDefectModal({ partId: lastScanned.partId, stageId: selectedStage, poId: lastScanned.poId })}
                    className="flex-1 py-4 text-sm font-bold uppercase tracking-widest text-red-600 hover:bg-red-50 transition-colors border-2 border-red-100 rounded-xl flex items-center justify-center gap-2"
                  >
                    <AlertCircle size={18} />
                    Báo lỗi (NG)
                  </button>
                )}
              </div>
            </motion.div>
          ) : (
            <div className="bg-gray-100 rounded-3xl border-2 border-dashed border-gray-300 h-full flex flex-col items-center justify-center p-12 text-center text-gray-400">
              <Package size={64} strokeWidth={1} />
              <p className="mt-6 font-medium">Chi tiết pallet sẽ hiển thị tại đây<br />ngay sau khi bạn thực hiện nhập kho.</p>
            </div>
          )}
        </AnimatePresence>
      </div>
    </motion.div>
  );
}

function SettingsView({ parts, onPartsChange, labelSettings, onLabelSettingsChange }: { 
  parts: Part[], 
  onPartsChange: () => void, 
  labelSettings: any,
  onLabelSettingsChange: (s: any) => void,
  key?: string 
}) {
  const [newPart, setNewPart] = useState<Part>({ id: '', name: '', unit: 'Cái', level: 1, skipLaser: false, skipBending: false, skipWelding: false, skipPainting: false, hasPaintingPO: false });
  const [editingId, setEditingId] = useState<string | null>(null);
  const [isImporting, setIsImporting] = useState(false);
  const [isImportingBOM, setIsImportingBOM] = useState(false);
  const [isImportingBOMV2, setIsImportingBOMV2] = useState(false);
  const [isImportingModelBOM, setIsImportingModelBOM] = useState(false);
  const [isImportingTransformations, setIsImportingTransformations] = useState(false);

  const [activeSettingsTab, setActiveSettingsTab] = useState<'parts' | 'bom' | 'label' | 'bom_v2' | 'model_bom' | 'transformations' | 'backup'>('parts');
  const [partSearchQuery, setPartSearchQuery] = useState('');
  const [partLevelFilter, setPartLevelFilter] = useState<'ALL' | 1 | 2 | 3>('ALL');

  const filteredParts = useMemo(() => {
    let result = parts;
    if (partLevelFilter !== 'ALL') {
      result = result.filter(p => (p.level || 1) === partLevelFilter);
    }
    if (partSearchQuery.trim()) {
      const q = partSearchQuery.trim().toLowerCase();
      result = result.filter(p => 
        (p.id && p.id.toLowerCase().includes(q)) || 
        (p.name && p.name.toLowerCase().includes(q)) || 
        (p.unit && p.unit.toLowerCase().includes(q))
      );
    }
    return [...result].sort((a, b) => (a.level || 0) - (b.level || 0));
  }, [parts, partSearchQuery, partLevelFilter]);

  // Backup & Restore LocalStorage State
  const [backupStats, setBackupStats] = useState(() => storageService.getStorageStats());
  const [downloadSuccessMessage, setDownloadSuccessMessage] = useState<string | null>(null);
  const [copiedJson, setCopiedJson] = useState(false);
  const [selectedBackupFile, setSelectedBackupFile] = useState<File | null>(null);
  const [importValidation, setImportValidation] = useState<{
    valid: boolean;
    error?: string;
    storageData?: Record<string, any>;
    metadata?: any;
    stats?: any;
  } | null>(null);
  const [importMode, setImportMode] = useState<'overwrite' | 'merge'>('overwrite');
  const [isProcessingImport, setIsProcessingImport] = useState(false);
  const [importResult, setImportResult] = useState<{
    success: boolean;
    message: string;
    keysRestored: number;
  } | null>(null);

  const [cleanupMessage, setCleanupMessage] = useState<string | null>(null);

  const refreshBackupStats = () => {
    setBackupStats(storageService.getStorageStats());
  };

  const handleManualCleanup = async () => {
    const res = await storageService.autoCleanupOldData();
    refreshBackupStats();
    onPartsChange();
    setCleanupMessage(`Đã dọn dẹp thành công: xóa ${res.labelsRemoved} nhãn QR (>24h), ${res.posRemoved} lệnh PO đã hoàn thành (>35 ngày) và ${res.txsRemoved} nhật ký quét (>35 ngày) trên máy và Supabase database!`);
    setTimeout(() => setCleanupMessage(null), 8000);
  };

  const handleDownloadBackup = () => {
    const filename = storageService.downloadBackupFile();
    setDownloadSuccessMessage(`Đã tải xuống thành công tệp "${filename}". Bạn có thể sao chép tệp này sang máy khác.`);
    refreshBackupStats();
    setTimeout(() => setDownloadSuccessMessage(null), 8000);
  };

  const handleCopyBackupJson = () => {
    const data = storageService.exportBackupData();
    navigator.clipboard.writeText(JSON.stringify(data, null, 2));
    setCopiedJson(true);
    setTimeout(() => setCopiedJson(false), 3000);
  };

  const handleBackupFileSelect = (file: File) => {
    setSelectedBackupFile(file);
    setImportResult(null);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const text = e.target?.result as string;
        const validation = storageService.validateBackupJson(text);
        setImportValidation(validation);
      } catch (err: any) {
        setImportValidation({
          valid: false,
          error: 'Không thể phân tích tệp JSON: ' + (err?.message || 'Lỗi không xác định')
        });
      }
    };
    reader.onerror = () => {
      setImportValidation({
        valid: false,
        error: 'Lỗi khi đọc file từ ổ đĩa máy tính.'
      });
    };
    reader.readAsText(file);
  };

  const handleExecuteImport = () => {
    if (!importValidation || !importValidation.valid || !importValidation.storageData) {
      setImportResult({
        success: false,
        keysRestored: 0,
        message: 'Vui lòng chọn tệp tin JSON hợp lệ trước.'
      });
      return;
    }

    setIsProcessingImport(true);
    setImportResult(null);

    setTimeout(() => {
      try {
        const res = storageService.importBackupData(importValidation.storageData!, importMode);
        setIsProcessingImport(false);
        setImportResult(res);

        if (res.success) {
          onPartsChange();
          const updatedLabelSettings = storageService.getLabelSettings();
          if (updatedLabelSettings) {
            onLabelSettingsChange(updatedLabelSettings);
          }
          refreshBackupStats();
        }
      } catch (err: any) {
        setIsProcessingImport(false);
        setImportResult({
          success: false,
          keysRestored: 0,
          message: 'Lỗi khi nạp dữ liệu: ' + (err?.message || String(err))
        });
      }
    }, 150);
  };

  const [showResetModal, setShowResetModal] = useState(false);
  const [resetPassword, setResetPassword] = useState('');

  const handleResetData = async () => {
    if (resetPassword === 'admin123') {
      await storageService.resetAllData();
      window.location.reload();
    } else {
      alert('Mật khẩu không chính xác!');
    }
  };

  const handleBOMImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsImportingBOM(true);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws) as any[];

        const bomMap = new Map<string, BOMDefinition>();
        
        data.forEach(row => {
          let childId = '';
          let parentId = '';
          let weightVal: any = 0;

          Object.keys(row).forEach(key => {
            const k = key.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, '').replace(/\n/g, '');
            if (k.includes('malinhkiencon') || k === 'malinhkien' || k.includes('childid') || k.includes('macon')) {
               childId = row[key];
            } else if (k.includes('malinhkiencha') || k.includes('malinhkienton') || k.includes('parentid') || k.includes('macha')) {
               parentId = row[key];
            } else if (k.includes('khoiluong') || k.includes('componentweight') || k.includes('kg') || k.includes('weight')) {
               weightVal = row[key];
            }
          });

          // Fallbacks for direct access just in case
          if (!childId) childId = row['Mã linh kiện con'] || row['Mã linh kiện'] || row['ChildID'] || row['Mã con'];
          if (!parentId) parentId = row['Mã linh kiện cha'] || row['Mã linh kiện tôn'] || row['ParentID'] || row['Mã cha'];
          if (!weightVal) weightVal = row['Khối lượng (Kg)'] ?? row['Khối lượng'] ?? row['ComponentWeight'] ?? row['Khối lượng linh kiện'];

          childId = String(childId || '').trim();
          parentId = String(parentId || '').trim();
          
          if (!childId) return;

          let weight = 0;
          if (typeof weightVal === 'string') {
            const clean = weightVal.replace(/[()]/g, '').replace(/,/g, '.');
            weight = parseFloat(clean) || 0;
          } else if (typeof weightVal === 'number') {
            weight = Math.abs(weightVal);
          }
          
          if (!bomMap.has(childId)) {
            bomMap.set(childId, {
              parentPartId: '',
              childPartId: childId,
              componentWeight: 0,
              scrapWeight: 0
            });
          }

          const def = bomMap.get(childId)!;
          
          const upperParentId = parentId.toUpperCase();
          if (upperParentId.includes('PL-') || upperParentId.includes('PHẾ') || upperParentId.includes('P-')) {
            def.scrapWeight += weight;
          } else {
            if (parentId && !def.parentPartId) def.parentPartId = parentId; // Keep the first valid parent ID
            def.componentWeight += weight;
          }
        });

        const importedBOM = Array.from(bomMap.values()).filter(b => b.childPartId && (b.componentWeight > 0 || b.scrapWeight > 0));

        if (importedBOM.length === 0) {
          alert('Không tìm thấy dữ liệu định mức hợp lệ. Vui lòng kiểm tra tiêu đề cột (Mã linh kiện, Mã linh kiện tôn, Khối lượng (Kg)).\n\nDữ liệu thô dòng 1:\n' + JSON.stringify(data[0] || {}));
          setIsImportingBOM(false);
          return;
        }

        storageService.saveBOM(importedBOM);
        onPartsChange();
        alert(`Đã nhập thành công ${importedBOM.length} định mức sản xuất!`);
      } catch (err) {
        if ((import.meta as any).env?.DEV) console.error('BOM Import Error:', err);
        alert('Lỗi khi đọc file định mức.');
      } finally {
        setIsImportingBOM(false);
        if (e.target) e.target.value = '';
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleExcelImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsImporting(true);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
        if (rows.length < 2) {
          alert('File Excel không có dữ liệu.');
          setIsImporting(false);
          return;
        }

        const headers = rows[0].map(h => String(h || '').trim().toLowerCase());
        const itemIdx = headers.findIndex(h => h === 'item' || h === 'mã linh kiện' || h === 'id');
        const descIdx = headers.findIndex(h => h === 'description' || h === 'tên linh kiện' || h === 'name');
        const unitIdx = headers.findIndex(h => h === 'unit' || h === 'đvt');
        const skipLaserIdx = headers.findIndex(h => h === 'skiplaser' || h === 'bỏ qua laser' || h === 'miễn laser' || h === 'skip laser');
        const skipBendIdx = headers.findIndex(h => h === 'skipbending' || h === 'bỏ qua chấn' || h === 'miễn chấn' || h === 'skip bend');
        const skipWeldIdx = headers.findIndex(h => h === 'skipwelding' || h === 'bỏ qua hàn' || h === 'miễn hàn' || h === 'skip weld');
        const skipPaintIdx = headers.findIndex(h => h === 'skippainting' || h === 'bỏ qua sơn' || h === 'miễn sơn' || h === 'skip paint');
        const hasPaintingPOIdx = headers.findIndex(h => h === 'haspaintingpo' || h === 'có po tại công đoạn sơn' || h === 'po sơn');

        if (itemIdx === -1 || descIdx === -1) {
          alert('Không tìm thấy cột Item hoặc Description trong file Excel.');
          setIsImporting(false);
          return;
        }

        const importedParts: Part[] = rows.slice(1).map(row => {
          // Determine level based on which of the first 3 columns has a value
          let level = 1;
          const col0 = String(row[0] || '').trim();
          const col1 = String(row[1] || '').trim();
          const col2 = String(row[2] || '').trim();

          if (col0 !== '') level = 1;
          else if (col1 !== '') level = 2;
          else if (col2 !== '') level = 3;

          const isTruthful = (val: any) => {
            const s = String(val || '').trim().toLowerCase();
            return s === 'x' || s === 'y' || s === 'yes' || s === 'v' || s === '1' || s === 'true' || s === 'đúng' || s === 'có';
          };

          return {
            id: String(row[itemIdx] || '').trim(),
            name: String(row[descIdx] || '').trim(),
            unit: String(unitIdx !== -1 ? row[unitIdx] : 'Cái').trim(),
            level: level,
            skipLaser: skipLaserIdx !== -1 ? isTruthful(row[skipLaserIdx]) : false,
            skipBending: skipBendIdx !== -1 ? isTruthful(row[skipBendIdx]) : false,
            skipWelding: skipWeldIdx !== -1 ? isTruthful(row[skipWeldIdx]) : false,
            skipPainting: skipPaintIdx !== -1 ? isTruthful(row[skipPaintIdx]) : false,
            hasPaintingPO: hasPaintingPOIdx !== -1 ? isTruthful(row[hasPaintingPOIdx]) : false
          };
        }).filter(p => p.id && p.name);

        if (importedParts.length === 0) {
          alert('Không tìm thấy dữ liệu hợp lệ trong file Excel.');
          setIsImporting(false);
          return;
        }

        // Update existing parts and add new ones from Excel
        const partsMap = new Map(parts.map(p => [p.id, p]));
        importedParts.forEach(p => {
          partsMap.set(p.id, p);
        });
        const newParts = Array.from(partsMap.values());

        storageService.saveParts(newParts);
        onPartsChange();
        alert(`Đã nhập thành công ${importedParts.length} linh kiện!`);
      } catch (err) {
        if ((import.meta as any).env?.DEV) console.error('Excel Import Error:', err);
        alert('Lỗi khi đọc file Excel. Vui lòng đảm bảo file đúng định dạng.');
      } finally {
        setIsImporting(false);
        if (e.target) e.target.value = '';
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleAddPart = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newPart.id || !newPart.name) return;
    
    const existing = parts.find(p => p.id === newPart.id);
    if (existing && !editingId) {
      alert('Mã linh kiện đã tồn tại!');
      return;
    }

    let updatedParts;
    if (editingId) {
      updatedParts = parts.map(p => p.id === editingId ? newPart : p);
    } else {
      updatedParts = [...parts, newPart];
    }

    storageService.saveParts(updatedParts);
    setNewPart({ id: '', name: '', unit: 'Cái', level: 1, hasPaintingPO: false });
    setEditingId(null);
    onPartsChange();
  };

  const handleEdit = (part: Part) => {
    setNewPart({
      ...part,
      level: part.level || 1,
      skipLaser: !!part.skipLaser,
      skipBending: !!part.skipBending,
      skipWelding: !!part.skipWelding,
      skipPainting: !!part.skipPainting,
      hasPaintingPO: !!part.hasPaintingPO
    });
    setEditingId(part.id);
  };

  const handleDelete = (id: string) => {
    if (confirm('Bạn có chắc chắn muốn xóa linh kiện này? Dữ liệu tồn kho liên quan sẽ không bị xóa nhưng có thể hiển thị không chính xác.')) {
      const updatedParts = parts.filter(p => p.id !== id);
      storageService.saveParts(updatedParts);
      onPartsChange();
    }
  };

  return (
    <motion.div 
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, y: -20 }}
      className="grid grid-cols-1 lg:grid-cols-3 gap-8"
    >
      <div className="lg:col-span-1 bg-white p-10 rounded-2xl border border-gray-200 shadow-sm space-y-8">
        <h2 className="text-2xl font-bold tracking-tight">{editingId ? 'Cập nhật linh kiện' : 'Thêm linh kiện mới'}</h2>
        <form onSubmit={handleAddPart} className="space-y-6">
          <div className="space-y-3">
            <label className="text-xs font-bold uppercase opacity-50">Mã linh kiện</label>
            <input 
              type="text" 
              value={newPart.id}
              onChange={e => setNewPart({...newPart, id: e.target.value})}
              disabled={!!editingId}
              className="w-full p-4 rounded-lg border border-gray-200 font-mono text-base focus:border-blue-600 outline-none disabled:bg-gray-50"
              placeholder="VD: LK001"
              required
            />
          </div>
          <div className="space-y-3">
            <label className="text-xs font-bold uppercase opacity-50">Tên linh kiện</label>
            <input 
              type="text" 
              value={newPart.name}
              onChange={e => setNewPart({...newPart, name: e.target.value})}
              className="w-full p-4 rounded-lg border border-gray-200 text-base focus:border-blue-600 outline-none"
              placeholder="VD: Khung thép A1"
              required
            />
          </div>
          <div className="grid grid-cols-2 gap-6">
            <div className="space-y-3">
              <label className="text-xs font-bold uppercase opacity-50">Đơn vị</label>
              <input 
                type="text" 
                value={newPart.unit}
                onChange={e => setNewPart({...newPart, unit: e.target.value})}
                className="w-full p-4 rounded-lg border border-gray-200 text-base focus:border-blue-600 outline-none"
                placeholder="Cái, Bộ..."
              />
            </div>
            <div className="space-y-3">
              <label className="text-xs font-bold uppercase opacity-50">Cấp BOM</label>
              <select 
                value={newPart.level || 1}
                onChange={e => setNewPart({...newPart, level: parseInt(e.target.value)})}
                className="w-full p-4 rounded-lg border border-gray-200 text-base focus:border-blue-600 outline-none bg-white"
              >
                <option value={1}>Cấp 1 (Hàn/Sơn)</option>
                <option value={2}>Cấp 2 (Laser/Chấn)</option>
                <option value={3}>Cấp 3 (Tôn tấm)</option>
              </select>
            </div>
          </div>
          <div className="space-y-4">
            <label className="text-xs font-bold uppercase opacity-50">Cấu hình Quy trình</label>
            <div className="grid grid-cols-4 gap-4">
              <label className="flex items-center gap-3 p-3 bg-gray-50 rounded-lg cursor-pointer hover:bg-gray-100 transition-colors">
                <input 
                  type="checkbox" 
                  checked={!!newPart.skipLaser} 
                  onChange={e => setNewPart({...newPart, skipLaser: e.target.checked})}
                  className="w-5 h-5 rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                />
                <span className="text-sm font-medium">Bỏ qua Laser</span>
              </label>
              <label className="flex items-center gap-3 p-3 bg-gray-50 rounded-lg cursor-pointer hover:bg-gray-100 transition-colors">
                <input 
                  type="checkbox" 
                  checked={!!newPart.skipBending} 
                  onChange={e => setNewPart({...newPart, skipBending: e.target.checked})}
                  className="w-5 h-5 rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                />
                <span className="text-sm font-medium">Bỏ qua Chấn</span>
              </label>
              <label className="flex items-center gap-3 p-3 bg-gray-50 rounded-lg cursor-pointer hover:bg-gray-100 transition-colors">
                <input 
                  type="checkbox" 
                  checked={!!newPart.skipWelding} 
                  onChange={e => setNewPart({...newPart, skipWelding: e.target.checked})}
                  className="w-5 h-5 rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                />
                <span className="text-sm font-medium">Bỏ qua Hàn</span>
              </label>
              <label className="flex items-center gap-3 p-3 bg-gray-50 rounded-lg cursor-pointer hover:bg-gray-100 transition-colors">
                <input 
                  type="checkbox" 
                  checked={!!newPart.skipPainting} 
                  onChange={e => setNewPart({...newPart, skipPainting: e.target.checked})}
                  className="w-5 h-5 rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                />
                <span className="text-sm font-medium">Bỏ qua Sơn</span>
              </label>
              <label className="flex items-center gap-3 p-3 bg-blue-50 rounded-lg cursor-pointer hover:bg-blue-100 transition-colors border border-blue-100 col-span-2">
                <input 
                  type="checkbox" 
                  checked={!!newPart.hasPaintingPO} 
                  onChange={e => setNewPart({...newPart, hasPaintingPO: e.target.checked})}
                  className="w-5 h-5 rounded border-blue-300 text-blue-600 focus:ring-blue-500"
                />
                <span className="text-sm font-bold text-blue-700">Có PO tại công đoạn sơn</span>
              </label>
            </div>
          </div>
          <div className="pt-6 flex gap-3">
            <button 
              type="submit"
              className="flex-1 bg-blue-600 text-white py-4 rounded-lg font-bold text-base uppercase hover:bg-blue-700 transition-all shadow-lg shadow-blue-100"
            >
              {editingId ? 'Cập nhật' : 'Thêm mới'}
            </button>
            {editingId && (
              <button 
                type="button"
                onClick={() => { setEditingId(null); setNewPart({ id: '', name: '', unit: 'Cái', level: 1 }); }}
                className="px-6 py-4 border border-gray-200 rounded-lg text-base font-bold uppercase hover:bg-gray-50"
              >
                Hủy
              </button>
            )}
          </div>
        </form>

        <div className="pt-6 border-t border-gray-100 mt-6 space-y-3">
          <div className="flex items-center justify-between">
            <h3 className="text-xs font-bold uppercase tracking-wider text-gray-700 flex items-center gap-1.5">
              <Database size={14} className="text-emerald-600" />
              Sao lưu Local Storage
            </h3>
            <span className="text-[11px] font-mono font-bold text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-full border border-emerald-200">
              {backupStats.sizeFormatted}
            </span>
          </div>
          <p className="text-xs text-gray-500">Tải tệp JSON về máy tính để sao lưu dữ liệu hoặc nạp dữ liệu từ máy khác vào.</p>
          <div className="grid grid-cols-2 gap-2">
            <button
              type="button"
              onClick={handleDownloadBackup}
              className="py-2.5 px-3 bg-emerald-600 text-white rounded-lg text-xs font-bold flex items-center justify-center gap-1.5 hover:bg-emerald-700 transition-all shadow-sm active:scale-95"
              title="Tải xuống toàn bộ dữ liệu hiện tại thành tệp JSON"
            >
              <Download size={14} />
              Tải file JSON
            </button>
            <button
              type="button"
              onClick={() => setActiveSettingsTab('backup')}
              className={cn(
                "py-2.5 px-3 border rounded-lg text-xs font-bold flex items-center justify-center gap-1.5 transition-all active:scale-95",
                activeSettingsTab === 'backup' 
                  ? "bg-blue-600 text-white border-blue-600 shadow-sm" 
                  : "bg-blue-50 text-blue-700 border-blue-200 hover:bg-blue-100"
              )}
              title="Mở tab nạp dữ liệu từ file JSON"
            >
              <FileJson size={14} />
              Nạp dữ liệu
            </button>
          </div>
          {downloadSuccessMessage && (
            <div className="p-2 bg-emerald-50 border border-emerald-200 rounded text-[11px] text-emerald-800 font-medium">
              ✓ {downloadSuccessMessage}
            </div>
          )}
        </div>

        <div className="pt-6 border-t border-gray-100 mt-6">
          <h3 className="text-xs font-bold uppercase opacity-50 mb-3 text-red-600">Khu vực nguy hiểm</h3>
          <button 
            onClick={() => setShowResetModal(true)}
            className="w-full py-2.5 border-2 border-red-100 text-red-600 rounded-lg text-xs font-bold uppercase hover:bg-red-50 transition-all flex items-center justify-center gap-2"
          >
            <Trash2 size={15} />
            Xóa toàn bộ dữ liệu hệ thống
          </button>
        </div>

        <AnimatePresence>
          {showResetModal && (
            <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
              <motion.div 
                initial={{ scale: 0.9, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                exit={{ scale: 0.9, opacity: 0 }}
                className="bg-white rounded-2xl p-8 max-w-md w-full shadow-2xl space-y-6"
              >
                <div className="flex items-center gap-4 text-red-600">
                  <AlertCircle size={32} />
                  <h3 className="text-xl font-bold">Xác nhận xóa dữ liệu</h3>
                </div>
                <p className="text-sm text-gray-500">Hành động này sẽ xóa vĩnh viễn toàn bộ danh mục linh kiện, tồn kho, định mức và nhật ký giao dịch. Không thể hoàn tác.</p>
                <div className="space-y-2">
                  <label className="text-[10px] font-bold uppercase opacity-50">Nhập mật khẩu xác nhận</label>
                  <input 
                    type="password"
                    value={resetPassword}
                    onChange={e => setResetPassword(e.target.value)}
                    className="w-full p-3 rounded-lg border border-gray-200 focus:border-red-500 outline-none"
                    placeholder="Nhập mật khẩu..."
                  />
                </div>
                <div className="flex gap-3 pt-2">
                  <button 
                    onClick={() => { setShowResetModal(false); setResetPassword(''); }}
                    className="flex-1 py-3 bg-gray-100 rounded-lg font-bold text-sm uppercase hover:bg-gray-200 transition-all"
                  >
                    Hủy bỏ
                  </button>
                  <button 
                    onClick={handleResetData}
                    className="flex-1 py-3 bg-red-600 text-white rounded-lg font-bold text-sm uppercase hover:bg-red-700 transition-all"
                  >
                    Xác nhận xóa
                  </button>
                </div>
              </motion.div>
            </div>
          )}
        </AnimatePresence>
      </div>

      <div className="lg:col-span-2 bg-white rounded-2xl border border-gray-200 shadow-sm overflow-hidden flex flex-col">
        <div className="flex border-b border-gray-100 flex-wrap bg-gray-50/50">
          <button 
            onClick={() => setActiveSettingsTab('parts')}
            className={cn(
              "px-5 py-4 text-sm font-bold uppercase tracking-wider transition-all border-b-2 whitespace-nowrap",
              activeSettingsTab === 'parts' ? "border-blue-600 text-blue-600 bg-white" : "border-transparent text-gray-400 hover:text-gray-600"
            )}
          >
            Danh mục linh kiện
          </button>
          <button 
            onClick={() => setActiveSettingsTab('bom')}
            className={cn(
              "px-5 py-4 text-sm font-bold uppercase tracking-wider transition-all border-b-2 whitespace-nowrap",
              activeSettingsTab === 'bom' ? "border-blue-600 text-blue-600 bg-white" : "border-transparent text-gray-400 hover:text-gray-600"
            )}
          >
            Định mức sản xuất (BOM)
          </button>
          <button 
            onClick={() => setActiveSettingsTab('bom_v2')}
            className={cn(
              "px-5 py-4 text-sm font-bold uppercase tracking-wider transition-all border-b-2 whitespace-nowrap",
              activeSettingsTab === 'bom_v2' ? "border-blue-600 text-blue-600 bg-white" : "border-transparent text-gray-400 hover:text-gray-600"
            )}
          >
            Định mức Hàn (BOM v2)
          </button>
          <button 
            onClick={() => setActiveSettingsTab('model_bom')}
            className={cn(
              "px-5 py-4 text-sm font-bold uppercase tracking-wider transition-all border-b-2 whitespace-nowrap",
              activeSettingsTab === 'model_bom' ? "border-blue-600 text-blue-600 bg-white" : "border-transparent text-gray-400 hover:text-gray-600"
            )}
          >
            Định mức Model
          </button>
          <button 
            onClick={() => setActiveSettingsTab('label')}
            className={cn(
              "px-5 py-4 text-sm font-bold uppercase tracking-wider transition-all border-b-2 whitespace-nowrap",
              activeSettingsTab === 'label' ? "border-blue-600 text-blue-600 bg-white" : "border-transparent text-gray-400 hover:text-gray-600"
            )}
          >
            Cấu hình khổ nhãn
          </button>
          <button 
            onClick={() => setActiveSettingsTab('transformations')}
            className={cn(
              "px-5 py-4 text-sm font-bold uppercase tracking-wider transition-all border-b-2 whitespace-nowrap",
              activeSettingsTab === 'transformations' ? "border-blue-600 text-blue-600 bg-white" : "border-transparent text-gray-400 hover:text-gray-600"
            )}
          >
            Chuyển đổi linh kiện
          </button>
          <button 
            onClick={() => {
              setActiveSettingsTab('backup');
              refreshBackupStats();
            }}
            className={cn(
              "px-5 py-4 text-sm font-bold uppercase tracking-wider transition-all border-b-2 whitespace-nowrap flex items-center gap-2",
              activeSettingsTab === 'backup' 
                ? "border-emerald-600 text-emerald-700 bg-white shadow-sm font-black" 
                : "border-transparent text-emerald-600 hover:text-emerald-700 hover:bg-emerald-50/50 font-bold"
            )}
          >
            <Database size={16} className="text-emerald-600" />
            Sao lưu & Nạp dữ liệu (JSON)
          </button>
        </div>

        {activeSettingsTab === 'bom_v2' ? (
          <div className="p-10 space-y-8">
            <div className="flex justify-between items-center">
              <div className="flex items-center gap-4">
                <div className="bg-orange-600 p-3 rounded-xl text-white">
                  <Layers size={24} />
                </div>
                <div>
                  <h2 className="text-2xl font-bold tracking-tight">Định mức Hàn (BOM v2)</h2>
                  <p className="text-sm text-gray-500">Tự động trừ linh kiện Level 2 khi nhập kho OUT công đoạn Hàn</p>
                </div>
              </div>
              <div className="flex gap-4">
                <label className="bg-white border-2 border-orange-600 text-orange-600 px-6 py-3 rounded-xl font-bold uppercase cursor-pointer hover:bg-orange-50 transition-all flex items-center gap-2">
                  <FileUp size={20} />
                  {isImportingBOMV2 ? 'Đang xử lý...' : 'Nhập Excel BOM v2'}
                  <input type="file" accept=".xlsx, .xls" className="hidden" onChange={(e) => {
                    const file = e.target.files?.[0];
                    if (!file) return;
                    setIsImportingBOMV2(true);
                    const reader = new FileReader();
                    reader.onload = (evt) => {
                      try {
                        const bstr = evt.target?.result;
                        const wb = XLSX.read(bstr, { type: 'binary' });
                        const wsname = wb.SheetNames[0];
                        const ws = wb.Sheets[wsname];
                        const data = XLSX.utils.sheet_to_json(ws) as any[];
                        
                        // Expected columns: ResultID, IngredientID, Quantity, SkipBending, ApplicableModel
                        const imported: BOMDefinitionV2[] = data.map(row => ({
                          resultPartId: String(row['ResultID'] || row['Mã thành phẩm'] || row['Mã cha'] || '').trim(),
                          ingredientPartId: String(row['IngredientID'] || row['Mã linh kiện'] || row['Mã con'] || '').trim(),
                          quantity: parseFloat(row['Quantity'] || row['Số lượng'] || '1'),
                          applicableModel: row['Model áp dụng'] || row['Model'] || row['ApplicableModel'] ? String(row['Model áp dụng'] || row['Model'] || row['ApplicableModel']).trim() : undefined,
                        })).filter(b => b.resultPartId && b.ingredientPartId && b.quantity > 0);

                        if (imported.length === 0) {
                          alert('Không tìm thấy dữ liệu hợp lệ. Cần các cột: ResultID, IngredientID, Quantity');
                        } else {
                          storageService.saveBOMV2(imported);
                          onPartsChange();
                          alert(`Đã nhập thành công ${imported.length} định mức Hàn!`);
                        }
                      } catch (err) {
                        alert('Lỗi khi đọc file.');
                      } finally {
                        setIsImportingBOMV2(false);
                        if (e.target) e.target.value = '';
                      }
                    };
                    reader.readAsBinaryString(file);
                  }} />
                </label>
              </div>
            </div>

            <div className="bg-white border border-gray-200 rounded-2xl overflow-hidden">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-gray-50 border-b border-gray-200">
                    <th className="p-6 text-xs font-bold uppercase opacity-50">Model áp dụng</th>
                    <th className="p-6 text-xs font-bold uppercase opacity-50">Thành phẩm (Level 1)</th>
                    <th className="p-6 text-xs font-bold uppercase opacity-50">Linh kiện thành phần (Level 2)</th>
                    <th className="p-6 text-xs font-bold uppercase opacity-50 text-center">Số lượng/1 đơn vị</th>
                  </tr>
                </thead>
                <tbody>
                  {storageService.getBOMV2().length === 0 ? (
                    <tr>
                      <td colSpan={4} className="p-20 text-center text-gray-400 italic">Chưa có dữ liệu định mức Hàn. Vui lòng nhập từ Excel.</td>
                    </tr>
                  ) : (
                    storageService.getBOMV2().map((b, i) => (
                      <tr key={i} className="border-b border-gray-100 hover:bg-gray-50 transition-colors">
                        <td className="p-6 font-mono text-sm opacity-80 whitespace-nowrap">
                          {b.applicableModel ? (
                            <span className="bg-purple-100 text-purple-700 px-3 py-1 rounded-full font-bold">{b.applicableModel}</span>
                          ) : (
                            <span className="opacity-40 italic">Tất cả Model</span>
                          )}
                        </td>
                        <td className="p-6">
                          <div className="font-bold">{getProcessValue(parts.find(p => p.id === b.resultPartId)?.name, parts.find(p => p.id === b.resultPartId), 'WELDING', 'OUT')}</div>
                          <div className="text-xs font-mono opacity-50">{getProcessValue(b.resultPartId, parts.find(p => p.id === b.resultPartId), 'WELDING', 'OUT')}</div>
                        </td>
                        <td className="p-6">
                          <div className="font-bold">{getProcessValue(parts.find(p => p.id === b.ingredientPartId)?.name, parts.find(p => p.id === b.ingredientPartId), 'BENDING', 'OUT')}</div>
                          <div className="text-xs font-mono opacity-50">{getProcessValue(b.ingredientPartId, parts.find(p => p.id === b.ingredientPartId), 'BENDING', 'OUT')}</div>
                        </td>
                        <td className="p-6 text-center font-mono font-bold text-orange-600 text-xl">
                          x{b.quantity}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        ) : activeSettingsTab === 'model_bom' ? (
          <div className="p-10 space-y-8">
            <div className="flex justify-between items-center">
              <div className="flex items-center gap-4">
                <div className="bg-purple-600 p-3 rounded-xl text-white">
                  <ClipboardList size={24} />
                </div>
                <div>
                  <h2 className="text-2xl font-bold tracking-tight">Định mức Model (Model BOM)</h2>
                  <p className="text-sm text-gray-500">Liên kết Model với các linh kiện Level 1</p>
                </div>
              </div>
              <div className="flex gap-4">
                <label className="bg-white border-2 border-purple-600 text-purple-600 px-6 py-3 rounded-xl font-bold uppercase cursor-pointer hover:bg-purple-50 transition-all flex items-center gap-2">
                  <FileUp size={20} />
                  {isImportingModelBOM ? 'Đang xử lý...' : 'Nhập Excel Model BOM'}
                  <input type="file" accept=".xlsx, .xls" className="hidden" onChange={(e) => {
                    const file = e.target.files?.[0];
                    if (!file) return;
                    setIsImportingModelBOM(true);
                    const reader = new FileReader();
                    reader.onload = (evt) => {
                      try {
                        const bstr = evt.target?.result;
                        const wb = XLSX.read(bstr, { type: 'binary' });
                        const wsname = wb.SheetNames[0];
                        const ws = wb.Sheets[wsname];
                        const data = XLSX.utils.sheet_to_json(ws) as any[];
                        
                        // Expected columns: ModelID, PartID, Quantity
                        const imported: ModelBOMDefinition[] = data.map(row => ({
                          modelId: String(row['ModelID'] || row['Mã model'] || '').trim(),
                          partId: String(row['PartID'] || row['Mã linh kiện'] || row['Mã con'] || '').trim(),
                          quantity: parseFloat(row['Quantity'] || row['Số lượng'] || '1')
                        })).filter(b => b.modelId && b.partId && b.quantity > 0);

                        if (imported.length === 0) {
                          alert('Không tìm thấy dữ liệu hợp lệ. Cần các cột: ModelID, PartID, Quantity');
                        } else {
                          storageService.saveModelBOM(imported);
                          onPartsChange();
                          alert(`Đã nhập thành công ${imported.length} định mức Model!`);
                        }
                      } catch (err) {
                        alert('Lỗi khi đọc file.');
                      } finally {
                        setIsImportingModelBOM(false);
                        if (e.target) e.target.value = '';
                      }
                    };
                    reader.readAsBinaryString(file);
                  }} />
                </label>
              </div>
            </div>

            <div className="bg-white border border-gray-200 rounded-2xl overflow-hidden">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-gray-50 border-b border-gray-200">
                    <th className="p-6 text-xs font-bold uppercase opacity-50">Tên Model</th>
                    <th className="p-6 text-xs font-bold uppercase opacity-50">Linh kiện Level 1</th>
                    <th className="p-6 text-xs font-bold uppercase opacity-50 text-center">Số lượng</th>
                  </tr>
                </thead>
                <tbody>
                  {storageService.getModelBOM().length === 0 ? (
                    <tr>
                      <td colSpan={3} className="p-20 text-center text-gray-400 italic">Chưa có dữ liệu định mức Model. Vui lòng nhập từ Excel.</td>
                    </tr>
                  ) : (
                    storageService.getModelBOM().map((b, i) => (
                      <tr key={i} className="border-b border-gray-100 hover:bg-gray-50 transition-colors">
                        <td className="p-6">
                          <div className="font-bold">{parts.find(p => p.id === b.modelId)?.name || b.modelId}</div>
                          <div className="text-xs font-mono opacity-50">{b.modelId}</div>
                        </td>
                        <td className="p-6">
                          <div className="font-bold">{getProcessValue(parts.find(p => p.id === b.partId)?.name, parts.find(p => p.id === b.partId), 'PAINTING', 'IN')}</div>
                          <div className="text-xs font-mono opacity-50">{getProcessValue(b.partId, parts.find(p => p.id === b.partId), 'PAINTING', 'IN')}</div>
                        </td>
                        <td className="p-6 text-center font-mono font-bold text-purple-600 text-xl">
                          x{b.quantity}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        ) : activeSettingsTab === 'label' ? (
          <div className="p-10 space-y-8">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
              <div className="flex items-center gap-4">
                <div className="bg-blue-600 p-3 rounded-xl text-white">
                  <Printer size={24} />
                </div>
                <div>
                  <h2 className="text-2xl font-bold tracking-tight">Cấu hình in nhãn trực tiếp (Khổ A7 75x100mm)</h2>
                  <p className="text-sm text-gray-500">Kích thước chuẩn cho máy in tem nhiệt (Chiều ngang 75mm x Chiều dọc 100mm)</p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => onLabelSettingsChange({ width: 75, height: 100, fontSize: 10, qrSize: 110 })}
                className="px-4 py-2.5 bg-[#F27D26] hover:bg-[#d96a1a] text-white rounded-xl text-xs font-bold flex items-center gap-2 transition-all shadow-md active:scale-95 shrink-0"
              >
                <RotateCcw size={16} />
                Khôi phục chuẩn A7 (75x100mm)
              </button>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 items-start">
              <div className="space-y-6 bg-white p-6 rounded-2xl border border-gray-200 shadow-sm">
                <h3 className="text-sm font-bold uppercase text-gray-700 tracking-wider">Thông số cài đặt</h3>
                <div className="grid grid-cols-2 gap-6">
                  <div className="space-y-2">
                    <label className="text-xs font-bold uppercase text-gray-600">Chiều rộng (mm)</label>
                    <input 
                      type="number"
                      value={labelSettings.width}
                      onChange={e => onLabelSettingsChange({...labelSettings, width: parseInt(e.target.value) || 75})}
                      className="w-full p-3.5 rounded-xl border border-gray-200 font-mono text-base outline-none focus:border-blue-600 focus:ring-2 focus:ring-blue-100"
                      placeholder="75"
                    />
                    <p className="text-[10px] text-gray-500">Chuẩn tem A7: 75mm</p>
                  </div>
                  <div className="space-y-2">
                    <label className="text-xs font-bold uppercase text-gray-600">Chiều cao (mm)</label>
                    <input 
                      type="number"
                      value={labelSettings.height}
                      onChange={e => onLabelSettingsChange({...labelSettings, height: parseInt(e.target.value) || 100})}
                      className="w-full p-3.5 rounded-xl border border-gray-200 font-mono text-base outline-none focus:border-blue-600 focus:ring-2 focus:ring-blue-100"
                      placeholder="100"
                    />
                    <p className="text-[10px] text-gray-500">Chuẩn tem A7: 100mm</p>
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-6">
                  <div className="space-y-2">
                    <label className="text-xs font-bold uppercase text-gray-600">Cỡ chữ cơ sở (px)</label>
                    <input 
                      type="number"
                      value={labelSettings.fontSize}
                      onChange={e => onLabelSettingsChange({...labelSettings, fontSize: parseInt(e.target.value) || 10})}
                      className="w-full p-3.5 rounded-xl border border-gray-200 font-mono text-base outline-none focus:border-blue-600 focus:ring-2 focus:ring-blue-100"
                      placeholder="10"
                    />
                    <p className="text-[10px] text-gray-500">Khuyên dùng: 8px - 12px</p>
                  </div>
                  <div className="space-y-2">
                    <label className="text-xs font-bold uppercase text-gray-600">Kích thước QR (px)</label>
                    <input 
                      type="number"
                      value={labelSettings.qrSize}
                      onChange={e => onLabelSettingsChange({...labelSettings, qrSize: parseInt(e.target.value) || 110})}
                      className="w-full p-3.5 rounded-xl border border-gray-200 font-mono text-base outline-none focus:border-blue-600 focus:ring-2 focus:ring-blue-100"
                      placeholder="110"
                    />
                    <p className="text-[10px] text-gray-500">Khuyên dùng: 90px - 120px</p>
                  </div>
                </div>
              </div>

              <div className="bg-gray-50 rounded-2xl p-6 border-2 border-dashed border-gray-200 flex flex-col items-center justify-center">
                <span className="text-xs font-bold uppercase text-gray-500 mb-4 tracking-wider">Xem trước nhãn thực tế ({labelSettings.width} x {labelSettings.height} mm)</span>
                
                <div className="shadow-2xl rounded-xl overflow-hidden bg-white border border-gray-300 transform scale-95 origin-top">
                  <A7QRLabelCard 
                    label={{
                      id: 'SAMPLE-A7',
                      partId: '04-29-08-SHA76622KL-0008',
                      partName: 'THANH GIẰNG TRÊN SHA76622KL',
                      stageId: 'LASER',
                      targetStageId: 'BENDING',
                      quantity: 22,
                      type: 'STAGE_OUT',
                      poId: 'PO-SHA76623-0906-LASER-JI6R8',
                      qrData: 'SAMPLE|04-29-08-SHA76622KL-0008|22|LASER|OUT|BENDING|PO-SHA76623-0906-LASER-JI6R8|PO-MASTER-001|400|100',
                      timestamp: Date.now()
                    }} 
                    parts={parts} 
                    labelSettings={labelSettings} 
                    isPrintArea={false} 
                  />
                </div>
                <p className="mt-4 text-xs text-gray-400 italic text-center">Hình ảnh hiển thị thực tế theo các thông số cấu hình ở trên</p>
              </div>
            </div>

            <div className="p-6 bg-blue-50 rounded-xl border border-blue-100 flex gap-4 items-start">
              <AlertCircle size={20} className="text-blue-600 mt-1" />
              <div className="text-sm text-blue-800 space-y-2">
                <p className="font-bold">Lưu ý khi in:</p>
                <ul className="list-disc pl-4 space-y-1">
                  <li>Trong hộp thoại in của trình duyệt, hãy chọn máy in nhãn của bạn.</li>
                  <li>Phần <strong>Khổ giấy (Paper Size)</strong> nên được đặt trùng với cấu hình ở trên.</li>
                  <li>Phần <strong>Lề (Margins)</strong> nên đặt là <strong>None</strong> để nhãn in ra chuẩn nhất.</li>
                  <li>Phần <strong>Tỷ lệ (Scale)</strong> nên đặt là <strong>Default</strong> hoặc <strong>100%</strong>.</li>
                </ul>
              </div>
            </div>
          </div>
        ) : activeSettingsTab === 'transformations' ? (
          <div className="p-10 space-y-8">
            <div className="flex justify-between items-center">
              <div className="flex items-center gap-4">
                <div className="bg-cyan-600 p-3 rounded-xl text-white">
                  <RotateCcw size={24} />
                </div>
                <div>
                  <h2 className="text-2xl font-bold tracking-tight">Cài đặt chuyển đổi linh kiện</h2>
                  <p className="text-sm text-gray-500">Tự động chuyển đổi tên linh kiện khi nhập kho công đoạn Sơn</p>
                </div>
              </div>
              <div className="flex gap-4">
                <button 
                  onClick={() => {
                    if (!confirm('Bạn có chắc chắn muốn đảo ngược tất cả dữ liệu Nguồn và Đích trong bảng này không?')) return;
                    const current = storageService.getTransformations();
                    const swapped = current.map(t => ({
                      ...t,
                      sourcePartId: t.targetPartId,
                      targetPartId: t.sourcePartId
                    }));
                    storageService.saveTransformations(swapped);
                    onPartsChange();
                    alert('Đã đảo ngược dữ liệu thành công!');
                  }}
                  className="bg-white border-2 border-orange-500 text-orange-500 px-6 py-3 rounded-xl font-bold uppercase hover:bg-orange-50 transition-all flex items-center gap-2"
                >
                  <RotateCcw size={20} />
                  Đảo ngược Nguồn/Đích
                </button>
                <label className="bg-white border-2 border-cyan-600 text-cyan-600 px-6 py-3 rounded-xl font-bold uppercase cursor-pointer hover:bg-cyan-100 transition-all flex items-center gap-2">
                  <FileUp size={20} />
                  {isImportingTransformations ? 'Đang xử lý...' : 'Nhập Excel Chuyển đổi'}
                  <input type="file" accept=".xlsx, .xls" className="hidden" onChange={(e) => {
                    const file = e.target.files?.[0];
                    if (!file) return;
                    setIsImportingTransformations(true);
                    const reader = new FileReader();
                    reader.onload = (evt) => {
                      try {
                        const bstr = evt.target?.result;
                        const wb = XLSX.read(bstr, { type: 'binary' });
                        const wsname = wb.SheetNames[0];
                        const ws = wb.Sheets[wsname];
                        const data = XLSX.utils.sheet_to_json(ws) as any[];
                        
                        // Support headers from user's image
                        const imported: PartTransformation[] = data.map(row => {
                          const source = String(row['Linh kiện nguồn (L2)'] || row['SourceID'] || row['Mã nguồn'] || row['Mã L2'] || '').trim();
                          const target = String(row['Linh kiện đích (L1)'] || row['TargetID'] || row['Mã đích'] || row['Mã L1'] || '').trim();
                          let stageStr = String(row['Công đoạn áp dụng'] || row['StageID'] || row['Công đoạn'] || 'PAINTING').trim().toUpperCase();
                          const applicableModel = row['Model áp dụng'] ? String(row['Model áp dụng']).trim() : undefined;
                          
                          // Map Vietnamese stage names to IDs
                          let stageId: StageId = 'PAINTING';
                          if (stageStr.includes('SƠN')) stageId = 'PAINTING';
                          else if (stageStr.includes('CHẤN') || stageStr.includes('DẬP')) stageId = 'BENDING';
                          else if (stageStr.includes('HÀN')) stageId = 'WELDING';
                          else if (stageStr.includes('LASER') || stageStr.includes('CẮT')) stageId = 'LASER';
                          else if (stageStr.includes('LẮP') || stageStr.includes('GÓI') || stageStr.includes('DCLR')) stageId = 'DCLR';
                          else {
                            const found = STAGES.find(s => s.id === stageStr || s.name.toUpperCase() === stageStr);
                            if (found) stageId = found.id;
                          }

                          let finalSource = source;
                          let finalTarget = target;
                          const normSource = source.toUpperCase().normalize('NFC').trim();
                          const normTarget = target.toUpperCase().normalize('NFC').trim();

                          const ptSrc = parts.find(p => p.id.toUpperCase().normalize('NFC').trim() === normSource || p.name.toUpperCase().normalize('NFC').trim() === normSource);
                          if (ptSrc) finalSource = ptSrc.id;

                          const ptTgt = parts.find(p => p.id.toUpperCase().normalize('NFC').trim() === normTarget || p.name.toUpperCase().normalize('NFC').trim() === normTarget || 
                                           (p.id.length > 5 && normTarget.replace(/\s+/g, '').includes(p.id.toUpperCase().normalize('NFC').replace(/\s+/g, ''))));
                          if (ptTgt) finalTarget = ptTgt.id;

                          return {
                            sourcePartId: finalSource, // Keep case but will handle insensitively in storage
                            targetPartId: finalTarget,
                            targetStageId: stageId,
                            applicableModel: applicableModel
                          };
                        }).filter(b => b.sourcePartId && b.targetPartId);

                        if (imported.length === 0) {
                          alert('Không tìm thấy dữ liệu hợp lệ. Cần các cột: Linh kiện nguồn (L2), Linh kiện đích (L1)');
                        } else {
                          storageService.saveTransformations(imported);
                          alert(`Đã nhập thành công ${imported.length} quy tắc chuyển đổi!`);
                          onPartsChange(); // Trigger refresh
                        }
                      } catch (err) {
                        alert('Lỗi khi đọc file.');
                      } finally {
                        setIsImportingTransformations(false);
                        if (e.target) e.target.value = '';
                      }
                    };
                    reader.readAsBinaryString(file);
                  }} />
                </label>
              </div>
            </div>

            <div className="bg-white border border-gray-200 rounded-2xl overflow-hidden">
               <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-gray-50 border-b border-gray-200">
                    <th className="p-6 text-xs font-bold uppercase opacity-50">Linh kiện Nguồn (Gốc)</th>
                    <th className="p-6 text-xs font-bold uppercase opacity-50">Model áp dụng</th>
                    <th className="p-6 text-xs font-bold uppercase opacity-50 text-center">
                      <ArrowRight size={16} className="mx-auto" />
                    </th>
                    <th className="p-6 text-xs font-bold uppercase opacity-50">Tên mới (Đích)</th>
                    <th className="p-6 text-xs font-bold uppercase opacity-50">Công đoạn Đích</th>
                    <th className="p-6 text-xs font-bold uppercase opacity-50 text-right">Thao tác</th>
                  </tr>
                </thead>
                <tbody>
                  {storageService.getTransformations().length === 0 ? (
                    <tr>
                      <td colSpan={6} className="p-20 text-center text-gray-400 italic">Chưa có quy tắc chuyển đổi nào.</td>
                    </tr>
                  ) : (
                    storageService.getTransformations().map((t, i) => {
                      const matchSource = parts.find(p => p.id.toUpperCase() === t.sourcePartId.toUpperCase() || p.name.toUpperCase() === t.sourcePartId.toUpperCase());
                      const matchTarget = parts.find(p => p.id.toUpperCase() === t.targetPartId.toUpperCase() || p.name.toUpperCase() === t.targetPartId.toUpperCase());
                      return (
                      <tr key={i} className="border-b border-gray-100 hover:bg-gray-50 transition-colors">
                        <td className="p-6">
                           <div className="font-bold">{matchSource?.name || t.sourcePartId}</div>
                           <div className="text-xs font-mono opacity-50">{matchSource?.id || t.sourcePartId}</div>
                        </td>
                        <td className="p-6 font-mono text-sm text-center">
                          {t.applicableModel ? (
                            <span className="bg-purple-100 text-purple-700 px-2 py-1 rounded line-clamp-1">{t.applicableModel}</span>
                          ) : (
                            <span className="text-gray-400 italic">Tất cả</span>
                          )}
                        </td>
                        <td className="p-6 text-center text-blue-600">
                           <ArrowRight size={20} className="mx-auto" />
                        </td>
                        <td className="p-6">
                           <div className="font-bold">{matchTarget?.name || t.targetPartId}</div>
                           <div className="text-xs font-mono opacity-50">{matchTarget?.id || t.targetPartId}</div>
                        </td>
                        <td className="p-6">
                           <span className="px-3 py-1 bg-blue-100 text-blue-700 rounded-full text-[10px] font-bold uppercase">
                             {STAGES.find(s => s.id === t.targetStageId)?.name || t.targetStageId}
                           </span>
                        </td>
                        <td className="p-6 text-right">
                           <button 
                             onClick={() => {
                               const current = storageService.getTransformations();
                               current.splice(i, 1);
                               storageService.saveTransformations(current);
                               onPartsChange();
                             }}
                             className="p-3 text-red-600 hover:bg-red-50 rounded-lg transition-colors"
                           >
                             <Trash2 size={20} />
                           </button>
                        </td>
                      </tr>
                    )})
                  )}
                </tbody>
              </table>
            </div>
            
            <div className="p-6 bg-cyan-50 rounded-xl border border-cyan-100 flex gap-4 items-start">
              <AlertCircle size={20} className="text-cyan-600 mt-1" />
              <div className="text-sm text-cyan-800 space-y-2">
                <p className="font-bold">Hướng dẫn cài đặt chuyển đổi:</p>
                <p>Quy tắc này sẽ tự động thay đổi mã linh kiện được ghi nhận vào kho IN của công đoạn chỉ định. 
                   Ví dụ: Bạn muốn linh kiện L2 (đã chấn xong) khi vào kho Sơn sẽ được tính là linh kiện L1.</p>
                <p className="text-xs opacity-70 italic">* Khi quét nhãn QR mã nguồn, hệ thống sẽ tự nhận diện và cộng tồn kho cho mã đích.</p>
              </div>
            </div>
          </div>
        ) : activeSettingsTab === 'backup' ? (
          <div className="p-8 lg:p-10 space-y-8">
            {/* Header Banner */}
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 pb-6 border-b border-gray-100">
              <div className="flex items-center gap-4">
                <div className="bg-emerald-600 p-3.5 rounded-2xl text-white shadow-lg shadow-emerald-600/20">
                  <Database size={28} />
                </div>
                <div>
                  <h2 className="text-2xl font-bold tracking-tight text-gray-900">
                    Lưu trữ & Nạp dữ liệu Local Storage (JSON)
                  </h2>
                  <p className="text-sm text-gray-500 mt-0.5">
                    Tải toàn bộ cơ sở dữ liệu hiện tại về máy tính dưới dạng file JSON để lưu trữ hoặc nạp sang máy tính khác mà không bị mất dữ liệu.
                  </p>
                </div>
              </div>
              <div className="flex items-center gap-3">
                <button
                  onClick={handleManualCleanup}
                  className="px-4 py-2.5 bg-red-50 hover:bg-red-100 text-red-700 border border-red-200 rounded-xl text-xs font-bold uppercase tracking-wider flex items-center gap-2 transition-all shadow-sm"
                  title="Xóa nhãn QR >24h, nhật ký >35 ngày và lệnh PO hoàn thành >35 ngày để giải phóng bộ nhớ"
                >
                  <Trash2 size={14} />
                  Dọn dẹp giải phóng bộ nhớ
                </button>
                <button
                  onClick={refreshBackupStats}
                  className="px-4 py-2.5 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-xl text-xs font-bold uppercase tracking-wider flex items-center gap-2 transition-all"
                  title="Cập nhật lại số liệu thống kê"
                >
                  <RefreshCw size={14} />
                  Làm mới
                </button>
              </div>
            </div>

            {cleanupMessage && (
              <div className="p-3.5 bg-emerald-50 border border-emerald-300 rounded-xl text-xs text-emerald-900 font-medium flex items-center gap-2">
                <CheckCircle2 size={16} className="text-emerald-700 shrink-0" />
                <span>{cleanupMessage}</span>
              </div>
            )}

            {/* Current Storage Stats Cards */}
            <div>
              <div className="flex items-center justify-between mb-3">
                <span className="text-xs font-bold uppercase tracking-wider text-gray-500">
                  Dữ liệu hiện có trên máy tính này
                </span>
                <span className="text-xs font-mono font-bold text-emerald-700 bg-emerald-50 px-2.5 py-1 rounded-full border border-emerald-200">
                  Dung lượng: {backupStats.sizeFormatted} ({backupStats.totalKeys} mục dữ liệu)
                </span>
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
                <div className="bg-gray-50/80 border border-gray-200 rounded-xl p-3.5">
                  <span className="text-xs text-gray-500 font-medium block">Linh kiện</span>
                  <span className="text-xl font-bold font-mono text-gray-900">{backupStats.partsCount}</span>
                  <span className="text-[10px] text-gray-400 block mt-0.5">mã sản phẩm</span>
                </div>
                <div className="bg-gray-50/80 border border-gray-200 rounded-xl p-3.5">
                  <span className="text-xs text-gray-500 font-medium block">Lệnh SX (PO)</span>
                  <span className="text-xl font-bold font-mono text-blue-600">{backupStats.posCount}</span>
                  <span className="text-[10px] text-gray-400 block mt-0.5">đơn hàng</span>
                </div>
                <div className="bg-gray-50/80 border border-gray-200 rounded-xl p-3.5">
                  <span className="text-xs text-gray-500 font-medium block">Tồn kho WIP</span>
                  <span className="text-xl font-bold font-mono text-emerald-600">{backupStats.inventoryCount}</span>
                  <span className="text-[10px] text-gray-400 block mt-0.5">dòng tồn</span>
                </div>
                <div className="bg-gray-50/80 border border-gray-200 rounded-xl p-3.5">
                  <span className="text-xs text-gray-500 font-medium block">Nhật ký quét</span>
                  <span className="text-xl font-bold font-mono text-purple-600">{backupStats.transactionsCount}</span>
                  <span className="text-[10px] text-gray-400 block mt-0.5">lượt giao dịch</span>
                </div>
                <div className="bg-gray-50/80 border border-gray-200 rounded-xl p-3.5">
                  <span className="text-xs text-gray-500 font-medium block">Nhãn QR</span>
                  <span className="text-xl font-bold font-mono text-orange-600">{backupStats.labelsCount}</span>
                  <span className="text-[10px] text-gray-400 block mt-0.5">nhãn đã in</span>
                </div>
                <div className="bg-gray-50/80 border border-gray-200 rounded-xl p-3.5">
                  <span className="text-xs text-gray-500 font-medium block">Định mức/BOM</span>
                  <span className="text-xl font-bold font-mono text-indigo-600">
                    {backupStats.bomCount + backupStats.bomV2Count + backupStats.modelBomCount + backupStats.normsCount}
                  </span>
                  <span className="text-[10px] text-gray-400 block mt-0.5">quy tắc cấu hình</span>
                </div>
              </div>
            </div>

            {/* Main 2-Column Actions */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
              {/* Left Column: Download / Export */}
              <div className="bg-gradient-to-br from-emerald-50/50 via-white to-emerald-50/30 border-2 border-emerald-200/80 rounded-2xl p-6 lg:p-7 flex flex-col justify-between shadow-sm">
                <div>
                  <div className="flex items-center justify-between mb-4">
                    <span className="px-3 py-1 bg-emerald-100 text-emerald-800 text-xs font-bold uppercase tracking-wider rounded-full flex items-center gap-1.5">
                      <Download size={13} />
                      Xuất dữ liệu ra máy tính
                    </span>
                    <span className="text-xs font-mono text-emerald-600 font-semibold">Định dạng .JSON</span>
                  </div>

                  <h3 className="text-xl font-bold text-gray-900 mb-2 flex items-center gap-2">
                    1. Tải xuống tệp sao lưu JSON
                  </h3>
                  <p className="text-sm text-gray-600 leading-relaxed mb-5">
                    Đóng gói toàn bộ cơ sở dữ liệu hiện có thành một tệp tin <strong className="text-emerald-700">.json</strong> duy nhất. Bạn có thể lưu vào máy tính, gửi qua Zalo, Email hoặc chuyển vào USB để nạp sang máy khác.
                  </p>

                  <div className="bg-white border border-emerald-100 rounded-xl p-4 mb-6 space-y-2 text-xs text-gray-600">
                    <p className="font-bold text-gray-700 uppercase tracking-wider text-[11px]">Dữ liệu được đóng gói trong tệp:</p>
                    <div className="grid grid-cols-2 gap-x-3 gap-y-1.5 pt-1">
                      <div className="flex items-center gap-1.5">
                        <Check size={13} className="text-emerald-600 shrink-0" />
                        <span>Toàn bộ danh mục linh kiện</span>
                      </div>
                      <div className="flex items-center gap-1.5">
                        <Check size={13} className="text-emerald-600 shrink-0" />
                        <span>Tồn kho WIP tất cả công đoạn</span>
                      </div>
                      <div className="flex items-center gap-1.5">
                        <Check size={13} className="text-emerald-600 shrink-0" />
                        <span>Danh sách Lệnh sản xuất (PO)</span>
                      </div>
                      <div className="flex items-center gap-1.5">
                        <Check size={13} className="text-emerald-600 shrink-0" />
                        <span>Lịch sử quét xuất/nhập kho</span>
                      </div>
                      <div className="flex items-center gap-1.5">
                        <Check size={13} className="text-emerald-600 shrink-0" />
                        <span>Định mức sản xuất BOM / v2 / Model</span>
                      </div>
                      <div className="flex items-center gap-1.5">
                        <Check size={13} className="text-emerald-600 shrink-0" />
                        <span>Cấu hình in nhãn, ca làm việc</span>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="space-y-3 pt-2">
                  <button
                    type="button"
                    onClick={handleDownloadBackup}
                    className="w-full py-4 px-6 bg-emerald-600 hover:bg-emerald-700 active:scale-[0.99] text-white rounded-xl font-bold uppercase tracking-wider flex items-center justify-center gap-3 shadow-lg shadow-emerald-600/20 transition-all text-sm"
                  >
                    <Download size={20} />
                    TẢI XUỐNG TỆP SAO LƯU (.JSON)
                  </button>

                  <button
                    type="button"
                    onClick={handleCopyBackupJson}
                    className="w-full py-2.5 px-4 bg-white hover:bg-emerald-50 text-emerald-800 border border-emerald-200 rounded-xl font-semibold text-xs flex items-center justify-center gap-2 transition-all"
                  >
                    <Copy size={14} />
                    {copiedJson ? '✓ Đã sao chép nội dung JSON vào bộ nhớ đệm!' : 'Sao chép toàn bộ mã JSON (Clipboard)'}
                  </button>

                  {downloadSuccessMessage && (
                    <div className="p-3 bg-emerald-100/80 border border-emerald-300 rounded-xl text-xs text-emerald-900 font-medium flex items-center gap-2 animate-fadeIn">
                      <CheckCircle2 size={16} className="text-emerald-700 shrink-0" />
                      <span>{downloadSuccessMessage}</span>
                    </div>
                  )}
                </div>
              </div>

              {/* Right Column: Upload / Import */}
              <div className="bg-gradient-to-br from-blue-50/50 via-white to-blue-50/30 border-2 border-blue-200/80 rounded-2xl p-6 lg:p-7 flex flex-col justify-between shadow-sm">
                <div>
                  <div className="flex items-center justify-between mb-4">
                    <span className="px-3 py-1 bg-blue-100 text-blue-800 text-xs font-bold uppercase tracking-wider rounded-full flex items-center gap-1.5">
                      <FileUp size={13} />
                      Nạp dữ liệu vào máy này
                    </span>
                    <span className="text-xs font-mono text-blue-600 font-semibold">Khôi phục / Chuyển máy</span>
                  </div>

                  <h3 className="text-xl font-bold text-gray-900 mb-2 flex items-center gap-2">
                    2. Nạp dữ liệu từ tệp JSON
                  </h3>
                  <p className="text-sm text-gray-600 leading-relaxed mb-5">
                    Chọn tệp tin sao lưu <strong className="text-blue-700">.json</strong> đã xuất từ máy tính khác để tải toàn bộ dữ liệu vào trình duyệt của máy này.
                  </p>

                  {/* Dropzone & File Selector */}
                  <label
                    onDragOver={(e) => e.preventDefault()}
                    onDrop={(e) => {
                      e.preventDefault();
                      const file = e.dataTransfer.files?.[0];
                      if (file) handleBackupFileSelect(file);
                    }}
                    className={cn(
                      "block w-full border-2 border-dashed rounded-xl p-5 text-center cursor-pointer transition-all",
                      selectedBackupFile 
                        ? "border-blue-500 bg-blue-50/40" 
                        : "border-gray-300 hover:border-blue-400 bg-white hover:bg-blue-50/20"
                    )}
                  >
                    <input
                      type="file"
                      accept=".json,application/json"
                      className="hidden"
                      onChange={(e) => {
                        const file = e.target.files?.[0];
                        if (file) handleBackupFileSelect(file);
                      }}
                    />
                    {selectedBackupFile ? (
                      <div className="flex items-center justify-center gap-3">
                        <div className="bg-blue-600 text-white p-2.5 rounded-lg">
                          <FileJson size={22} />
                        </div>
                        <div className="text-left">
                          <p className="text-sm font-bold text-gray-900 truncate max-w-[240px] sm:max-w-[320px]">
                            {selectedBackupFile.name}
                          </p>
                          <p className="text-xs text-gray-500 font-mono">
                            {(selectedBackupFile.size / 1024).toFixed(1)} KB • Nhấp để chọn tệp khác
                          </p>
                        </div>
                      </div>
                    ) : (
                      <div className="space-y-1.5">
                        <div className="mx-auto w-10 h-10 rounded-full bg-blue-100 flex items-center justify-center text-blue-600 mb-2">
                          <Upload size={20} />
                        </div>
                        <p className="text-sm font-bold text-gray-800">
                          Nhấp để chọn tệp tin JSON hoặc kéo thả vào đây
                        </p>
                        <p className="text-xs text-gray-400">
                          Chấp nhận các tệp sao lưu .json của hệ thống
                        </p>
                      </div>
                    )}
                  </label>

                  {/* Validation Info */}
                  {importValidation && !importValidation.valid && (
                    <div className="mt-4 p-3.5 bg-red-50 border border-red-200 rounded-xl text-xs text-red-700 flex items-start gap-2">
                      <AlertCircle size={16} className="text-red-600 shrink-0 mt-0.5" />
                      <div>
                        <strong className="block">Tệp JSON không hợp lệ:</strong>
                        <span>{importValidation.error}</span>
                      </div>
                    </div>
                  )}

                  {importValidation && importValidation.valid && (
                    <div className="mt-4 bg-white border border-blue-200 rounded-xl p-4 space-y-3">
                      <div className="flex items-center justify-between border-b border-gray-100 pb-2">
                        <span className="text-xs font-bold text-gray-700 uppercase tracking-wider flex items-center gap-1.5">
                          <CheckCircle2 size={15} className="text-emerald-600" />
                          Phát hiện dữ liệu hợp lệ trong file
                        </span>
                        {importValidation.metadata?.exportDateFormatted && (
                          <span className="text-[11px] font-mono text-gray-500">
                            Ngày xuất: {importValidation.metadata.exportDateFormatted}
                          </span>
                        )}
                      </div>

                      <div className="grid grid-cols-3 gap-2 text-center text-xs">
                        <div className="bg-blue-50/60 p-2 rounded-lg">
                          <span className="text-gray-500 block text-[11px]">Linh kiện</span>
                          <span className="font-bold text-sm text-blue-700 font-mono">{importValidation.stats?.partsCount || 0}</span>
                        </div>
                        <div className="bg-blue-50/60 p-2 rounded-lg">
                          <span className="text-gray-500 block text-[11px]">Lệnh SX (PO)</span>
                          <span className="font-bold text-sm text-blue-700 font-mono">{importValidation.stats?.posCount || 0}</span>
                        </div>
                        <div className="bg-blue-50/60 p-2 rounded-lg">
                          <span className="text-gray-500 block text-[11px]">Tồn kho WIP</span>
                          <span className="font-bold text-sm text-blue-700 font-mono">{importValidation.stats?.inventoryCount || 0}</span>
                        </div>
                      </div>

                      {/* Mode selection */}
                      <div className="pt-2 border-t border-gray-100 space-y-2">
                        <label className="text-[11px] font-bold uppercase tracking-wider text-gray-600 block">
                          Chế độ nạp dữ liệu:
                        </label>
                        <div className="grid grid-cols-1 gap-2">
                          <label className={cn(
                            "flex items-start gap-3 p-2.5 rounded-lg border cursor-pointer transition-all",
                            importMode === 'overwrite' 
                              ? "border-blue-600 bg-blue-50/50" 
                              : "border-gray-200 hover:bg-gray-50"
                          )}>
                            <input 
                              type="radio" 
                              name="importMode" 
                              value="overwrite" 
                              checked={importMode === 'overwrite'}
                              onChange={() => setImportMode('overwrite')}
                              className="mt-0.5"
                            />
                            <div className="text-xs">
                              <span className="font-bold text-gray-900 block">
                                Ghi đè toàn bộ (Khuyên dùng khi chuyển máy mới)
                              </span>
                              <span className="text-gray-500 text-[11px]">
                                Xóa dữ liệu cũ trên máy này và thay thế 100% bằng dữ liệu từ tệp JSON.
                              </span>
                            </div>
                          </label>

                          <label className={cn(
                            "flex items-start gap-3 p-2.5 rounded-lg border cursor-pointer transition-all",
                            importMode === 'merge' 
                              ? "border-blue-600 bg-blue-50/50" 
                              : "border-gray-200 hover:bg-gray-50"
                          )}>
                            <input 
                              type="radio" 
                              name="importMode" 
                              value="merge" 
                              checked={importMode === 'merge'}
                              onChange={() => setImportMode('merge')}
                              className="mt-0.5"
                            />
                            <div className="text-xs">
                              <span className="font-bold text-gray-900 block">
                                Hợp nhất dữ liệu (Merge)
                              </span>
                              <span className="text-gray-500 text-[11px]">
                                Giữ lại các bản ghi hiện tại và bổ sung thêm các bản ghi mới từ tệp JSON.
                              </span>
                            </div>
                          </label>
                        </div>
                      </div>
                    </div>
                  )}

                  {importResult && (
                    <div className={cn(
                      "mt-4 p-4 rounded-xl border text-xs flex flex-col gap-2 animate-fadeIn",
                      importResult.success 
                        ? "bg-emerald-50 border-emerald-300 text-emerald-900" 
                        : "bg-red-50 border-red-300 text-red-900"
                    )}>
                      <div className="flex items-center gap-2">
                        {importResult.success ? (
                          <CheckCircle2 size={18} className="text-emerald-600 shrink-0" />
                        ) : (
                          <AlertCircle size={18} className="text-red-600 shrink-0" />
                        )}
                        <span className="font-bold text-sm">{importResult.message}</span>
                      </div>
                      {importResult.success && (
                        <div className="flex items-center gap-2 pt-1">
                          <button
                            type="button"
                            onClick={() => window.location.reload()}
                            className="px-3 py-1.5 bg-emerald-700 hover:bg-emerald-800 text-white rounded-lg font-bold text-xs flex items-center gap-1.5 transition-all shadow-sm"
                          >
                            <RefreshCw size={12} />
                            Tải lại trang (F5) để áp dụng toàn diện
                          </button>
                          <span className="text-[11px] text-emerald-700">Dữ liệu đã được áp dụng vào giao diện ngay lập tức.</span>
                        </div>
                      )}
                    </div>
                  )}
                </div>

                <div className="pt-4">
                  <button
                    type="button"
                    disabled={!importValidation?.valid || isProcessingImport}
                    onClick={handleExecuteImport}
                    className={cn(
                      "w-full py-4 px-6 rounded-xl font-bold uppercase tracking-wider flex items-center justify-center gap-3 shadow-lg transition-all text-sm",
                      importValidation?.valid && !isProcessingImport
                        ? "bg-blue-600 hover:bg-blue-700 text-white shadow-blue-600/20 active:scale-[0.99] cursor-pointer"
                        : "bg-gray-200 text-gray-400 cursor-not-allowed shadow-none"
                    )}
                  >
                    <HardDrive size={18} />
                    {isProcessingImport ? 'ĐANG NẠP DỮ LIỆU...' : 'XÁC NHẬN NẠP DỮ LIỆU VÀO MÁY NÀY'}
                  </button>
                </div>
              </div>
            </div>

            {/* Step-by-Step Migration Guide */}
            <div className="bg-white border border-gray-200 rounded-2xl p-6">
              <h4 className="text-sm font-bold uppercase tracking-wider text-gray-800 mb-4 flex items-center gap-2">
                <FileText size={18} className="text-blue-600" />
                Hướng dẫn chuyển dữ liệu từ máy này sang máy khác:
              </h4>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div className="p-4 bg-gray-50 rounded-xl border border-gray-100 flex gap-3.5">
                  <span className="w-7 h-7 rounded-full bg-emerald-600 text-white font-bold font-mono text-sm flex items-center justify-center shrink-0">
                    1
                  </span>
                  <div>
                    <strong className="text-sm font-bold text-gray-900 block mb-1">
                      Tại máy cũ (có dữ liệu)
                    </strong>
                    <p className="text-xs text-gray-600 leading-relaxed">
                      Mở mục này và bấm nút <strong>"TẢI XUỐNG TỆP SAO LƯU (.JSON)"</strong>. Một tệp tin JSON sẽ được tải về thư mục Downloads của máy tính.
                    </p>
                  </div>
                </div>

                <div className="p-4 bg-gray-50 rounded-xl border border-gray-100 flex gap-3.5">
                  <span className="w-7 h-7 rounded-full bg-blue-600 text-white font-bold font-mono text-sm flex items-center justify-center shrink-0">
                    2
                  </span>
                  <div>
                    <strong className="text-sm font-bold text-gray-900 block mb-1">
                      Sao chép tệp tin
                    </strong>
                    <p className="text-xs text-gray-600 leading-relaxed">
                      Sao chép tệp JSON vừa tải sang máy tính mới qua USB, Zalo, Email hoặc Google Drive.
                    </p>
                  </div>
                </div>

                <div className="p-4 bg-gray-50 rounded-xl border border-gray-100 flex gap-3.5">
                  <span className="w-7 h-7 rounded-full bg-purple-600 text-white font-bold font-mono text-sm flex items-center justify-center shrink-0">
                    3
                  </span>
                  <div>
                    <strong className="text-sm font-bold text-gray-900 block mb-1">
                      Tại máy tính mới
                    </strong>
                    <p className="text-xs text-gray-600 leading-relaxed">
                      Mở ứng dụng trên máy mới, vào mục <strong>"Sao lưu & Nạp dữ liệu (JSON)"</strong>, chọn tệp tin JSON và bấm <strong>"XÁC NHẬN NẠP DỮ LIỆU"</strong>.
                    </p>
                  </div>
                </div>
              </div>
            </div>
          </div>
        ) : activeSettingsTab === 'parts' ? (
          <>
            <div className="p-6 sm:p-8 border-b border-gray-100 bg-gray-50/50 space-y-4">
              <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                <div className="flex items-center gap-3">
                  <h2 className="font-bold text-2xl tracking-tight text-gray-900">
                    Linh kiện ({filteredParts.length}{filteredParts.length !== parts.length ? `/${parts.length}` : ''})
                  </h2>
                  {partSearchQuery && (
                    <span className="text-xs bg-blue-50 text-blue-700 font-bold px-2.5 py-1 rounded-full border border-blue-200">
                      Đang lọc: &ldquo;{partSearchQuery}&rdquo;
                    </span>
                  )}
                </div>
                <label className={cn(
                  "flex items-center gap-2.5 px-5 py-2.5 bg-green-600 text-white rounded-xl text-xs font-bold uppercase tracking-widest cursor-pointer hover:bg-green-700 transition-all shadow-sm active:scale-95 shrink-0",
                  isImporting && "opacity-50 cursor-not-allowed"
                )}>
                  <FileUp size={18} />
                  {isImporting ? 'Đang nhập...' : 'Nhập Excel'}
                  <input 
                    type="file" 
                    accept=".xlsx, .xls" 
                    className="hidden" 
                    onChange={handleExcelImport}
                    disabled={isImporting}
                  />
                </label>
              </div>

              {/* Fast Search & Level Filter Controls */}
              <div className="flex flex-col sm:flex-row gap-3 items-stretch sm:items-center">
                <div className="relative flex-1">
                  <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
                  <input 
                    type="text" 
                    value={partSearchQuery}
                    onChange={(e) => setPartSearchQuery(e.target.value)}
                    placeholder="Tìm nhanh theo Mã LK, Tên linh kiện, Đơn vị..."
                    className="w-full pl-10 pr-9 py-2.5 bg-white border border-gray-200 rounded-xl text-sm focus:border-blue-500 focus:ring-4 focus:ring-blue-100 outline-none transition-all placeholder:text-gray-400 font-medium"
                  />
                  {partSearchQuery && (
                    <button 
                      onClick={() => setPartSearchQuery('')}
                      className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600 p-0.5 rounded-full hover:bg-gray-100 cursor-pointer"
                      title="Xóa tìm kiếm"
                    >
                      <X size={15} />
                    </button>
                  )}
                </div>

                <div className="flex items-center gap-1 bg-gray-200/60 p-1 rounded-xl shrink-0 overflow-x-auto no-scrollbar">
                  {(['ALL', 1, 2, 3] as const).map(lvl => {
                    const count = lvl === 'ALL' ? parts.length : parts.filter(p => (p.level || 1) === lvl).length;
                    return (
                      <button
                        key={lvl}
                        onClick={() => setPartLevelFilter(lvl)}
                        className={cn(
                          "px-3 py-1.5 rounded-lg text-xs font-bold transition-all whitespace-nowrap cursor-pointer",
                          partLevelFilter === lvl 
                            ? "bg-white text-gray-900 shadow-xs" 
                            : "text-gray-500 hover:text-gray-800"
                        )}
                      >
                        {lvl === 'ALL' ? `Tất cả (${count})` : `Cấp ${lvl} (${count})`}
                      </button>
                    );
                  })}
                </div>
              </div>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-gray-50/50 border-b border-gray-100">
                    <th className="p-6 pl-10 text-xs font-mono uppercase opacity-50">Cấp</th>
                    <th className="p-6 text-xs font-mono uppercase opacity-50">Mã LK</th>
                    <th className="p-6 text-xs font-mono uppercase opacity-50">Linh kiện</th>
                    <th className="p-6 text-xs font-mono uppercase opacity-50">Quy trình</th>
                    <th className="p-6 text-xs font-mono uppercase opacity-50">ĐVT</th>
                    <th className="p-6 pr-10 text-xs font-mono uppercase opacity-50 text-right">Thao tác</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-50">
                  {filteredParts.map(part => (
                    <tr key={part.id} className="hover:bg-gray-50/30 transition-colors">
                      <td className="p-6 pl-10">
                        <span className={cn(
                          "px-3 py-1 rounded text-xs font-bold",
                          part.level === 1 ? "bg-purple-100 text-purple-700" :
                          part.level === 2 ? "bg-blue-100 text-blue-700" :
                          "bg-green-100 text-green-700"
                        )}>
                          L{part.level || 1}
                        </span>
                      </td>
                      <td className="p-6 font-mono text-base font-bold">{part.id}</td>
                      <td className="p-6">
                        <div className="text-base font-medium">{part.name}</div>
                        <div className="flex gap-2 mt-1">
                          {part.skipLaser && <span className="px-1.5 py-0.5 bg-yellow-50 text-yellow-600 text-[9px] font-bold rounded border border-yellow-100">MIỄN LASER</span>}
                          {part.skipBending && <span className="px-1.5 py-0.5 bg-red-50 text-red-600 text-[9px] font-bold rounded border border-red-100">MIỄN CHẤN</span>}
                          {part.skipWelding && <span className="px-1.5 py-0.5 bg-orange-50 text-orange-600 text-[9px] font-bold rounded border border-orange-100">MIỄN HÀN</span>}
                          {part.skipPainting && <span className="px-1.5 py-0.5 bg-gray-100 text-gray-600 text-[9px] font-bold rounded border border-gray-200">MIỄN SƠN</span>}
                          {part.hasPaintingPO && <span className="px-1.5 py-0.5 bg-blue-600 text-white text-[9px] font-bold rounded shadow-sm">PAINT PO</span>}
                        </div>
                      </td>
                      <td className="p-6">
                        <div className="flex flex-col gap-1 text-[10px] uppercase font-bold text-gray-400">
                          <span className={cn(part.skipLaser ? "line-through opacity-30" : "text-blue-600")}>Laser</span>
                          <span className={cn(part.skipBending ? "line-through opacity-30" : "text-blue-600")}>Chấn</span>
                          <span className={cn(part.skipWelding ? "line-through opacity-30" : "text-blue-600")}>Hàn</span>
                          <span className={cn(part.skipPainting ? "line-through opacity-30" : "text-blue-600")}>Sơn</span>
                        </div>
                      </td>
                      <td className="p-6 text-base text-gray-500">{part.unit}</td>
                      <td className="p-6 pr-10 text-right">
                        <div className="flex justify-end gap-3">
                          <button 
                            onClick={() => handleEdit(part)}
                            className="p-3 text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
                          >
                            <Edit2 size={20} />
                          </button>
                          <button 
                            onClick={() => handleDelete(part.id)}
                            className="p-3 text-red-600 hover:bg-red-50 rounded-lg transition-colors"
                          >
                            <Trash2 size={20} />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                  {filteredParts.length === 0 && (
                    <tr>
                      <td colSpan={6} className="p-16 text-center text-gray-400">
                        <Package size={40} className="mx-auto mb-2 opacity-30" />
                        <p className="text-sm font-medium">
                          {parts.length === 0 ? 'Chưa có linh kiện nào trong danh mục.' : `Không tìm thấy linh kiện nào phù hợp với tìm kiếm "${partSearchQuery}".`}
                        </p>
                        {parts.length > 0 && (partSearchQuery || partLevelFilter !== 'ALL') && (
                          <button 
                            type="button"
                            onClick={() => { setPartSearchQuery(''); setPartLevelFilter('ALL'); }}
                            className="mt-3 px-3.5 py-1.5 text-xs font-bold text-blue-600 bg-blue-50 hover:bg-blue-100 rounded-lg transition cursor-pointer"
                          >
                            Xóa bộ lọc tìm kiếm
                          </button>
                        )}
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </>
        ) : (
          <>
            <div className="p-8 border-b border-gray-100 flex justify-between items-center bg-gray-50/50">
              <div className="flex flex-col">
                <h2 className="font-bold text-2xl tracking-tight">Định mức sản xuất (KG)</h2>
                <p className="text-xs text-gray-400">Cột yêu cầu: Mã linh kiện, Mã linh kiện tôn, Khối lượng (Kg)</p>
              </div>
              <label className={cn(
                "flex items-center gap-3 px-6 py-3 bg-blue-600 text-white rounded-lg text-sm font-bold uppercase tracking-widest cursor-pointer hover:bg-blue-700 transition-all",
                isImportingBOM && "opacity-50 cursor-not-allowed"
              )}>
                <FileUp size={20} />
                {isImportingBOM ? 'Đang nhập...' : 'Nhập Định Mức'}
                <input 
                  type="file" 
                  accept=".xlsx, .xls" 
                  className="hidden" 
                  onChange={handleBOMImport}
                  disabled={isImportingBOM}
                />
              </label>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-gray-50/50 border-b border-gray-100">
                    <th className="p-6 pl-10 text-xs font-mono uppercase opacity-50">Mã linh kiện tôn</th>
                    <th className="p-6 text-xs font-mono uppercase opacity-50">Mã linh kiện</th>
                    <th className="p-6 text-xs font-mono uppercase opacity-50">KL Linh kiện (kg)</th>
                    <th className="p-6 text-xs font-mono uppercase opacity-50">KL Phế (kg)</th>
                    <th className="p-6 text-xs font-mono uppercase opacity-50">Tổng (kg)</th>
                    <th className="p-6 pr-10 text-xs font-mono uppercase opacity-50 text-right">Thao tác</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-50">
                  {storageService.getBOM().map((bom, idx) => (
                    <tr key={idx} className="hover:bg-gray-50/30 transition-colors">
                      <td className="p-6 pl-10 font-mono text-base">{bom.parentPartId}</td>
                      <td className="p-6 font-mono text-base">{bom.childPartId}</td>
                      <td className="p-6 font-mono text-base font-bold text-blue-600">{bom.componentWeight.toFixed(4)}</td>
                      <td className="p-6 font-mono text-base font-bold text-orange-600">{bom.scrapWeight.toFixed(4)}</td>
                      <td className="p-6 font-mono text-base font-bold">{(bom.componentWeight + bom.scrapWeight).toFixed(4)}</td>
                      <td className="p-6 pr-10 text-right">
                        <button 
                          onClick={() => {
                            const current = storageService.getBOM();
                            current.splice(idx, 1);
                            storageService.saveBOM(current);
                            onPartsChange();
                          }}
                          className="p-3 text-red-600 hover:bg-red-50 rounded-lg transition-colors"
                        >
                          <Trash2 size={20} />
                        </button>
                      </td>
                    </tr>
                  ))}
                  {storageService.getBOM().length === 0 && (
                    <tr>
                      <td colSpan={6} className="p-12 text-center text-gray-400 italic">Chưa có dữ liệu định mức. Vui lòng nhập file Excel.</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </>
        )}
      </div>
    </motion.div>
  );
}

function NormsView({ parts, onNormsChange }: { parts: Part[], onNormsChange: () => void }) {
  const [activeTab, setActiveTab] = useState<StageId | 'NESTING' | 'BW_HSQD' | 'LASER_HSQD' | 'GLAZING_STAGE' | 'PAINTING_HSQD'>('NESTING');
  const [bendingWeldingHSQD, setBendingWeldingHSQD] = useState<{partId: string, hsqd: number, partName?: string}[]>(() => storageService.getBendingWeldingHSQD());
  const [laserHSQD, setLaserHSQD] = useState<{partId: string, hsqd: number, partName?: string}[]>(() => storageService.getLaserHSQD());
  const [glazingStageHSQD, setGlazingStageHSQD] = useState<{partId: string, hsqd: number, partName?: string}[]>(() => storageService.getGlazingStageHSQD());
  const [paintingHSQD, setPaintingHSQD] = useState<{partId: string, hsqd: number, partName?: string}[]>(() => storageService.getPaintingHSQD());
  const [isImporting, setIsImporting] = useState(false);
  const norms = storageService.getNorms();
  const nesting = storageService.getLaserNesting();

  const handleImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsImporting(true);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const dataBuffer = new Uint8Array(evt.target?.result as ArrayBuffer);
        const wb = XLSX.read(dataBuffer, { type: 'array' });
        
        // Helper chuẩn hóa tên cột loại bỏ dấu tiếng Việt, NBSP, khoảng trắng và ký tự đặc biệt
        const cleanKey = (s: any): string => {
          if (!s) return '';
          return String(s)
            .replace(/\u00A0/g, ' ')
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/đ/g, 'd')
            .replace(/Đ/g, 'd')
            .toLowerCase()
            .replace(/[^a-z0-9]/g, '');
        };

        const parseNum = (val: any): number => {
          if (typeof val === 'number') return isNaN(val) ? 0 : val;
          if (val === undefined || val === null || val === '') return 0;
          const s = String(val).replace(/\s/g, '').replace(',', '.');
          const n = parseFloat(s);
          return isNaN(n) ? 0 : n;
        };

        if (activeTab === 'NESTING') {
          // Xử lý Tổ hợp Laser (Nesting)
          let rawImported: any[] = [];
          for (const sheetName of wb.SheetNames) {
            const ws = wb.Sheets[sheetName];
            const rawGrid: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });
            if (!rawGrid || rawGrid.length === 0) continue;

            let headerRowIdx = -1;
            let nestCol = -1, partCol = -1, qtyCol = -1, secCol = -1, modelCol = -1;

            for (let r = 0; r < Math.min(rawGrid.length, 10); r++) {
              const row = rawGrid[r] || [];
              let nC = -1, pC = -1, qC = -1, sC = -1, mC = -1;
              row.forEach((cell, cIdx) => {
                const k = cleanKey(cell);
                if (k.includes('nesting') || k.includes('maban') || k.includes('matohop') || k.includes('maphoi') || k.includes('matam') || k.includes('idtohop')) nC = cIdx;
                if (k.includes('partid') || k.includes('malinhkien') || k.includes('malk') || k.includes('masp') || k.includes('mahang') || k.includes('linhkien') || k.includes('part')) pC = cIdx;
                if (k.includes('qty') || k.includes('soluong') || k.includes('sltam')) qC = cIdx;
                if (k.includes('giay') || k.includes('seconds') || k.includes('thoigian') || k.includes('dinhmuc')) sC = cIdx;
                if (k.includes('model')) mC = cIdx;
              });

              if (nC !== -1 && (pC !== -1 || qC !== -1)) {
                headerRowIdx = r;
                nestCol = nC;
                partCol = pC !== -1 ? pC : 1;
                qtyCol = qC !== -1 ? qC : 2;
                secCol = sC !== -1 ? sC : 3;
                modelCol = mC;
                break;
              }
            }

            if (headerRowIdx !== -1) {
              for (let r = headerRowIdx + 1; r < rawGrid.length; r++) {
                const row = rawGrid[r];
                if (!row) continue;
                const nestId = String(row[nestCol] || '').trim();
                const partIdRaw = String(row[partCol] || '').trim();
                const qty = parseNum(row[qtyCol]);
                const seconds = parseNum(row[secCol]);
                const model = modelCol !== -1 && row[modelCol] ? String(row[modelCol]).trim() : undefined;

                if (nestId && (partIdRaw || qty > 0)) {
                  let finalPartId = partIdRaw;
                  if (!finalPartId) {
                    const found = parts.find(p => cleanKey(p.name) === cleanKey(partIdRaw));
                    if (found) finalPartId = found.id;
                  }
                  rawImported.push({
                    nestingId: nestId,
                    partId: finalPartId,
                    qtyPerSheet: qty > 0 ? qty : 1,
                    secondsPerSheet: seconds,
                    applicableModel: model
                  });
                }
              }
              if (rawImported.length > 0) break;
            }
          }

          if (rawImported.length === 0) {
            alert('Không tìm thấy dữ liệu hợp lệ. Cần các cột: Mã bàn (Nesting ID), Linh kiện kết hợp, Số lượng / Tấm, Thời gian 1 tấm, Model áp dụng');
          } else {
            const nestingTotals = new Map<string, number>();
            const nestingModels = new Map<string, string>();
            rawImported.forEach(row => {
              if (row.secondsPerSheet > 0) nestingTotals.set(row.nestingId, row.secondsPerSheet);
              if (row.applicableModel) nestingModels.set(row.nestingId, row.applicableModel);
            });
            const imported: LaserNesting[] = rawImported.map(row => ({
              nestingId: row.nestingId,
              partId: row.partId,
              qtyPerSheet: row.qtyPerSheet,
              secondsPerUnit: 0,
              secondsPerSheet: nestingTotals.get(row.nestingId) || row.secondsPerSheet || 0,
              applicableModel: row.applicableModel || nestingModels.get(row.nestingId)
            }));
            storageService.saveLaserNesting(imported);
            alert(`Đã nhập thành công ${imported.length} linh kiện vào ${nestingTotals.size} tổ hợp Laser!`);
            onNormsChange();
          }
        } else if (activeTab === 'LASER_HSQD' || activeTab === 'BW_HSQD' || activeTab === 'GLAZING_STAGE' || activeTab === 'PAINTING_HSQD') {
          // Xử lý HSQĐ Laser, HSQĐ Chấn/Hàn, HSQĐ Dán Kính 2 hoặc HSQĐ Sơn
          interface SheetCandidate {
            sheetName: string;
            score: number;
            items: { partId: string; partName?: string; hsqd: number }[];
          }

          const candidates: SheetCandidate[] = [];

          for (const sheetName of wb.SheetNames) {
            const ws = wb.Sheets[sheetName];
            const rawGrid: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });
            if (!rawGrid || rawGrid.length < 2) continue;

            const sNameClean = cleanKey(sheetName);
            let sheetBaseScore = 0;
            if (sNameClean.includes('hsqd') || sNameClean.includes('quydoi') || sNameClean.includes('heso')) sheetBaseScore += 100;
            if (activeTab === 'LASER_HSQD' && sNameClean.includes('laser')) sheetBaseScore += 50;
            if (activeTab === 'BW_HSQD' && (sNameClean.includes('chan') || sNameClean.includes('han') || sNameClean.includes('dap'))) sheetBaseScore += 50;
            if (activeTab === 'GLAZING_STAGE' && (sNameClean.includes('kinh') || sNameClean.includes('dan'))) sheetBaseScore += 60;
            if (activeTab === 'PAINTING_HSQD' && (sNameClean.includes('son') || sNameClean.includes('paint'))) sheetBaseScore += 60;

            let bestHeaderRowIdx = -1;
            let bestIdCol = -1, bestNameCol = -1, bestHsqdCol = -1;
            let bestHeaderScore = -1;

            for (let r = 0; r < Math.min(rawGrid.length, 15); r++) {
              const row = rawGrid[r] || [];
              let foundId = -1, foundName = -1, foundHsqd = -1;
              let hasExactHsqd = false;
              let hasExactId = false;

              row.forEach((cell, cIdx) => {
                const k = cleanKey(cell);
                if (!k) return;

                // STT / Order TUYỆT ĐỐI không phải là Mã linh kiện / Mã thành phẩm
                if (k === 'stt' || k === 'sothutu' || k === 'tt' || k === 'no' || k === 'order' || k === 'index') {
                  return;
                }

                // Nhận diện Mã thành phẩm / Mã linh kiện
                if (k.includes('mathanhpham') || k.includes('malinhkien') || k.includes('partid') || k.includes('masanpham') || k.includes('malk') || k.includes('matp') || k.includes('masp') || k.includes('mahang') || k.includes('partcode') || k.includes('itemcode')) {
                  foundId = cIdx;
                  hasExactId = true;
                } else if (foundId === -1 && (k.startsWith('ma') || k.endsWith('code')) && !k.includes('dinhmuc') && !k.includes('may') && !k.includes('mau')) {
                  foundId = cIdx;
                }

                // Nhận diện Tên thành phẩm / Tên linh kiện
                if (k.includes('tenthanhpham') || k.includes('tenlinhkien') || k.includes('partname') || k.includes('tensanpham') || k.includes('tenlk') || k.includes('tentp') || k.includes('tensp') || k.includes('tenhang') || k.includes('description') || k.startsWith('ten')) {
                  foundName = cIdx;
                }

                // Nhận diện HSQĐ: Ưu tiên tuyệt đối cột có chữ HSQĐ / Hệ số / Quy đổi
                if (k.includes('hsqd') || k.includes('heso') || k.includes('quydoi') || k.includes('coefficient') || k.includes('ratio')) {
                  foundHsqd = cIdx;
                  hasExactHsqd = true;
                } else if (foundHsqd === -1 && (k.includes('dinhmuc') || k.includes('norm') || k.includes('tyle'))) {
                  // Chỉ dùng tạm nếu chưa có cột hsqd thực sự
                  foundHsqd = cIdx;
                }
              });

              if (foundId !== -1 && foundHsqd !== -1) {
                let rowScore = (hasExactHsqd ? 120 : 20) + (hasExactId ? 60 : 10) + (foundName !== -1 ? 20 : 0);
                if (rowScore > bestHeaderScore) {
                  bestHeaderScore = rowScore;
                  bestHeaderRowIdx = r;
                  bestIdCol = foundId;
                  bestNameCol = foundName;
                  bestHsqdCol = foundHsqd;
                }
              }
            }

            // Fallback nếu không phát hiện được chữ header
            if (bestHeaderRowIdx === -1 && rawGrid.length >= 2) {
              const r1 = rawGrid[1] || [];
              if (r1.length >= 2) {
                bestHeaderRowIdx = 0;
                // Kiểm tra xem cột 0 có phải STT (1, 2, 3...) không
                const isCol0Seq = typeof r1[0] === 'number' || (typeof r1[0] === 'string' && /^\d+$/.test(r1[0].trim()));
                if (isCol0Seq && r1.length >= 3) {
                  // Cột 0 là STT -> Bỏ qua cột 0, lấy cột 1 làm Mã
                  bestIdCol = 1;
                  bestNameCol = r1.length >= 4 ? 2 : -1;
                  bestHsqdCol = r1.length >= 4 ? 3 : 2;
                } else {
                  bestIdCol = 0;
                  bestNameCol = r1.length >= 3 ? 1 : -1;
                  bestHsqdCol = r1.length >= 3 ? 2 : 1;
                }
                bestHeaderScore = 5;
              }
            }

            if (bestHeaderRowIdx !== -1 && bestIdCol !== -1 && bestHsqdCol !== -1) {
              const sheetItems: { partId: string; partName?: string; hsqd: number }[] = [];
              let isIdNumericSequence = true;
              let sampleCount = 0;

              for (let r = bestHeaderRowIdx + 1; r < rawGrid.length; r++) {
                const row = rawGrid[r];
                if (!row) continue;
                let partIdRaw = String(row[bestIdCol] || '').trim();
                const partNameRaw = bestNameCol !== -1 && row[bestNameCol] !== undefined ? String(row[bestNameCol]).trim() : '';
                const hsqd = parseNum(row[bestHsqdCol]);

                // Nếu partIdRaw là số đơn giản 1, 2, 3... (như STT) nhưng cột khác có mã sản phẩm dạng '04-29-...'
                if (/^\d{1,4}$/.test(partIdRaw)) {
                  for (let c = 0; c < row.length; c++) {
                    if (c !== bestIdCol && c !== bestHsqdCol) {
                      const val = String(row[c] || '').trim();
                      if (val.length > 5 && (val.includes('-') || /[a-zA-Z]/.test(val))) {
                        partIdRaw = val; // Đây mới là mã sản phẩm thực sự!
                        break;
                      }
                    }
                  }
                }

                let finalPartId = partIdRaw;
                if (!finalPartId && partNameRaw) {
                  const found = parts.find(p => cleanKey(p.name) === cleanKey(partNameRaw));
                  if (found) finalPartId = found.id;
                }

                if (finalPartId && hsqd > 0) {
                  sampleCount++;
                  if (!/^\d+$/.test(finalPartId)) {
                    isIdNumericSequence = false;
                  }

                  const existingIdx = sheetItems.findIndex(x => x.partId.toUpperCase() === finalPartId.toUpperCase());
                  if (existingIdx >= 0) {
                    sheetItems[existingIdx] = { partId: finalPartId, partName: partNameRaw || sheetItems[existingIdx].partName, hsqd };
                  } else {
                    sheetItems.push({ partId: finalPartId, partName: partNameRaw || undefined, hsqd });
                  }
                }
              }

              if (sheetItems.length > 0) {
                let totalScore = sheetBaseScore + bestHeaderScore;
                // Nếu mã sản phẩm là mã thực sự (có chữ cái hoặc dấu gạch như '04-29-08-SHA76210KL-0017'), cộng điểm rất lớn
                if (!isIdNumericSequence && sampleCount > 0) {
                  totalScore += 250;
                } else if (isIdNumericSequence && sampleCount > 0) {
                  // Nếu toàn là số STT 1, 2, 3... thì trừ điểm nặng vì chắc chắn không phải bảng HSQĐ chuẩn
                  totalScore -= 200;
                }

                candidates.push({
                  sheetName,
                  score: totalScore,
                  items: sheetItems
                });
              }
            }
          }

          // Sắp xếp các sheet theo điểm cao nhất trước
          candidates.sort((a, b) => b.score - a.score);
          const bestCandidate = candidates[0];

          if (!bestCandidate || bestCandidate.items.length === 0) {
            alert(`Không tìm thấy dữ liệu hợp lệ. Cần các cột: Mã thành phẩm / Mã linh kiện, Tên thành phẩm (tùy chọn), HSQĐ`);
          } else {
            const imported = bestCandidate.items;
            if (activeTab === 'LASER_HSQD') {
              storageService.saveLaserHSQD(imported);
              setLaserHSQD(imported);
              alert(`Đã nhập thành công ${imported.length} HSQĐ cho Cắt Laser (từ sheet "${bestCandidate.sheetName}")!`);
            } else if (activeTab === 'GLAZING_STAGE') {
              storageService.saveGlazingStageHSQD(imported);
              setGlazingStageHSQD(imported);
              alert(`Đã nhập thành công ${imported.length} HSQĐ cho Dán Kính 2 (từ sheet "${bestCandidate.sheetName}")!`);
            } else if (activeTab === 'PAINTING_HSQD') {
              storageService.savePaintingHSQD(imported);
              setPaintingHSQD(imported);
              alert(`Đã nhập thành công ${imported.length} HSQĐ cho Sơn (từ sheet "${bestCandidate.sheetName}")!`);
            } else {
              storageService.saveBendingWeldingHSQD(imported);
              setBendingWeldingHSQD(imported);
              alert(`Đã nhập thành công ${imported.length} HSQĐ cho Chấn/Hàn (từ sheet "${bestCandidate.sheetName}")!`);
            }
            onNormsChange();
          }
        } else {
          // Standard Stage Norm Import (BENDING, WELDING, PAINTING, DCLR, GLAZING)
          interface NormCandidate {
            sheetName: string;
            score: number;
            items: ProductivityNorm[];
          }

          const normCandidates: NormCandidate[] = [];

          for (const sheetName of wb.SheetNames) {
            const ws = wb.Sheets[sheetName];
            const rawGrid: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });
            if (!rawGrid || rawGrid.length < 2) continue;

            const sNameClean = cleanKey(sheetName);
            let sheetBaseScore = 0;
            const currentStageNameClean = cleanKey(STAGES.find(s => s.id === activeTab)?.name || '');
            if (sNameClean.includes(currentStageNameClean)) sheetBaseScore += 50;

            let headerRowIdx = -1;
            let idCol = -1, nameCol = -1, normCol = -1;
            let bestScore = -1;

            for (let r = 0; r < Math.min(rawGrid.length, 15); r++) {
              const row = rawGrid[r] || [];
              let foundId = -1, foundName = -1, foundNorm = -1;
              let hasExactId = false;

              row.forEach((cell, cIdx) => {
                const k = cleanKey(cell);
                if (!k) return;

                // STT / Order không phải là Mã linh kiện
                if (k === 'stt' || k === 'sothutu' || k === 'tt' || k === 'no' || k === 'order' || k === 'index') {
                  return;
                }

                if (k.includes('mathanhpham') || k.includes('malinhkien') || k.includes('partid') || k.includes('masanpham') || k.includes('malk') || k.includes('matp') || k.includes('masp') || k.includes('mahang')) {
                  foundId = cIdx;
                  hasExactId = true;
                } else if (foundId === -1 && (k.startsWith('ma') || k.endsWith('code')) && !k.includes('dinhmuc') && !k.includes('may')) {
                  foundId = cIdx;
                }

                if (k.includes('tenthanhpham') || k.includes('tenlinhkien') || k.includes('partname') || k.includes('tensanpham') || k.includes('tenlk') || k.includes('tentp') || k.includes('tenhang') || k.startsWith('ten')) {
                  foundName = cIdx;
                }

                if (k.includes('seconds') || k.includes('giay') || k.includes('thoigian') || k.includes('dinhmuc') || k.includes('hsqd') || k.includes('heso') || k.includes('quydoi') || k.includes('norm')) {
                  foundNorm = cIdx;
                }
              });

              if (foundId !== -1 && foundNorm !== -1) {
                const rScore = (hasExactId ? 50 : 10) + (foundName !== -1 ? 20 : 0);
                if (rScore > bestScore) {
                  bestScore = rScore;
                  headerRowIdx = r;
                  idCol = foundId;
                  nameCol = foundName;
                  normCol = foundNorm;
                }
              }
            }

            if (headerRowIdx === -1 && rawGrid.length >= 2) {
              const r1 = rawGrid[1] || [];
              if (r1.length >= 2) {
                headerRowIdx = 0;
                const isCol0Seq = typeof r1[0] === 'number' || (typeof r1[0] === 'string' && /^\d+$/.test(r1[0].trim()));
                if (isCol0Seq && r1.length >= 3) {
                  idCol = 1;
                  nameCol = r1.length >= 4 ? 2 : -1;
                  normCol = r1.length >= 4 ? 3 : 2;
                } else {
                  idCol = 0;
                  nameCol = r1.length >= 3 ? 1 : -1;
                  normCol = r1.length >= 3 ? 2 : 1;
                }
                bestScore = 5;
              }
            }

            if (headerRowIdx !== -1 && idCol !== -1 && normCol !== -1) {
              const sheetNorms: ProductivityNorm[] = [];
              let isIdNumericSequence = true;
              let sampleCount = 0;

              for (let r = headerRowIdx + 1; r < rawGrid.length; r++) {
                const row = rawGrid[r];
                if (!row) continue;
                let partIdRaw = String(row[idCol] || '').trim();
                const partNameRaw = nameCol !== -1 && row[nameCol] !== undefined ? String(row[nameCol]).trim() : '';
                const normVal = parseNum(row[normCol]);

                if (/^\d{1,4}$/.test(partIdRaw)) {
                  for (let c = 0; c < row.length; c++) {
                    if (c !== idCol && c !== normCol) {
                      const val = String(row[c] || '').trim();
                      if (val.length > 5 && (val.includes('-') || /[a-zA-Z]/.test(val))) {
                        partIdRaw = val;
                        break;
                      }
                    }
                  }
                }

                let finalPartId = partIdRaw;
                if (!finalPartId && partNameRaw) {
                  const found = parts.find(p => cleanKey(p.name) === cleanKey(partNameRaw));
                  if (found) finalPartId = found.id;
                }

                if (finalPartId && normVal > 0) {
                  sampleCount++;
                  if (!/^\d+$/.test(finalPartId)) {
                    isIdNumericSequence = false;
                  }
                  sheetNorms.push({
                    partId: finalPartId,
                    stageId: activeTab as StageId,
                    secondsPerUnit: normVal
                  });
                }
              }

              if (sheetNorms.length > 0) {
                let totalScore = sheetBaseScore + bestScore;
                if (!isIdNumericSequence && sampleCount > 0) totalScore += 200;
                else if (isIdNumericSequence && sampleCount > 0) totalScore -= 150;
                normCandidates.push({
                  sheetName,
                  score: totalScore,
                  items: sheetNorms
                });
              }
            }
          }

          normCandidates.sort((a, b) => b.score - a.score);
          const bestCandidate = normCandidates[0];

          if (!bestCandidate || bestCandidate.items.length === 0) {
            alert(`Không tìm thấy dữ liệu hợp lệ. Cần các cột: Mã linh kiện / Mã thành phẩm hoặc Tên linh kiện, ${(activeTab === 'DCLR' || activeTab === 'GLAZING') ? 'Hệ số quy đổi' : 'Định mức (Giây)'}`);
          } else {
            const imported = bestCandidate.items;
            const currentNorms = storageService.getNorms();
            const updated = [...currentNorms];
            imported.forEach(newNorm => {
              const index = updated.findIndex(n => n.partId === newNorm.partId && n.stageId === newNorm.stageId);
              if (index > -1) updated[index] = newNorm;
              else updated.push(newNorm);
            });
            storageService.saveNorms(updated);
            alert(`Đã nhập thành công ${imported.length} định mức cho ${STAGES.find(s => s.id === activeTab)?.name} (từ sheet "${bestCandidate.sheetName}")!`);
            onNormsChange();
          }
        }
      } catch (err: any) {
        alert('Lỗi khi đọc file: ' + (err?.message || err));
      } finally {
        setIsImporting(false);
        if (e.target) e.target.value = '';
      }
    };
    reader.readAsArrayBuffer(file);
  };

  return (
    <motion.div 
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, y: -20 }}
      className="space-y-8"
    >
      <div className="flex bg-white p-2 rounded-2xl border border-gray-200 shadow-sm gap-2 overflow-x-auto">
        <button
          onClick={() => setActiveTab('NESTING')}
          className={cn(
            "flex-1 py-4 px-6 rounded-xl font-bold uppercase tracking-widest transition-all whitespace-nowrap",
            activeTab === 'NESTING' ? "bg-orange-600 text-white shadow-lg" : "text-gray-400 hover:bg-gray-50"
          )}
        >
          Tổ hợp Laser
        </button>
        <button
          onClick={() => setActiveTab('LASER_HSQD')}
          className={cn(
            "flex-1 py-4 px-6 rounded-xl font-bold uppercase tracking-widest transition-all whitespace-nowrap",
            activeTab === 'LASER_HSQD' ? "bg-amber-600 text-white shadow-lg" : "text-gray-400 hover:bg-gray-50"
          )}
        >
          HSQĐ Laser
        </button>
        {STAGES.filter(s => s.id !== 'LASER').map(stage => {
          const tabLabel = stage.id === 'GLAZING' ? 'HSQĐ Dán Kính (Xưởng CK)' : (stage.id === 'DCLR' ? 'HSQĐ DCLR' : stage.name);
          return (
            <React.Fragment key={stage.id}>
              <button
                onClick={() => setActiveTab(stage.id)}
                className={cn(
                  "flex-1 py-4 px-6 rounded-xl font-bold uppercase tracking-widest transition-all whitespace-nowrap",
                  activeTab === stage.id ? "bg-blue-600 text-white shadow-lg" : "text-gray-400 hover:bg-gray-50"
                )}
              >
                {tabLabel}
              </button>
              {stage.id === 'PAINTING' && (
                <button
                  onClick={() => setActiveTab('PAINTING_HSQD')}
                  className={cn(
                    "flex-1 py-4 px-6 rounded-xl font-bold uppercase tracking-widest transition-all whitespace-nowrap",
                    activeTab === 'PAINTING_HSQD' ? "bg-yellow-500 text-white shadow-lg" : "text-gray-400 hover:bg-gray-50"
                  )}
                >
                  HSQĐ Sơn
                </button>
              )}
            </React.Fragment>
          );
        })}
        <button
          onClick={() => setActiveTab('GLAZING_STAGE')}
          className={cn(
            "flex-1 py-4 px-6 rounded-xl font-bold uppercase tracking-widest transition-all whitespace-nowrap",
            activeTab === 'GLAZING_STAGE' ? "bg-cyan-600 text-white shadow-lg" : "text-gray-400 hover:bg-gray-50"
          )}
        >
          HSQĐ Dán Kính 2
        </button>
        <button
          onClick={() => setActiveTab('BW_HSQD')}
          className={cn(
            "flex-1 py-4 px-6 rounded-xl font-bold uppercase tracking-widest transition-all whitespace-nowrap",
            activeTab === 'BW_HSQD' ? "bg-purple-600 text-white shadow-lg" : "text-gray-400 hover:bg-gray-50"
          )}
        >
          HSQĐ Chấn/Hàn
        </button>
      </div>

      <div className="bg-white rounded-3xl border border-gray-200 shadow-xl overflow-hidden">
        <div className="p-10 border-b border-gray-100 flex justify-between items-center bg-gray-50/30">
          <div>
            <h2 className="text-2xl font-bold tracking-tight">
              {activeTab === 'NESTING' ? 'Định mức Tổ hợp (Nesting) Laser' 
                : activeTab === 'LASER_HSQD' ? 'Hệ số quy đổi (HSQĐ) Cắt Laser'
                : activeTab === 'BW_HSQD' ? 'Hệ số quy đổi (HSQĐ) Chấn/Hàn'
                : activeTab === 'PAINTING_HSQD' ? 'Hệ số quy đổi (HSQĐ) Sơn (Biểu đồ công đoạn Sơn)'
                : activeTab === 'GLAZING_STAGE' ? 'Hệ số quy đổi (HSQĐ) Dán Kính 2 (Biểu đồ công đoạn riêng lẻ)'
                : activeTab === 'GLAZING' ? 'Hệ số quy đổi (HSQĐ) Dán Kính (Biểu đồ Năng suất Xưởng Cơ Khí)'
                : activeTab === 'DCLR'
                ? `Hệ số quy đổi (HSQĐ): ${STAGES.find(s => s.id === activeTab)?.name}`
                : `Định mức công đoạn: ${STAGES.find(s => s.id === activeTab)?.name}`}
            </h2>
            <p className="text-sm text-gray-500 mt-1">
              {activeTab === 'NESTING' 
                ? 'Khai báo các linh kiện được cắt chung trên một bàn máy (tấm tôn) để tối ưu thời gian' 
                : activeTab === 'LASER_HSQD'
                ? 'Quản lý hệ số quy đổi sản phẩm/linh kiện cho công đoạn Cắt Laser'
                : activeTab === 'BW_HSQD'
                ? 'Quản lý hệ số quy đổi sản phẩm/linh kiện cho công đoạn Chấn/Dập và Hàn'
                : activeTab === 'PAINTING_HSQD'
                ? 'Quản lý hệ số quy đổi sản phẩm/linh kiện tính riêng cho đường biểu đồ năng suất lao động của công đoạn Sơn'
                : activeTab === 'GLAZING_STAGE'
                ? 'Quản lý hệ số quy đổi tính riêng cho đường biểu đồ năng suất lao động của công đoạn Dán Kính'
                : activeTab === 'GLAZING'
                ? 'Quản lý hệ số quy đổi Dán Kính dùng để tính vào Năng suất Lao động toàn Xưởng Cơ Khí'
                : activeTab === 'DCLR'
                ? `Quản lý hệ số quy đổi (HSQĐ) cho công đoạn ${STAGES.find(s => s.id === activeTab)?.name}`
                : 'Quản lý thời gian sản xuất dự kiến trên mỗi đơn vị linh kiện'}
            </p>
          </div>
          <div className="flex flex-wrap items-center gap-3">
            {/* Nút Tải File Mẫu cho từng Tab */}
            <button
              onClick={() => downloadNormTemplate(activeTab)}
              className="flex items-center gap-2 px-5 py-3.5 bg-white border-2 border-emerald-600 text-emerald-700 hover:bg-emerald-50 rounded-xl text-xs font-bold uppercase tracking-wider transition-all shadow-sm cursor-pointer"
              title="Tải file Excel mẫu chuẩn cho tab hiện tại"
            >
              <FileDown size={18} className="text-emerald-600" />
              <span>Tải file mẫu Excel</span>
            </button>

            {/* Nút Xóa Tất Cả cho từng tab */}
            {((activeTab === 'NESTING' && nesting.length > 0) ||
              (activeTab === 'LASER_HSQD' && laserHSQD.length > 0) ||
              (activeTab === 'BW_HSQD' && bendingWeldingHSQD.length > 0) ||
              (activeTab === 'PAINTING_HSQD' && paintingHSQD.length > 0) ||
              (activeTab === 'GLAZING_STAGE' && glazingStageHSQD.length > 0) ||
              (activeTab !== 'NESTING' && activeTab !== 'LASER_HSQD' && activeTab !== 'BW_HSQD' && activeTab !== 'GLAZING_STAGE' && activeTab !== 'PAINTING_HSQD' && norms.filter(n => n.stageId === activeTab).length > 0)) && (
              <button 
                onClick={() => {
                  const tabName = activeTab === 'NESTING' ? 'định mức tổ hợp Laser'
                    : activeTab === 'LASER_HSQD' ? 'toàn bộ HSQĐ Laser'
                    : activeTab === 'BW_HSQD' ? 'toàn bộ HSQĐ Chấn/Hàn'
                    : activeTab === 'PAINTING_HSQD' ? 'toàn bộ HSQĐ Sơn'
                    : activeTab === 'GLAZING_STAGE' ? 'toàn bộ HSQĐ Dán Kính 2'
                    : activeTab === 'GLAZING' ? 'toàn bộ HSQĐ Dán Kính (Xưởng Cơ Khí)'
                    : activeTab === 'DCLR'
                    ? `toàn bộ HSQĐ của ${STAGES.find(s => s.id === activeTab)?.name}`
                    : `toàn bộ định mức của ${STAGES.find(s => s.id === activeTab)?.name}`;
                  if (confirm(`Bạn có chắc chắn muốn xóa ${tabName}?`)) {
                    if (activeTab === 'NESTING') {
                      storageService.saveLaserNesting([]);
                      onNormsChange();
                    } else if (activeTab === 'LASER_HSQD') {
                      storageService.saveLaserHSQD([]);
                      setLaserHSQD([]);
                      onNormsChange();
                    } else if (activeTab === 'PAINTING_HSQD') {
                      storageService.savePaintingHSQD([]);
                      setPaintingHSQD([]);
                      onNormsChange();
                    } else if (activeTab === 'GLAZING_STAGE') {
                      storageService.saveGlazingStageHSQD([]);
                      setGlazingStageHSQD([]);
                      onNormsChange();
                    } else if (activeTab === 'BW_HSQD') {
                      storageService.saveBendingWeldingHSQD([]);
                      setBendingWeldingHSQD([]);
                      onNormsChange();
                    } else {
                      const updated = norms.filter(n => n.stageId !== activeTab);
                      storageService.saveNorms(updated);
                      onNormsChange();
                    }
                  }
                }}
                className="px-5 py-3.5 border border-red-200 text-red-600 rounded-xl text-xs font-bold uppercase hover:bg-red-50 transition-all font-mono cursor-pointer"
              >
                Xóa tất cả
              </button>
            )}

            <label className={cn(
              "flex items-center gap-2.5 px-6 py-3.5 bg-[#141414] text-white rounded-xl text-xs font-bold uppercase tracking-wider cursor-pointer hover:bg-black transition-all shadow-md font-mono",
              isImporting && "opacity-50 cursor-not-allowed"
            )}>
              <FileUp size={18} />
              {isImporting ? 'Đang nhập...' : `Nhập ${activeTab === 'NESTING' ? 'file Tổ hợp' : activeTab === 'LASER_HSQD' ? 'HSQĐ Laser' : activeTab === 'BW_HSQD' ? 'HSQĐ Chấn/Hàn' : activeTab === 'PAINTING_HSQD' ? 'HSQĐ Sơn' : activeTab === 'GLAZING_STAGE' ? 'HSQĐ Dán Kính 2' : (activeTab === 'DCLR' || activeTab === 'GLAZING') ? `HSQĐ ${activeTab === 'GLAZING' ? 'Dán Kính (Xưởng CK)' : STAGES.find(s => s.id === activeTab)?.name}` : `định mức ${activeTab}`}`}
              <input type="file" accept=".xlsx, .xls" className="hidden" onChange={handleImport} disabled={isImporting} />
            </label>
          </div>
        </div>

        <div className="overflow-x-auto">
          {activeTab === 'NESTING' ? (
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-gray-50/50 border-b border-gray-100">
                  <th className="p-8 pl-12 text-sm font-mono uppercase opacity-80">Mã bàn (Nesting ID)</th>
                  <th className="p-8 text-sm font-mono uppercase opacity-80">Model áp dụng</th>
                  <th className="p-8 text-sm font-mono uppercase opacity-80">Linh kiện kết hợp</th>
                  <th className="p-8 text-sm font-mono uppercase opacity-80 text-center">Số lượng / Tấm</th>
                  <th className="p-8 text-sm font-mono uppercase opacity-80 text-center bg-orange-50/50">Thời gian 1 bàn (Giây)</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-50">
                {Array.from(new Set(nesting.map(n => n.nestingId))).map((nestId, idx) => {
                  const items = nesting.filter(n => n.nestingId === nestId);
                  const totalSeconds = items[0]?.secondsPerSheet || 0;
                  const maxModel = items[0]?.applicableModel;

                  return (
                    <tr key={idx} className="hover:bg-gray-50/30 transition-colors">
                      <td className="p-8 pl-12 font-mono text-lg font-bold text-orange-600">{nestId}</td>
                      <td className="p-8 font-mono text-sm opacity-80 whitespace-nowrap">
                        {maxModel ? (
                          <span className="bg-orange-100 text-orange-700 px-3 py-1 rounded-full font-bold">{maxModel}</span>
                        ) : (
                          <span className="opacity-40 italic">Tất cả Model</span>
                        )}
                      </td>
                      <td className="p-8">
                        <div className="flex flex-wrap gap-2">
                          {items.map((it, i) => (
                            <span key={i} className="px-3 py-1 bg-gray-100 rounded-full text-xs font-bold border border-gray-200">
                              {parts.find(p => p.id === it.partId)?.name || it.partId}
                            </span>
                          ))}
                        </div>
                      </td>
                      <td className="p-8 text-center text-gray-500 font-mono">
                        <div className="flex flex-col gap-1 items-center">
                          {items.map((it, i) => (
                            <span key={i} className="text-sm font-bold">
                              x{it.qtyPerSheet}
                            </span>
                          ))}
                        </div>
                      </td>
                      <td className="p-8 text-center bg-orange-50/20">
                        <span className="font-mono text-2xl font-black text-orange-600">{totalSeconds}</span>
                        <span className="ml-2 text-xs font-bold uppercase opacity-40 text-orange-900">giây</span>
                      </td>
                    </tr>
                  );
                })}
                {nesting.length === 0 && (
                  <tr>
                    <td colSpan={5} className="p-24 text-center text-gray-400 italic">
                      <div className="flex flex-col items-center">
                        <Layers size={64} className="opacity-10 mb-6" />
                        <p className="text-xl">Chưa có định mức tổ hợp linh kiện Laser.</p>
                        <p className="text-sm mt-2 font-mono uppercase">Vui lòng nhập file Excel đúng cấu trúc các cột: Mã bàn (Nesting ID), Linh kiện kết hợp, Số lượng / Tấm, Thời gian / LK (Giây).</p>
                      </div>
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          ) : activeTab === 'LASER_HSQD' ? (
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-gray-50/50 border-b border-gray-100">
                  <th className="p-8 pl-12 text-sm font-mono uppercase opacity-75">Mã linh kiện / Thành phẩm</th>
                  <th className="p-8 text-sm font-mono uppercase opacity-75">Tên linh kiện / Thành phẩm</th>
                  <th className="p-8 text-sm font-mono uppercase opacity-75 text-center">HSQĐ</th>
                  <th className="p-8 pr-12 text-sm font-mono uppercase opacity-75 text-right">Thao tác</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-50">
                {laserHSQD.map((item, idx) => {
                  const part = parts.find(p => p.id === item.partId || p.id.toUpperCase() === item.partId.toUpperCase());
                  return (
                    <tr key={idx} className="hover:bg-gray-50/30 transition-colors">
                      <td className="p-8 pl-12 font-mono text-lg font-bold">{item.partId}</td>
                      <td className="p-8 text-lg font-medium text-gray-600">{item.partName || part?.name || 'N/A'}</td>
                      <td className="p-8 text-center">
                        <span className="font-mono text-2xl font-black text-amber-600">{item.hsqd}</span>
                      </td>
                      <td className="p-8 pr-12 text-right">
                        <button 
                          onClick={() => {
                            const updated = laserHSQD.filter(n => n.partId !== item.partId);
                            storageService.saveLaserHSQD(updated);
                            setLaserHSQD(updated);
                            onNormsChange();
                          }}
                          className="p-3 text-red-400 hover:text-red-600 hover:bg-red-50 rounded-xl transition-all cursor-pointer"
                        >
                          <Trash2 size={24} />
                        </button>
                      </td>
                    </tr>
                  );
                })}
                {laserHSQD.length === 0 && (
                  <tr>
                    <td colSpan={4} className="p-24 text-center text-gray-400 italic">
                      <p className="text-xl">Chưa có hệ số quy đổi cho Laser.</p>
                      <p className="text-sm mt-2 font-mono uppercase">Vui lòng nhập file Excel chứa cột Mã linh kiện / Mã thành phẩm, Tên linh kiện, HSQĐ.</p>
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          ) : activeTab === 'BW_HSQD' ? (
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-gray-50/50 border-b border-gray-100">
                  <th className="p-8 pl-12 text-sm font-mono uppercase opacity-75">Mã thành phẩm</th>
                  <th className="p-8 text-sm font-mono uppercase opacity-75">Tên thành phẩm</th>
                  <th className="p-8 text-sm font-mono uppercase opacity-75 text-center">HSQĐ</th>
                  <th className="p-8 pr-12 text-sm font-mono uppercase opacity-75 text-right">Thao tác</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-50">
                {bendingWeldingHSQD.map((item, idx) => {
                  const part = parts.find(p => p.id === item.partId || p.id.toUpperCase() === item.partId.toUpperCase());
                  return (
                    <tr key={idx} className="hover:bg-gray-50/30 transition-colors">
                      <td className="p-8 pl-12 font-mono text-lg font-bold">{item.partId}</td>
                      <td className="p-8 text-lg font-medium text-gray-600">{item.partName || part?.name || 'N/A'}</td>
                      <td className="p-8 text-center">
                        <span className="font-mono text-2xl font-black text-purple-600">{item.hsqd}</span>
                      </td>
                      <td className="p-8 pr-12 text-right">
                        <button 
                          onClick={() => {
                            const updated = bendingWeldingHSQD.filter(n => n.partId !== item.partId);
                            storageService.saveBendingWeldingHSQD(updated);
                            setBendingWeldingHSQD(updated);
                            onNormsChange();
                          }}
                          className="p-3 text-red-400 hover:text-red-600 hover:bg-red-50 rounded-xl transition-all"
                        >
                          <Trash2 size={24} />
                        </button>
                      </td>
                    </tr>
                  );
                })}
                {bendingWeldingHSQD.length === 0 && (
                  <tr>
                    <td colSpan={4} className="p-24 text-center text-gray-400 italic">
                      <p className="text-xl">Chưa có hệ số quy đổi cho Chấn/Hàn.</p>
                      <p className="text-sm mt-2 font-mono uppercase">Vui lòng nhập file Excel chứa cột Mã thành phẩm, Tên thành phẩm, HSQĐ.</p>
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          ) : activeTab === 'GLAZING_STAGE' ? (
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-gray-50/50 border-b border-gray-100">
                  <th className="p-8 pl-12 text-sm font-mono uppercase opacity-75">Mã linh kiện / Thành phẩm</th>
                  <th className="p-8 text-sm font-mono uppercase opacity-75">Tên linh kiện / Thành phẩm</th>
                  <th className="p-8 text-sm font-mono uppercase opacity-75 text-center">HSQĐ Dán Kính 2</th>
                  <th className="p-8 pr-12 text-sm font-mono uppercase opacity-75 text-right">Thao tác</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-50">
                {glazingStageHSQD.map((item, idx) => {
                  const part = parts.find(p => p.id === item.partId || p.id.toUpperCase() === item.partId.toUpperCase());
                  return (
                    <tr key={idx} className="hover:bg-gray-50/30 transition-colors">
                      <td className="p-8 pl-12 font-mono text-lg font-bold">{item.partId}</td>
                      <td className="p-8 text-lg font-medium text-gray-600">{item.partName || part?.name || 'N/A'}</td>
                      <td className="p-8 text-center">
                        <span className="font-mono text-2xl font-black text-cyan-600">{item.hsqd}</span>
                      </td>
                      <td className="p-8 pr-12 text-right">
                        <button 
                          onClick={() => {
                            const updated = glazingStageHSQD.filter(n => n.partId !== item.partId);
                            storageService.saveGlazingStageHSQD(updated);
                            setGlazingStageHSQD(updated);
                            onNormsChange();
                          }}
                          className="p-3 text-red-400 hover:text-red-600 hover:bg-red-50 rounded-xl transition-all cursor-pointer"
                        >
                          <Trash2 size={24} />
                        </button>
                      </td>
                    </tr>
                  );
                })}
                {glazingStageHSQD.length === 0 && (
                  <tr>
                    <td colSpan={4} className="p-24 text-center text-gray-400 italic">
                      <p className="text-xl">Chưa có hệ số quy đổi cho Dán Kính 2 (Biểu đồ công đoạn riêng lẻ).</p>
                      <p className="text-sm mt-2 font-mono uppercase">Vui lòng nhập file Excel chứa cột Mã linh kiện / Mã thành phẩm, Tên linh kiện, Hệ số quy đổi (HSQĐ).</p>
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          ) : activeTab === 'PAINTING_HSQD' ? (
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-gray-50/50 border-b border-gray-100">
                  <th className="p-8 pl-12 text-sm font-mono uppercase opacity-75">Mã linh kiện / Thành phẩm</th>
                  <th className="p-8 text-sm font-mono uppercase opacity-75">Tên linh kiện / Thành phẩm</th>
                  <th className="p-8 text-sm font-mono uppercase opacity-75 text-center">HSQĐ Sơn</th>
                  <th className="p-8 pr-12 text-sm font-mono uppercase opacity-75 text-right">Thao tác</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-50">
                {paintingHSQD.map((item, idx) => {
                  const part = parts.find(p => p.id === item.partId || p.id.toUpperCase() === item.partId.toUpperCase());
                  return (
                    <tr key={idx} className="hover:bg-gray-50/30 transition-colors">
                      <td className="p-8 pl-12 font-mono text-lg font-bold">{item.partId}</td>
                      <td className="p-8 text-lg font-medium text-gray-600">{item.partName || part?.name || 'N/A'}</td>
                      <td className="p-8 text-center">
                        <span className="font-mono text-2xl font-black text-amber-500">{item.hsqd}</span>
                      </td>
                      <td className="p-8 pr-12 text-right">
                        <button 
                          onClick={() => {
                            const updated = paintingHSQD.filter(n => n.partId !== item.partId);
                            storageService.savePaintingHSQD(updated);
                            setPaintingHSQD(updated);
                            onNormsChange();
                          }}
                          className="p-3 text-red-400 hover:text-red-600 hover:bg-red-50 rounded-xl transition-all cursor-pointer"
                        >
                          <Trash2 size={24} />
                        </button>
                      </td>
                    </tr>
                  );
                })}
                {paintingHSQD.length === 0 && (
                  <tr>
                    <td colSpan={4} className="p-24 text-center text-gray-400 italic">
                      <p className="text-xl">Chưa có hệ số quy đổi cho công đoạn Sơn.</p>
                      <p className="text-sm mt-2 font-mono uppercase">Vui lòng nhập file Excel chứa cột Mã linh kiện, Tên linh kiện, Hệ số quy đổi (HSQĐ).</p>
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          ) : (
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-gray-50/50 border-b border-gray-100">
                  <th className="p-8 pl-12 text-sm font-mono uppercase opacity-75">Mã linh kiện</th>
                  <th className="p-8 text-sm font-mono uppercase opacity-75">Tên linh kiện</th>
                    <th className="p-8 pb-4 text-sm font-mono uppercase opacity-75 text-center">
                      {(activeTab === 'DCLR' || activeTab === 'GLAZING') ? 'Hệ số quy đổi' : 'Thời gian (Giây/Đơn vị)'}
                    </th>
                  <th className="p-8 pr-12 text-sm font-mono uppercase opacity-75 text-right">Thao tác</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-50">
                {norms.filter(n => n.stageId === activeTab).map((norm, idx) => {
                  const part = parts.find(p => 
                    p.id === norm.partId || 
                    p.id.toUpperCase() === norm.partId.toUpperCase() ||
                    p.name.trim().toLowerCase() === norm.partId.trim().toLowerCase()
                  );
                  return (
                    <tr key={idx} className="hover:bg-gray-50/30 transition-colors">
                      <td className="p-8 pl-12 font-mono text-lg font-bold">{norm.partId}</td>
                      <td className="p-8 text-lg font-medium text-gray-700">
                        {part?.name ? (
                          <span className="font-semibold text-gray-800">{part.name}</span>
                        ) : (
                          <span className="text-gray-400 italic font-mono text-sm">Chưa có tên trong danh mục</span>
                        )}
                      </td>
                      <td className="p-8 text-center">
                        <span className={cn("font-mono text-2xl font-black", (activeTab === 'DCLR' || activeTab === 'GLAZING') ? "text-emerald-600" : "text-blue-600")}>
                          {norm.secondsPerUnit}
                        </span>
                        {activeTab !== 'DCLR' && activeTab !== 'GLAZING' && (
                          <span className="ml-2 text-xs font-bold uppercase opacity-40">giây</span>
                        )}
                      </td>
                      <td className="p-8 pr-12 text-right">
                        <button 
                          onClick={() => {
                            const updated = norms.filter(n => !(n.partId === norm.partId && n.stageId === norm.stageId));
                            storageService.saveNorms(updated);
                            onNormsChange();
                          }}
                          className="p-3 text-red-400 hover:text-red-600 hover:bg-red-50 rounded-xl transition-all cursor-pointer"
                        >
                          <Trash2 size={24} />
                        </button>
                      </td>
                    </tr>
                  );
                })}
                {norms.filter(n => n.stageId === activeTab).length === 0 && (
                  <tr>
                    <td colSpan={4} className="p-24 text-center text-gray-400 italic">
                      <div className="flex flex-col items-center">
                        <ClipboardList size={64} className="opacity-10 mb-6" />
                        <p className="text-xl">Chưa có định mức cho công đoạn này.</p>
                        <p className="text-sm mt-2">Vui lòng nhập file Excel định mức linh kiện.</p>
                      </div>
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          )}
        </div>
      </div>
    </motion.div>
  );
}

interface HistoryProps {
  transactions: Transaction[];
  parts: Part[];
  key?: string;
}

function HistoryView({ transactions, parts }: HistoryProps) {
  // Only show transactions from the last 3 days
  const cutoff = Date.now() - 3 * 24 * 60 * 60 * 1000;
  const filteredTransactions = transactions.filter(t => t.timestamp >= cutoff);

  return (
    <motion.div 
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, y: -20 }}
      className="bg-white rounded-2xl border border-gray-200 shadow-sm overflow-hidden"
    >
      <div className="p-10 border-b border-gray-100 flex justify-between items-center">
        <h2 className="font-bold text-2xl tracking-tight">Nhật ký giao dịch hệ thống <span className="text-sm font-normal text-gray-500 bg-gray-100 px-3 py-1 rounded-full ml-4">3 ngày gần nhất</span></h2>
        <div className="flex gap-4">
          <button className="px-6 py-3 text-sm font-bold uppercase border border-gray-200 rounded-lg hover:bg-gray-50">Xuất Excel</button>
          <button className="px-6 py-3 text-sm font-bold uppercase border border-gray-200 rounded-lg hover:bg-gray-50">Lọc dữ liệu</button>
        </div>
      </div>
      
      <div className="overflow-x-auto">
        <table className="w-full text-left border-collapse">
          <thead>
            <tr className="bg-gray-50 border-b border-gray-100">
              <th className="p-6 pl-10 text-sm font-mono uppercase opacity-80">Thời gian</th>
              <th className="p-6 text-sm font-mono uppercase opacity-80">Loại GD</th>
              <th className="p-6 text-sm font-mono uppercase opacity-80">Mã Linh kiện</th>
              <th className="p-6 text-sm font-mono uppercase opacity-80">Số lượng</th>
              <th className="p-6 text-sm font-mono uppercase opacity-80">Công đoạn</th>
              <th className="p-6 text-sm font-mono uppercase opacity-80">Nguồn</th>
              <th className="p-6 pr-10 text-sm font-mono uppercase opacity-80">ID Giao dịch</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-50">
            {filteredTransactions.length === 0 ? (
              <tr>
                <td colSpan={7} className="p-24 text-center text-gray-400">
                  <div className="flex flex-col items-center">
                    <History size={64} strokeWidth={1} className="opacity-20 mb-6" />
                    <p className="font-mono text-base">Hệ thống chưa ghi nhận giao dịch nào</p>
                  </div>
                </td>
              </tr>
            ) : (
              filteredTransactions.map((tx) => (
                <tr key={tx.id} className="hover:bg-gray-50/50 transition-colors">
                  <td className="p-6 pl-10 font-mono text-base">{format(tx.timestamp, 'dd/MM/yyyy HH:mm:ss')}</td>
                  <td className="p-6">
                    <span className={cn(
                      "text-sm font-bold uppercase px-3 py-1.5 rounded-full",
                      tx.type === 'STAGE_OUT' ? "bg-orange-100 text-orange-700" : "bg-green-100 text-green-700"
                    )}>
                      {tx.type === 'STAGE_OUT' ? 'Hoàn thành' : 'Nhập kho'}
                    </span>
                  </td>
                  <td className="p-6 font-bold text-lg">
                    <div className="flex items-center gap-2">
                       {getProcessValue(tx.partId, parts.find(p => p.id === tx.partId), tx.stageId, tx.type === 'STAGE_OUT' ? 'OUT' : 'IN')}
                       {tx.originalPartId && (
                         <span className="text-[10px] bg-blue-50 text-blue-600 px-2 py-0.5 rounded-full uppercase">Từ {tx.originalPartId}</span>
                       )}
                    </div>
                    <span className="block text-sm font-normal opacity-80">
                      {getProcessValue(parts.find(p => p.id === tx.partId)?.name, parts.find(p => p.id === tx.partId), tx.stageId, tx.type === 'STAGE_OUT' ? 'OUT' : 'IN')}
                    </span>
                  </td>
                  <td className="p-6 font-mono font-bold text-2xl">
                    {tx.type === 'STAGE_OUT' ? '+' : ''}{tx.quantity}
                  </td>
                  <td className="p-6 font-bold text-lg">{STAGES.find(s => s.id === tx.stageId)?.name}</td>
                  <td className="p-6 text-base opacity-80">{tx.sourceStageId ? STAGES.find(s => s.id === tx.sourceStageId)?.name : '-'}</td>
                  <td className="p-6 pr-10 font-mono text-sm opacity-70">{tx.id.split('-')[0].toUpperCase()}</td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </motion.div>
  );
}

function WorkingHoursView() {
  const [configs, setConfigs] = useState<ShiftConfig[]>([]);
  
  useEffect(() => {
    setConfigs(storageService.getShiftConfigs());
  }, []);

  const handleSave = () => {
    storageService.saveShiftConfigs(configs);
    alert('Đã lưu cài đặt ca làm việc & nghỉ ngơi! \nHệ thống sẽ tự động áp dụng nguồn lực ngoại lệ khi tính toán kế hoạch sản xuất cho PO mới.');
  };

  const handleReset = () => {
    if (confirm('Bạn có chắc muốn khôi phục về cài đặt mặc định từ hệ thống?')) {
      storageService.resetShiftConfigs();
      setConfigs(storageService.getShiftConfigs());
    }
  };

  const updateShift = (stageId: StageId, shiftIdx: number, field: 'start' | 'end', val: string) => {
    setConfigs(prev => prev.map(c => {
      if (c.stageId === stageId) {
        const newShifts = [...c.shifts];
        newShifts[shiftIdx] = { ...newShifts[shiftIdx], [field]: val };
        return { ...c, shifts: newShifts };
      }
      return c;
    }));
  };

  const updateBreak = (stageId: StageId, breakIdx: number, field: 'start' | 'end', val: string) => {
    setConfigs(prev => prev.map(c => {
      if (c.stageId === stageId) {
        const newBreaks = [...c.breaks];
        newBreaks[breakIdx] = { ...newBreaks[breakIdx], [field]: val };
        return { ...c, breaks: newBreaks };
      }
      return c;
    }));
  };

  const addShift = (stageId: StageId) => {
    setConfigs(prev => prev.map(c => {
      if (c.stageId === stageId) {
        return { ...c, shifts: [...c.shifts, { start: '08:00', end: '17:00' }] };
      }
      return c;
    }));
  };

  const addBreak = (stageId: StageId) => {
    setConfigs(prev => prev.map(c => {
      if (c.stageId === stageId) {
        return { ...c, breaks: [...c.breaks, { start: '12:00', end: '13:00' }] };
      }
      return c;
    }));
  };

  const removeShift = (stageId: StageId, idx: number) => {
    setConfigs(prev => prev.map(c => {
      if (c.stageId === stageId) {
        return { ...c, shifts: c.shifts.filter((_, i) => i !== idx) };
      }
      return c;
    }));
  };

  const removeBreak = (stageId: StageId, idx: number) => {
    setConfigs(prev => prev.map(c => {
      if (c.stageId === stageId) {
        return { ...c, breaks: c.breaks.filter((_, i) => i !== idx) };
      }
      return c;
    }));
  };

  const updateWorkerCount = (stageId: StageId, val: number) => {
    setConfigs(prev => prev.map(c => {
      if (c.stageId === stageId) {
        return { ...c, workerCount: Math.max(1, val) };
      }
      return c;
    }));
  };

  const addOverride = (stageId: StageId) => {
    setConfigs(prev => prev.map(c => {
      if (c.stageId === stageId) {
        return { ...c, workerOverrides: [...(c.workerOverrides || []), { modelId: 'Gói khung tủ', workerCount: 2 }] };
      }
      return c;
    }));
  };

  const updateOverride = (stageId: StageId, idx: number, key: 'modelId'|'workerCount', val: string|number) => {
    setConfigs(prev => prev.map(c => {
      if (c.stageId === stageId) {
        const newOverrides = [...(c.workerOverrides || [])];
        newOverrides[idx] = { ...newOverrides[idx], [key]: val };
        return { ...c, workerOverrides: newOverrides };
      }
      return c;
    }));
  };

  const removeOverride = (stageId: StageId, idx: number) => {
    setConfigs(prev => prev.map(c => {
      if (c.stageId === stageId) {
        return { ...c, workerOverrides: (c.workerOverrides || []).filter((_, i) => i !== idx) };
      }
      return c;
    }));
  };

  const handleImportOverrides = (e: React.ChangeEvent<HTMLInputElement>, stageId: StageId) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws) as any[];

      const imported = data.map(row => ({
        modelId: String(row['Tên linh kiện'] || row['Model'] || row['Mã linh kiện'] || row['ModelId'] || '').trim(),
        workerCount: parseInt(row['Định mức ngoại lệ'] || row['Số lượng nhân sự'] || row['WorkerCount'] || '1')
      })).filter(o => o.modelId && o.workerCount > 0);

      if (imported.length > 0) {
        setConfigs(prev => prev.map(c => {
          if (c.stageId === stageId) {
            // Merge or replace? Replacing is cleaner for bulk updates
            return { ...c, workerOverrides: imported };
          }
          return c;
        }));
        alert(`Đã nhập thành công ${imported.length} ngoại lệ định mức cho công đoạn ${stageId}!`);
      } else {
        alert('Không tìm thấy dữ hiệu hợp lệ. Cần các cột: Tên linh kiện, Định mức ngoại lệ');
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = ''; // Reset input
  };

  return (
    <motion.div 
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, y: -20 }}
      className="space-y-8"
    >
      <div className="flex justify-between items-center mb-6">
        <div>
          <h2 className="text-2xl font-bold tracking-tight">Cấu hình thời gian làm việc</h2>
          <p className="text-sm text-gray-500 mt-1">Cài đặt số ca và các khoảng nghỉ ngơi để hệ thống tính toán kế hoạch sản xuất PO chính xác nhất.</p>
        </div>
        <div className="flex gap-4">
          <button 
            onClick={handleReset}
            className="flex items-center gap-2 px-6 py-4 bg-gray-100 text-gray-600 rounded-xl font-bold uppercase tracking-widest hover:bg-gray-200 transition-all"
          >
            <RotateCcw size={20} />
            Mặc định
          </button>
          <button 
            onClick={handleSave}
            className="flex items-center gap-2 px-8 py-4 bg-blue-600 text-white rounded-xl font-bold uppercase tracking-widest hover:bg-blue-700 transition-all shadow-lg"
          >
            <Save size={20} />
            Lưu cài đặt
          </button>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        {STAGES.map(stage => {
          const config = configs.find(c => c.stageId === stage.id);
          if (!config) return null;

          return (
            <div key={stage.id} className={cn(
              "bg-white rounded-2xl border border-gray-200 shadow-sm overflow-hidden border-t-4",
              stage.id === 'LASER' ? "border-t-orange-500" : 
              stage.id === 'GLAZING' ? "border-t-indigo-500" :
              stage.id === 'PAINTING' ? "border-t-pink-500" : "border-t-blue-500"
            )}>
              <div className="p-6 border-b border-gray-100 flex justify-between items-center bg-gray-50/50">
                <h3 className="text-xl font-bold flex items-center gap-3 italic">
                  {stage.id === 'LASER' && <Flame className="text-orange-500" />}
                  {stage.id === 'GLAZING' && <Square className="text-indigo-500" />}
                  {stage.id === 'PAINTING' && <Monitor className="text-pink-500" />}
                  {(stage.id !== 'LASER' && stage.id !== 'GLAZING' && stage.id !== 'PAINTING') && <Monitor className="text-blue-500" />}
                  {stage.name}
                </h3>
                <span className="text-xs font-mono font-bold bg-gray-200 px-2 py-1 rounded">
                  {config.shifts.length} CA LÀM VIỆC
                </span>
              </div>
              
              <div className="p-8 space-y-8">
                {/* Worker Count */}
                <div className="flex flex-col gap-3 bg-blue-50 p-4 rounded-xl border border-blue-100">
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-3">
                      <div className="p-2 bg-white rounded-lg shadow-sm">
                        <Users size={20} className="text-blue-600" />
                      </div>
                      <div>
                        <h4 className="text-sm font-bold text-gray-700">Số lượng nhân sự / nguồn lực</h4>
                        <p className="text-xs text-gray-500">Mặc định cho công đoạn này</p>
                      </div>
                    </div>
                    <div className="flex items-center gap-6">
                      <label className="flex items-center gap-2 cursor-pointer bg-white px-3 py-2 rounded-lg border border-gray-200 shadow-sm hover:border-blue-300 transition-all">
                        <input
                          type="checkbox"
                          checked={config.workOnSunday || false}
                          onChange={(e) => {
                            setConfigs(prev => prev.map(c => c.stageId === stage.id ? { ...c, workOnSunday: e.target.checked } : c));
                          }}
                          className="w-4 h-4 text-blue-600 rounded border-gray-300 focus:ring-blue-500"
                        />
                        <span className="text-xs font-bold text-gray-700 select-none">LÀM CHỦ NHẬT</span>
                      </label>
                      <div className="flex items-center gap-3">
                        <input 
                          type="number" 
                          min="1"
                          value={config.workerCount || 1}
                          onChange={(e) => updateWorkerCount(stage.id, parseInt(e.target.value))}
                          className="w-20 bg-white border border-gray-200 rounded-lg p-2 font-mono text-center font-bold text-blue-600 focus:border-blue-500 outline-none"
                        />
                        <span className="text-xs font-bold text-gray-400">NGƯỜI</span>
                      </div>
                    </div>
                  </div>

                  {/* Overrides Table */}
                  <div className="mt-2 border-t border-blue-200 pt-3">
                    <div className="flex justify-between items-center mb-2">
                      <h5 className="text-xs font-bold text-blue-700 uppercase">Ngoại lệ theo Model</h5>
                      <div className="flex gap-2">
                        <label className="text-xs font-bold bg-white text-blue-600 px-2 py-1 rounded shadow-sm hover:bg-blue-100 transition-colors cursor-pointer border border-blue-200">
                          <Upload size={12} className="inline mr-1" />
                          Nhập Excel
                          <input 
                            type="file" 
                            accept=".xlsx, .xls"
                            className="hidden"
                            onChange={(e) => handleImportOverrides(e, stage.id)}
                          />
                        </label>
                        <button 
                          onClick={() => addOverride(stage.id)}
                          className="text-xs font-bold bg-white text-blue-600 px-2 py-1 rounded shadow-sm hover:bg-blue-100 transition-colors border border-blue-200"
                        >
                          + Thêm thủ công
                        </button>
                      </div>
                    </div>
                    {config.workerOverrides && config.workerOverrides.length > 0 ? (
                      <div className="space-y-2">
                        {config.workerOverrides.map((ov, oIdx) => (
                          <div key={oIdx} className="flex items-center gap-2 bg-white p-2 rounded-lg shadow-sm border border-blue-100">
                            <input 
                              type="text"
                              value={ov.modelId}
                              onChange={e => updateOverride(stage.id, oIdx, 'modelId', e.target.value)}
                              placeholder="Tên model (VD: Gói khung tủ)"
                              className="flex-1 bg-transparent border-none text-sm outline-none px-2 font-bold text-gray-700"
                            />
                            <div className="flex items-center gap-2">
                              <input 
                                type="number"
                                min="1"
                                value={ov.workerCount}
                                onChange={e => updateOverride(stage.id, oIdx, 'workerCount', parseInt(e.target.value))}
                                className="w-16 bg-gray-50 border border-gray-200 rounded p-1 font-mono text-center text-sm focus:border-blue-500 outline-none"
                              />
                              <button 
                                onClick={() => removeOverride(stage.id, oIdx)}
                                className="p-1 text-gray-400 hover:text-red-500 transition-colors"
                              >
                                <Trash2 size={16} />
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <p className="text-xs text-blue-500/70 italic">Chưa có cài đặt ngoại lệ.</p>
                    )}
                  </div>
                </div>

                {/* Shifts */}
                <div>
                  <div className="flex justify-between items-center mb-4">
                    <h4 className="text-sm font-bold uppercase tracking-widest text-gray-400 flex items-center gap-2">
                       Ca làm việc
                    </h4>
                    <button 
                      onClick={() => addShift(stage.id)}
                      className="p-2 bg-blue-50 text-blue-600 rounded-lg hover:bg-blue-100 transition-colors"
                    >
                      <Plus size={16} />
                    </button>
                  </div>
                  <div className="space-y-3">
                    {config.shifts.map((shift, sIdx) => (
                      <div key={sIdx} className="flex items-center gap-4 bg-gray-50 p-4 rounded-xl border border-gray-100">
                        <span className="font-mono text-xs opacity-40">CA {sIdx + 1}</span>
                        <input 
                          type="time" 
                          value={shift.start} 
                          onChange={(e) => updateShift(stage.id, sIdx, 'start', e.target.value)}
                          className="bg-white border border-gray-200 rounded-lg p-2 font-mono text-sm focus:border-blue-500 outline-none"
                        />
                        <ArrowRight size={16} className="text-gray-300" />
                        <input 
                          type="time" 
                          value={shift.end} 
                          onChange={(e) => updateShift(stage.id, sIdx, 'end', e.target.value)}
                          className="bg-white border border-gray-200 rounded-lg p-2 font-mono text-sm focus:border-blue-500 outline-none"
                        />
                        <button 
                          onClick={() => removeShift(stage.id, sIdx)}
                          className="ml-auto text-gray-300 hover:text-red-500 transition-colors"
                        >
                          <Trash2 size={18} />
                        </button>
                      </div>
                    ))}
                    {config.shifts.length === 0 && <p className="text-xs text-red-500 italic">Chưa cài đặt ca làm việc!</p>}
                  </div>
                </div>

                {/* Breaks */}
                <div>
                  <div className="flex justify-between items-center mb-4">
                    <h4 className="text-sm font-bold uppercase tracking-widest text-gray-400 flex items-center gap-2">
                      Giờ nghỉ ngơi
                    </h4>
                    <button 
                      onClick={() => addBreak(stage.id)}
                      className="p-2 bg-orange-50 text-orange-600 rounded-lg hover:bg-orange-100 transition-colors"
                    >
                      <Plus size={16} />
                    </button>
                  </div>
                  <div className="space-y-3">
                    {config.breaks.map((brk, bIdx) => (
                      <div key={bIdx} className="flex items-center gap-4 bg-orange-50/30 p-4 rounded-xl border border-orange-100">
                        <span className="font-mono text-xs opacity-40">Nghỉ {bIdx + 1}</span>
                        <input 
                          type="time" 
                          value={brk.start} 
                          onChange={(e) => updateBreak(stage.id, bIdx, 'start', e.target.value)}
                          className="bg-white border border-gray-200 rounded-lg p-2 font-mono text-sm focus:border-blue-500 outline-none"
                        />
                        <ArrowRight size={16} className="text-orange-200" />
                        <input 
                          type="time" 
                          value={brk.end} 
                          onChange={(e) => updateBreak(stage.id, bIdx, 'end', e.target.value)}
                          className="bg-white border border-gray-200 rounded-lg p-2 font-mono text-sm focus:border-blue-500 outline-none"
                        />
                        <button 
                          onClick={() => removeBreak(stage.id, bIdx)}
                          className="ml-auto text-gray-300 hover:text-red-500 transition-colors"
                        >
                          <Trash2 size={18} />
                        </button>
                      </div>
                    ))}
                    {config.breaks.length === 0 && <p className="text-xs text-gray-400 italic">Không có giờ nghỉ.</p>}
                  </div>
                </div>
              </div>
            </div>
          );
        })}
      </div>
    </motion.div>
  );
}

function DefectModal({ data, onClose, onDefectRecorded, inventory, setLastTransaction }: { data: any, onClose: () => void, onDefectRecorded: () => void, inventory: InventoryItem[], setLastTransaction: (tx: Transaction) => void }) {
  const [qty, setQty] = useState(0);
  const [reasonId, setReasonId] = useState(DEFECT_REASONS[0].id);
  const [note, setNote] = useState("");
  const [selectedPartId, setSelectedPartId] = useState(data.partId || "");
  const [errorMsg, setErrorMsg] = useState("");
  const [recordedTx, setRecordedTx] = useState<Transaction | null>(null);

  const allParts = storageService.getParts();
  
  // Filter parts if no specific partId is provided but stageId is
  const availableParts = useMemo(() => {
    if (data.partId) return allParts;
    if (!data.stageId) return allParts;
    
    const partsWithStock = inventory
      .filter(i => i.stageId === data.stageId && i.location === 'IN' && i.quantity > 0)
      .map(i => i.partId);
      
    if (partsWithStock.length === 0) return allParts; // Fallback to all if somehow empty
    
    return allParts.filter(p => partsWithStock.includes(p.id));
  }, [allParts, data.partId, data.stageId, inventory]);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const finalPartId = selectedPartId || data.partId;
    if (!finalPartId) {
      setErrorMsg('Vui lòng chọn mã linh kiện');
      return;
    }
    if (qty <= 0) {
      setErrorMsg('Vui lòng nhập số lượng lỗi hợp lệ');
      return;
    }
    try {
      const reasonName = DEFECT_REASONS.find(r => r.id === reasonId)?.name || reasonId;
      const targetLocation = data.location || 'IN';
      const tx = storageService.recordDefect(finalPartId, data.stageId, targetLocation, qty, reasonName, note, data.poId);
      
      if (data.stageId === 'GLAZING') {
        setRecordedTx(tx);
        onDefectRecorded();
      } else {
        onDefectRecorded();
        onClose();
      }
    } catch (err) {
      setErrorMsg(err instanceof Error ? err.message : 'Lỗi');
    }
  };

  const handlePrint = () => {
    if (recordedTx) {
      setLastTransaction(recordedTx);
      setTimeout(() => {
        window.print();
      }, 200);
    }
  };

  if (recordedTx) {
    return (
      <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[200] flex items-center justify-center p-4 no-print">
        <motion.div initial={{ scale: 0.9, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} className="bg-white rounded-3xl p-8 max-w-md w-full shadow-2xl relative text-gray-900">
          <button onClick={onClose} className="absolute right-6 top-6 text-gray-400 hover:text-black"><X size={24} /></button>
          <div className="text-center">
            <div className="w-20 h-20 bg-green-100 text-green-600 rounded-full flex items-center justify-center mx-auto mb-6">
              <CheckCircle2 size={40} />
            </div>
            <h2 className="text-2xl font-bold mb-2">Báo lỗi thành công!</h2>
            <p className="text-gray-500 mb-8">Vui lòng in mã QR dưới đây để dán lên hàng lỗi (GLAZING).</p>
            
            <div className="bg-gray-50 p-6 rounded-2xl mb-8 border border-gray-100 inline-block mx-auto w-full">
              <div className="flex justify-center mb-4">
                <QRCodeSVG value={recordedTx.qrData || ''} size={200} level="H" />
              </div>
              <div className="font-mono font-bold text-sm uppercase">{recordedTx.id}</div>
              <div className="text-[10px] text-gray-500 uppercase mt-1 tracking-widest font-bold">NHÃN LỖI (DEFECT)</div>
              {recordedTx.defectReason && (
                <div className="mt-4 p-3 bg-red-50 text-red-700 rounded-xl border border-red-100">
                  <span className="text-[10px] uppercase tracking-widest font-bold block mb-1">Lỗi:</span>
                  <span className="font-black text-sm uppercase">{recordedTx.defectReason}</span>
                  {recordedTx.defectCategory && (
                    <div className="text-xs mt-1 italic break-words">Ghi chú: {recordedTx.defectCategory}</div>
                  )}
                </div>
              )}
            </div>

            <div className="flex gap-4 line-clamp-1">
              <button onClick={handlePrint} className="flex-1 bg-blue-600 text-white py-4 rounded-xl font-bold hover:bg-blue-700 shadow-lg shadow-blue-200 uppercase tracking-tight">
                IN MÃ HỦY
              </button>
              <button onClick={onClose} className="flex-1 bg-gray-100 text-gray-600 py-4 rounded-xl font-bold hover:bg-gray-200 uppercase tracking-tight">
                ĐÓNG
              </button>
            </div>
          </div>
        </motion.div>
      </div>
    );
  }

  return (
    <div className="fixed inset-0 z-[200] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm animate-in fade-in duration-200 text-gray-900">
      <motion.div 
        initial={{ scale: 0.9, opacity: 0 }}
        animate={{ scale: 1, opacity: 1 }}
        className="bg-white rounded-3xl shadow-2xl w-full max-w-lg overflow-hidden border-4 border-red-500/10"
      >
        <div className="bg-red-600 p-8 text-white flex justify-between items-center">
          <div className="flex items-center gap-3">
            <AlertCircle size={32} />
            <div>
              <h3 className="text-2xl font-bold uppercase tracking-tight leading-none mb-1">Khai báo hàng lỗi</h3>
              <p className="text-xs font-mono opacity-80 uppercase tracking-widest">Defect / NG Record</p>
            </div>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-white/20 rounded-full transition-colors"><X size={28} /></button>
        </div>

        <form onSubmit={handleSubmit} className="p-8 space-y-6">
          <div className="p-4 bg-gray-50 rounded-xl border border-gray-100 flex justify-between items-center text-left">
            <div className="flex-1">
              <div className="text-xs font-bold text-gray-400 uppercase mb-1">Linh kiện / Stage</div>
              {data.partId ? (
                <div className="font-bold text-gray-900">{data.partId}</div>
              ) : (
                <div className="w-full">
                  <SearchableSelect 
                    options={availableParts.map(p => ({ id: p.id, label: `${p.id} - ${p.name}` }))}
                    value={selectedPartId}
                    onChange={setSelectedPartId}
                    placeholder="Chọn linh kiện..."
                  />
                </div>
              )}
              <div className="text-xs text-red-600 font-bold uppercase">{STAGES.find(s => s.id === data.stageId)?.name}</div>
            </div>
            {data.poId && (
              <div className="text-right">
                <div className="text-xs font-bold text-gray-400 uppercase mb-1">Lệnh PO</div>
                <div className="font-mono font-bold text-gray-700">{data.poId}</div>
              </div>
            )}
          </div>

          <div className="space-y-3">
            <label className="text-sm font-bold uppercase tracking-widest text-gray-500">1. Số lượng lỗi (NG)</label>
            <div className="flex gap-4">
              <button type="button" onClick={() => setQty(Math.max(0, qty - 1))} className="w-16 h-16 rounded-xl border-2 border-gray-100 flex items-center justify-center font-bold text-2xl hover:bg-gray-50 text-gray-900">-</button>
              <input 
                type="number" 
                step="any"
                autoFocus
                value={qty || ''} 
                onChange={(e) => setQty(parseFloat(e.target.value) || 0)}
                className="flex-1 h-16 text-center border-2 border-gray-900 rounded-xl font-mono font-bold text-2xl outline-none focus:ring-4 ring-red-500/5 text-gray-900"
                placeholder="0"
              />
              <button type="button" onClick={() => setQty(qty + 1)} className="w-16 h-16 rounded-xl border-2 border-gray-100 flex items-center justify-center font-bold text-2xl hover:bg-gray-50 text-gray-900">+</button>
            </div>
          </div>

          <div className="space-y-3">
            <label className="text-sm font-bold uppercase tracking-widest text-gray-500">2. Lý do lỗi</label>
            <select 
              value={reasonId}
              onChange={(e) => setReasonId(e.target.value)}
              className="w-full p-4 rounded-xl border-2 border-gray-100 font-bold text-lg focus:border-red-600 outline-none bg-white text-gray-900"
            >
              {DEFECT_REASONS.map(r => (
                <option key={r.id} value={r.id}>{r.name}</option>
              ))}
            </select>
          </div>

          <div className="space-y-3">
            <label className="text-sm font-bold uppercase tracking-widest text-gray-500">3. Ghi chú thêm (Tùy chọn)</label>
            <textarea 
              value={note}
              onChange={(e) => setNote(e.target.value)}
              className="w-full p-4 rounded-xl border-2 border-gray-100 font-medium text-lg focus:border-red-600 outline-none h-24 resize-none text-gray-900 bg-white"
              placeholder="Mô tả chi tiết tình trạng..."
            />
          </div>

          <button 
            type="submit"
            className="w-full h-16 bg-red-600 text-white rounded-xl font-black text-xl uppercase shadow-xl hover:bg-red-700 active:scale-95 transition-all flex items-center justify-center gap-3"
          >
            <AlertCircle size={24} /> Xác nhận báo lỗi
          </button>
        </form>
      </motion.div>

      {/* Error Message Modal */}
      <AnimatePresence>
        {errorMsg && (
          <div className="fixed inset-0 z-[300] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="bg-white rounded-3xl p-8 max-w-sm w-full shadow-2xl flex flex-col items-center text-center space-y-6 border-4 border-red-500/10"
            >
              <div className="w-16 h-16 bg-red-100 text-red-500 rounded-full flex items-center justify-center">
                <AlertCircle size={32} />
              </div>
              <div>
                <h4 className="text-xl font-bold text-gray-900 mb-2">Lỗi thao tác</h4>
                <p className="text-gray-600 font-medium">{errorMsg}</p>
              </div>
              <button 
                onClick={() => setErrorMsg("")}
                className="w-full py-4 bg-gray-900 text-white rounded-xl font-bold uppercase tracking-widest hover:bg-gray-800 transition-colors"
              >
                Đóng
              </button>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

function DisposalModal({ data, onClose, onDisposed, setLastTransaction }: { data: any, onClose: () => void, onDisposed: (tx: Transaction) => void, setLastTransaction: (tx: Transaction) => void }) {
  const [qty, setQty] = useState(data.quantity);
  const [recordedTx, setRecordedTx] = useState<Transaction | null>(null);

  const handleDispose = () => {
    try {
      const tx = storageService.recordDisposal(data.partId, data.stageId, qty, data.defectReason, data.defectCategory);
      setRecordedTx(tx);
      onDisposed(tx);
    } catch (err) {
      alert(err instanceof Error ? err.message : 'Lỗi khi xuất hủy');
    }
  };

  const handlePrint = () => {
    if (recordedTx) {
      setLastTransaction(recordedTx);
      setTimeout(() => {
        window.print();
      }, 200);
    }
  };

  if (recordedTx) {
    return (
      <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4 no-print">
        <motion.div initial={{ scale: 0.9, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} className="bg-white rounded-3xl p-8 max-w-md w-full shadow-2xl relative text-gray-900">
          <button onClick={onClose} className="absolute right-6 top-6 text-gray-400 hover:text-black"><X size={24} /></button>
          <div className="text-center">
            <div className="w-20 h-20 bg-green-100 text-green-600 rounded-full flex items-center justify-center mx-auto mb-6">
              <CheckCircle2 size={40} />
            </div>
            <h2 className="text-2xl font-bold mb-2">Đã xuất hủy thành công!</h2>
            <p className="text-gray-500 mb-8">Vui lòng in mã QR dưới đây để dán lên hàng đi hủy.</p>
            
            <div className="bg-gray-50 p-6 rounded-2xl mb-8 border border-gray-100 inline-block mx-auto w-full">
              <div className="flex justify-center mb-4">
                <QRCodeSVG value={recordedTx.qrData || ''} size={200} level="H" />
              </div>
              <div className="font-mono font-bold text-sm uppercase">{recordedTx.id}</div>
              <div className="text-[10px] text-gray-500 uppercase mt-1 tracking-widest font-bold">NHÃN ĐEM HỦY (DISPOSAL)</div>
              {recordedTx.defectReason && (
                <div className="mt-4 p-3 bg-red-50 text-red-700 rounded-xl border border-red-100">
                  <span className="text-[10px] uppercase tracking-widest font-bold block mb-1">Lỗi:</span>
                  <span className="font-black text-sm uppercase">{recordedTx.defectReason}</span>
                  {recordedTx.defectCategory && (
                    <div className="text-xs mt-1 italic break-words">Ghi chú: {recordedTx.defectCategory}</div>
                  )}
                </div>
              )}
            </div>

            <div className="flex gap-4 line-clamp-1">
              <button onClick={handlePrint} className="flex-1 bg-blue-600 text-white py-4 rounded-xl font-bold hover:bg-blue-700 shadow-lg shadow-blue-200">
                IN MÃ HỦY
              </button>
              <button onClick={onClose} className="flex-1 bg-gray-100 text-gray-600 py-4 rounded-xl font-bold hover:bg-gray-200">
                ĐÓNG
              </button>
            </div>
          </div>
        </motion.div>
      </div>
    );
  }

  return (
     <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
        <motion.div initial={{ y: 20, opacity: 0 }} animate={{ y: 0, opacity: 1 }} className="bg-white rounded-3xl p-8 max-w-md w-full shadow-2xl relative text-gray-900">
          <button onClick={onClose} className="absolute right-6 top-6 text-gray-400 hover:text-black"><X size={24} /></button>
          <h2 className="text-2xl font-bold mb-2 flex items-center gap-3">
            <Trash2 className="text-red-600" />
            Xác nhận xuất hủy
          </h2>
          <p className="text-gray-500 mb-6 font-medium">Bạn đang thực hiện xuất hủy hàng lỗi khỏi kho. Hành động này không thể hoàn tác.</p>
          
          <div className="bg-red-50 p-6 rounded-2xl border border-red-100 mb-8">
             <div className="text-xs font-bold text-red-400 uppercase tracking-widest mb-1">Chi tiết hàng lỗi</div>
             <div className="text-lg font-bold text-red-900">{data.partId}</div>
             <div className="text-sm font-medium text-red-700 opacity-70 mb-4">{data.partName}</div>
             <div className="flex justify-between items-center bg-white/50 p-3 rounded-lg border border-red-100">
                <span className="text-xs font-bold text-red-900">Tại kho:</span>
                <span className="text-sm font-bold text-red-900 bg-red-100 px-2 py-0.5 rounded italic">{data.stageName} (DEFECT)</span>
             </div>
          </div>

          <div className="space-y-4 mb-8">
            <label className="text-xs font-bold text-gray-500 uppercase tracking-widest">Số lượng xuất hủy ({data.unit})</label>
            <div className="flex items-center gap-4">
              <input 
                type="number"
                value={qty}
                onChange={(e) => setQty(parseFloat(e.target.value) || 0)}
                max={data.quantity}
                className="flex-1 p-5 rounded-2xl border-2 border-gray-100 font-bold text-2xl focus:border-red-600 outline-none text-center"
              />
            </div>
            <p className="text-[10px] text-gray-400 italic font-medium">* Tồn kho lỗi hiện có: {data.quantity} {data.unit}</p>
          </div>

          <div className="flex gap-4">
            <button onClick={handleDispose} className="flex-1 bg-red-600 text-white py-5 rounded-2xl font-bold uppercase tracking-widest hover:bg-red-700 shadow-xl shadow-red-200 transition-all active:scale-95">
              XUẤT HỦY & IN QR
            </button>
          </div>
        </motion.div>
     </div>
  );
}

function DefectView({ parts, transactions, inventory, refreshData, setLastTransaction }: any) {
  const [searchTerm, setSearchTerm] = useState("");
  const [disposalModal, setDisposalModal] = useState<any>(null);

  const defectItems = inventory.filter((i: any) => i.location === 'DEFECT' && i.quantity > 0);
  const combinedTxs = transactions.filter((t: any) => t.type === 'DEFECT' || t.type === 'DISPOSAL')
    .sort((a: any, b: any) => b.timestamp - a.timestamp);

  const filteredItems = defectItems.filter((i: any) => i.partId.toLowerCase().includes(searchTerm.toLowerCase()));
  const filteredTxs = combinedTxs.filter((tx: any) => 
    tx.partId.toLowerCase().includes(searchTerm.toLowerCase()) || 
    (tx.defectReason || "").toLowerCase().includes(searchTerm.toLowerCase())
  );

  return (
    <motion.div 
      initial={{ opacity: 0, x: 20 }}
      animate={{ opacity: 1, x: 0 }}
      exit={{ opacity: 0, x: -20 }}
      className="space-y-8"
    >
      <div className="bg-white p-8 rounded-3xl border border-gray-200 shadow-sm text-gray-900">
        <div className="flex items-center gap-4 mb-8">
          <div className="w-12 h-12 bg-red-100 text-red-600 rounded-2xl flex items-center justify-center">
            <AlertCircle size={24} />
          </div>
          <div className="text-left">
            <h2 className="text-2xl font-bold tracking-tight">Tồn kho hàng lỗi hiện tại</h2>
            <p className="text-xs text-gray-400 font-bold uppercase tracking-widest">Danh sách các linh kiện đang chờ xử lý hủy</p>
          </div>
        </div>

        {defectItems.length > 0 ? (
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            {filteredItems.map((item: any) => {
              const part = parts.find((p: any) => p.id === item.partId);
              const lastDefectTx = combinedTxs.find((tx: any) => tx.partId === item.partId && tx.stageId === item.stageId && tx.type === 'DEFECT');
              return (
                <div key={`${item.partId}-${item.stageId}`} className="p-6 rounded-2xl bg-gray-50 border border-gray-100 flex flex-col justify-between text-left">
                  <div>
                    <div className="flex justify-between items-start mb-4">
                      <span className="px-2 py-1 bg-white text-gray-500 text-[10px] font-bold rounded border border-gray-200 uppercase">
                        {STAGES.find(s => s.id === item.stageId)?.name}
                      </span>
                      <span className="font-mono text-2xl font-bold text-red-600">{item.quantity}</span>
                    </div>
                    <h3 className="font-bold text-gray-900 mb-1">{item.partId}</h3>
                    <p className="text-xs text-gray-500 uppercase truncate mb-4">{part?.name}</p>
                    {lastDefectTx && lastDefectTx.defectReason && (
                      <div className="mb-4 text-xs font-medium text-red-600 bg-red-50 p-2 rounded truncate break-all">
                        Lỗi: {lastDefectTx.defectReason}
                        {lastDefectTx.defectCategory && ` - ${lastDefectTx.defectCategory}`}
                      </div>
                    )}
                  </div>
                  <button 
                    onClick={() => setDisposalModal({ 
                      partId: item.partId, 
                      partName: part?.name,
                      stageId: item.stageId,
                      stageName: STAGES.find(s => s.id === item.stageId)?.name,
                      quantity: item.quantity,
                      unit: part?.unit,
                      defectReason: lastDefectTx?.defectReason,
                      defectCategory: lastDefectTx?.defectCategory
                    })}
                    className="w-full bg-red-600 text-white py-3 rounded-xl font-bold uppercase tracking-widest text-xs hover:bg-red-700 transition-all flex items-center justify-center gap-2"
                  >
                    <Trash2 size={16} />
                    Xuất hủy & In QR
                  </button>
                </div>
              );
            })}
          </div>
        ) : (
          <div className="p-12 text-center bg-gray-50 rounded-2xl border border-dashed border-gray-200">
            <p className="text-gray-400 italic">Không có hàng lỗi tồn kho.</p>
          </div>
        )}
      </div>

      <div className="bg-white p-8 rounded-3xl border border-gray-200 shadow-sm text-gray-900">
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6 mb-8">
          <div className="text-left">
            <h2 className="text-2xl font-bold tracking-tight">Cập nhật & Nhật ký (NG Log)</h2>
            <p className="text-xs text-gray-400 font-bold uppercase tracking-widest">Lịch sử ghi nhận và xuất hủy hàng lỗi</p>
          </div>
          <div className="relative w-full md:w-64">
            <input 
              type="text"
              placeholder="Tìm theo mã LK, Lý do..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full pl-10 pr-4 py-3 rounded-xl border border-gray-200 focus:border-blue-600 outline-none"
            />
            <Search size={18} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="border-b-2 border-gray-100 bg-gray-50/50">
                <th className="p-4 text-xs font-mono uppercase opacity-60">Thời gian</th>
                <th className="p-4 text-xs font-mono uppercase opacity-60">Loại</th>
                <th className="p-4 text-xs font-mono uppercase opacity-60">Linh kiện</th>
                <th className="p-4 text-xs font-mono uppercase opacity-60">Công đoạn</th>
                <th className="p-4 text-xs font-mono uppercase opacity-60 text-right">Số lượng</th>
                <th className="p-4 text-xs font-mono uppercase opacity-60">Chi tiết</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50 text-sm">
              {filteredTxs.map((tx: any) => (
                <tr key={tx.id} className="hover:bg-gray-50/80 transition-colors">
                  <td className="p-4 font-mono text-xs opacity-60">{format(tx.timestamp, 'dd/MM HH:mm')}</td>
                  <td className="p-4">
                    {tx.type === 'DEFECT' ? (
                      <span className="px-2 py-1 bg-red-100 text-red-700 text-[10px] font-bold rounded uppercase whitespace-nowrap">Ghi nhận lỗi</span>
                    ) : (
                      <span className="px-2 py-1 bg-gray-600 text-white text-[10px] font-bold rounded uppercase whitespace-nowrap">Xuất hủy</span>
                    )}
                  </td>
                  <td className="p-4 text-left">
                    <div className="font-bold">{tx.partId}</div>
                    <div className="text-[10px] text-gray-500 uppercase truncate max-w-[150px]">{parts.find((p: any) => p.id === tx.partId)?.name}</div>
                  </td>
                  <td className="p-4">
                    <span className="px-2 py-1 bg-gray-100 text-gray-600 text-[10px] font-bold rounded uppercase">
                      {STAGES.find(s => s.id === tx.stageId)?.name}
                    </span>
                  </td>
                  <td className={`p-4 font-mono font-bold text-lg text-right ${tx.type === 'DEFECT' ? 'text-red-600' : 'text-gray-900'}`}>
                    {tx.type === 'DEFECT' ? '+' : '-'}{tx.quantity}
                  </td>
                  <td className="p-4 text-left">
                    {tx.type === 'DEFECT' ? (
                      <>
                        <div className="font-bold text-gray-900">{tx.defectReason}</div>
                        {tx.defectCategory && <div className="text-[10px] text-gray-500 italic mt-0.5">Note: {tx.defectCategory}</div>}
                      </>
                    ) : (
                      <div className="text-gray-500 font-medium">Xuất disposal (QR: {tx.id})</div>
                    )}
                  </td>
                </tr>
              ))}
              {filteredTxs.length === 0 && (
                <tr>
                  <td colSpan={6} className="p-20 text-center text-gray-400 italic">
                    Chưa có nhật ký nào.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {disposalModal && (
        <DisposalModal 
          data={disposalModal} 
          setLastTransaction={setLastTransaction}
          onClose={() => setDisposalModal(null)} 
          onDisposed={(tx) => {
            refreshData();
          }} 
        />
      )}
    </motion.div>
  );
}
