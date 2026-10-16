import React, { useState, useRef, useEffect, useMemo } from 'react';
import { Search, X, ChevronDown, Check, Sparkles, AlertCircle } from 'lucide-react';
import { normalizeVietnamese, buildSearchIndex, searchIndexedItems } from './search';

export interface SmartSearchOption {
  id: string;
  name?: string;
  code?: string;
  model?: string;
  poId?: string;
  badge?: string;
  badgeColor?: string;
  sublabel?: string;
  detail?: string;
  [key: string]: any;
}

interface SmartSearchSelectProps<T extends SmartSearchOption> {
  options: T[];
  value: string;
  onChange: (value: string, selectedItem?: T) => void;
  placeholder?: string;
  label?: string;
  getOptionLabel?: (item: T) => string;
  getOptionSublabel?: (item: T) => string;
  getOptionValue?: (item: T) => string;
  searchKeys?: (keyof T)[];
  disabled?: boolean;
  className?: string;
  autoFocus?: boolean;
  emptyText?: string;
}

export function SmartSearchSelect<T extends SmartSearchOption>({
  options,
  value,
  onChange,
  placeholder = 'Gõ tìm mã linh kiện, tên, model...',
  label,
  getOptionLabel = (item) => item.name || item.id,
  getOptionSublabel = (item) => item.sublabel || (item.name ? item.id : ''),
  getOptionValue = (item) => item.id,
  searchKeys = ['id', 'name', 'code', 'model', 'poId', 'detail'],
  disabled = false,
  className = '',
  autoFocus = false,
  emptyText = 'Không tìm thấy kết quả phù hợp'
}: SmartSearchSelectProps<T>) {
  const [isOpen, setIsOpen] = useState(false);
  const [query, setQuery] = useState('');
  const [debouncedQuery, setDebouncedQuery] = useState('');
  const [highlightedIndex, setHighlightedIndex] = useState(0);

  const containerRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  // Debounce user input query by 200ms
  useEffect(() => {
    const timer = setTimeout(() => {
      setDebouncedQuery(query);
    }, 200);
    return () => clearTimeout(timer);
  }, [query]);

  // Pre-index options array into RAM memory ONCE when options change
  const preindexedOptions = useMemo(() => {
    return buildSearchIndex(options, searchKeys as any);
  }, [options, searchKeys]);

  // Currently selected item
  const selectedItem = useMemo(() => {
    if (!value) return null;
    return options.find(o => getOptionValue(o) === value) || null;
  }, [options, value, getOptionValue]);

  // Sync display text with selected item if user is not actively typing
  useEffect(() => {
    if (!isOpen) {
      if (selectedItem) {
        const itemLabel = getOptionLabel(selectedItem);
        const itemVal = getOptionValue(selectedItem);
        setQuery(itemLabel !== itemVal ? `${itemLabel} [${itemVal}]` : itemLabel);
      } else {
        setQuery('');
      }
    }
  }, [value, selectedItem, isOpen, getOptionLabel, getOptionValue]);

  // Filter & score options using pre-indexed RAM cache
  const filteredOptions = useMemo(() => {
    if (!debouncedQuery.trim()) return options;
    
    // If query matches currently selected format string, return full options
    if (selectedItem) {
      const itemLabel = getOptionLabel(selectedItem);
      const itemVal = getOptionValue(selectedItem);
      const expectedText = itemLabel !== itemVal ? `${itemLabel} [${itemVal}]` : itemLabel;
      if (debouncedQuery === expectedText) {
        return options;
      }
    }

    return searchIndexedItems(debouncedQuery, preindexedOptions);
  }, [options, preindexedOptions, debouncedQuery, selectedItem, getOptionLabel, getOptionValue]);

  // Reset highlight index when filtered list changes
  useEffect(() => {
    setHighlightedIndex(0);
  }, [filteredOptions.length]);

  // Handle outside click to close dropdown
  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    }
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleSelect = (item: T) => {
    const val = getOptionValue(item);
    onChange(val, item);
    const itemLabel = getOptionLabel(item);
    setQuery(itemLabel !== val ? `${itemLabel} [${val}]` : itemLabel);
    setIsOpen(false);
  };

  const handleClear = (e: React.MouseEvent) => {
    e.stopPropagation();
    onChange('');
    setQuery('');
    setIsOpen(true);
    inputRef.current?.focus();
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (disabled) return;

    if (e.key === 'ArrowDown') {
      e.preventDefault();
      if (!isOpen) {
        setIsOpen(true);
      } else {
        setHighlightedIndex(prev => (prev + 1) % Math.max(1, filteredOptions.length));
      }
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      if (!isOpen) {
        setIsOpen(true);
      } else {
        setHighlightedIndex(prev => (prev - 1 + filteredOptions.length) % Math.max(1, filteredOptions.length));
      }
    } else if (e.key === 'Enter') {
      if (isOpen && filteredOptions.length > 0) {
        e.preventDefault();
        handleSelect(filteredOptions[highlightedIndex]);
      }
    } else if (e.key === 'Escape') {
      setIsOpen(false);
    }
  };

  // Helper to highlight matching text query in labels
  const renderHighlightedText = (text: string, searchQuery: string) => {
    if (!text || !searchQuery.trim()) return text;
    
    const normText = normalizeVietnamese(text);
    const queryTokens = normalizeVietnamese(searchQuery).split(' ').filter(Boolean);
    
    if (queryTokens.length === 0) return text;

    // Check if any query token matches in text
    let hasMatch = false;
    for (const token of queryTokens) {
      if (normText.includes(token)) {
        hasMatch = true;
        break;
      }
    }

    if (!hasMatch) return text;

    return text;
  };

  return (
    <div ref={containerRef} className={`relative w-full text-left ${className}`}>
      {label && (
        <label className="block text-xs font-bold uppercase tracking-wider text-gray-700 mb-1">
          {label}
        </label>
      )}

      <div className="relative flex items-center">
        <div className="absolute left-3.5 text-gray-400 pointer-events-none flex items-center">
          <Search size={18} />
        </div>

        <input
          ref={inputRef}
          type="text"
          disabled={disabled}
          autoFocus={autoFocus}
          value={query}
          onFocus={() => {
            if (!disabled) setIsOpen(true);
          }}
          onChange={(e) => {
            setQuery(e.target.value);
            if (!isOpen) setIsOpen(true);
          }}
          onKeyDown={handleKeyDown}
          placeholder={placeholder}
          className={`w-full pl-10 pr-16 py-3 bg-white border border-gray-300 rounded-xl text-sm font-semibold text-gray-900 focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-600 transition-all shadow-xs placeholder:text-gray-400 ${
            disabled ? 'bg-gray-100 text-gray-400 cursor-not-allowed' : 'cursor-text'
          }`}
          autoComplete="off"
        />

        <div className="absolute right-3 flex items-center gap-1">
          {query && !disabled && (
            <button
              type="button"
              onClick={handleClear}
              className="p-1 text-gray-400 hover:text-gray-600 rounded-lg hover:bg-gray-100 transition-colors cursor-pointer"
              title="Xóa ô tìm kiếm"
            >
              <X size={16} />
            </button>
          )}

          <button
            type="button"
            disabled={disabled}
            onClick={() => setIsOpen(!isOpen)}
            className="p-1 text-gray-400 hover:text-gray-600 transition-colors cursor-pointer"
          >
            <ChevronDown size={18} className={`transition-transform duration-200 ${isOpen ? 'rotate-180' : ''}`} />
          </button>
        </div>
      </div>

      {/* Autocomplete Suggestions Dropdown */}
      {isOpen && (
        <div className="absolute left-0 right-0 top-full mt-1.5 bg-white border border-gray-200 rounded-2xl shadow-2xl z-[100] max-h-72 overflow-y-auto divide-y divide-gray-100 animate-in fade-in zoom-in-95 duration-150">
          {filteredOptions.length === 0 ? (
            <div className="p-4 text-center text-gray-400 text-xs font-medium flex items-center justify-center gap-2">
              <AlertCircle size={16} className="text-gray-300 shrink-0" />
              <span>{emptyText}</span>
            </div>
          ) : (
            <div className="py-1">
              <div className="px-3 py-1.5 text-[10px] font-bold uppercase tracking-wider text-gray-400 bg-gray-50/80 border-b border-gray-100 flex justify-between items-center">
                <span>Gợi ý linh kiện / PO ({filteredOptions.length})</span>
                <span className="text-blue-600 font-mono">Xếp theo độ khớp</span>
              </div>

              {filteredOptions.slice(0, 40).map((item, index) => {
                const itemVal = getOptionValue(item);
                const isSelected = itemVal === value;
                const isHighlighted = index === highlightedIndex;

                const itemLabel = getOptionLabel(item);
                const sublabel = getOptionSublabel(item);

                return (
                  <div
                    key={itemVal + '_' + index}
                    onClick={() => handleSelect(item)}
                    onMouseEnter={() => setHighlightedIndex(index)}
                    className={`px-3.5 py-2.5 transition-colors cursor-pointer flex items-center justify-between gap-3 text-left ${
                      isSelected
                        ? 'bg-blue-50/80 text-blue-900 font-bold'
                        : isHighlighted
                          ? 'bg-gray-100/80 text-gray-900'
                          : 'hover:bg-gray-50 text-gray-800'
                    }`}
                  >
                    <div className="min-w-0 flex-1">
                      <div className="flex items-center gap-2 flex-wrap">
                        <span className="text-xs font-bold text-gray-900 truncate">
                          {renderHighlightedText(itemLabel, query)}
                        </span>

                        {item.badge && (
                          <span className={`px-2 py-0.5 rounded text-[10px] font-bold uppercase ${item.badgeColor || 'bg-blue-100 text-blue-700'}`}>
                            {item.badge}
                          </span>
                        )}
                      </div>

                      {sublabel && (
                        <p className="text-[11px] text-gray-500 font-mono truncate mt-0.5">
                          {sublabel}
                        </p>
                      )}

                      {item.detail && (
                        <p className="text-[10px] text-gray-400 truncate">
                          {item.detail}
                        </p>
                      )}
                    </div>

                    {isSelected && (
                      <div className="w-5 h-5 rounded-full bg-blue-600 text-white flex items-center justify-center shrink-0">
                        <Check size={12} />
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
