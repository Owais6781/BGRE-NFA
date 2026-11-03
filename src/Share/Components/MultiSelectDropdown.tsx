import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import { ChevronDown, X, Check } from 'lucide-react';

export interface IMultiSelectOption {
  label: string;
  value: string;
}

interface MultiSelectDropdownProps {
  label: string;
  options: IMultiSelectOption[];
  selectedValues: string[];
  onChange: (values: string[]) => void;
  placeholder?: string;
  disabled?: boolean;
  required?: boolean;
  error?: string;
}

const MultiSelectDropdown: React.FC<MultiSelectDropdownProps> = ({
  label,
  options,
  selectedValues,
  onChange,
  placeholder = 'All',
  disabled = false,
  required = false,
  error,
}) => {
  const [isOpen, setIsOpen] = useState(false);
  const [searchText, setSearchText] = useState('');
  const dropdownRef = useRef<HTMLDivElement>(null);

  // Close dropdown when clicking outside
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setIsOpen(false);
        setSearchText('');
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Toggle a single value
  const toggleValue = (value: string) => {
    if  (selectedValues.indexOf(value) !== -1) {
      onChange(selectedValues.filter(v => v !== value));
    } else {
      onChange([...selectedValues, value]);
    }
  };

  // Remove a single tag
  const removeValue = (value: string, e: React.MouseEvent) => {
    e.stopPropagation();
    onChange(selectedValues.filter(v => v !== value));
  };

  // Select All / Deselect All
  const toggleSelectAll = () => {
    if (selectedValues.length === options.length) {
      onChange([]);
    } else {
      onChange(options.map(o => o.value));
    }
  };

  // Clear All
  const clearAll = (e: React.MouseEvent) => {
    e.stopPropagation();
    onChange([]);
  };

  // Filter options by search
  const filteredOptions = options.filter(o =>
    o.label.toLowerCase().includes(searchText.toLowerCase())
  );

  // Get label by value
  const getLabel = (value: string): string => {
    const option = options.find(o => o.value === value);
    return option ? option.label : value;
  };

  return (
    <div className="multiSelect" ref={dropdownRef}>

      {/* Label */}
      <label className="multiSelect__label flex items-center gap-1.5" style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
        {label}
        {required && <span className="multiSelect__required"> *</span>}
        {error && (
          <span className="inline-flex h-1.5 w-1.5 relative" style={{ display: 'inline-flex' }}>
            <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-red-400 opacity-75"></span>
            <span className="relative inline-flex rounded-full h-1.5 w-1.5 bg-red-500"></span>
          </span>
        )}
      </label>

      {/* Trigger Box */}
      <div
        className={`multiSelect__trigger ${disabled ? 'multiSelect__trigger--disabled' : ''} ${isOpen ? 'multiSelect__trigger--active' : ''} ${error ? 'border-red-400' : ''}`}
        onClick={() => !disabled && setIsOpen(!isOpen)}
      >
        <div className="multiSelect__tags">
          {selectedValues.length === 0 && (
            <span className="multiSelect__placeholder">{placeholder}</span>
          )}

          {selectedValues.map(val => (
            <span key={val} className="multiSelect__tag">
              {getLabel(val)}
              <span className="multiSelect__tagRemove" onClick={(e) => removeValue(val, e)}>
                <X size={12} />
              </span>
            </span>
          ))}
        </div>

        <div className="multiSelect__actions">
          {selectedValues.length > 0 && (
            <span className="multiSelect__clear" onClick={clearAll}>
              <X size={14} />
            </span>
          )}
          <ChevronDown
            size={16}
            className={`multiSelect__chevron ${isOpen ? 'multiSelect__chevron--open' : ''}`}
          />
        </div>
      </div>

      {/* Dropdown Panel */}
      {isOpen && !disabled && (
        <div className="multiSelect__dropdown">

          {/* Search Box */}
          <div className="multiSelect__search">
            <input
              type="text"
              className="multiSelect__searchInput"
              placeholder="Search..."
              value={searchText}
              onChange={(e) => setSearchText(e.target.value)}
              autoFocus
            />
          </div>

          {/* Select All */}
          {options.length > 0 && (
            <div className="multiSelect__selectAll" onClick={toggleSelectAll}>
              <div className={`multiSelect__checkbox ${selectedValues.length === options.length ? 'multiSelect__checkbox--checked' : ''}`}>
                {selectedValues.length === options.length && <Check size={12} />}
              </div>
              <span>
                {selectedValues.length === options.length ? 'Deselect All' : 'Select All'}
              </span>
            </div>
          )}

          {/* Options List */}
          <div className="multiSelect__options">
            {filteredOptions.length === 0 ? (
              <div className="multiSelect__noData">No options found</div>
            ) : (
              filteredOptions.map(option => {
                const isSelected = selectedValues.indexOf(option.value) !== -1;
                return (
                  <div
                    key={option.value}
                    className={`multiSelect__option ${isSelected ? 'multiSelect__option--selected' : ''}`}
                    onClick={() => toggleValue(option.value)}
                  >
                    <div className={`multiSelect__checkbox ${isSelected ? 'multiSelect__checkbox--checked' : ''}`}>
                      {isSelected && <Check size={12} />}
                    </div>
                    <span>{option.label}</span>
                  </div>
                );
              })
            )}
          </div>

          {/* Footer */}
          <div className="multiSelect__footer">
            <span>{selectedValues.length} of {options.length} selected</span>
          </div>
        </div>
      )}
      {error && (
        <span className="text-red-500 text-[11px] font-medium mt-1 block">
          {error}
        </span>
      )}
    </div>
  );
};

export default MultiSelectDropdown;
