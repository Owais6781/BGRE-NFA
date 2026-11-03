import * as React from 'react';
import "../styles/global.scss";

interface DropdownFieldProps {
  label: string;
  value: string;
  onChange: (value: string) => void;
  options: { label: string; value: string }[];
  required?: boolean;
  disabled?: boolean;
  emptyLabel?: string;
  showEmptyOption?: boolean;
  error?: string;
}

const DropdownField: React.FC<DropdownFieldProps> = ({
  label, value, onChange, options, required = false, disabled = false, emptyLabel, showEmptyOption = true, error,
}) => {
  return (
    <div className="inputWrapper">
      <label className="formLabel flex items-center gap-1.5" style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
        {label}
        {required && <span className="text-red-500 ml-[2px]">*</span>}
        {error && (
          <span className="inline-flex h-1.5 w-1.5 relative" style={{ display: 'inline-flex' }}>
            <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-red-400 opacity-75"></span>
            <span className="relative inline-flex rounded-full h-1.5 w-1.5 bg-red-500"></span>
          </span>
        )}
      </label>
      <select
        value={value}
        onChange={(e) => onChange(e.target.value)}
        disabled={disabled}
        className={`formInput formSelect${error ? ' border-red-400' : ''}`}
      >
        {showEmptyOption && <option value="">{emptyLabel || 'All'}</option>}
        {options.map((item) => (
          <option key={item.value} value={item.value}>{item.label}</option>
        ))}
      </select>
      {error && (
        <span className="text-red-500 text-[11px] font-medium mt-1 block">
          {error}
        </span>
      )}
    </div>
  );
};

export default DropdownField;
