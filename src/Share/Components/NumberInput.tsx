import * as React from 'react';
import "../styles/global.scss";

interface NumberInputProps {
  label: string;
  value: string | number;
  onChange: (value: string) => void;
  onBlur?: (value: string) => void;
  placeholder?: string;
  required?: boolean;
  disabled?: boolean;
  min?: number;
  step?: number;
  error?: string;
}

export const NumberInput: React.FC<NumberInputProps> = ({
  label, value, onChange, onBlur, placeholder = '0.00',
  required = false, disabled = false, min, step = 0.01, error,
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
      <input
        type="text"
        value={value}
        placeholder={placeholder}
        disabled={disabled}
        onBlur={(e) => {
          if (onBlur) {
            onBlur(e.target.value);
          }
        }}
        onChange={(e) => {
          // Allow only numbers and a single decimal point
          const val = e.target.value.replace(/[^0-9.]/g, '');
          if ((val.match(/\./g) || []).length <= 1) {
            onChange(val);
          }
        }}
        className={`formInput${error ? ' border-red-400' : ''}`}
      />
    </div>
  );
};

export default NumberInput;
