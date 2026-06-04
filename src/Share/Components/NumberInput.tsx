import * as React from 'react';
import "../styles/global.scss";

interface NumberInputProps {
  label: string;
  value: string | number;
  onChange: (value: string) => void;
  onBlur?: () => void;
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
      <label className="formLabel">
        {label}
        {required && <span className="text-red-500 ml-[2px]">*</span>}
      </label>
      <input
        type="text"
        value={value}
        placeholder={placeholder}
        disabled={disabled}
        onBlur={onBlur}
        onChange={(e) => {
          // Allow only numbers and a single decimal point
          const val = e.target.value.replace(/[^0-9.]/g, '');
          if ((val.match(/\./g) || []).length <= 1) {
            onChange(val);
          }
        }}
        className={`formInput${error ? ' border-red-400' : ''}`}
      />
      {error && (
        <span className="text-[11px] text-red-500 mt-0.5">{error}</span>
      )}
    </div>
  );
};

export default NumberInput;
