import * as React from 'react';
import "../styles/global.scss";

interface NumberInputProps {
  label: string;
  value: string | number;
  onChange: (value: string) => void;
  placeholder?: string;
  required?: boolean;
  disabled?: boolean;
  min?: number;
  step?: number;
  error?: string;
}

export const NumberInput: React.FC<NumberInputProps> = ({
  label, value, onChange, placeholder = '0.00',
  required = false, disabled = false, min, step = 0.01, error,
}) => {
  return (
    <div className="inputWrapper">
      <label className="formLabel">
        {label}
        {required && <span className="text-red-500 ml-[2px]">*</span>}
      </label>
      <input
        type="number"
        value={value}
        placeholder={placeholder}
        disabled={disabled}
        min={min}
        step={step}
        onChange={(e) => onChange(e.target.value)}
        className={`formInput${error ? ' border-red-400' : ''}`}
      />
      {error && (
        <span className="text-[11px] text-red-500 mt-0.5">{error}</span>
      )}
    </div>
  );
};

export default NumberInput;
