import * as React from 'react';
import "../styles/global.scss";

interface DatePickerFieldProps {
  label: string;
  value: string;
  onChange: (value: string) => void;
  required?: boolean;
  min?: string;
  max?: string;
}

const DatePickerField: React.FC<DatePickerFieldProps> = ({
  label, value, onChange, required = false, min, max,
}) => {
  return (
    <div className="flex flex-col gap-[6px] w-full">
      <label className="formLabel">
        {label}
        {required && <span className="text-red-500 ml-[2px]">*</span>}
      </label>
      <div className="relative w-full">
        <input
          type="date"
          value={value}
          min={min}
          max={max}
          onChange={(e) => onChange(e.target.value)}
          className="formInput w-full pr-9"
        />
        <span className="absolute right-3 top-1/2 -translate-y-1/2 pointer-events-none text-[#666]">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/>
            <line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>
          </svg>
        </span>
      </div>
    </div>
  );
};

export default DatePickerField;