import * as React from 'react';
import "../styles/global.scss";

interface DropdownFieldProps {
  label: string;
  value: string;
  onChange: (value: string) => void;
  options: { label: string; value: string }[];
  required?: boolean;
  disabled?: boolean;
}

const DropdownField: React.FC<DropdownFieldProps> = ({
  label, value, onChange, options, required = false, disabled = false,
}) => {
  return (
    <div className="inputWrapper">
      <label className="formLabel">
        {label}
        {required && <span className="text-red-500 ml-[2px]">*</span>}
      </label>
      <select
        value={value}
        onChange={(e) => onChange(e.target.value)}
        disabled={disabled}
        className="formInput formSelect"
      >
        <option value="">All</option>
        {options.map((item) => (
          <option key={item.value} value={item.value}>{item.label}</option>
        ))}
      </select>
    </div>
  );
};

export default DropdownField;
