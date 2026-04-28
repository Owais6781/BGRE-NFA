import * as React from 'react';
import "../styles/global.scss";

interface IInputFieldProps {
  label: string;
  value: string;
  onChange: (value: string) => void;
  type?: string;
  placeholder?: string;
  required?: boolean;
  disabled?: boolean;
  error?: string;
}

export const InputField: React.FC<IInputFieldProps> = ({
  label, value, onChange, type = 'text',
  placeholder, required = false, disabled = false, error,
}) => {
  return (
    <div className="inputWrapper">
      <label className="formLabel">
        {label}
        {required && <span className="text-red-500 ml-[2px]">*</span>}
      </label>
      <input
        type={type}
        value={value}
        placeholder={placeholder}
        disabled={disabled}
        onChange={(e) => onChange(e.target.value)}
        className={`formInput${error ? ' border-red-400 focus:border-red-500' : ''}`}
      />
      {error && (
        <span className="text-[11px] text-red-500 flex items-center gap-1 mt-0.5">
          {error}
        </span>
      )}
    </div>
  );
};

export default InputField;
