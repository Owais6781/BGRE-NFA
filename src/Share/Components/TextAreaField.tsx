import * as React from 'react';
import "../styles/global.scss";

interface TextAreaFieldProps {
  label: string;
  value: string;
  onChange: (value: string) => void;
  rows?: number;
  placeholder?: string;
  required?: boolean;
  disabled?: boolean;
  error?: string;
}

export const TextAreaField: React.FC<TextAreaFieldProps> = ({
  label, value, onChange, rows = 3,
  placeholder, required = false, disabled = false, error,
}) => {
  return (
    <div className="inputWrapper">
      <label className="formLabel">
        {label}
        {required && <span className="text-red-500 ml-[2px]">*</span>}
      </label>
      <textarea
        rows={rows}
        value={value || ''}
        placeholder={placeholder}
        disabled={disabled}
        onChange={(e) => onChange(e.target.value)}
        className={`formInput resize-y${error ? ' border-red-400' : ''}`}
      />
      {error && (
        <span className="text-[11px] text-red-500 mt-0.5">{error}</span>
      )}
    </div>
  );
};

export default TextAreaField;
