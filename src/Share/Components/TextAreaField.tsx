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
      <textarea
        rows={rows}
        value={value || ''}
        placeholder={placeholder}
        disabled={disabled}
        onChange={(e) => onChange(e.target.value)}
        className={`formInput resize-y${error ? ' border-red-400' : ''}`}
      />
    </div>
  );
};

export default TextAreaField;
