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
        type={type}
        value={value || ''}
        placeholder={placeholder}
        disabled={disabled}
        onChange={(e) => onChange(e.target.value)}
        className={`formInput${error ? ' border-red-400 focus:border-red-500' : ''}`}
      />
      {error && (
        <span className="text-red-500 text-[11px] font-medium mt-1 block">
          {error}
        </span>
      )}
    </div>
  );
};

export default InputField;
