import * as React from 'react';
import "../styles/global.scss";

interface SearchInputProps {
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
}

export const SearchInput: React.FC<SearchInputProps> = ({
  value,
  onChange,
  placeholder = 'Search…',
}) => {
  return (
    <div className="flex items-center gap-2">
      <label className="formLabel" style={{ marginBottom: 0, whiteSpace: 'nowrap' }}>
        Search:
      </label>
      <div className="relative">
        <span className="absolute left-3 top-1/2 -translate-y-1/2 text-[#94a3b8] pointer-events-none">
          <svg width="13" height="13" viewBox="0 0 20 20" fill="none">
            <circle cx="9" cy="9" r="7" stroke="currentColor" strokeWidth="2" />
            <path d="M14 14l4 4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
          </svg>
        </span>
        <input
          type="text"
          value={value}
          onChange={(e) => onChange(e.target.value)}
          placeholder={placeholder}
          className="formInput"
          style={{ paddingLeft: '30px', width: '200px' }}
        />
      </div>
    </div>
  );
};

export default SearchInput;
