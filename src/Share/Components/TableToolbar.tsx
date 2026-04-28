import * as React from 'react';
import { SearchInput } from './SearchInput';

const DEFAULT_PAGE_SIZE_OPTIONS = [10, 25, 50, 100];

interface TableToolbarProps {
  /** Built-in page size selector rendered on the left */
  pageSize?: number;
  pageSizeOptions?: number[];
  onPageSizeChange?: (size: number) => void;
  /** Built-in search input rendered on the right */
  search?: string;
  onSearch?: (val: string) => void;
  searchPlaceholder?: string;
  /** Custom slot overrides */
  left?: React.ReactNode;
  right?: React.ReactNode;
}

const TableToolbar: React.FC<TableToolbarProps> = ({
  pageSize,
  pageSizeOptions = DEFAULT_PAGE_SIZE_OPTIONS,
  onPageSizeChange,
  search,
  onSearch,
  searchPlaceholder,
  left,
  right,
}) => {
  const leftContent: React.ReactNode = left !== undefined
    ? left
    : pageSize !== undefined && onPageSizeChange
      ? (
        <div className="flex items-center gap-2">
          <span className="formLabel" style={{ marginBottom: 0 }}>Show</span>
          <select
            value={pageSize}
            onChange={(e) => onPageSizeChange(Number(e.target.value))}
            className="formInput formSelect"
            style={{ width: 70 }}
          >
            {pageSizeOptions.map((n) => <option key={n} value={n}>{n}</option>)}
          </select>
          <span className="formLabel" style={{ marginBottom: 0 }}>entries</span>
        </div>
      )
      : null;

  const rightContent: React.ReactNode = right !== undefined
    ? right
    : search !== undefined && onSearch
      ? <SearchInput value={search} onChange={onSearch} placeholder={searchPlaceholder} />
      : null;

  if (!leftContent && !rightContent) return null;

  return (
    <div className="tableToolbar">
      {leftContent ?? <span />}
      {rightContent}
    </div>
  );
};

export default TableToolbar;
