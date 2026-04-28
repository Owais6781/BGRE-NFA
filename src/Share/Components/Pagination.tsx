import * as React from 'react';
import "../styles/global.scss";

interface PaginationProps {
  page: number;
  total: number;
  pageSize: number;
  totalRecords: number;
  pageSizeOptions?: number[];
  onPageChange: (page: number) => void;
  onPageSizeChange: (size: number) => void;
}

export const Pagination: React.FC<PaginationProps> = ({
  page,
  total,
  pageSize,
  totalRecords,
  pageSizeOptions = [10, 25, 50, 100],
  onPageChange,
  onPageSizeChange,
}) => {
  const from = totalRecords === 0 ? 0 : (page - 1) * pageSize + 1;
  const to   = Math.min(page * pageSize, totalRecords);

  return (
    <div className="flex items-center justify-between flex-wrap gap-3 mt-4 pt-3 border-t border-[#e2e8f0]">
      <div className="flex items-center gap-3">
        <span className="formLabel" style={{ marginBottom: 0 }}>
          Rows per page:
        </span>
        <select
          value={pageSize}
          onChange={(e) => onPageSizeChange(Number(e.target.value))}
          className="formInput formSelect"
          style={{ width: 'auto', minWidth: 70 }}
        >
          {pageSizeOptions.map((s) => (
            <option key={s} value={s}>{s}</option>
          ))}
        </select>
        <span className="formLabel" style={{ marginBottom: 0, color: '#94a3b8' }}>
          Showing {from}–{to} of {totalRecords}
        </span>
      </div>

      <div className="flex items-center gap-1">
        <button
          onClick={() => onPageChange(1)}
          disabled={page === 1}
          className="btn btnSecondary"
          title="First"
        >
          «
        </button>
        <button
          onClick={() => onPageChange(page - 1)}
          disabled={page === 1}
          className="btn btnSecondary"
        >
          Previous
        </button>
        <span className="btn btnPrimary" style={{ cursor: 'default', pointerEvents: 'none' }}>{page}</span>
        <button
          onClick={() => onPageChange(page + 1)}
          disabled={page === total}
          className="btn btnSecondary"
        >
          Next
        </button>
        <button
          onClick={() => onPageChange(total)}
          disabled={page === total}
          className="btn btnSecondary"
          title="Last"
        >
          »
        </button>
      </div>
    </div>
  );
};

export default Pagination;
