import * as React from 'react';
import "../styles/global.scss";

interface ColumnDef {
  key: string;
  label: string;
  align?: 'left' | 'center' | 'right';
  className?: string;
  render?: (value: any, row: any) => React.ReactNode;
}

interface CommonTableProps {
  columns: ColumnDef[];
  data: any[];
  onRowClick?: (row: any) => void;
  loading?: boolean;
  emptyMessage?: string;
  stickyHeader?: boolean;
  stickyFirstColumn?: boolean;
  maxHeight?: string;
  minWidth?: string;
  footer?: React.ReactNode;
}

export const CommonTable: React.FC<CommonTableProps> = ({
  columns,
  data,
  onRowClick,
  loading = false,
  emptyMessage = 'No records found',
  stickyHeader = true,
  stickyFirstColumn = false,
  maxHeight,
  minWidth,
  footer,
}) => {
  return (
    <div className="tableWrapper" style={{ maxHeight }}>
      <table className="commonTable" style={{ minWidth }}>
        <thead>
          <tr>
            {columns.map((col, index) => (
              <th
                key={col.key}
                className={`tableHeader ${stickyHeader ? 'stickyHeader' : ''} ${stickyFirstColumn && index === 0 ? 'stickyCol' : ''} ${col.align === 'center' ? 'text-center' : col.align === 'right' ? 'text-right' : ''} ${col.className || ''}`.trim()}
                style={col.align ? { textAlign: col.align } : undefined}
              >
                {col.label}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {loading ? (
            <tr>
              <td colSpan={columns.length} className="tableCell text-center py-8 text-[#94a3b8]">
                <span className="inline-flex items-center gap-2">
                  <span className="inline-block w-4 h-4 border-2 border-[#3b82f6] border-t-transparent rounded-full animate-spin" />
                  Loading…
                </span>
              </td>
            </tr>
          ) : data.length === 0 ? (
            <tr>
              <td colSpan={columns.length} className="tableCell text-center py-10 text-[#94a3b8]">
                {emptyMessage}
              </td>
            </tr>
          ) : (
            <>
              {data.map((row, i) => (
                <tr
                  key={i}
                  className={`tableRow${onRowClick ? ' cursor-pointer' : ''}`}
                  onClick={() => onRowClick?.(row)}
                >
                  {columns.map((col, index) => (
                    <td
                      key={col.key}
                      className={`tableCell ${stickyFirstColumn && index === 0 ? 'stickyCol' : ''} ${col.align === 'center' ? 'text-center' : col.align === 'right' ? 'text-right' : ''} ${col.className || ''}`.trim()}
                      style={col.align ? { textAlign: col.align } : undefined}
                    >
                      {col.render ? col.render(row[col.key], row) : row[col.key] ?? '—'}
                    </td>
                  ))}
                </tr>
              ))}
              {footer}
            </>
          )}
        </tbody>
      </table>
    </div>
  );
};

export default CommonTable;
