import * as React from 'react';
import "../styles/global.scss";

interface ColumnDef {
  key: string;
  label: string;
  render?: (value: any, row: any) => React.ReactNode;
}

interface CommonTableProps {
  columns: ColumnDef[];
  data: any[];
  onRowClick?: (row: any) => void;
  loading?: boolean;
  emptyMessage?: string;
}

export const CommonTable: React.FC<CommonTableProps> = ({
  columns,
  data,
  onRowClick,
  loading = false,
  emptyMessage = 'No records found',
}) => {
  return (
    <div className="tableWrapper">
      <table className="commonTable">
        <thead>
          <tr>
            {columns.map((col) => (
              <th key={col.key} className="tableHeader">
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
            data.map((row, i) => (
              <tr
                key={i}
                className={`tableRow${onRowClick ? ' cursor-pointer' : ''}`}
                onClick={() => onRowClick?.(row)}
              >
                {columns.map((col) => (
                  <td key={col.key} className="tableCell">
                    {col.render ? col.render(row[col.key], row) : row[col.key] ?? '—'}
                  </td>
                ))}
              </tr>
            ))
          )}
        </tbody>
      </table>
    </div>
  );
};

export default CommonTable;
