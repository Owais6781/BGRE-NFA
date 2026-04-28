import * as React from 'react';
import { useState } from 'react';
import "../styles/global.scss";

interface ReportTableProps {
  columns:        { key: string; label: string }[];
  data:           any[];
  emptyMessage?:  string;
  loading?:       boolean;
}

const ReportTable: React.FC<ReportTableProps> = ({
  columns,
  data,
  emptyMessage = 'Please click on search button to download report',
  loading      = false,
}) => {
  return (
    <div className="tableWrapper mt-3">
      <table className="commonTable">
        <thead>
          <tr>
            {columns.map((col) => (
              <th key={col.key} className="reportTableHeader">
                {col.label}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {loading ? (
            <tr>
              <td
                colSpan={columns.length}
                className="reportTableCell text-center py-4"
              >
                Loading...
              </td>
            </tr>
          ) : data.length === 0 ? (
            <tr>
              <td
                colSpan={columns.length}
                className="reportTableCell py-3 px-3"
              >
                {emptyMessage}
              </td>
            </tr>
          ) : (
            data.map((row, i) => (
              <tr key={i} className="tableRow">
                {columns.map((col) => (
                  <td key={col.key} className="reportTableCell">
                    {row[col.key] ?? '—'}
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

export default ReportTable;