import * as React from 'react';
import * as XLSX from 'xlsx';
import "../styles/global.scss";

interface ColumnDef {
  key: string;
  label: string;
}

interface ExportToExcelProps {
  data: any[];
  columns: ColumnDef[];
  fileName?: string;
  sheetName?: string;
}

const ExportToExcel: React.FC<ExportToExcelProps> = ({
  data,
  columns,
  fileName = 'export',
  sheetName = 'Sheet1',
}) => {
  if (!data || data.length === 0) {
    return null;
  }

  const handleExport = () => {
    const exportRows = data.map((row) => {
      const formattedRow: Record<string, any> = {};
      columns.forEach((col) => {
        formattedRow[col.label] = row[col.key] ?? '';
      });
      return formattedRow;
    });

    const worksheet = XLSX.utils.json_to_sheet(exportRows);

    worksheet['!cols'] = columns.map((col) => ({
      wch: Math.max(col.label.length + 4, 15),
    }));

    const headerRange = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    for (let col = headerRange.s.c; col <= headerRange.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (worksheet[cellAddress]) {
        worksheet[cellAddress].s = {
          font: { bold: true },
          fill: { fgColor: { rgb: '1e3a5f' } },
        };
      }
    }

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    XLSX.writeFile(workbook, `${fileName}.xlsx`);
  };

  return (
    <button onClick={handleExport} className="exportBtn">
      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
        <polyline points="7 10 12 15 17 10"/>
        <line x1="12" y1="15" x2="12" y2="3"/>
      </svg>
      Export to Excel
    </button>
  );
};

export default ExportToExcel;
