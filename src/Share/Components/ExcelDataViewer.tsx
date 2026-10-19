import * as React from 'react';
import { useState, useMemo } from 'react';
import { Copy, Check, Maximize2, X, FileSpreadsheet } from 'lucide-react';
import '../styles/global.scss';

export interface ExcelDataViewerProps {
  html?: string;
  label?: string;
  className?: string;
}

interface IParsedExcelData {
  hasTable: boolean;
  rowCount?: number;
  colCount?: number;
  tableHtml?: string;
  textBefore?: string;
  textAfter?: string;
  tsv?: string;
  rawHtml?: string;
}

export const ExcelDataViewer: React.FC<ExcelDataViewerProps> = ({
  html,
  label = 'Excel Data',
  className = ''
}) => {
  const [isExpanded, setIsExpanded] = useState<boolean>(false);
  const [copied, setCopied] = useState<boolean>(false);

  // Parse HTML content to detect <table> or tab-separated text
  const parsedData = useMemo<IParsedExcelData | null>(() => {
    if (!html || html.trim() === '' || html === '-') {
      return null;
    }

    // 1. Check if HTML contains <table>
    if (html.includes('<table')) {
      try {
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');
        const tables = doc.querySelectorAll('table');

        if (tables.length > 0) {
          const mainTable = tables[0];

          // Calculate row & col counts
          const trs = Array.from(mainTable.querySelectorAll('tr'));
          const rowCount = trs.length;
          let maxCols = 0;
          trs.forEach(tr => {
            let cols = 0;
            Array.from(tr.querySelectorAll('th, td')).forEach(c => {
              cols += parseInt(c.getAttribute('colspan') || '1', 10);
            });
            if (cols > maxCols) maxCols = cols;
          });

          // Build TSV for clipboard copy
          const tsvRows: string[] = [];
          trs.forEach(tr => {
            const cellTexts = Array.from(tr.querySelectorAll('th, td')).map(c =>
              (c.textContent || '').trim().replace(/\t/g, ' ')
            );
            tsvRows.push(cellTexts.join('\t'));
          });
          const tsv = tsvRows.join('\n');

          // Enhance the table DOM classes
          mainTable.classList.add('excelDataTable');

          trs.forEach((tr, rIdx) => {
            const cells = Array.from(tr.querySelectorAll('th, td'));
            if (cells.length > 0) {
              // Pin first column (e.g. Tax Invoice Number)
              cells[0].classList.add('excelCellPinned');
            }
            if (rIdx === 0 || tr.parentElement?.tagName.toLowerCase() === 'thead') {
              tr.classList.add('excelHeaderRow');
            }
            cells.forEach(cell => {
              const text = cell.textContent?.trim() || '';
              // Detect numeric/currency values: e.g. 14,859 or ₹17,533 or 0 or -50.25 or 12%
              // Ignore standard date formats like DD-MM-YYYY or DD/MM/YYYY
              if (/^[-+]?₹?\s*[\d,]+(\.\d+)?%?$/.test(text) && !/^\d{2}[-/]\d{2}[-/]\d{4}$/.test(text)) {
                cell.classList.add('excelCellNumeric');
              }
            });
          });

          // Extract any text before or after the table
          const textBefore: string[] = [];
          const textAfter: string[] = [];
          let seenTable = false;
          Array.from(doc.body.children).forEach(child => {
            if (child === mainTable) {
              seenTable = true;
            } else if (!seenTable) {
              textBefore.push(child.outerHTML);
            } else {
              textAfter.push(child.outerHTML);
            }
          });

          return {
            hasTable: true,
            rowCount,
            colCount: maxCols,
            tableHtml: mainTable.outerHTML,
            textBefore: textBefore.join(''),
            textAfter: textAfter.join(''),
            tsv
          };
        }
      } catch (err) {
        console.error('Error parsing HTML table in ExcelDataViewer:', err);
      }
    }

    // 2. Check if raw tab-separated text (pasted into a plain textarea)
    if (html.includes('\t') && html.includes('\n')) {
      const lines = html.trim().split('\n').filter(l => l.trim().length > 0);
      if (lines.length > 1) {
        const rows = lines.map(line => line.split('\t').map(c => c.trim()));
        const rowCount = rows.length;
        const colCount = Math.max(...rows.map(r => r.length));

        const tableHtml = `
          <table class="excelDataTable">
            <thead>
              <tr class="excelHeaderRow">
                ${rows[0].map((h, i) => `<th class="${i === 0 ? 'excelCellPinned' : ''}">${h}</th>`).join('')}
              </tr>
            </thead>
            <tbody>
              ${rows.slice(1).map(r => `
                <tr>
                  ${r.map((c, i) => {
                    const isNum = /^[-+]?₹?\s*[\d,]+(\.\d+)?%?$/.test(c) && !/^\d{2}[-/]\d{2}[-/]\d{4}$/.test(c);
                    const classes = [i === 0 ? 'excelCellPinned' : '', isNum ? 'excelCellNumeric' : ''].filter(Boolean).join(' ');
                    return `<td class="${classes}">${c}</td>`;
                  }).join('')}
                </tr>
              `).join('')}
            </tbody>
          </table>
        `;

        return {
          hasTable: true,
          rowCount,
          colCount,
          tableHtml,
          textBefore: '',
          textAfter: '',
          tsv: html
        };
      }
    }

    // 3. Regular rich text without table
    return {
      hasTable: false,
      rawHtml: html
    };
  }, [html]);

  // Handle copying TSV data to clipboard
  const handleCopy = (e: React.MouseEvent) => {
    e.stopPropagation();
    if (parsedData?.tsv) {
      navigator.clipboard.writeText(parsedData.tsv).then(() => {
        setCopied(true);
        setTimeout(() => setCopied(false), 2000);
      }).catch(err => {
        console.error('Failed to copy to clipboard', err);
      });
    }
  };

  // Close modal on Escape key
  React.useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && isExpanded) {
        setIsExpanded(false);
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isExpanded]);

  // If no content
  if (!parsedData) {
    return <span className="text-slate-400 italic text-sm">-</span>;
  }

  // If regular rich text without tables
  if (!parsedData.hasTable) {
    return (
      <div className={`richTextContent text-sm text-slate-700 ${className}`}>
        <div dangerouslySetInnerHTML={{ __html: parsedData.rawHtml || '-' }} />
      </div>
    );
  }

  const {
    rowCount = 0,
    colCount = 0,
    tableHtml = '',
    textBefore = '',
    textAfter = ''
  } = parsedData;

  // Render Excel table grid
  return (
    <div className={`flex flex-col gap-2 w-full ${className}`}>
      {/* Any leading text before table */}
      {textBefore ? (
        <div 
          className="richTextContent text-sm text-slate-700 mb-2"
          dangerouslySetInnerHTML={{ __html: textBefore }} 
        />
      ) : null}

      {/* Dedicated Excel Data Card */}
      <div className="excelViewerCard">
        {/* Table Toolbar */}
        <div className="excelViewerToolbar">
          <div className="flex items-center gap-2.5 min-w-0">
            <div className="w-6 h-6 rounded-md bg-emerald-50 text-emerald-600 flex items-center justify-center shrink-0 border border-emerald-100">
              <FileSpreadsheet size={14} />
            </div>
            <span className="excelViewerTitle truncate">{label}</span>
            <span className="excelViewerBadge">
              {rowCount} {rowCount === 1 ? 'Row' : 'Rows'} • {colCount} Columns
            </span>
          </div>

          <div className="flex items-center gap-1.5 shrink-0">
            <button
              type="button"
              onClick={handleCopy}
              className="excelViewerBtn"
              title="Copy table to clipboard for Excel"
            >
              {copied ? <Check size={13} className="text-emerald-600" /> : <Copy size={13} />}
              <span>{copied ? 'Copied!' : 'Copy'}</span>
            </button>
            <button
              type="button"
              onClick={() => setIsExpanded(true)}
              className="excelViewerBtn"
              title="Expand table to full screen"
            >
              <Maximize2 size={13} />
              <span>Expand</span>
            </button>
          </div>
        </div>

        {/* Scrollable Table Viewport */}
        <div className="excelTableScrollContainer">
          <div dangerouslySetInnerHTML={{ __html: tableHtml }} />
        </div>

        {/* Scroll Hint / Footer */}
        {colCount > 5 ? (
          <div className="excelScrollHint">
            <span className="flex items-center gap-1 text-[11px] text-slate-400">
              ← Scroll horizontally to view all {colCount} columns →
            </span>
            <span className="text-[11px] font-medium text-emerald-600 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-100/60">
              100% Data Preserved
            </span>
          </div>
        ) : null}
      </div>

      {/* Any trailing text after table */}
      {textAfter ? (
        <div 
          className="richTextContent text-sm text-slate-700 mt-2"
          dangerouslySetInnerHTML={{ __html: textAfter }} 
        />
      ) : null}

      {/* ── Full Screen / Expanded Modal View ── */}
      {isExpanded && (
        <div 
          className="fixed inset-0 z-[99999] bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-3 sm:p-6"
          onClick={() => setIsExpanded(false)}
        >
          <div 
            className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-[96vw] h-[92vh] flex flex-col overflow-hidden"
            onClick={e => e.stopPropagation()}
          >
            {/* Modal Header */}
            <div className="flex items-center justify-between px-6 py-3.5 border-b border-slate-200 bg-slate-50 shrink-0">
              <div className="flex items-center gap-3">
                <div className="w-8 h-8 rounded-lg bg-emerald-50 text-emerald-600 flex items-center justify-center shrink-0 border border-emerald-100">
                  <FileSpreadsheet size={18} />
                </div>
                <div>
                  <h3 className="font-bold text-slate-800 text-sm">{label} — Full View</h3>
                  <span className="text-xs text-slate-400">
                    {rowCount} {rowCount === 1 ? 'Row' : 'Rows'} • {colCount} Columns
                  </span>
                </div>
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={handleCopy}
                  className="excelViewerBtn"
                  title="Copy table to clipboard for Excel"
                >
                  {copied ? <Check size={14} className="text-emerald-600" /> : <Copy size={14} />}
                  <span>{copied ? 'Copied to Clipboard!' : 'Copy Data'}</span>
                </button>
                <button
                  type="button"
                  onClick={() => setIsExpanded(false)}
                  className="excelViewerCloseBtn"
                  title="Close full view"
                >
                  <X size={16} />
                  <span>Close</span>
                </button>
              </div>
            </div>

            {/* Modal Scrollable Table Body */}
            <div className="flex-1 overflow-auto p-4 bg-slate-50/50">
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <div className="excelTableScrollContainer !max-h-[calc(92vh-130px)]">
                  <div dangerouslySetInnerHTML={{ __html: tableHtml }} />
                </div>
              </div>
            </div>

            {/* Modal Footer */}
            <div className="px-6 py-2 bg-white border-t border-slate-100 flex items-center justify-between text-xs text-slate-400 shrink-0">
              <span>Use mouse wheel or scrollbar to navigate rows and columns</span>
              <span>Press Escape or click Close to return</span>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default ExcelDataViewer;
