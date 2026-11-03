import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import { Editor, EditorTextChangeEvent } from 'primereact/editor';

import "../styles/global.scss";

interface RichTextEditorProps {
  label: string;
  value: string;
  placeholder?: string;
  onChange: (value: string) => void;
  required?: boolean;
  height?: number;
  readOnly?: boolean;
  error?: string;
  maxTableColumns?: number;
}

/**
 * Counts the maximum number of columns in any table found within HTML or plain text.
 */
export const getTableMaxColumns = (htmlOrText: string): number => {
  if (!htmlOrText) return 0;

  let maxCols = 0;

  // 1. Check HTML tables
  if (/<\s*table/i.test(htmlOrText)) {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(htmlOrText, 'text/html');
      const tables = doc.querySelectorAll('table');

      tables.forEach((table) => {
        const rows = table.querySelectorAll('tr');
        rows.forEach((row) => {
          let colsInRow = 0;
          const cells = Array.from(row.children).filter(
            (c) => c.tagName.toLowerCase() === 'th' || c.tagName.toLowerCase() === 'td'
          );
          cells.forEach((cell) => {
            const colspanAttr = cell.getAttribute('colspan');
            const colspan = colspanAttr ? parseInt(colspanAttr, 10) : 1;
            colsInRow += isNaN(colspan) || colspan < 1 ? 1 : colspan;
          });
          if (colsInRow > maxCols) {
            maxCols = colsInRow;
          }
        });
      });
    } catch (e) {
      console.warn('Error parsing HTML table:', e);
    }
  }

  // 2. Check tab-separated text (e.g. copied directly from Excel cells)
  if (htmlOrText.includes('\t')) {
    const lines = htmlOrText.split(/\r?\n/).filter((line) => line.trim().length > 0);
    for (const line of lines) {
      const colCount = line.split('\t').length;
      if (colCount > maxCols) {
        maxCols = colCount;
      }
    }
  }

  return maxCols;
};

export const RichTextEditor: React.FC<RichTextEditorProps> = ({
  label,
  value,
  placeholder,
  onChange,
  required = false,
  height = 200,
  readOnly = false,
  error,
  maxTableColumns = 8,
}) => {
  const [showValidationModal, setShowValidationModal] = useState<boolean>(false);
  const [exceededColumns, setExceededColumns] = useState<number>(0);
  const containerRef = useRef<HTMLDivElement>(null);
  const lastValidValueRef = useRef<string>(value || '');

  useEffect(() => {
    lastValidValueRef.current = value || '';
  }, [value]);

  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    const handlePaste = (e: ClipboardEvent): void => {
      const clipboardData = e.clipboardData;
      if (!clipboardData) return;

      const html = clipboardData.getData('text/html');
      const text = clipboardData.getData('text/plain');

      const colCountFromHtml = html ? getTableMaxColumns(html) : 0;
      const colCountFromText = text ? getTableMaxColumns(text) : 0;
      const maxDetectedCols = Math.max(colCountFromHtml, colCountFromText);

      if (maxDetectedCols > maxTableColumns) {
        e.preventDefault();
        e.stopPropagation();
        setExceededColumns(maxDetectedCols);
        setShowValidationModal(true);
      }
    };

    container.addEventListener('paste', handlePaste, true);
    return () => {
      container.removeEventListener('paste', handlePaste, true);
    };
  }, [maxTableColumns]);

  const handleTextChange = (e: EditorTextChangeEvent): void => {
    const newHtml = e.htmlValue || '';
    if (newHtml && (newHtml.includes('<table') || newHtml.includes('<TABLE'))) {
      const colCount = getTableMaxColumns(newHtml);
      if (colCount > maxTableColumns) {
        setExceededColumns(colCount);
        setShowValidationModal(true);
        // Do not update parent with invalid table exceeding column limit
        return;
      }
    }
    lastValidValueRef.current = newHtml;
    onChange(newHtml);
  };

  return (
    <div ref={containerRef} className="flex flex-col gap-[6px] w-full relative">
      <label
        className="formLabel flex items-center gap-1.5"
        style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}
      >
        {label}
        {required && <span className="text-red-500 ml-[2px]">*</span>}
        {error && (
          <span className="inline-flex h-1.5 w-1.5 relative" style={{ display: 'inline-flex' }}>
            <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-red-400 opacity-75"></span>
            <span className="relative inline-flex rounded-full h-1.5 w-1.5 bg-red-500"></span>
          </span>
        )}
      </label>

      <Editor
        value={value || ''}
        placeholder={placeholder}
        onTextChange={handleTextChange}
        style={{ height }}
        readOnly={readOnly}
      />
      {error && (
        <span className="text-red-500 text-[11px] font-medium mt-1 block">
          {error}
        </span>
      )}

      {/* ── Table Column Limit Validation Modal ── */}
      {showValidationModal && (
        <div
          className="fixed inset-0 flex items-center justify-center p-4"
          style={{
            position: 'fixed',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: 'rgba(0, 0, 0, 0.45)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 99999
          }}
          onClick={(e) => {
            if (e.target === e.currentTarget) setShowValidationModal(false);
          }}
        >
          <div
            style={{
              backgroundColor: '#ffffff',
              borderRadius: '16px',
              padding: '28px 32px',
              maxWidth: '520px',
              width: '90%',
              boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04)',
              boxSizing: 'border-box'
            }}
          >
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
              <svg
                xmlns="http://www.w3.org/2000/svg"
                viewBox="0 0 24 24"
                fill="none"
                stroke="#eab308"
                strokeWidth="2.5"
                strokeLinecap="round"
                strokeLinejoin="round"
                style={{ width: '22px', height: '22px', color: '#eab308', flexShrink: 0 }}
              >
                <path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z" />
                <line x1="12" y1="9" x2="12" y2="13" />
                <line x1="12" y1="17" x2="12.01" y2="17" />
              </svg>
              <h3
                style={{
                  fontSize: '18px',
                  fontWeight: 700,
                  color: '#6b1222',
                  margin: 0,
                  lineHeight: 1.2
                }}
              >
                Validation
              </h3>
            </div>

            <p
              style={{
                fontSize: '13px',
                lineHeight: '1.6',
                color: '#475569',
                marginTop: '16px',
                marginBottom: '24px'
              }}
            >
              The copied table contains {exceededColumns} columns, which exceeds the maximum limit of {maxTableColumns} columns for A4 Landscape document display. Wide spreadsheets cannot fit legibly on official approval notes. Please attach this spreadsheet in the Attachments section below, or reduce the table to {maxTableColumns} columns or fewer.
            </p>

            <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
              <button
                type="button"
                onClick={() => setShowValidationModal(false)}
                style={{
                  backgroundColor: '#5d1020',
                  color: '#ffffff',
                  fontWeight: 700,
                  fontSize: '13px',
                  padding: '8px 26px',
                  borderRadius: '6px',
                  border: 'none',
                  cursor: 'pointer',
                  transition: 'background-color 0.15s ease'
                }}
                onMouseOver={(e) => (e.currentTarget.style.backgroundColor = '#450a16')}
                onMouseOut={(e) => (e.currentTarget.style.backgroundColor = '#5d1020')}
              >
                OK
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default RichTextEditor;
