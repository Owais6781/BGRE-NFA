import * as React from 'react';
import { Editor, EditorTextChangeEvent } from 'primereact/editor';

import "../styles/global.scss";

interface RichTextEditorProps {
  label:     string;
  value:     string;
  onChange:  (value: string) => void;
  required?: boolean;
  height?:   number;
  readOnly?: boolean;
  error?:    string;
}

export const RichTextEditor: React.FC<RichTextEditorProps> = ({
  label,
  value,
  onChange,
  required = false,
  height   = 200,
  readOnly = false,
  error,
}) => {
  return (
    <div className="flex flex-col gap-[6px] w-full">

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

      <Editor
        value={value || ''}
        onTextChange={(e: EditorTextChangeEvent) => onChange(e.htmlValue || '')}
        style={{ height }}
        readOnly={readOnly}
      />

    </div>
  );
};

export default RichTextEditor;