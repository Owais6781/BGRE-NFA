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
}

export const RichTextEditor: React.FC<RichTextEditorProps> = ({
  label,
  value,
  onChange,
  required = false,
  height   = 200,
  readOnly = false,
}) => {
  return (
    <div className="flex flex-col gap-[6px] w-full">

      <label className="formLabel">
        {label}
        {required && <span className="text-red-500 ml-[2px]">*</span>}
      </label>

      <Editor
        value={value}
        onTextChange={(e: EditorTextChangeEvent) => onChange(e.htmlValue || '')}
        style={{ height }}
        readOnly={readOnly}
      />

    </div>
  );
};

export default RichTextEditor;