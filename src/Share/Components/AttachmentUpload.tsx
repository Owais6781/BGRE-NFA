import * as React from 'react';
import "../styles/global.scss";

interface AttachmentUploadProps {
  label: string;
  onChange: (files: FileList | null) => void;
  multiple?: boolean;
}

export const AttachmentUpload: React.FC<AttachmentUploadProps> = ({
  label, onChange, multiple = true,
}) => {
  return (
    <div className="flex items-center gap-2">
      <label className="formLabel mb-0">{label}</label>
      <label className="flex items-center gap-1 cursor-pointer">
        {/* paperclip icon */}
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none"
          stroke="#555" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
          <path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19
            a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48"/>
        </svg>
        <input
          type="file"
          multiple={multiple}
          onChange={(e) => onChange(e.target.files)}
          className="hidden"
        />
      </label>
    </div>
  );
};

export default AttachmentUpload;