import * as React from 'react';

export interface IAttachmentValidationError {
  fileName: string;
  message: string;
}

export const MAX_ATTACHMENT_SIZE_BYTES = 10 * 1024 * 1024; // 10 MB

/**
 * Validates a list of files against SharePoint attachment constraints:
 * 1. Duplicate check (cannot attach files with duplicate names).
 * 2. File size must not exceed 10 MB.
 * 3. File names must not contain prohibited characters: ~ " # % & * : < > ? / \ { | } +
 */
export const validateAttachmentFiles = (
  files: File[],
  existingFileNames: string[] = []
): { validFiles: File[]; errors: IAttachmentValidationError[] } => {
  const validFiles: File[] = [];
  const errors: IAttachmentValidationError[] = [];
  const seenLowerNames = new Set<string>(
    existingFileNames.filter(Boolean).map((name) => name.trim().toLowerCase())
  );

  for (const file of files) {
    const lowerName = file.name.trim().toLowerCase();

    // 1. Duplicate check
    if (seenLowerNames.has(lowerName)) {
      errors.push({
        fileName: file.name,
        message: `A file named "${file.name}" is already attached. Duplicate attachments are not allowed.`
      });
      continue;
    }

    // 2. File size check
    if (file.size > MAX_ATTACHMENT_SIZE_BYTES) {
      const sizeMB = (file.size / (1024 * 1024)).toFixed(2);
      errors.push({
        fileName: file.name,
        message: `File "${file.name}" exceeds the maximum allowed size of 10 MB (${sizeMB} MB). Please select a file under 10 MB.`
      });
      continue;
    }

    // 3. Prohibited characters check
    if (/[~"#%&*:<>?/\\{}|+]/.test(file.name)) {
      errors.push({
        fileName: file.name,
        message: `File "${file.name}" contains invalid characters (~ " # % & * : < > ? / \\ { | } +). Please rename the file before attaching.`
      });
      continue;
    }

    seenLowerNames.add(lowerName);
    validFiles.push(file);
  }

  return { validFiles, errors };
};

interface AttachmentValidationModalProps {
  isOpen: boolean;
  errors: IAttachmentValidationError[];
  onClose: () => void;
}

export const AttachmentValidationModal: React.FC<AttachmentValidationModalProps> = ({
  isOpen,
  errors,
  onClose
}) => {
  if (!isOpen || errors.length === 0) return null;

  return (
    <div
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
        zIndex: 999999
      }}
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        style={{
          backgroundColor: '#ffffff',
          borderRadius: '16px',
          padding: '28px 32px',
          maxWidth: '560px',
          width: '90%',
          boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04)',
          boxSizing: 'border-box'
        }}
      >
        {/* Header */}
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
            Attachment Validation
          </h3>
        </div>

        {/* Subtitle */}
        <p
          style={{
            fontSize: '13px',
            fontWeight: 600,
            color: '#334155',
            marginTop: '16px',
            marginBottom: '10px'
          }}
        >
          {errors.length} file(s) could not be attached:
        </p>

        {/* Bullet List */}
        <ul
          style={{
            margin: '0 0 24px 0',
            paddingLeft: '18px',
            listStyleType: 'disc',
            display: 'flex',
            flexDirection: 'column',
            gap: '8px'
          }}
        >
          {errors.map((err, idx) => (
            <li
              key={idx}
              style={{
                fontSize: '13px',
                lineHeight: '1.5',
                color: '#475569'
              }}
            >
              {err.message}
            </li>
          ))}
        </ul>

        {/* OK Button */}
        <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
          <button
            type="button"
            onClick={onClose}
            style={{
              backgroundColor: '#5d1020',
              color: '#ffffff',
              fontWeight: 700,
              fontSize: '13px',
              padding: '8px 28px',
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
  );
};

export default AttachmentValidationModal;
