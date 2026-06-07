import * as React from 'react';
import { useState, useRef } from 'react';
import { X, Upload, FileText } from 'lucide-react';
import DropdownField from './DropdownField';
import CommonButton from './CommonButton';
import '../styles/global.scss';

// ── Static Document Types ──
const documentTypeOptions = [
  { label: 'Budget Doc', value: 'Budget Doc' },
  { label: 'Purposely Doc', value: 'Purposely Doc' },
  { label: 'Invoice', value: 'Invoice' },
  { label: 'Purchase Order', value: 'Purchase Order' },
  { label: 'Quotation', value: 'Quotation' },
  { label: 'Supporting Doc', value: 'Supporting Doc' },
];

// ── Interfaces ──
export interface IAttachmentEntry {
  documentType: string;
  file: File;
}

interface IAttachmentPopupProps {
  isOpen: boolean;
  onClose: () => void;
  onSave: (entries: IAttachmentEntry[]) => void;
}

const AttachmentPopup: React.FC<IAttachmentPopupProps> = ({
  isOpen,
  onClose,
  onSave,
}) => {
  const [documentType, setDocumentType] = useState<string>(documentTypeOptions[0].value);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [entries, setEntries] = useState<IAttachmentEntry[]>([]);
  const [error, setError] = useState<string>('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  // ── Reset Form ──
  const resetForm = (): void => {
    setDocumentType(documentTypeOptions[0].value);
    setSelectedFiles([]);
    setEntries([]);
    setError('');
  };

  // ── Handle File Selection ──
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    if (e.target.files && e.target.files.length > 0) {
      const filesArray = Array.from(e.target.files);
      setSelectedFiles(prev => [...prev, ...filesArray]);
      setError('');
    }
    // Reset input so same file can be selected again
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  // ── Remove Selected File ──
  const removeFile = (index: number): void => {
    setSelectedFiles(prev => prev.filter((_, i) => i !== index));
  };

  // ── Add to Entries List ──
  const handleAddEntry = (): void => {
    if (!documentType) {
      setError('Please select a Document Type.');
      return;
    }
    if (selectedFiles.length === 0) {
      setError('Please attach at least one file.');
      return;
    }

    const newEntries: IAttachmentEntry[] = selectedFiles.map(file => ({
      documentType,
      file,
    }));

    setEntries(prev => [...prev, ...newEntries]);
    setDocumentType('');
    setSelectedFiles([]);
    setError('');
  };

  // ── Remove Entry ──
  const removeEntry = (index: number): void => {
    setEntries(prev => prev.filter((_, i) => i !== index));
  };

  // ── Save & Close ──
  const handleSave = (): void => {
    if (entries.length === 0) {
      setError('Please add at least one document before saving.');
      return;
    }
    onSave(entries);
    resetForm();
  };

  // ── Exit ──
  const handleExit = (): void => {
    resetForm();
    onClose();
  };

  // ── Format File Size ──
  const formatSize = (bytes: number): string => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1048576) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / 1048576).toFixed(1)} MB`;
  };

  if (!isOpen) return null;

  return (
    <div className="popupOverlay">
      <div className="popupContainer">

        {/* ── Header ── */}
        <div className="popupHeader">
          <div className="popupHeader__title">
            <Upload size={16} />
            <span>Upload Documents</span>
          </div>
          <button className="popupHeader__close" onClick={handleExit}>
            <X size={18} />
          </button>
        </div>

        {/* ── Body ── */}
        <div className="popupBody">

          {/* Document Type + File Select Row */}
          <div className="flex flex-wrap gap-x-4 gap-y-4 mb-4">
            <div className="flex-1 min-w-[200px]">
              <DropdownField
                label="Document Type"
                value={documentType}
                onChange={(v: string) => {
                  setDocumentType(v);
                  setError('');
                }}
                options={documentTypeOptions}
                required
                showEmptyOption={false}
              />
            </div>
            <div className="flex-1 min-w-[200px]">
              <label className="fieldLabel">Attachment</label>
              <div className="fileUploadWrapper">
                <input
                  type="file"
                  ref={fileInputRef}
                  multiple
                  onChange={handleFileChange}
                  className="fileInput"
                  id="popupFileInput"
                />
                <label htmlFor="popupFileInput" className="fileUploadBtn">
                  <Upload size={14} />
                  <span>Browse Files</span>
                </label>
              </div>
            </div>
          </div>

          {/* ── Selected Files Preview ── */}
          {selectedFiles.length > 0 && (
            <div className="selectedFilesPreview mb-4">
              <p className="fieldLabel">Selected Files:</p>
              <div className="fileChipList">
                {selectedFiles.map((file, idx) => (
                  <div key={idx} className="fileChip">
                    <FileText size={12} />
                    <span className="fileChip__name">{file.name}</span>
                    <span className="fileChip__size">({formatSize(file.size)})</span>
                    <button
                      className="fileChip__remove"
                      onClick={() => removeFile(idx)}
                    >
                      <X size={12} />
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* ── Add Button ── */}
          <div className="flex justify-end mb-4">
            <CommonButton
              text="+ Add to List"
              variant="secondary"
              onClick={handleAddEntry}
            />
          </div>

          {/* ── Error Message ── */}
          {error && <p className="popupError">{error}</p>}

          {/* ── Entries Table ── */}
          {entries.length > 0 && (
            <div className="popupTableWrapper">
              <table className="popupTable">
                <thead>
                  <tr>
                    <th>#</th>
                    <th>Document Type</th>
                    <th>File Name</th>
                    <th>Size</th>
                    <th>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {entries.map((entry, idx) => (
                    <tr key={idx}>
                      <td>{idx + 1}</td>
                      <td>{entry.documentType}</td>
                      <td>{entry.file.name}</td>
                      <td>{formatSize(entry.file.size)}</td>
                      <td>
                        <button
                          className="removeBtn"
                          onClick={() => removeEntry(idx)}
                        >
                          <X size={14} />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>

        {/* ── Footer ── */}
        <div className="popupFooter">
          <CommonButton text="Save Document" variant="primary" onClick={handleSave} />
          <CommonButton text="Exit" variant="danger" onClick={handleExit} />
        </div>

      </div>
    </div>
  );
};

export default AttachmentPopup;