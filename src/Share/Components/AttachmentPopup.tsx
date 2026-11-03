import * as React from 'react';
import { useState, useRef, useEffect, useMemo } from 'react';
import { X, Upload, FileText, Info } from 'lucide-react';
import CommonButton from './CommonButton';
import AttachmentValidationModal, {
  validateAttachmentFiles,
  IAttachmentValidationError
} from './AttachmentValidationModal';
import '../styles/global.scss';

// ── Standard Attachment Types ──
export const ATTACHMENT_TYPES: string[] = [
  'Commercial Proposal',
  'Technical Proposal',
  'Comparative Statement',
  'Quotation',
  'Purchase Order / Work Order',
  'Contract / Agreement',
  'Budget Approval',
  'Justification Note',
  'Invoice ',
  'Scope of Work (SOW)',
  'Bill of Quantities (BOQ)',
  'Drawings / Specifications',
  'Compliance / Certificates',
  'Annexure',
  'Supporting Document',
  'Others'
];

// ── Interfaces ──
export interface IAttachmentEntry {
  documentType?: string;
  file: File;
}

export interface IExistingDocument {
  id?: number;
  fileName: string;
  documentType: string;
  isPending?: boolean;
  pendingIndex?: number;
}

interface IAttachmentPopupProps {
  isOpen: boolean;
  onClose: () => void;
  onSave: (entries: IAttachmentEntry[]) => void;
  existingDocument?: IExistingDocument | null;
  onUpdateExisting?: (newType: string) => Promise<void> | void;
  existingFileNames?: string[];
}

const AttachmentPopup: React.FC<IAttachmentPopupProps> = ({
  isOpen,
  onClose,
  onSave,
  existingDocument,
  onUpdateExisting,
  existingFileNames = [],
}) => {
  const [selectedType, setSelectedType] = useState<string>('');
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [entries, setEntries] = useState<IAttachmentEntry[]>([]);
  const [error, setError] = useState<string>('');
  const [isUpdating, setIsUpdating] = useState<boolean>(false);
  const [validationErrors, setValidationErrors] = useState<IAttachmentValidationError[]>([]);
  const [showValidationModal, setShowValidationModal] = useState<boolean>(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const isEditMode = Boolean(existingDocument);

  // Dynamic type options that include existing type if custom
  const typeOptions = useMemo(() => {
    const types = [...ATTACHMENT_TYPES];
    if (existingDocument?.documentType && !types.includes(existingDocument.documentType)) {
      types.unshift(existingDocument.documentType);
    }
    return types;
  }, [existingDocument]);

  // Synchronize state when opened or document changes
  useEffect(() => {
    if (isOpen) {
      if (existingDocument) {
        setSelectedType(existingDocument.documentType || '');
      } else {
        setSelectedType('');
      }
      setSelectedFiles([]);
      setEntries([]);
      setError('');
      setIsUpdating(false);
      setValidationErrors([]);
      setShowValidationModal(false);
    }
  }, [isOpen, existingDocument]);

  // ── Reset Form ──
  const resetForm = (): void => {
    setSelectedType('');
    setSelectedFiles([]);
    setEntries([]);
    setError('');
    setIsUpdating(false);
    setValidationErrors([]);
    setShowValidationModal(false);
  };

  // ── Process Incoming Files with Validation ──
  const handleIncomingFiles = (incomingFiles: File[]): void => {
    if (!incomingFiles || incomingFiles.length === 0) return;

    const allExistingNames = [
      ...(existingFileNames || []),
      ...selectedFiles.map(f => f.name),
      ...entries.map(e => e.file.name)
    ];

    const { validFiles, errors } = validateAttachmentFiles(incomingFiles, allExistingNames);

    if (errors.length > 0) {
      setValidationErrors(errors);
      setShowValidationModal(true);
    }

    if (validFiles.length > 0) {
      setSelectedFiles(prev => [...prev, ...validFiles]);
      setError('');
    }
  };

  // ── Handle File Selection ──
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    if (e.target.files && e.target.files.length > 0) {
      const filesArray = Array.from(e.target.files);
      handleIncomingFiles(filesArray);
    }
    // Reset input so same file can be selected again
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  // ── Handle Clipboard Paste ──
  const handlePaste = (e: React.ClipboardEvent): void => {
    if (e.clipboardData && e.clipboardData.files && e.clipboardData.files.length > 0) {
      e.preventDefault();
      e.stopPropagation();
      const filesArray = Array.from(e.clipboardData.files);
      handleIncomingFiles(filesArray);
    }
  };

  // ── Handle Drag & Drop ──
  const handleDrop = (e: React.DragEvent): void => {
    e.preventDefault();
    e.stopPropagation();
    if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const filesArray = Array.from(e.dataTransfer.files);
      handleIncomingFiles(filesArray);
    }
  };

  const handleDragOver = (e: React.DragEvent): void => {
    e.preventDefault();
    e.stopPropagation();
  };

  // ── Remove Selected File ──
  const removeFile = (index: number): void => {
    setSelectedFiles(prev => prev.filter((_, i) => i !== index));
  };

  // ── Add to Entries List ──
  const handleAddEntry = (): void => {
    if (!selectedType || !selectedType.trim()) {
      setError('Please select a Document Type.');
      return;
    }

    if (selectedFiles.length === 0) {
      setError('Please attach at least one file.');
      return;
    }

    const newEntries: IAttachmentEntry[] = selectedFiles.map(file => ({
      documentType: selectedType,
      file,
    }));

    setEntries(prev => [...prev, ...newEntries]);
    setSelectedFiles([]);
    setError('');
  };

  // ── Remove Entry ──
  const removeEntry = (index: number): void => {
    setEntries(prev => prev.filter((_, i) => i !== index));
  };

  // ── Save & Close for New Uploads ──
  const handleSave = (): void => {
    if (entries.length === 0) {
      setError('Please add at least one document before saving.');
      return;
    }
    onSave(entries);
    resetForm();
  };

  // ── Save Existing Document Type Update ──
  const handleUpdateExisting = async (): Promise<void> => {
    if (!selectedType || !selectedType.trim()) {
      setError('Please select a Document Type.');
      return;
    }

    if (onUpdateExisting) {
      try {
        setIsUpdating(true);
        await onUpdateExisting(selectedType);
        resetForm();
      } catch (err) {
        console.error('Error updating document type:', err);
        setError('Failed to update attachment type.');
      } finally {
        setIsUpdating(false);
      }
    }
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
      <div
        className="popupContainer"
        onPaste={handlePaste}
        onDrop={handleDrop}
        onDragOver={handleDragOver}
      >

        {/* ── Header ── */}
        <div className="popupHeader">
          <div className="popupHeader__title">
            {isEditMode ? <FileText size={16} /> : <Upload size={16} />}
            <span>{isEditMode ? 'Edit Attachment Type' : 'Upload Documents'}</span>
          </div>
          <button className="popupHeader__close" onClick={handleExit} type="button">
            <X size={18} />
          </button>
        </div>

        {/* ── Body ── */}
        <div className="popupBody">

          {/* Edit Mode: Display Existing File Name (left) and Document Type (right) in ONE ROW */}
          {isEditMode && existingDocument && (
            <div
              className="popupRow"
              style={{
                display: 'flex',
                flexDirection: 'row',
                alignItems: 'flex-end',
                gap: '16px',
                width: '100%',
                marginBottom: '16px',
                boxSizing: 'border-box'
              }}
            >
              {/* Left Column: File Name */}
              <div className="popupCol" style={{ flex: '1 1 50%', minWidth: 0 }}>
                <label className="fieldLabel" style={{ display: 'block', marginBottom: '6px' }}>File Name</label>
                <div
                  className="flex items-center gap-2 p-2.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-semibold text-slate-700"
                  style={{ height: '38px', boxSizing: 'border-box', overflow: 'hidden' }}
                >
                  <FileText size={14} className="text-blue-600 shrink-0" />
                  <span className="truncate">{existingDocument.fileName}</span>
                </div>
              </div>

              {/* Right Column: Document Type Dropdown */}
              <div className="popupCol" style={{ flex: '1 1 50%', minWidth: 0 }}>
                <label className="fieldLabel flex items-center gap-1.5" style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', marginBottom: '6px' }}>
                  Document Type <span className="text-red-500">*</span>
                </label>
                <select
                  value={selectedType}
                  onChange={(e) => {
                    setSelectedType(e.target.value);
                    setError('');
                  }}
                  className="formInput formSelect"
                  style={{ width: '100%', height: '38px', boxSizing: 'border-box' }}
                >
                  <option value="">-- Select Type --</option>
                  {typeOptions.map((type) => (
                    <option key={type} value={type}>
                      {type}
                    </option>
                  ))}
                </select>
              </div>
            </div>
          )}

          {/* Add Mode: Document Type (left) and Attachment (right) side-by-side in ONE ROW */}
          {!isEditMode && (
            <>
              {/* Attachment Guidelines Banner */}
              <div
                style={{
                  display: 'flex',
                  alignItems: 'flex-start',
                  gap: '8px',
                  padding: '10px 12px',
                  backgroundColor: 'rgba(254, 242, 242, 0.7)',
                  border: '1px solid #fee2e2',
                  borderRadius: '8px',
                  fontSize: '12px',
                  color: '#334155',
                  marginBottom: '16px'
                }}
              >
                <Info size={15} style={{ color: '#d97706', flexShrink: 0, marginTop: '2px' }} />
                <div style={{ lineHeight: '1.45' }}>
                  <strong style={{ color: '#0f172a' }}>Attachment Guidelines:</strong> Max 10 MB per file · File names cannot contain:{' '}
                  <span
                    style={{
                      fontFamily: 'monospace',
                      color: '#b91c1c',
                      backgroundColor: '#fff',
                      padding: '1px 5px',
                      borderRadius: '4px',
                      border: '1px solid #fecaca',
                      fontSize: '11px'
                    }}
                  >
                    ~ &quot; # % &amp; * : &lt; &gt; ? / \ &#123; | &#125; +
                  </span>
                </div>
              </div>

              <div
                className="popupRow"
                style={{
                  display: 'flex',
                  flexDirection: 'row',
                  alignItems: 'flex-end',
                  gap: '16px',
                  width: '100%',
                  marginBottom: '16px',
                  boxSizing: 'border-box'
                }}
              >
                {/* Left Column: Document Type Dropdown */}
                <div className="popupCol" style={{ flex: '1 1 50%', minWidth: 0 }}>
                  <label className="fieldLabel flex items-center gap-1.5" style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', marginBottom: '6px' }}>
                    Document Type <span className="text-red-500">*</span>
                  </label>
                  <select
                    value={selectedType}
                    onChange={(e) => {
                      setSelectedType(e.target.value);
                      setError('');
                    }}
                    className="formInput formSelect"
                    style={{ width: '100%', height: '38px', boxSizing: 'border-box' }}
                  >
                    <option value="">-- Select Type --</option>
                    {typeOptions.map((type) => (
                      <option key={type} value={type}>
                        {type}
                      </option>
                    ))}
                  </select>
                </div>

                {/* Right Column: Attachment File Select */}
                <div className="popupCol" style={{ flex: '1 1 50%', minWidth: 0 }}>
                  <label className="fieldLabel flex items-center gap-1.5" style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', marginBottom: '6px' }}>
                    Attachment <span className="text-red-500">*</span>
                  </label>
                  <div className="fileUploadWrapper" style={{ width: '100%' }}>
                    <input
                      type="file"
                      ref={fileInputRef}
                      multiple
                      onChange={handleFileChange}
                      className="fileInput"
                      id="popupFileInput"
                    />
                    <label
                      htmlFor="popupFileInput"
                      className="fileUploadBtn cursor-pointer"
                      style={{
                        width: '100%',
                        height: '38px',
                        boxSizing: 'border-box',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        gap: '8px'
                      }}
                    >
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
                          type="button"
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
                        <th>Type</th>
                        <th>File Name</th>
                        <th>Size</th>
                        <th>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {entries.map((entry, idx) => (
                        <tr key={idx}>
                          <td>{idx + 1}</td>
                          <td>
                            <span className="inline-block px-2 py-0.5 rounded text-xs font-bold bg-slate-100 text-slate-700">
                              {entry.documentType || 'Attachment'}
                            </span>
                          </td>
                          <td>{entry.file.name}</td>
                          <td>{formatSize(entry.file.size)}</td>
                          <td>
                            <button
                              type="button"
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
            </>
          )}

          {/* Edit Mode Error Message */}
          {isEditMode && error && <p className="popupError">{error}</p>}

        </div>

        {/* ── Footer ── */}
        <div className="popupFooter">
          {isEditMode ? (
            <>
              <CommonButton
                text={isUpdating ? "Updating..." : "Update Type"}
                variant="primary"
                disabled={isUpdating}
                onClick={handleUpdateExisting}
              />
              <CommonButton text="Cancel" variant="danger" onClick={handleExit} />
            </>
          ) : (
            <>
              <CommonButton text="Save Document" variant="primary" onClick={handleSave} />
              <CommonButton text="Exit" variant="danger" onClick={handleExit} />
            </>
          )}
        </div>

      </div>

      {/* ── Attachment Validation Modal ── */}
      <AttachmentValidationModal
        isOpen={showValidationModal}
        errors={validationErrors}
        onClose={() => setShowValidationModal(false)}
      />
    </div>
  );
};

export default AttachmentPopup;
