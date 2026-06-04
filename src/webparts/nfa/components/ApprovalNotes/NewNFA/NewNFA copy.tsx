import * as React from 'react';
import { useState,useEffect  } from 'react';
import {
  Building2, FileText, Wrench, DollarSign,
  Users, Paperclip, Settings, Star,ChevronDown,Plus, Trash2 
} from 'lucide-react';
import { InputField }       from '../../../../../Share/Components/InputField';
import { TextAreaField }    from '../../../../../Share/Components/TextAreaField';
import { NumberInput }      from '../../../../../Share/Components/NumberInput';
import { AttachmentUpload } from '../../../../../Share/Components/AttachmentUpload';
import { RichTextEditor }   from '../../../../../Share/Components/RichTextEditor';
import DropdownField        from '../../../../../Share/Components/DropdownField';
import CommonButton         from '../../../../../Share/Components/CommonButton';
import "../../../../../Share/styles/global.scss";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { DropdownService, IDropdownOption } from '../../../../../Share/Services/DropdownService';
import MultiSelectDropdown from '../../../../../Share/Components/MultiSelectDropdown';
import AttachmentPopup, { IAttachmentEntry } from '../../../../../Share/Components/AttachmentPopup';
import { NFADocService, ISavedDocument } from '../../../../../Share/Services/NFADocService';

const approverOptions  = [{ label: 'User 1', value: 'user1' }, { label: 'User 2', value: 'user2' }, { label: 'User 3', value: 'user3' }];
const workflowOptions  = [{ label: 'MANUAL', value: 'MANUAL' }, { label: 'AUTO', value: 'AUTO' }];
interface CollapsibleSectionProps {
  icon: React.ReactNode;
  title: string;
  children: React.ReactNode;
  defaultOpen?: boolean;
}

const CollapsibleSection: React.FC<CollapsibleSectionProps> = ({
  icon, title, children, defaultOpen = true,
}) => {
  const [isOpen, setIsOpen] = useState(defaultOpen);

  return (
    <div className={`formSection ${isOpen ? 'formSection--open' : 'formSection--closed'}`}>
      <div className="formSectionHeader formSectionHeader--toggle"
        onClick={() => setIsOpen(!isOpen)}>
        <div className="formSectionHeader__left">{icon}<span>{title}</span></div>
        <div className={`formSectionHeader__chevron ${isOpen ? 'chevron--open' : 'chevron--closed'}`}>
          <ChevronDown size={16} />
        </div>
      </div>
      <div className={`formSectionBody--collapsible ${isOpen ? 'formSectionBody--expanded' : 'formSectionBody--collapsed'}`}>
        <div className="formSectionBody">{children}</div>
      </div>
    </div>
  );
};




// ── Props Interface ──
interface INewNFAProps {
  context: WebPartContext;
  sp: SPFI;
}

const NewNFA: React.FC<INewNFAProps> = ({ context, sp }) => {
  
  const today = new Date().toLocaleDateString('en-IN');

 
  const [form, setForm] = useState({
    originatorDept: '', date: today, project: '',  projectSubEntity: [] as string[],
    referenceSubject: '', introduction: '', chronology: '',
    commercialProposal: '', paymentTerms: '',
    liquidatedDamage: 'NA', contractPerf: 'NA',
    defectLiability: 'NA', workCompletion: 'Work must be completed and handed over to the satisfaction of Owner on or before -------',
    budgetRemarks: '', packageAwardedCost: '0.00', budget: '0.00',
    capexSerialNo: '', recommendation: '', contractorName: '',
    annexure: '', note1: '', note2: '', note3: '', note4: '',
    workflow: 'MANUAL', doaMatrix: '',
    approvers: Array(12).fill('') as string[],
    attachments: null as FileList | null,
  });

  const set = (key: string, val: any) => setForm(f => ({ ...f, [key]: val }));
  const setApprover = (i: number, val: string) => {
    const arr = [...form.approvers]; arr[i] = val;
    setForm(f => ({ ...f, approvers: arr }));
  };



   // ── Dropdown Options State  and Fetch Dropdown Data on Component Mount   ──
  const [projectOptions, setProjectOptions] = useState<IDropdownOption[]>([]);
  const [loadingProjects, setLoadingProjects] = useState<boolean>(true);
  const [subEntityOptions, setSubEntityOptions] = useState<IDropdownOption[]>([]);
  const [loadingSubEntity, setLoadingSubEntity] = useState<boolean>(false);
  const [functionOptions, setFunctionOptions] = useState<IDropdownOption[]>([]);
  const [loadingFunction, setLoadingFunction] = useState<boolean>(true);

  const [isPopupOpen, setIsPopupOpen] = useState<boolean>(false);
  const [savedDocuments, setSavedDocuments] = useState<ISavedDocument[]>([]);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [uniqId] = useState<string>(() => {
  const service = new NFADocService(context);
  return service.generateUniqId();
});
  
  useEffect(() => {
  const fetchDropdowns = async () => {
    try {
      const service = new DropdownService(context);

      // Fetch Project 
      const projects = await service.getProjects();
      setProjectOptions(projects);

      // Fetch Function Master ──
      const functions = await service.getFunctionMaster();
      setFunctionOptions(functions);

    } catch (error) {
      console.error("Error loading dropdown data:", error);
    } finally {
      setLoadingProjects(false);
      setLoadingFunction(false);  // ← ADD THIS
    }
  };

  fetchDropdowns();
}, [context]);


   // ──  THIS NEW useEffect Bind SubEntity ──
  const { project } = form;
useEffect(() => {
  const fetchSubEntities = async () => {
    setForm(f => ({ ...f, projectSubEntity: [] }));
    setSubEntityOptions([]);

    if (!form.project) {
      return;
    }

    setLoadingSubEntity(true);
    try {
      const service = new DropdownService(context);
      const subEntities = await service.getSubEntity(form.project);
      setSubEntityOptions(subEntities);
    } catch (error) {
      console.error("Error loading SubEntity data:", error);
    } finally {
      setLoadingSubEntity(false);
    }
  };

  fetchSubEntities();
}, [project, context]);

// ── THIS Document HANDLER ──

const handleSaveDocuments = async (entries: IAttachmentEntry[]): Promise<void> => {
  setIsSaving(true);
  try {
    const service = new NFADocService(context);
    const saved = await service.saveDocuments(entries, uniqId);
    setSavedDocuments(prev => [...prev, ...saved]);
    setIsPopupOpen(false);
  } catch (error) {
    console.error('Error saving documents:', error);
    alert('Failed to save documents. Please try again.');
  } finally {
    setIsSaving(false);
  }
};

const handleDeleteDocument = async (itemId: number): Promise<void> => {
  const confirmDelete = confirm('Are you sure you want to remove this document?');
  if (!confirmDelete) return;

  try {
    const service = new NFADocService(context);
    await service.deleteDocument(itemId);
    setSavedDocuments(prev => prev.filter(doc => doc.id !== itemId));
  } catch (error) {
    console.error('Error deleting document:', error);
    alert('Failed to delete document.');
  }
};

const formatFileSize = (bytes: number): string => {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1048576) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1048576).toFixed(1)} MB`;
};

  return (
    <div className="pageContainer">

      {/* Page Title */}
      <p className="formPageTitle">Fill Below Details to Submit Note for Approval</p>

      {/* ── Section 1: Basic Information ── */}
      <CollapsibleSection icon={<Building2 size={13} />} title="Basic Information" defaultOpen={true}>
          <div className="flex flex-wrap gap-x-6 gap-y-4 mb-4">
            <div className="flex-1 min-w-[220px]">
              <DropdownField label="Project" value={form.project}
                onChange={v => set('project', v)}
                options={projectOptions} 
                disabled={loadingProjects}
                />
            </div>

            <div className="flex-1 min-w-[220px]">
               <MultiSelectDropdown
                  label="Sub-Entity"
                  options={subEntityOptions}
                  selectedValues={form.projectSubEntity}
                  onChange={(values: string[]) => set('projectSubEntity', values)}
                  disabled={loadingSubEntity || !form.project}
                  placeholder={
                  !form.project
                  ? "Select Project first"
                  : loadingSubEntity
                  ? "Loading..."
                  : "Select Sub-Entities"
                 }
              />

            </div>
          </div>
          <div className="flex flex-wrap gap-x-6 gap-y-4 mb-4">
            <div className="flex-1 min-w-[220px]">
                              <DropdownField
                  label="Originator Department"
                  value={form.originatorDept}
                  onChange={(v: string) => set('originatorDept', v)}
                  options={functionOptions}
                  disabled={loadingFunction}
                />
            </div>
            <div className="flex-1 min-w-[220px]">
              
                <InputField label="Date" value={form.date} onChange={v => set('date', v)} disabled />
            </div>
          </div>
        
    </CollapsibleSection>

      {/* ── Section 2: NFA Details ── */}
      <CollapsibleSection icon={<FileText size={13} />} title="NFA Details" defaultOpen={false}>
       
          <div className="mb-4">
            <TextAreaField label="Reference / Subject" value={form.referenceSubject}
              onChange={v => set('referenceSubject', v)} rows={2} placeholder="Reference / Subject" required />
          </div>
          <div className="mb-4">
            <RichTextEditor label="Introduction / Preamble / Background"
              value={form.introduction} onChange={v => set('introduction', v)} />
          </div>
          <div className="mb-4">
            <RichTextEditor label="Chronology" value={form.chronology} onChange={v => set('chronology', v)} />
          </div>
          <div className="mb-4">
            <RichTextEditor label="Commercial Proposal/s" value={form.commercialProposal} onChange={v => set('commercialProposal', v)} />
          </div>
          <div>
            <RichTextEditor label="Payment Terms" value={form.paymentTerms} onChange={v => set('paymentTerms', v)} />
          </div>
        
      </CollapsibleSection>

      {/* ── Section 3: Contract Terms ── */}
      <CollapsibleSection icon={<Wrench size={13} />} title="Contract Terms" defaultOpen={false}>
       
          <div className="mb-4">
            <TextAreaField label="Liquidated Damage" value={form.liquidatedDamage}
              onChange={v => set('liquidatedDamage', v)} rows={3} required />
          </div>
          <div className="mb-4">
            <TextAreaField label="Contract Performance Guarantee" value={form.contractPerf}
              onChange={v => set('contractPerf', v)} rows={3} required />
          </div>
          <div className="mb-4">
            <TextAreaField label="Defect Liability Period" value={form.defectLiability}
              onChange={v => set('defectLiability', v)} rows={3} required />
          </div>
          <div>
            <TextAreaField label="Work Completion Period" value={form.workCompletion}
              onChange={v => set('workCompletion', v)} rows={3} required />
          </div>
        
      </CollapsibleSection>

      {/* ── Section 4: Financial Details ── */}
      <CollapsibleSection icon={<DollarSign size={13} />} title="Financial Details" defaultOpen={false}>
        
          <div className="mb-4">
            <TextAreaField label="Budget Remarks" value={form.budgetRemarks}
              onChange={v => set('budgetRemarks', v)} rows={3} />
          </div>
          <div className="flex flex-wrap gap-x-6 gap-y-4 mb-4">
            <div className="flex-1 min-w-[220px]">
              <NumberInput label="Package Awarded Cost" value={form.packageAwardedCost}
                onChange={v => set('packageAwardedCost', v)} required />
            </div>
            <div className="flex-1 min-w-[220px]">
              <NumberInput label="Budget" value={form.budget} onChange={v => set('budget', v)} required />
            </div>
          </div>
          <div>
            <TextAreaField label="Capex Serial No" value={form.capexSerialNo}
              onChange={v => set('capexSerialNo', v)} rows={3} />
          </div>
      
      </CollapsibleSection>

      {/* ── Section 5: Recommendation ── */}
      <CollapsibleSection icon={<Star size={13} />} title="Recommendation" defaultOpen={false}>
        
          <div className="mb-4">
            <RichTextEditor label="Recommendation with Justification" value={form.recommendation}
              onChange={v => set('recommendation', v)} />
          </div>
          <div>
            <InputField label="Contractor Name" value={form.contractorName}
              onChange={v => set('contractorName', v)} placeholder="Enter Contractor Name" required />
          </div>
       
      </CollapsibleSection>

      {/* ── Section 6: Supplementary ── */}
      <CollapsibleSection icon={<FileText size={13} />} title="Annexure & Notes" defaultOpen={false}>
       
          <div className="mb-4">
            <TextAreaField label="Annexure / Enclosure (if any)" value={form.annexure}
              onChange={v => set('annexure', v)} rows={3} placeholder="Enter Annexure / Enclosure (if any)" />
          </div>
          {(['note1', 'note2', 'note3', 'note4'] as const).map((key, i) => (
            <div key={key} className="mb-4">
              <TextAreaField label={`Note-${i + 1}`} value={(form as any)[key]}
                onChange={v => set(key, v)} rows={3} placeholder={`Enter Note-${i + 1}`} />
            </div>
          ))}
        
      </CollapsibleSection>

      {/* ── Section 7: Workflow & Approvers ── */}
      <CollapsibleSection icon={<Settings size={13} />} title="Workflow & Approvers" defaultOpen={false}>
       
          <div className="flex flex-wrap gap-x-6 gap-y-4 mb-5">
            <div className="flex-1 min-w-[220px]">
              <DropdownField label="Workflow Execution" value={form.workflow}
                onChange={v => set('workflow', v)} options={workflowOptions} required />
            </div>
            <div className="flex-1 min-w-[220px]">
              <DropdownField label="DOA Matrix" value={form.doaMatrix}
                onChange={v => set('doaMatrix', v)}
                options={[{ label: 'Matrix A', value: 'A' }, { label: 'Matrix B', value: 'B' }]} />
          </div>

          <div className="formSectionHeader" style={{ margin: '0 -24px 16px', borderRadius: 0 }}>
            <Users size={13} /> Approver Levels
          </div>

          {[0, 4, 8].map((startIdx) => (
            <div key={startIdx} className="flex flex-wrap gap-x-4 gap-y-4 mb-4">
              {[0, 1, 2, 3].map((offset) => {
                const idx = startIdx + offset;
                return (
                  <div key={idx} className="flex-1 min-w-[180px]">
                    <DropdownField
                      label={`Approver Level-${idx + 1}`}
                      value={form.approvers[idx]}
                      onChange={v => setApprover(idx, v)}
                      options={approverOptions}
                      required={idx === 0}
                    />
                  </div>
                );
              })}
            </div>
          ))}
        </div>
      </CollapsibleSection>

      {/* ── Section 8: Attachments ── */}

        <CollapsibleSection
          icon={<Paperclip size={13} />}
          title="Attachments"
          defaultOpen={false}
        >
          {/* Add Attachment Button */}
          <div className="flex items-center justify-between mb-4">
            <p style={{ fontSize: '13px', color: '#718096' }}>
              UniqID: <strong>{uniqId}</strong>
            </p>
            <CommonButton
              text="+ Add Attachment"
              variant="primary"
              onClick={() => setIsPopupOpen(true)}
            />
          </div>

          {/* Saved Documents Table */}
          {savedDocuments.length > 0 ? (
            <div className="savedDocsSection">
              <table className="savedDocsTable">
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
                  {savedDocuments.map((doc, idx) => (
                    <tr key={doc.id}>
                      <td>{idx + 1}</td>
                      <td>{doc.documentType}</td>
                      <td>{doc.fileName}</td>
                      <td>{formatFileSize(doc.fileSize)}</td>
                      <td>
                        <button
                          className="removeBtn"
                          onClick={() => handleDeleteDocument(doc.id)}
                          title="Remove"
                        >
                          <Trash2 size={14} />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <p className="noDocsMessage">
              No documents attached yet. Click "+ Add Attachment" to upload.
            </p>
          )}
        </CollapsibleSection>

        {/* ── Attachment Popup ── */}
        <AttachmentPopup
          isOpen={isPopupOpen}
          onClose={() => setIsPopupOpen(false)}
          onSave={handleSaveDocuments}
        />

      {/* ── Action Buttons ── */}
      <div className="flex items-center justify-center gap-4 pt-2 pb-4">
        <CommonButton text="Submit"       variant="primary"   onClick={() => console.log('Submit', form)} />
        <CommonButton text="Save as Draft" variant="secondary" onClick={() => console.log('Draft', form)} />
        <CommonButton text="Cancel"       variant="danger"    onClick={() => console.log('Cancel')} />
      </div>

    </div>
  );
};

export default NewNFA;
