import * as React from 'react';
import { SPHttpClient } from "@microsoft/sp-http";
import { useState, useEffect, useRef } from 'react';
import { Building2, FileText, Wrench, DollarSign, Paperclip, Settings, Star, ChevronDown, Trash2
 } from 'lucide-react';
import { Skeleton } from 'primereact/skeleton'; 
import { InputField } from '../../../../../Share/Components/InputField';
import { TextAreaField } from '../../../../../Share/Components/TextAreaField';
import { NumberInput } from '../../../../../Share/Components/NumberInput';
import { RichTextEditor } from '../../../../../Share/Components/RichTextEditor';
import DropdownField from '../../../../../Share/Components/DropdownField';
import CommonButton from '../../../../../Share/Components/CommonButton';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { DropdownService, IDropdownOption } from '../../../../../Share/Services/DropdownService';
import MultiSelectDropdown from '../../../../../Share/Components/MultiSelectDropdown';
import AttachmentPopup, { IAttachmentEntry } from '../../../../../Share/Components/AttachmentPopup';
import { NFADocService, ISavedDocument } from '../../../../../Share/Services/NFADocService';
import ValidationErrors from '../../../../../Share/Components/ValidationErrors';
import MessagePopup, { MessageType } from '../../../../../Share/Components/MessagePopup';
import { NFAService, NFAStatus, IValidationError, INFAFormData } from '../../../../../Share/Services/NFAService';
import { useNavigate, useParams } from 'react-router-dom';
import "../../../../../Share/styles/global.scss";

// Options
const workflowOptions = [{ label: 'MANUAL', value: 'MANUAL' }, { label: 'AUTO', value: 'AUTO' }];

// --- Component: Collapsible Section ---
interface CollapsibleSectionProps {
  icon: React.ReactNode;
  title: string;
  children: React.ReactNode;
  defaultOpen?: boolean;
}

const CollapsibleSection: React.FC<CollapsibleSectionProps> = ({ icon, title, children, defaultOpen = true }) => {
  const [isOpen, setIsOpen] = useState(defaultOpen);
  return (
    <div className={`formSection ${isOpen ? 'formSection--open' : 'formSection--closed'}`}>
      <div className="formSectionHeader formSectionHeader--toggle" onClick={() => setIsOpen(!isOpen)}>
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

const NFASkeleton = () => {
  return (
    <div className="pageContainer">
      <div className="flex justify-center mb-8">
        <Skeleton width="50%" height="2.5rem" />
      </div>
      {[1, 2, 3].map((i) => (
        <div key={i} className="formSection mb-6">
          <div className="formSectionHeader">
            <div className="flex items-center gap-2">
              <Skeleton shape="circle" size="1.2rem" />
              <Skeleton width="120px" height="1rem" />
            </div>
          </div>
          <div className="formSectionBody">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div>
                <Skeleton width="80px" height="0.8rem" className="mb-2" />
                <Skeleton width="100%" height="2.5rem" borderRadius="6px" />
              </div>
              <div>
                <Skeleton width="80px" height="0.8rem" className="mb-2" />
                <Skeleton width="100%" height="2.5rem" borderRadius="6px" />
              </div>
            </div>
          </div>
        </div>
      ))}
    </div>
  );
};

// --- Main Component ---
interface INewNFAProps {
  context: WebPartContext;
  sp: SPFI;
}

const NewNFA: React.FC<INewNFAProps> = ({ context, sp }) => {
  const { id } = useParams();
  const navigate = useNavigate();
  const today = new Date().toLocaleDateString('en-IN');

  // Logic Locks
  const hasInitialized = useRef(false);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);

  // Data States
  const [projectOptions, setProjectOptions] = useState<IDropdownOption[]>([]);
  const [subEntityOptions, setSubEntityOptions] = useState<IDropdownOption[]>([]);
  const [functionOptions, setFunctionOptions] = useState<IDropdownOption[]>([]);
  const [loadingSubEntity, setLoadingSubEntity] = useState(false);
  const [savedDocuments, setSavedDocuments] = useState<ISavedDocument[]>([]);
  const [pendingAttachments, setPendingAttachments] = useState<IAttachmentEntry[]>([]);
  const [errors, setErrors] = useState<IValidationError[]>([]);
  const [isPopupOpen, setIsPopupOpen] = useState(false);
  const [matrixApproverOptions, setMatrixApproverOptions] = useState<IDropdownOption[]>([]);
  const [doaMatrixOptions, setDoaMatrixOptions] = useState<IDropdownOption[]>([]);
  const [doaMatrixApprovers, setDoaMatrixApprovers] = useState<IDropdownOption[]>([]);
  const lastDoaCheck = useRef({ dept: '', amount: '' });
  const [doaAlert, setDoaAlert] = useState<{ show: boolean, type: MessageType, title: string, messages: string[] }>({
    show: false, type: 'info', title: '', messages: []
  });

  const [form, setForm] = useState<INFAFormData>({
    originatorDept: '', date: today, project: '', projectSubEntity: [],
    referenceSubject: '', introduction: '', chronology: '',
    commercialProposal: '', paymentTerms: '',
    liquidatedDamage: 'NA', contractPerf: 'NA',
    defectLiability: 'NA', workCompletion: 'Work must be completed and handed over to the satisfaction of Owner on or before -------',
    budgetRemarks: '', packageAwardedCost: '0.00', budget: '0.00',
    capexSerialNo: '', recommendation: '', contractorName: '',
    annexure: '', note1: '', note2: '', note3: '', note4: '',
    workflow: 'MANUAL', doaMatrix: '',
    approvers: Array(12).fill(null),
    uniqId: '',
  });

  // --- 1. THE MASTER INITIALIZATION (Runs exactly once) ---
  useEffect(() => {
    if (hasInitialized.current) return;

    const setupPage = async () => {
      const dropdownService = new DropdownService(context);
      const nfaService = new NFAService(context);
      const docService = new NFADocService(context);

      try {
        const [projects, dpts] = await Promise.all([
          dropdownService.getProjects(),
          dropdownService.getFunctionMaster()
        ]);
        setProjectOptions(projects);
        setFunctionOptions(dpts);

        if (id) {
          const nfaData = await nfaService.getNFA(parseInt(id));
          if (nfaData) {
            const matchedProject = projects.find(p => p.label === nfaData.project);
            if (matchedProject) nfaData.project = matchedProject.value;

            const subs1 = await dropdownService.getSubEntity(nfaData.project);
            setSubEntityOptions(subs1);

            if (nfaData.projectSubEntity && nfaData.projectSubEntity.length > 0) {
              const subIds = nfaData.projectSubEntity.map(name => {
                const match = subs1.find(s => s.label === name);
                return match ? match.value : name;
              });
              nfaData.projectSubEntity = subIds;
            }
            const [docs, subs] = await Promise.all([
              docService.getDocumentsByUniqId(nfaData.uniqId),
              dropdownService.getSubEntity(nfaData.project)
            ]);
            setSavedDocuments(docs);
            setSubEntityOptions(subs);
            setForm(nfaData);
          }
        } else {
          const newId = docService.generateUniqId();
          setForm(prev => ({ ...prev, uniqId: newId }));
        }

        hasInitialized.current = true;
      } catch (err) {
        console.error("Initialization Failed:", err);
      } finally {
        setTimeout(() => setIsLoading(false), 100);
      }
    };
    setupPage();
  }, [id]);

  // --- 2. APPROVER MATRIX FILL ---
  useEffect(() => {
    if (!form.originatorDept) {
      setMatrixApproverOptions([]);
      return;
    }
    const getOptionsFromMatrix = async () => {
      try {
        const siteUrl = context.pageContext.web.absoluteUrl;
        const endpoint = `${siteUrl}/_api/web/lists/getbytitle('ApproverMatrix')/items?$filter=Title eq '${form.originatorDept}' and IsActive eq 1&$select=Approver/Id,Approver/Title&$expand=Approver`;
        const response = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
        const data = await response.json();
        if (data.value) {
          const options = data.value.map((item: any) => ({
            label: item.Approver.Title,
            value: String(item.Approver.Id)
          }));
          setMatrixApproverOptions(options);
        }
      } catch (err) {
        console.error("Error fetching Matrix options:", err);
      }
    };
    getOptionsFromMatrix();
  }, [form.originatorDept, id]);

  // --- 4. AUTO WORKFLOW: DOA MATRIX LOOKUP ---
  const fetchMatchingMatrix = async (manualTrigger = false) => {
    if (form.workflow !== 'AUTO' || !form.originatorDept || !form.packageAwardedCost) {
      if (form.workflow === 'MANUAL') {
        setDoaMatrixOptions([]);
        setDoaMatrixApprovers([]);
        setForm(prev => ({ ...prev, doaMatrix: '' }));
      }
      return;
    }

    try {
      const rawAmount = String(form.packageAwardedCost || "0").replace(/,/g, '');
      const amount = parseFloat(rawAmount) || 0;
      const shouldAlert = manualTrigger;

      const siteUrl = context.pageContext.web.absoluteUrl;
      const filter = `Function eq '${form.originatorDept}' and MinAmount le ${amount} and MaxAmount ge ${amount}`;
      const endpoint = `${siteUrl}/_api/web/lists/getbytitle('DOAMatrix')/items?$filter=${filter}`;

      const response = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
      if (!response.ok) return;
      const data = await response.json();

      if (data.value && data.value.length > 0) {
        const options = data.value.map((item: any) => ({
          label: item.Title,
          value: String(item.Id)
        }));
        setDoaMatrixOptions(options);
        if (options.length === 1 && form.doaMatrix !== options[0].value) {
          setForm(prev => ({ ...prev, doaMatrix: options[0].value }));
        }
        lastDoaCheck.current = { dept: form.originatorDept, amount: rawAmount };
      } else {
        setDoaMatrixOptions([]);
        setDoaMatrixApprovers([]);
        updateField('doaMatrix', '');
        if (shouldAlert && amount > 0) {
          setDoaAlert({
            show: true,
            type: 'warning',
            title: 'DOA Matrix Not Found',
            messages: [`No DOA Matrix rule found for '${form.originatorDept}' at the amount of ${amount}.`, "Please verify your DOA list or switch to MANUAL workflow."]
          });
          lastDoaCheck.current = { dept: form.originatorDept, amount: rawAmount };
        }
      }
    } catch (err) {
      console.error("Error in fetchMatchingMatrix:", err);
    }
  };

  useEffect(() => {
    fetchMatchingMatrix(false);
  }, [form.workflow, form.originatorDept]);

  // --- 5. AUTO WORKFLOW: POPULATE APPROVERS ---
  useEffect(() => {
    if (form.workflow !== 'AUTO' || !form.doaMatrix) return;
    const fetchApproversFromMatrix = async () => {
      try {
        const siteUrl = context.pageContext.web.absoluteUrl;
        const levels = Array.from({ length: 12 }, (_, i) => `Approver${i + 1}`);
        const select = levels.map(l => `${l}/Id,${l}/Title`).join(',');
        const expand = levels.join(',');
        const endpoint = `${siteUrl}/_api/web/lists/getbytitle('DOAMatrix')/items(${form.doaMatrix})?$select=${select}&$expand=${expand}`;
        const response = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
        if (!response.ok) return;
        const item = await response.json();
        if (item) {
          const matrixApprovers = Array(12).fill(null);
          const approverOptionsList: IDropdownOption[] = [];
          for (let i = 1; i <= 12; i++) {
            const field = `Approver${i}`;
            if (item[field] && item[field].Id) {
              const userId = String(item[field].Id);
              matrixApprovers[i - 1] = userId;
              if (!approverOptionsList.find(o => o.value === userId)) {
                approverOptionsList.push({ label: item[field].Title, value: userId });
              }
            }
          }
          setDoaMatrixApprovers(approverOptionsList);
          setForm(prev => ({ ...prev, approvers: matrixApprovers }));
        }
      } catch (err) {
        console.error("Error in fetchApproversFromMatrix:", err);
      }
    };
    fetchApproversFromMatrix();
  }, [form.doaMatrix, form.workflow]);

  // --- HELPERS ---
  const handleProjectChange = async (projectId: string) => {
    setForm(prev => ({ ...prev, project: projectId, projectSubEntity: [] }));
    if (!projectId) {
      setSubEntityOptions([]);
      return;
    }
    setLoadingSubEntity(true);
    try {
      const service = new DropdownService(context);
      const subs = await service.getSubEntity(projectId);
      setSubEntityOptions(subs);
    } finally {
      setLoadingSubEntity(false);
    }
  };

  const updateField = (key: string, val: any) => setForm(f => ({ ...f, [key]: val }));
  const updateApprover = (i: number, val: string) => {
    const arr = [...form.approvers];
    arr[i] = val;
    setForm(f => ({ ...f, approvers: arr }));
  };

  const validate = (status: NFAStatus) => {
    setErrors([]);
    const service = new NFAService(context);
    const result = service.validateForm(form, status);
    if (!result.isValid) {
      setErrors(result.errors);
      return false;
    }
    return true;
  };

  if (isLoading) return <NFASkeleton />;

  return (
    <div className="pageContainer">
      <p className="formPageTitle">{id ? "Edit Note for Approval Request" : "Note for Approval Request Form"}</p>

      <CollapsibleSection icon={<Building2 size={13} />} title="Basic Information">
        <div className="flex flex-wrap gap-x-6 gap-y-4 mb-4">
          <div className="flex-1 min-w-[220px]">
            <DropdownField label="Project" value={form.project} options={projectOptions} onChange={handleProjectChange} />
          </div>
          <div className="flex-1 min-w-[220px]">
            <MultiSelectDropdown label="Sub-Entity" options={subEntityOptions} selectedValues={form.projectSubEntity} onChange={v => updateField('projectSubEntity', v)} disabled={loadingSubEntity || !form.project} />
          </div>
        </div>
        <div className="flex flex-wrap gap-x-6 gap-y-4">
          <div className="flex-1 min-w-[220px]">
            <DropdownField label="Function" value={form.originatorDept} options={functionOptions} onChange={v => updateField('originatorDept', v)} />
          </div>
          <div className="flex-1 min-w-[220px]">
            <InputField label="Date" value={form.date} onChange={() => { }} disabled />
          </div>
        </div>
      </CollapsibleSection>

      <CollapsibleSection icon={<FileText size={13} />} title="NFA Details" defaultOpen={false}>
        <div className="mb-4"><TextAreaField label="Subject" value={form.referenceSubject} onChange={v => updateField('referenceSubject', v)} rows={2} required /></div>
        <div className="mb-4"><RichTextEditor label="Introduction" value={form.introduction} onChange={v => updateField('introduction', v)} /></div>
        <div className="mb-4"><RichTextEditor label="Chronology" value={form.chronology} onChange={v => updateField('chronology', v)} /></div>
        <div className="mb-4"><RichTextEditor label="Commercial Proposals" value={form.commercialProposal} onChange={v => updateField('commercialProposal', v)} /></div>
        <RichTextEditor label="Payment Terms" value={form.paymentTerms} onChange={v => updateField('paymentTerms', v)} />
      </CollapsibleSection>

      <CollapsibleSection icon={<Wrench size={13} />} title="Contract Terms" defaultOpen={false}>
        <div className="mb-4"><TextAreaField label="LD" value={form.liquidatedDamage} onChange={v => updateField('liquidatedDamage', v)} rows={3} /></div>
        <div className="mb-4"><TextAreaField label="CPG" value={form.contractPerf} onChange={v => updateField('contractPerf', v)} rows={3} /></div>
        <div className="mb-4"><TextAreaField label="DLP" value={form.defectLiability} onChange={v => updateField('defectLiability', v)} rows={3} /></div>
        <TextAreaField label="Work Completion" value={form.workCompletion} onChange={v => updateField('workCompletion', v)} rows={3} />
      </CollapsibleSection>

      <CollapsibleSection icon={<DollarSign size={13} />} title="Financials" defaultOpen={false}>
        <div className="mb-4"><TextAreaField label="Remarks" value={form.budgetRemarks} onChange={v => updateField('budgetRemarks', v)} rows={3} /></div>
        <div className="flex gap-x-6 mb-4">
          <div className="flex-1"><NumberInput label="Awarded Cost" value={form.packageAwardedCost} onChange={v => updateField('packageAwardedCost', v)} onBlur={() => fetchMatchingMatrix(true)} /></div>
          <div className="flex-1"><NumberInput label="Budget" value={form.budget} onChange={v => updateField('budget', v)} /></div>
        </div>
        <TextAreaField label="Capex No" value={form.capexSerialNo} onChange={v => updateField('capexSerialNo', v)} rows={3} />
      </CollapsibleSection>

      <CollapsibleSection icon={<Star size={13} />} title="Recommendation" defaultOpen={false}>
        <div className="mb-4"><RichTextEditor label="Justification" value={form.recommendation} onChange={v => updateField('recommendation', v)} /></div>
        <InputField label="Contractor Name" value={form.contractorName} onChange={v => updateField('contractorName', v)} />
      </CollapsibleSection>

      <CollapsibleSection icon={<Settings size={13} />} title="Workflow & Approvers" defaultOpen={false}>
        <div className="flex gap-x-6 mb-6">
          <div className="flex-1"><DropdownField label="Workflow" value={form.workflow} options={workflowOptions} onChange={v => updateField('workflow', v)} /></div>
          <div className="flex-1"><DropdownField label="DOA Matrix" value={form.doaMatrix} options={doaMatrixOptions} onChange={v => updateField('doaMatrix', v)} disabled={form.workflow !== 'AUTO'} /></div>
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '16px' }}>
          {form.approvers.map((app, idx) => (
            <DropdownField key={idx} label={`Approver Level ${idx + 1}`} value={String(app ?? "")} options={form.workflow === 'AUTO' ? doaMatrixApprovers : matrixApproverOptions} onChange={(v) => updateApprover(idx, v)} disabled={form.workflow === 'AUTO'} />
          ))}
        </div>
      </CollapsibleSection>

      <CollapsibleSection icon={<Paperclip size={13} />} title="Attachments" defaultOpen={false}>
        <div className="flex justify-end mb-4"><CommonButton text="+ Add Attachment" variant="primary" onClick={() => setIsPopupOpen(true)} /></div>
        {(savedDocuments.length > 0 || pendingAttachments.length > 0) ? (
          <table className="savedDocsTable w-full border text-left">
            <thead><tr><th className="p-2">#</th><th className="p-2">Type</th><th className="p-2">File</th><th className="p-2">Status</th><th className="p-2">Del</th></tr></thead>
            <tbody>
              {savedDocuments.map((doc, i) => (
                <tr key={doc.id} className="border-t">
                  <td className="p-2">{i + 1}</td>
                  <td className="p-2">{doc.documentType}</td>
                  <td className="p-2 text-blue-600">{doc.fileName}</td>
                  <td className="p-2 text-green-600 font-semibold">Saved</td>
                  <td className="p-2"><button onClick={async () => { if (confirm("Are you sure?")) { const s = new NFADocService(context); await s.deleteDocument(doc.id); setSavedDocuments(prev => prev.filter(d => d.id !== doc.id)); } }} className="text-red-500"><Trash2 size={14} /></button></td>
                </tr>
              ))}
              {pendingAttachments.map((entry, i) => (
                <tr key={`pending-${i}`} className="border-t bg-blue-50/30">
                  <td className="p-2">{savedDocuments.length + i + 1}</td>
                  <td className="p-2">{entry.documentType}</td>
                  <td className="p-2 text-gray-600 italic">{entry.file.name}</td>
                  <td className="p-2 text-amber-600 font-semibold">Pending</td>
                  <td className="p-2"><button onClick={() => setPendingAttachments(prev => prev.filter((_, idx) => idx !== i))} className="text-red-500"><Trash2 size={14} /></button></td>
                </tr>
              ))}
            </tbody>
          </table>
        ) : <p className="noDocsMessage">No documents attached.</p>}
      </CollapsibleSection>

      <AttachmentPopup isOpen={isPopupOpen} onClose={() => setIsPopupOpen(false)} onSave={(entries) => { setPendingAttachments(prev => [...prev, ...entries]); setIsPopupOpen(false); }} />
      <MessagePopup isOpen={doaAlert.show} type={doaAlert.type} title={doaAlert.title} messages={doaAlert.messages} onClose={() => setDoaAlert(prev => ({ ...prev, show: false }))} />
      {errors.length > 0 && <ValidationErrors errors={errors} onClose={() => setErrors([])} />}

      <div className="flex justify-center gap-4 py-8">
        <CommonButton text={id ? "Update NFA" : "Submit NFA"} variant="primary" disabled={isSaving} onClick={async () => {
          if (!validate('Pending')) return;
          setIsSaving(true);
          try {
            const docService = new NFADocService(context);
            if (pendingAttachments.length > 0) await docService.saveDocuments(pendingAttachments, form.uniqId);
            const service = new NFAService(context);
            const res = id ? await service.updateNFA(parseInt(id), form, 'Pending', projectOptions, subEntityOptions) : await service.saveNFA(form, 'Pending', projectOptions, subEntityOptions);
            alert(res.message); navigate('/');
          } catch (err) { alert("Error saving NFA or Uploading Documents."); } finally { setIsSaving(false); }
        }} />
        <CommonButton text="Save Draft" variant="secondary" disabled={isSaving} onClick={async () => {
          if (!validate('Draft')) return;
          setIsSaving(true);
          try {
            const docService = new NFADocService(context);
            if (pendingAttachments.length > 0) await docService.saveDocuments(pendingAttachments, form.uniqId);
            const service = new NFAService(context);
            const res = id ? await service.updateNFA(parseInt(id), form, 'Draft', projectOptions, subEntityOptions) : await service.saveNFA(form, 'Draft', projectOptions, subEntityOptions);
            alert(res.message); navigate('/');
          } catch (err) { alert("Error saving Draft or Uploading Documents."); } finally { setIsSaving(false); }
        }} />
        <CommonButton text="Cancel" variant="danger" onClick={() => navigate('/')} />
      </div>
    </div>
  );
};

export default NewNFA;
