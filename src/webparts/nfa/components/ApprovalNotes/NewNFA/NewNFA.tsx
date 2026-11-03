import * as React from 'react';
import { SPHttpClient } from "@microsoft/sp-http";
import { useState, useEffect, useRef, useMemo } from 'react';
import {
  Building2, FileText, Wrench, DollarSign, Paperclip, Settings, Star, ChevronDown, Trash2, History, MessageSquare, Info
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
import AttachmentValidationModal, {
  validateAttachmentFiles,
  IAttachmentValidationError
} from '../../../../../Share/Components/AttachmentValidationModal';
import { NFADocService, ISavedDocument } from '../../../../../Share/Services/NFADocService';
import ValidationErrors from '../../../../../Share/Components/ValidationErrors';
import MessagePopup, { MessageType } from '../../../../../Share/Components/MessagePopup';
import { NFAService, NFAStatus, IValidationError, INFAFormData } from '../../../../../Share/Services/NFAService';
import { useNavigate, useParams } from 'react-router-dom';
import "../../../../../Share/styles/global.scss";
import PageTitle from '../../../../../Share/Components/PageTitle';
import { forceScrollToTop } from '../../../../../Share/utils';

// Options
const workflowOptions = [{ label: 'MANUAL', value: 'MANUAL' }, { label: 'AUTO', value: 'AUTO' }];

// --- Component: Collapsible Section ---
interface CollapsibleSectionProps {
  icon: React.ReactNode;
  title: string;
  children: React.ReactNode;
  defaultOpen?: boolean;
  hasError?: boolean;
}

const CollapsibleSection: React.FC<CollapsibleSectionProps> = ({ icon, title, children, defaultOpen = true, hasError = false }) => {
  const [isOpen, setIsOpen] = useState(defaultOpen);
  return (
    <div className={`formSection ${isOpen ? 'formSection--open' : 'formSection--closed'} ${hasError ? 'border-red-300 shadow-[0_0_10px_rgba(239,68,68,0.05)]' : ''}`}>
      <div
        className={`formSectionHeader formSectionHeader--toggle flex items-center justify-between cursor-pointer ${hasError ? 'bg-red-50/30' : ''}`}
        onClick={() => setIsOpen(!isOpen)}
      >
        <div className="formSectionHeader__left flex items-center gap-2">
          {icon}
          <span>{title}</span>
          {hasError && (
            <span className="flex h-2 w-2 relative">
              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-red-400 opacity-75"></span>
              <span className="relative inline-flex rounded-full h-2 w-2 bg-red-500"></span>
            </span>
          )}
        </div>
        <div className={`formSectionHeader__chevron ${isOpen ? 'chevron--open' : 'chevron--closed'}`}>
          <ChevronDown size={16} className={hasError ? 'text-red-500' : ''} />
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
  const [savingProgress, setSavingProgress] = useState<{
    step: 'attachments' | 'form' | 'complete';
    currentFileIndex: number;
    totalFiles: number;
    currentFileName: string;
    percent: number;
  } | null>(null);

  // Data States
  const [projectOptions, setProjectOptions] = useState<IDropdownOption[]>([]);
  const [subEntityOptions, setSubEntityOptions] = useState<IDropdownOption[]>([]);
  const [functionOptions, setFunctionOptions] = useState<IDropdownOption[]>([]);
  const [loadingSubEntity, setLoadingSubEntity] = useState(false);
  const [savedDocuments, setSavedDocuments] = useState<ISavedDocument[]>([]);
  const [pendingAttachments, setPendingAttachments] = useState<IAttachmentEntry[]>([]);
  const [errors, setErrors] = useState<IValidationError[]>([]);
  const [isPopupOpen, setIsPopupOpen] = useState(false);
  const [attachmentValidationErrors, setAttachmentValidationErrors] = useState<IAttachmentValidationError[]>([]);
  const [showAttachmentValidationModal, setShowAttachmentValidationModal] = useState<boolean>(false);
  const [manualSlotsCount, setManualSlotsCount] = useState(4);
  const [matrixApproverOptions, setMatrixApproverOptions] = useState<IDropdownOption[]>([]);
  const [doaMatrixOptions, setDoaMatrixOptions] = useState<IDropdownOption[]>([]);
  const [doaMatrixApprovers, setDoaMatrixApprovers] = useState<IDropdownOption[]>([]);
  const lastDoaCheck = useRef({ dept: '', amount: '' });
  const isInitialMount = useRef(true);
  const [doaAlert, setDoaAlert] = useState<{ show: boolean, type: MessageType, title: string, messages: string[], onSuccessRedirect?: boolean }>({
    show: false, type: 'info', title: '', messages: [], onSuccessRedirect: false
  });
  const [history, setHistory] = useState<any[]>([]);
  const [comments, setComments] = useState<any[]>([]);

  const [form, setForm] = useState<INFAFormData>({
    originatorDept: '', date: today, project: '', projectSubEntity: [],
    referenceSubject: '', introduction: '', chronology: '',
    completeIndentPRReceivedFromDate: '',
    boqSpecificationsRevisionsByDate: '',
    RFQRefNo: '',
    remark: '',
    commercialProposal: '', paymentTerms: '',
    liquidatedDamage: '', contractPerf: '',
    defectLiability: '', workCompletion: 'Work must be completed and handed over to the satisfaction of Owner on or before -------',
    budgetRemarks: '', packageAwardedCost: '0.00', budget: '0.00',
    budgetedNonBudgeted: "",
    capexSerialNo: '',
    recommendation: '', contractorName: '',
    annexure: ' ', note1: '', note2: ' ', note3: ' ', note4: ' ',
    workflow: 'MANUAL', doaMatrix: '',
    approvers: Array(12).fill(null),
    uniqId: '',
    level: 0,
    pendingAtId: null,
    Approvedrejectedby: "",
    Approvedrejectdatetime: "",
    PendingFrom: "",
  });

  const handlePopupClose = () => {
    const redirect = doaAlert.onSuccessRedirect;
    setDoaAlert(prev => ({ ...prev, show: false }));
    if (redirect) {
      navigate('/');
    }
  };

  const handleAttachmentPasteOrDrop = (files: File[]): void => {
    if (!files || files.length === 0) return;
    const existingNames = [
      ...savedDocuments.map(d => d.fileName),
      ...pendingAttachments.map(p => p.file.name)
    ];
    const { validFiles, errors: valErrors } = validateAttachmentFiles(files, existingNames);
    if (valErrors.length > 0) {
      setAttachmentValidationErrors(valErrors);
      setShowAttachmentValidationModal(true);
    }
    if (validFiles.length > 0) {
      const newEntries: IAttachmentEntry[] = validFiles.map(file => ({
        documentType: 'Supporting Document',
        file
      }));
      setPendingAttachments(prev => [...prev, ...newEntries]);
    }
  };

  const hasSectionErrors = (fields: string[]) => {
    return errors.some(e => fields.includes(e.field));
  };

  const financialsSummary = useMemo(() => {
    const cost = parseFloat(String(form.packageAwardedCost || "0").replace(/,/g, '')) || 0;
    const budget = parseFloat(String(form.budget || "0").replace(/,/g, '')) || 0;
    const variance = budget - cost;
    const percent = budget > 0 ? Math.round((cost / budget) * 100) : 0;
    return { cost, budget, variance, percent };
  }, [form.packageAwardedCost, form.budget]);

  const getFileBadge = (fileName: string) => {
    const ext = fileName.split('.').pop()?.toLowerCase();
    const map: Record<string, { bg: string, text: string, border: string, label: string }> = {
      'pdf': { bg: 'bg-red-50', text: 'text-red-700', border: 'border-red-200', label: 'PDF' },
      'xlsx': { bg: 'bg-emerald-50', text: 'text-emerald-700', border: 'border-emerald-200', label: 'EXCEL' },
      'xls': { bg: 'bg-emerald-50', text: 'text-emerald-700', border: 'border-emerald-200', label: 'EXCEL' },
      'docx': { bg: 'bg-blue-50', text: 'text-blue-700', border: 'border-blue-200', label: 'WORD' },
      'doc': { bg: 'bg-blue-50', text: 'text-blue-700', border: 'border-blue-200', label: 'WORD' },
      'png': { bg: 'bg-purple-50', text: 'text-purple-700', border: 'border-purple-200', label: 'IMAGE' },
      'jpg': { bg: 'bg-purple-50', text: 'text-purple-700', border: 'border-purple-200', label: 'IMAGE' },
      'jpeg': { bg: 'bg-purple-50', text: 'text-purple-700', border: 'border-purple-200', label: 'IMAGE' },
      'svg': { bg: 'bg-purple-50', text: 'text-purple-700', border: 'border-purple-200', label: 'IMAGE' }
    };
    const style = map[ext || ''] || { bg: 'bg-slate-50', text: 'text-slate-700', border: 'border-slate-200', label: 'FILE' };
    return (
      <span className={`inline-flex items-center px-2 py-0.5 rounded text-[10px] font-bold border ${style.bg} ${style.text} ${style.border}`}>
        {style.label}
      </span>
    );
  };

  const shouldShowApprover = (idx: number) => {
    if (form.workflow === 'AUTO') {
      return form.approvers[idx] !== null && form.approvers[idx] !== '';
    } else {
      if (idx === 0) return true;
      return (form.approvers[idx - 1] !== null && form.approvers[idx - 1] !== '') || (form.approvers[idx] !== null && form.approvers[idx] !== '');
    }
  };


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
          console.log("URL ID:", id);
          console.log("Parsed ID:", parseInt(id,));

          const nfaData = await nfaService.getNFA(parseInt(id));



          if (nfaData) {
            const matchedProject = projects.find(p => p.label === nfaData.project);
            if (matchedProject) nfaData.project = matchedProject.value;
            const matchedFunc = dpts.find(d => d.value === String(nfaData.originatorDept) || d.label === nfaData.originatorDept);
            if (matchedFunc) nfaData.originatorDept = matchedFunc.value;

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

            const parseCommentHistory = (commentHistoryText: string): any[] => {
              if (!commentHistoryText) return [];
              const lines = commentHistoryText.split("\n");
              const parsed: any[] = [];
              for (const line of lines) {
                const cleanLine = line
                  .replace(/<[^>]*>/g, "")
                  .replace(/&#58;/g, ":")
                  .replace(/&nbsp;/g, " ")
                  .replace(/&amp;/g, "&")
                  .replace(/&lt;/g, "<")
                  .replace(/&gt;/g, ">")
                  .trim();
                if (!cleanLine) continue;
                const match = cleanLine.match(/^\[(.*?)\]\s+(.*?)\s+\((.*?)\)(?:\s+\[Level\s+(.*?)\])?:\s+(.*)$/);
                if (match) {
                  parsed.push({
                    id: Math.random(),
                    date: match[1],
                    user: match[2],
                    action: match[3],
                    level: match[4] || "1",
                    remarks: match[5]
                  });
                } else {
                  parsed.push({
                    id: Math.random(),
                    date: "",
                    user: "System",
                    action: "Log",
                    remarks: cleanLine
                  });
                }
              }
              return parsed;
            };

            const parseWorkflowHistory = (workflowHistoryText: string): any[] => {
              if (!workflowHistoryText) return [];
              const lines = workflowHistoryText.split("\n");
              const parsed: any[] = [];
              for (const line of lines) {
                const cleanLine = line
                  .replace(/<[^>]*>/g, "")
                  .replace(/&#58;/g, ":")
                  .replace(/&nbsp;/g, " ")
                  .replace(/&amp;/g, "&")
                  .replace(/&lt;/g, "<")
                  .replace(/&gt;/g, ">")
                  .trim();
                if (!cleanLine) continue;

                const appMatch = cleanLine.match(/^NFA Request (.*?) by (.*?) at Level\s+(.*?)\.\s+on\s+(.*)$/);
                if (appMatch) {
                  parsed.push({
                    id: Math.random(),
                    date: appMatch[4].trim(),
                    user: appMatch[2].trim(),
                    action: appMatch[1].trim(),
                    rawText: cleanLine,
                    remarks: ""
                  });
                  continue;
                }

                const initMatch = cleanLine.match(/^NFA Request Initiated by (.*?) on (.*)$/);
                if (initMatch) {
                  parsed.push({
                    id: Math.random(),
                    date: initMatch[2].trim(),
                    user: initMatch[1].trim(),
                    action: "Initiated",
                    rawText: cleanLine,
                    remarks: ""
                  });
                  continue;
                }

                parsed.push({
                  id: Math.random(),
                  date: "",
                  user: "System",
                  action: "Log",
                  rawText: cleanLine,
                  remarks: cleanLine
                });
              }
              return parsed;
            };

            if (nfaData.workflowHistory) {
              setHistory(parseWorkflowHistory(nfaData.workflowHistory).reverse());
            } else if (nfaData.commentHistory) {
              setHistory(parseCommentHistory(nfaData.commentHistory).reverse());
            }

            if (nfaData.commentHistory) {
              setComments(parseCommentHistory(nfaData.commentHistory));
            }

            // Calculate initial manual slots count
            const lastFilledIdx = nfaData.approvers.map(a => a !== null && a !== '').lastIndexOf(true);
            setManualSlotsCount(lastFilledIdx >= 0 ? Math.max(4, lastFilledIdx + 1) : 4);
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
        //const endpoint = `${siteUrl}/_api/web/lists/getbytitle('ApproverMatrix')/items?$filter=Title eq '${form.originatorDept}' and IsActive eq 1&$select=Approver/Id,Approver/Title&$expand=Approver`;
        const funcName = functionOptions.find(f => f.value === String(form.originatorDept))?.label || form.originatorDept;
        const endpoint = `${siteUrl}/_api/web/lists/getbytitle('ApproverMaster')/items?$filter=(FunctionName eq '${funcName}' or FunctionName eq '${form.originatorDept}' or FunctionID eq '${form.originatorDept}')&$select=Id,ApproverName/Id,ApproverName/Title&$expand=ApproverName`;
        const response = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
        const data = await response.json();

        console.log("ApproverMaster Data:", data);
        if (data.value) {
          const options = data.value.map((item: any) => ({
            //  label: item.Approver.Title,
            //  value: String(item.Approver.Id)
            label: item.ApproverName.Title,
            value: String(item.ApproverName.Id)
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
  const fetchMatchingMatrix = async (manualTrigger = false, amountVal?: string) => {
    const costToUse = amountVal !== undefined ? amountVal : form.packageAwardedCost;
    const rawAmount = String(costToUse || "0").replace(/,/g, '');
    const amount = parseFloat(rawAmount) || 0;

    if (form.workflow !== 'AUTO' || !form.originatorDept || amount === 0) {
      setDoaMatrixOptions([]);
      setDoaMatrixApprovers([]);
      updateField('doaMatrix', '');
      return;
    }

    try {
      const shouldAlert = manualTrigger;

      const siteUrl = context.pageContext.web.absoluteUrl;
      const funcName = functionOptions.find(f => f.value === String(form.originatorDept))?.label || form.originatorDept;
      const filter = `(Function eq '${funcName}' or Function eq '${form.originatorDept}' or FunctionID eq '${form.originatorDept}') and MinAmount le ${amount} and MaxAmount ge ${amount}`;
      const endpoint = `${siteUrl}/_api/web/lists/getbytitle('DOAMatrix')/items?$filter=${filter}`;

      const response = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
      if (!response.ok) return;
      const data = await response.json();
      if (data.value && data.value.length > 0) {
        const options = data.value.map((item: any) => ({
          label: item.Title,
          value: item.Title
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
            messages: [`No DOA Matrix rule found for '${funcName}' at the amount of ${amount}.`, "Please verify your DOA list or switch to MANUAL workflow."]
          });
          lastDoaCheck.current = { dept: form.originatorDept, amount: rawAmount };
        }
      }
    } catch (err) {
      console.error("Error in fetchMatchingMatrix:", err);
    }
  };

  useEffect(() => {
    if (isInitialMount.current) {
      isInitialMount.current = false;
      fetchMatchingMatrix(false);
      return;
    }

    if (form.workflow === 'AUTO') {
      fetchMatchingMatrix(true);
    }
  }, [form.workflow, form.originatorDept]);

  useEffect(() => {
    const timer = setTimeout(() => {
      fetchMatchingMatrix(false);
    }, 500);
    return () => clearTimeout(timer);
  }, [form.packageAwardedCost]);

  // --- 5. AUTO WORKFLOW: POPULATE APPROVERS ---
  useEffect(() => {
    if (form.workflow !== 'AUTO' || !form.doaMatrix) return;
    const fetchApproversFromMatrix = async () => {
      try {
        const siteUrl = context.pageContext.web.absoluteUrl;
        const levels = Array.from({ length: 12 }, (_, i) => `Approver${i + 1}`);
        const select = levels.map(l => `${l}/Id,${l}/Title`).join(',');
        const expand = levels.join(',');
        const filter = `Title eq '${form.doaMatrix}'`;
        const endpoint = `${siteUrl}/_api/web/lists/getbytitle('DOAMatrix')/items?$filter=${filter}&$select=${select}&$expand=${expand}`;
        const response = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
        if (!response.ok) return;
        const data = await response.json();
        if (data.value && data.value.length > 0) {
          const item = data.value[0];
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
          if (matrixApprovers[0]) {
            clearFieldError('approvers');
          }
        }
      } catch (err) {
        console.error("Error in fetchApproversFromMatrix:", err);
      }
    };
    fetchApproversFromMatrix();
  }, [form.doaMatrix, form.workflow]);

  // --- HELPERS ---
  const clearFieldError = (field: string) => {
    setErrors(prev => prev.filter(e => e.field !== field));
  };

  const handleProjectChange = async (projectId: string) => {
    clearFieldError('project');
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

  const updateField = (key: string, val: any) => {
    clearFieldError(key);
    if (key === 'workflow') {
      clearFieldError('approvers');
    }
    setForm(f => ({ ...f, [key]: val }));
  };

  const updateApprover = (i: number, val: string) => {
    if (i === 0) {
      clearFieldError('approvers');
    }
    const arr = [...form.approvers];
    arr[i] = val;
    setForm(f => ({ ...f, approvers: arr }));
  };

  const addApproverSlot = () => {
    setManualSlotsCount(prev => Math.min(12, prev + 1));
  };

  const removeApproverSlot = (idxToRemove: number) => {
    const updated = [...form.approvers];
    updated.splice(idxToRemove, 1); // Remove the element at the index
    updated.push(null); // Keep length 12
    setForm(prev => ({ ...prev, approvers: updated }));
    setManualSlotsCount(prev => Math.max(1, prev - 1));
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

  const getFieldError = (field: string) => {
    return errors.find(e => e.field === field)?.message;
  };

  if (isLoading) return <NFASkeleton />;

  return (
    <div className="pageContainer">
      <PageTitle
        title={id ? "Edit NFA" : "New NFA"}
        subtitle={id ? "Modify existing Note for Approval details." : "Create a new Note for Approval."}
      />

      <CollapsibleSection icon={<Building2 size={13} />} title="Basic Information" hasError={hasSectionErrors(['project', 'projectSubEntity', 'originatorDept'])}>
        <div className="flex flex-wrap gap-x-6 gap-y-4 mb-4">
          <div className="flex-1 min-w-[220px]">
            <DropdownField label="Entity" value={form.project} options={projectOptions} onChange={handleProjectChange} emptyLabel="Select Entity..." error={getFieldError('project')} />
          </div>
          <div className="flex-1 min-w-[220px]">
            <MultiSelectDropdown label="Sub-Entity" options={subEntityOptions} selectedValues={form.projectSubEntity} onChange={v => updateField('projectSubEntity', v)} disabled={loadingSubEntity || !form.project} required error={getFieldError('projectSubEntity')} />
          </div>
        </div>
        <div className="flex flex-wrap gap-x-6 gap-y-4">
          <div className="flex-1 min-w-[220px]">
            <DropdownField label="Function" value={form.originatorDept} options={functionOptions} onChange={v => updateField('originatorDept', v)} emptyLabel="Select Function..." error={getFieldError('originatorDept')} />
          </div>
          <div className="flex-1 min-w-[220px]">
            <InputField label="Date" value={form.date} onChange={() => { }} disabled />
          </div>
        </div>
      </CollapsibleSection>

      <CollapsibleSection icon={<FileText size={13} />} title="NFA Details" defaultOpen={false} hasError={hasSectionErrors(['referenceSubject', 'introduction', 'chronology', 'commercialProposal', 'paymentTerms'])}>
        <div className="mb-4"><TextAreaField label="Subject" value={form.referenceSubject} onChange={v => updateField('referenceSubject', v)} rows={2} required error={getFieldError('referenceSubject')} /></div>
        <div className="mb-4"><RichTextEditor label="Introduction" value={form.introduction} onChange={v => updateField('introduction', v)}  error={getFieldError('introduction')} /></div>
        <div className="flex gap-4 mb-4">
          <div className="flex-1 min-w-0">
            <InputField label="Complete Indent/PR received from & Date" type="date" value={form.completeIndentPRReceivedFromDate || ''} onChange={(v) => updateField('completeIndentPRReceivedFromDate', v)}  error={getFieldError('completeIndentPRReceivedFromDate')} />

          </div>
          <div className="flex-1 min-w-0">
            <InputField label="BOQ/Specifications Revisions by & Date " type="date" value={form.boqSpecificationsRevisionsByDate || ''} onChange={(v) => updateField('boqSpecificationsRevisionsByDate', v)} />
          </div>
          <div className="flex-1 min-w-0">
            <InputField label="RFQ Ref No" value={form.RFQRefNo || ''} onChange={(v) => updateField('RFQRefNo', v)} />
          </div>
        </div>
        <div className="mb-4"><InputField label="Remarks" value={form.remark || ''} onChange={(v) => updateField('remark', v)} /></div>

        <div className="mb-4"><RichTextEditor label="Chronology" value={form.chronology} onChange={v => updateField('chronology', v)}  error={getFieldError('chronology')} /></div>
        <div className="mb-4"><RichTextEditor label="Commercial Proposals" value={form.commercialProposal} onChange={v => updateField('commercialProposal', v)}  error={getFieldError('commercialProposal')} /></div>
        <RichTextEditor label="Payment Terms" value={form.paymentTerms} onChange={v => updateField('paymentTerms', v)} />
      </CollapsibleSection>

      <CollapsibleSection icon={<Wrench size={13} />} title="Contract Terms" defaultOpen={false} hasError={hasSectionErrors(['liquidatedDamage', 'contractPerf', 'defectLiability', 'workCompletion'])}>
        <div className="mb-4"><TextAreaField label="Liquidated Damage" value={form.liquidatedDamage} onChange={v => updateField('liquidatedDamage', v)} required rows={3} /></div>
        <div className="mb-4"><TextAreaField label="Contract Performance Gurantee" value={form.contractPerf} onChange={v => updateField('contractPerf', v)} required rows={3} /></div>
        <div className="mb-4"><TextAreaField label="Defect Liability Period" value={form.defectLiability} onChange={v => updateField('defectLiability', v)} required rows={3} /></div>
        <TextAreaField label="Work Completion" value={form.workCompletion} onChange={v => updateField('workCompletion', v)} required rows={3} />
      </CollapsibleSection>

      <CollapsibleSection icon={<DollarSign size={13} />} title="Financials" defaultOpen={false} hasError={hasSectionErrors(['budgetRemarks', 'packageAwardedCost', 'budget'])}>
        <div className="mb-4"><TextAreaField label="Budget Remarks" value={form.budgetRemarks} onChange={v => updateField('budgetRemarks', v)} rows={3} /></div>
        <div className="flex gap-x-6 mb-4">
          <div className="flex-1"><NumberInput label="Packeage Awarded Cost" value={form.packageAwardedCost} onChange={v => updateField('packageAwardedCost', v)} onBlur={v => fetchMatchingMatrix(true, v)} required error={getFieldError('packageAwardedCost')} /></div>
          <div className="flex-1"><NumberInput label="Budget" value={form.budget} onChange={v => updateField('budget', v)} required error={getFieldError('budget')} /></div>
          <div className="flex-1">
            <DropdownField
              label="Budgeted/ Non Budgeted"
              value={form.budgetedNonBudgeted}
              options={[
                { label: "Budgeted", value: "Budgeted" },
                { label: "Non Budgeted", value: "Non Budgeted" }
              ]}
              onChange={(value) =>
                setForm((prev) => ({
                  ...prev,
                  budgetedNonBudgeted: value
                }))
              }
              emptyLabel="Select Budget Type..."
              error={getFieldError("budgetedNonBudgeted")}
              required 
            />
          </div>
        </div>
        {/* <div className="mb-4"><TextAreaField label="Capex No" value={form.capexSerialNo} onChange={v => updateField('capexSerialNo', v)} rows={3} /></div> */}
        {financialsSummary.budget > 0 && (
          <div className="mt-4 p-4 bg-slate-50/50 rounded-2xl border border-slate-200/80 flex flex-wrap gap-4 items-center justify-between">
            <div className="flex flex-col">
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Variance (Budget - Cost)</span>
              <span className={`text-sm font-extrabold ${financialsSummary.variance >= 0 ? 'text-emerald-600' : 'text-rose-600'}`}>
                {financialsSummary.variance >= 0 ? '+' : ''}
                {financialsSummary.variance.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
              </span>
            </div>
            <div className="flex flex-col items-end">
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Budget Utilization</span>
              <div className="flex items-center gap-2 mt-1">
                <div className="w-24 h-2 bg-slate-200 rounded-full overflow-hidden">
                  <div
                    className={`h-full rounded-full transition-all duration-500 ${financialsSummary.percent > 100 ? 'bg-rose-500' : 'bg-[#C6A854]'}`}
                    style={{ width: `${Math.min(100, financialsSummary.percent)}%` }}
                  />
                </div>
                <span className="text-xs font-bold text-slate-700">{financialsSummary.percent}%</span>
              </div>
            </div>
          </div>
        )}
      </CollapsibleSection>

      <CollapsibleSection icon={<Star size={13} />} title="Recommendation" defaultOpen={false} hasError={hasSectionErrors(['recommendation', 'contractorName'])}>
        <div className="mb-4"><RichTextEditor label="Recommendation with Justification" value={form.recommendation} onChange={v => updateField('recommendation', v)} /></div>
        <InputField label="Contractor Name" value={form.contractorName} placeholder='Enter Contractor Name' onChange={v => updateField('contractorName', v)} />
        <div className="mb-4"><TextAreaField label="Annexure/ Enclosur(if any)" placeholder='Enter Annexure/ Enclosur (if any)' value={form.annexure} onChange={v => updateField('annexure', v)} rows={3} /></div>
        <div className="mb-4"><TextAreaField label="Note 1" placeholder='Enter Note 1' value={form.note1} onChange={v => updateField('note1', v)} rows={2} /></div>
        <div className="mb-4"><TextAreaField label="Note 2" placeholder='Enter Note 2' value={form.note2} onChange={v => updateField('note2', v)} rows={2} /></div>
        <div className="mb-4"><TextAreaField label="Note 3" placeholder='Enter Note 3' value={form.note3} onChange={v => updateField('note3', v)} rows={2} /></div>
        <div className="mb-4"><TextAreaField label="Note 4" placeholder='Enter Note 4' value={form.note4} onChange={v => updateField('note4', v)} rows={2} /></div>

      </CollapsibleSection>

      <CollapsibleSection icon={<Settings size={13} />} title="Workflow & Approvers" defaultOpen={false} hasError={hasSectionErrors(['workflow', 'doaMatrix', 'approvers'])}>
        <div className="flex gap-x-6 mb-6">
          <div className="flex-1">
            <DropdownField label="Workflow" value={form.workflow} options={workflowOptions} onChange={v => updateField('workflow', v)} showEmptyOption={false} />
          </div>
          <div className="flex-1">
            <DropdownField
              label="DOA Matrix"
              value={form.doaMatrix}
              options={doaMatrixOptions}
              onChange={v => updateField('doaMatrix', v)}
              disabled={form.workflow !== 'AUTO'}
              emptyLabel="Select DOA Matrix..."
              error={form.workflow === 'AUTO' ? getFieldError('approvers') : undefined}
            />
          </div>
        </div>

        {form.workflow === 'AUTO' ? (
          <>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '20px' }}>
              {form.approvers.map((app, idx) => {
                if (!app) return null;
                return (
                  <DropdownField
                    key={idx}
                    label={`Approver Level ${idx + 1}`}
                    value={String(app ?? "")}
                    options={doaMatrixApprovers}
                    onChange={(v) => updateApprover(idx, v)}
                    disabled
                    emptyLabel="None"
                  />
                );
              })}
            </div>
            {form.doaMatrix && (
              <div className="mt-6 p-4 bg-slate-50/50 rounded-2xl border border-slate-200/80">
                <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-3">Auto Approval Chain Preview</p>
                <div className="flex flex-col gap-2">
                  {form.approvers.map((app, idx) => {
                    if (!app) return null;
                    const appObj = doaMatrixApprovers.find(o => o.value === app);
                    const name = appObj ? appObj.label : `Approver ID: ${app}`;
                    return (
                      <div key={idx} className="flex items-center gap-3 py-1">
                        <div className="w-6 h-6 rounded-full bg-[#C6A854]/10 text-[#C6A854] border border-[#C6A854]/20 flex items-center justify-center font-bold text-[10px] shrink-0">
                          {idx + 1}
                        </div>
                        <div className="flex items-center gap-1.5">
                          <span className="text-xs font-semibold text-slate-500">Level {idx + 1}:</span>
                          <span className="text-xs font-bold text-slate-800">{name}</span>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </>
        ) : (
          <div className="flex flex-col gap-4">
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '20px' }}>
              {Array.from({ length: manualSlotsCount }).map((_, idx) => {
                const app = form.approvers[idx];
                return (
                  <div key={idx} className="flex items-end gap-2 relative group">
                    <div className="flex-1">
                      <DropdownField
                        label={`Approver Level ${idx + 1}`}
                        value={String(app ?? "")}
                        options={matrixApproverOptions}
                        onChange={(v) => updateApprover(idx, v)}
                        emptyLabel="Select Approver..."
                        error={idx === 0 ? getFieldError('approvers') : undefined}
                      />
                    </div>
                    {manualSlotsCount > 1 && idx === manualSlotsCount - 1 && (
                      <button
                        type="button"
                        onClick={() => removeApproverSlot(idx)}
                        className="mb-1 p-2 rounded-lg text-slate-400 hover:text-rose-500 hover:bg-rose-50 border border-slate-200 hover:border-rose-100 transition-all cursor-pointer"
                        title={`Remove Level ${idx + 1}`}
                      >
                        <Trash2 size={16} />
                      </button>
                    )}
                  </div>
                );
              })}
            </div>
            {manualSlotsCount < 12 && (
              <div className="flex justify-start mt-2">
                <button
                  type="button"
                  onClick={addApproverSlot}
                  className="flex items-center gap-2 px-4 py-2 text-xs font-bold text-[#C6A854] bg-[#C6A854]/5 hover:bg-[#C6A854]/10 rounded-xl border border-[#C6A854]/20 hover:border-[#C6A854]/30 transition-all cursor-pointer"
                >
                  + Add Approver Level
                </button>
              </div>
            )}
          </div>
        )}
      </CollapsibleSection>

      <CollapsibleSection icon={<Paperclip size={13} />} title="Attachments" >
        {/* Attachment Guidelines Banner */}
        <div
          style={{
            display: 'flex',
            alignItems: 'flex-start',
            gap: '8px',
            padding: '10px 14px',
            backgroundColor: 'rgba(254, 242, 242, 0.7)',
            border: '1px solid #fee2e2',
            borderRadius: '10px',
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
          onPaste={(e) => {
            if (e.clipboardData && e.clipboardData.files && e.clipboardData.files.length > 0) {
              e.preventDefault();
              e.stopPropagation();
              handleAttachmentPasteOrDrop(Array.from(e.clipboardData.files));
            }
          }}
          onDrop={(e) => {
            e.preventDefault();
            e.stopPropagation();
            if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
              handleAttachmentPasteOrDrop(Array.from(e.dataTransfer.files));
            }
          }}
          onDragOver={(e) => {
            e.preventDefault();
            e.stopPropagation();
          }}
        >
          <div className="flex justify-end mb-4"><CommonButton text="+ Add Attachment" variant="primary" onClick={() => setIsPopupOpen(true)} /></div>
          {(savedDocuments.length > 0 || pendingAttachments.length > 0) ? (
            <div className="border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
              <table className="savedDocsTable w-full border-collapse text-left">
                <thead>
                  <tr className="bg-slate-50/80 border-b border-slate-200 text-xs font-bold text-slate-500">
                    <th className="p-3">#</th>
                    <th className="p-3">Type</th>
                    <th className="p-3">File Name</th>
                    <th className="p-3">Status</th>
                    <th className="p-3 text-center">Action</th>
                  </tr>
                </thead>
                <tbody>
                  {savedDocuments.map((doc, i) => (
                    <tr key={doc.id} className="border-t border-slate-100 hover:bg-slate-50/50 transition-colors text-sm">
                      <td className="p-3 font-semibold text-slate-500">{i + 1}</td>
                      <td className="p-3 flex items-center gap-2">
                        {getFileBadge(doc.fileName)}
                        <span className="text-xs font-bold text-slate-600">{doc.documentType}</span>
                      </td>
                      <td className="p-3 font-semibold text-blue-600 hover:underline cursor-pointer">{doc.fileName}</td>
                      <td className="p-3">
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-bold bg-emerald-50 text-emerald-700 border border-emerald-100">
                          Saved
                        </span>
                      </td>
                      <td className="p-3 text-center">
                        <button onClick={async () => { if (confirm("Are you sure you want to delete this document?")) { const s = new NFADocService(context); await s.deleteDocument(doc.id); setSavedDocuments(prev => prev.filter(d => d.id !== doc.id)); } }} className="text-red-500 hover:text-red-700 transition-colors">
                          <Trash2 size={14} />
                        </button>
                      </td>
                    </tr>
                  ))}
                  {pendingAttachments.map((entry, i) => (
                    <tr key={`pending-${i}`} className="border-t border-slate-100 bg-blue-50/10 hover:bg-blue-50/20 transition-colors text-sm">
                      <td className="p-3 font-semibold text-slate-500">{savedDocuments.length + i + 1}</td>
                      <td className="p-3 flex items-center gap-2">
                        {getFileBadge(entry.file.name)}
                        <span className="text-xs font-bold text-slate-600">{entry.documentType}</span>
                      </td>
                      <td className="p-3 font-medium text-slate-500 italic">{entry.file.name}</td>
                      <td className="p-3">
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-bold bg-amber-50 text-amber-700 border border-amber-100 animate-pulse">
                          Pending
                        </span>
                      </td>
                      <td className="p-3 text-center">
                        <button onClick={() => setPendingAttachments(prev => prev.filter((_, idx) => idx !== i))} className="text-red-500 hover:text-red-700 transition-colors">
                          <Trash2 size={14} />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : <p className="noDocsMessage px-2">No documents attached.</p>}
        </div>
      </CollapsibleSection>

      {id && history.length > 0 && (
        <CollapsibleSection icon={<History size={13} />} title="Workflow History" defaultOpen={false}>
          <div className="space-y-4 max-w-2xl px-2 py-1">
            {history.map((item, idx) => {
              const isApproved = item.action === 'Approved';
              const isRejected = item.action === 'Rejected';
              const isQuery = item.action === 'Query';

              let dotClass = 'bg-blue-500';
              if (isApproved) {
                dotClass = 'bg-emerald-500';
              } else if (isRejected) {
                dotClass = 'bg-rose-500';
              } else if (isQuery) {
                dotClass = 'bg-amber-500';
              }

              return (
                <div key={item.id || idx} className="relative pl-7 pb-2 last:pb-0">
                  {idx < history.length - 1 && (
                    <div className="absolute left-[7px] top-5 bottom-0 w-[1.5px] bg-slate-100" />
                  )}
                  <div className={`absolute left-0 top-1.5 w-3.5 h-3.5 rounded-full border border-white flex items-center justify-center z-10 ${dotClass} shadow-sm`} />
                  <div className="flex flex-col min-w-0">
                    <p className="text-[11px] text-slate-600 font-medium leading-relaxed">
                      {item.rawText || `${item.user} (${item.action})${item.remarks ? `: ${item.remarks}` : ""}`}
                    </p>
                  </div>
                </div>
              );
            })}
          </div>
        </CollapsibleSection>
      )}

      {id && comments.filter(item => item.remarks && item.remarks !== "No remarks" && item.remarks.trim() !== "").length > 0 && (
        <CollapsibleSection icon={<MessageSquare size={13} />} title="Comments History" defaultOpen={false}>
          <div className="space-y-4 max-w-2xl px-2 py-1">
            {comments.filter(item => item.remarks && item.remarks !== "No remarks" && item.remarks.trim() !== "").map((item, idx) => (
              <div key={item.id || idx} className="flex flex-col gap-1 p-2.5 bg-slate-50/60 hover:bg-slate-50 rounded-xl border border-slate-100 transition-colors">
                <div className="flex items-center justify-between gap-2">
                  <span className="text-[10px] font-medium text-slate-600 text-left">Comment: {item.remarks || ""}</span>
                </div>
                <div className="p-2 bg-white rounded-lg border border-slate-100 relative shadow-sm">
                  <div className="absolute -top-1 left-3 w-2.5 h-2.5 bg-white border-t border-l border-slate-100 rotate-45" />
                  <p className="text-[10px] text-slate-500 leading-normal relative z-10">
                    Remarks Added by {item.user} at Level {item.level || "1"} on {item.date}
                  </p>
                </div>
              </div>
            ))}
          </div>
        </CollapsibleSection>
      )}

      <AttachmentPopup
        isOpen={isPopupOpen}
        onClose={() => setIsPopupOpen(false)}
        onSave={(entries) => {
          setPendingAttachments(prev => [...prev, ...entries]);
          setIsPopupOpen(false);
        }}
        existingFileNames={[
          ...savedDocuments.map(d => d.fileName),
          ...pendingAttachments.map(p => p.file.name)
        ]}
      />

      {isSaving && (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          {!savingProgress ? (
            <>
              <style>{`
                @keyframes indeterminate-progress {
                  0% { left: -35%; right: 100%; }
                  60% { left: 100%; right: -90%; }
                  100% { left: 100%; right: -90%; }
                }
                @keyframes indeterminate-progress-short {
                  0% { left: -200%; right: 100%; }
                  60% { left: 107%; right: -8%; }
                  100% { left: 107%; right: -8%; }
                }
                .animate-progress-1 {
                  animation: indeterminate-progress 2s infinite ease-out;
                }
                .animate-progress-2 {
                  animation: indeterminate-progress-short 2s infinite ease-out;
                  animation-delay: 1.15s;
                }
              `}</style>

              <div className="bg-white/95 rounded-3xl p-8 max-w-sm w-full border border-slate-200/80 shadow-2xl flex flex-col items-center justify-center gap-5 text-center transform scale-100 transition-all duration-300">
                <div className="relative flex items-center justify-center h-16 w-16">
                  <span className="animate-ping absolute inline-flex h-12 w-12 rounded-full bg-blue-400 opacity-20"></span>
                  <div className="h-14 w-14 rounded-full border-4 border-slate-100 border-t-blue-600 animate-spin flex items-center justify-center">
                    <FileText size={20} className="text-blue-600" />
                  </div>
                </div>

                <div className="space-y-1.5">
                  <h3 className="font-black text-slate-800 text-base">Saving NFA</h3>
                  <p className="text-[11px] text-slate-450 font-bold uppercase tracking-wider">
                    Please wait while we upload your attachments and save the request
                  </p>
                </div>

                <div className="relative w-full h-1.5 bg-slate-100 rounded-full overflow-hidden">
                  <div className="absolute top-0 bottom-0 left-0 bg-gradient-to-r from-blue-500 to-indigo-600 rounded-full animate-progress-1"></div>
                  <div className="absolute top-0 bottom-0 left-0 bg-gradient-to-r from-blue-400 to-blue-500 rounded-full animate-progress-2"></div>
                </div>
              </div>
            </>
          ) : (
            <div className="bg-white/95 rounded-3xl p-8 max-w-md w-full border border-slate-200/80 shadow-2xl flex flex-col items-center justify-center gap-5 text-center transform scale-100 transition-all duration-300">
              <div className="relative flex items-center justify-center h-16 w-16">
                <span className="animate-ping absolute inline-flex h-12 w-12 rounded-full bg-blue-400 opacity-20"></span>
                <div className="h-14 w-14 rounded-full border-4 border-slate-100 border-t-blue-600 animate-spin flex items-center justify-center">
                  <FileText size={20} className="text-blue-600" />
                </div>
              </div>

              <div className="space-y-1.5 w-full">
                <h3 className="font-black text-slate-800 text-base">
                  {savingProgress.step === 'attachments' ? "Uploading Attachments" : "Saving Form Data"}
                </h3>
                <div className="text-[11px] text-slate-400 font-bold uppercase tracking-wider">
                  {savingProgress.step === 'attachments' ? (
                    <div className="space-y-1">
                      <div>Uploading file {savingProgress.currentFileIndex} of {savingProgress.totalFiles}</div>
                      <div className="text-slate-600 font-semibold truncate max-w-xs mx-auto normal-case text-xs">
                        "{savingProgress.currentFileName}"
                      </div>
                    </div>
                  ) : (
                    <div>Saving details and updating status...</div>
                  )}
                </div>
              </div>

              <div className="w-full space-y-1">
                <div className="relative w-full h-2 bg-slate-100 rounded-full overflow-hidden">
                  <div
                    className="absolute top-0 bottom-0 left-0 bg-gradient-to-r from-blue-500 to-indigo-600 rounded-full transition-all duration-300"
                    style={{ width: `${savingProgress.percent}%` }}
                  ></div>
                </div>
                <div className="flex justify-end text-[10px] font-black text-slate-500">
                  {savingProgress.percent}%
                </div>
              </div>
            </div>
          )}
        </div>
      )}

      <MessagePopup isOpen={doaAlert.show} type={doaAlert.type} title={doaAlert.title} messages={doaAlert.messages} onClose={handlePopupClose} />
      {errors.length > 0 && <ValidationErrors errors={errors} onClose={() => setErrors([])} />}

      <div className="flex justify-center gap-4 py-8">
        <CommonButton text={id ? "Update NFA" : "Submit NFA"} variant="primary" disabled={isSaving} onClick={async () => {
          if (!validate('Pending')) return;
          setIsSaving(true);
          setSavingProgress({
            step: 'attachments',
            currentFileIndex: 1,
            totalFiles: pendingAttachments.length,
            currentFileName: pendingAttachments[0]?.file.name || '',
            percent: 0
          });
          try {
            const docService = new NFADocService(context);
            if (pendingAttachments.length > 0) {
              await docService.saveDocuments(pendingAttachments, form.uniqId,undefined, (current: number, total: number, fileName: string) => {
                const percent = Math.round((current / total) * 100);
                setSavingProgress({
                  step: 'attachments',
                  currentFileIndex: Math.min(current + 1, total),
                  totalFiles: total,
                  currentFileName: fileName,
                  percent
                });
              });
            }
            setSavingProgress({
              step: 'form',
              currentFileIndex: 0,
              totalFiles: 0,
              currentFileName: '',
              percent: 90
            });
            const service = new NFAService(context);
            const res = id ? await service.updateNFA(parseInt(id), form, 'Pending', projectOptions, subEntityOptions, functionOptions) : await service.saveNFA(form, 'Pending', projectOptions, subEntityOptions, functionOptions);

            setSavingProgress({
              step: 'complete',
              currentFileIndex: 0,
              totalFiles: 0,
              currentFileName: '',
              percent: 100
            });
            setDoaAlert({
              show: true,
              type: res.success ? 'success' : 'error',
              title: res.success ? 'Success' : 'Error',
              messages: [res.message],
              onSuccessRedirect: res.success
            });
          } catch (err) {
            console.error("Error submitting NFA:", err);
            const errMsg = err instanceof Error ? err.message : "Unknown error occurred.";
            setDoaAlert({
              show: true,
              type: 'error',
              title: 'Error Saving NFA',
              messages: ["Error saving NFA or Uploading Documents.", errMsg],
              onSuccessRedirect: false
            });
          } finally {
            setIsSaving(false);
            setSavingProgress(null);
          }
        }} />
        <CommonButton text="Save Draft" variant="secondary" disabled={isSaving} onClick={async () => {
          if (!validate('Draft')) return;
          setIsSaving(true);
          setSavingProgress({
            step: 'attachments',
            currentFileIndex: 1,
            totalFiles: pendingAttachments.length,
            currentFileName: pendingAttachments[0]?.file.name || '',
            percent: 0
          });
          try {
            const docService = new NFADocService(context);
            if (pendingAttachments.length > 0) {
              await docService.saveDocuments(pendingAttachments, form.uniqId,undefined, (current, total, fileName) => {
                const percent = Math.round((current / total) * 100);
                setSavingProgress({
                  step: 'attachments',
                  currentFileIndex: Math.min(current + 1, total),
                  totalFiles: total,
                  currentFileName: fileName,
                  percent
                });
              });
            }
            setSavingProgress({
              step: 'form',
              currentFileIndex: 0,
              totalFiles: 0,
              currentFileName: '',
              percent: 90
            });
            const service = new NFAService(context);
            const res = id ? await service.updateNFA(parseInt(id), form, 'Draft', projectOptions, subEntityOptions, functionOptions) : await service.saveNFA(form, 'Draft', projectOptions, subEntityOptions, functionOptions);

            setSavingProgress({
              step: 'complete',
              currentFileIndex: 0,
              totalFiles: 0,
              currentFileName: '',
              percent: 100
            });
            setDoaAlert({
              show: true,
              type: res.success ? 'success' : 'error',
              title: res.success ? 'Success' : 'Error',
              messages: [res.message],
              onSuccessRedirect: res.success
            });
          } catch (err) {
            console.error("Error saving NFA draft:", err);
            const errMsg = err instanceof Error ? err.message : "Unknown error occurred.";
            setDoaAlert({
              show: true,
              type: 'error',
              title: 'Error Saving Draft',
              messages: ["Error saving Draft or Uploading Documents.", errMsg],
              onSuccessRedirect: false
            });
          } finally {
            setIsSaving(false);
            setSavingProgress(null);
          }
        }} />
        <CommonButton text="Cancel" variant="danger" onClick={() => {
          forceScrollToTop();
          navigate('/');
        }} />
      </div>
    </div>
  );
};

export default NewNFA;
