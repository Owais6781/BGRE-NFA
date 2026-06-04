import * as React from 'react';
import { SPHttpClient } from "@microsoft/sp-http";
import { useState, useEffect, useRef } from 'react';
import { Building2, FileText, Wrench, DollarSign, Paperclip, Settings, Star, ChevronDown, ArrowLeft
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
import { NFADocService, ISavedDocument } from '../../../../../Share/Services/NFADocService';
import { NFAService, INFAFormData } from '../../../../../Share/Services/NFAService';
import { useNavigate, useParams } from 'react-router-dom';
import "../../../../../Share/styles/global.scss";

// --- Component: Collapsible Section (Shared) ---
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

const NFASkeleton = () => (
    <div className="pageContainer">
      <div className="flex justify-center mb-8"><Skeleton width="50%" height="2.5rem" /></div>
      {[1, 2, 3].map((i) => (
        <div key={i} className="formSection mb-6">
          <div className="formSectionHeader"><div className="flex items-center gap-2"><Skeleton shape="circle" size="1.2rem" /><Skeleton width="120px" height="1rem" /></div></div>
          <div className="formSectionBody"><div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div><Skeleton width="80px" height="0.8rem" className="mb-2" /><Skeleton width="100%" height="2.5rem" borderRadius="6px" /></div>
              <div><Skeleton width="80px" height="0.8rem" className="mb-2" /><Skeleton width="100%" height="2.5rem" borderRadius="6px" /></div>
          </div></div>
        </div>
      ))}
    </div>
);

interface IViewNFAProps {
  context: WebPartContext;
  sp: SPFI;
}

const ViewNFA: React.FC<IViewNFAProps> = ({ context, sp }) => {
  const { id } = useParams();
  const navigate = useNavigate();

  const hasInitialized = useRef(false);
  const [isLoading, setIsLoading] = useState(true);
  const [form, setForm] = useState<INFAFormData | null>(null);
  const [savedDocuments, setSavedDocuments] = useState<ISavedDocument[]>([]);
  const [projectOptions, setProjectOptions] = useState<IDropdownOption[]>([]);
  const [subEntityOptions, setSubEntityOptions] = useState<IDropdownOption[]>([]);
  const [functionOptions, setFunctionOptions] = useState<IDropdownOption[]>([]);
  const [approverOptions, setApproverOptions] = useState<IDropdownOption[]>([]);

  useEffect(() => {
    if (hasInitialized.current) return;

    const loadData = async () => {
      console.log("ViewNFA: Starting to load data for ID:", id);
      const dropdownService = new DropdownService(context);
      const nfaService = new NFAService(context);
      const docService = new NFADocService(context);

      try {
        if (!id) { console.error("ViewNFA: No ID found in URL"); navigate('/'); return; }

        const [projects, funcs, nfaData] = await Promise.all([
          dropdownService.getProjects(),
          dropdownService.getFunctionMaster(),
          nfaService.getNFA(parseInt(id))
        ]);

        console.log("ViewNFA: Basic data loaded. NFA Data found:", !!nfaData);

        if (nfaData) {
          // 1. Find the Project ID that matches the saved Name (Project Name is stored in list)
          const matchedProject = projects.find(p => p.label === nfaData.project);
          const pId = matchedProject ? matchedProject.value : "";
          
          console.log("ViewNFA: Matched Project ID:", pId);

          // 2. Load specific sub-entities and documents
          const [docs, subs] = await Promise.all([
            docService.getDocumentsByUniqId(nfaData.uniqId),
            pId ? dropdownService.getSubEntity(pId) : Promise.resolve([])
          ]);

          console.log("ViewNFA: Related data loaded. Docs:", docs.length, "Subs:", subs.length);

          // 3. Fetch all site users for labels (since we just have IDs in approvers)
          const siteUrl = context.pageContext.web.absoluteUrl;
          const userEndpoint = `${siteUrl}/_api/web/siteusers?$filter=PrincipalType eq 1 and IsHiddenInUI eq false`;
          const userResponse = await context.spHttpClient.get(userEndpoint, SPHttpClient.configurations.v1);
          const userData = await userResponse.json();
          const userOpts = userData.value.map((u: any) => ({ label: u.Title, value: String(u.Id) }));

          setProjectOptions(projects);
          setFunctionOptions(funcs);
          setSavedDocuments(docs);
          setSubEntityOptions(subs);
          setApproverOptions(userOpts);
          setForm(nfaData);
          console.log("ViewNFA: State binding complete.");
        } else {
          console.warn("ViewNFA: NFA Record not found for ID:", id);
          alert("NFA Record not found.");
          navigate('/');
        }

        hasInitialized.current = true;
      } catch (err) {
        console.error("ViewNFA Initialization Failed:", err);
      } finally {
        setTimeout(() => setIsLoading(false), 100);
      }
    };
    loadData();
  }, [id]);

  if (isLoading || !form) return <NFASkeleton />;

  return (
    <div className="pageContainer">
      <div className="flex items-center gap-4 mb-6">
        <button className="iconBtn" onClick={() => navigate('/')}><ArrowLeft size={20}/></button>
        <p className="sectionTitle" style={{ marginBottom: 0 }}>View Note for Approval — {form.referenceSubject}</p>
      </div>

      <CollapsibleSection icon={<Building2 size={13} />} title="Basic Information">
        <div className="flex flex-wrap gap-x-6 gap-y-4 mb-4">
          <div className="flex-1 min-w-[220px]">
            <InputField label="Project" value={form.project} onChange={() => {}} disabled />
          </div>
          <div className="flex-1 min-w-[220px]">
             <InputField label="Sub-Entity" value={form.projectSubEntity.join("; ")} onChange={() => {}} disabled />
          </div>
        </div>
        <div className="flex flex-wrap gap-x-6 gap-y-4">
          <div className="flex-1 min-w-[220px]">
            <InputField label="Function" value={form.originatorDept} onChange={() => {}} disabled />
          </div>
          <div className="flex-1 min-w-[220px]">
            <InputField label="Date" value={form.date} onChange={() => {}} disabled />
          </div>
        </div>
      </CollapsibleSection>

      <CollapsibleSection icon={<FileText size={13} />} title="NFA Details" defaultOpen={false}>
        <div className="mb-4"><TextAreaField label="Subject" value={form.referenceSubject} onChange={() => {}} disabled rows={2} /></div>
        <div className="mb-4"><RichTextEditor label="Introduction" value={form.introduction} onChange={() => {}} readOnly /></div>
        <div className="mb-4"><RichTextEditor label="Chronology" value={form.chronology} onChange={() => {}} readOnly /></div>
        <div className="mb-4"><RichTextEditor label="Commercial Proposals" value={form.commercialProposal} onChange={() => {}} readOnly /></div>
        <RichTextEditor label="Payment Terms" value={form.paymentTerms} onChange={() => {}} readOnly />
      </CollapsibleSection>

      <CollapsibleSection icon={<Wrench size={13} />} title="Contract Terms" defaultOpen={false}>
        <div className="mb-4"><TextAreaField label="LD" value={form.liquidatedDamage} onChange={() => {}} disabled rows={3} /></div>
        <div className="mb-4"><TextAreaField label="CPG" value={form.contractPerf} onChange={() => {}} disabled rows={3} /></div>
        <div className="mb-4"><TextAreaField label="DLP" value={form.defectLiability} onChange={() => {}} disabled rows={3} /></div>
        <TextAreaField label="Work Completion" value={form.workCompletion} onChange={() => {}} disabled rows={3} />
      </CollapsibleSection>

      <CollapsibleSection icon={<DollarSign size={13} />} title="Financials" defaultOpen={false}>
        <div className="mb-4"><TextAreaField label="Remarks" value={form.budgetRemarks} onChange={() => {}} disabled rows={3} /></div>
        <div className="flex gap-x-6 mb-4">
          <div className="flex-1"><InputField label="Awarded Cost" value={form.packageAwardedCost} onChange={() => {}} disabled /></div>
          <div className="flex-1"><InputField label="Budget" value={form.budget} onChange={() => {}} disabled /></div>
        </div>
        <TextAreaField label="Capex No" value={form.capexSerialNo} onChange={() => {}} disabled rows={3} />
      </CollapsibleSection>

      <CollapsibleSection icon={<Star size={13} />} title="Recommendation" defaultOpen={false}>
        <div className="mb-4"><RichTextEditor label="Justification" value={form.recommendation} onChange={() => {}} readOnly /></div>
        <InputField label="Contractor Name" value={form.contractorName} onChange={() => {}} disabled />
      </CollapsibleSection>

      <CollapsibleSection icon={<Settings size={13} />} title="Workflow & Approvers" defaultOpen={false}>
        <div className="flex gap-x-6 mb-6">
          <div className="flex-1"><InputField label="Workflow" value={form.workflow} onChange={() => {}} disabled /></div>
          <div className="flex-1"><InputField label="DOA Matrix" value={form.doaMatrix} onChange={() => {}} disabled /></div>
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '16px' }}>
          {form.approvers.map((appId, idx) => {
             const user = approverOptions.find(o => o.value === String(appId));
             return (
               <InputField 
                 key={idx} 
                 label={`Approver Level ${idx + 1}`} 
                 value={user ? user.label : (appId ? `User ID: ${appId}` : "Not Assigned")} 
                 onChange={() => {}} 
                 disabled 
               />
             );
          })}
        </div>
      </CollapsibleSection>

      <CollapsibleSection icon={<Paperclip size={13} />} title="Attachments" defaultOpen={false}>
        {savedDocuments.length > 0 ? (
          <table className="savedDocsTable w-full border text-left">
            <thead><tr><th className="p-2">#</th><th className="p-2">Type</th><th className="p-2">File</th><th className="p-2">Status</th></tr></thead>
            <tbody>
              {savedDocuments.map((doc, i) => (
                <tr key={doc.id} className="border-t">
                  <td className="p-2">{i + 1}</td>
                  <td className="p-2">{doc.documentType}</td>
                  <td className="p-2 text-blue-600 font-medium">
                  
                  </td>
                  <td className="p-2 text-green-600 font-semibold">Attached</td>
                </tr>
              ))}
            </tbody>
          </table>
        ) : <p className="noDocsMessage">No documents attached.</p>}
      </CollapsibleSection>

      <div className="flex justify-center gap-4 py-8">
        <CommonButton text="Back to Home" variant="secondary" onClick={() => navigate('/')} />
      </div>
    </div>
  );
};

export default ViewNFA;
