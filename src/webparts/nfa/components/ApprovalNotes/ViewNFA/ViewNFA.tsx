import * as React from 'react';
import { useState, useEffect, useRef, useMemo } from 'react';
import { 
  Building2, FileText, Wrench, DollarSign, Paperclip, Settings, Star, ChevronDown, ArrowLeft,
  History, MessageSquare, ChevronRight
} from 'lucide-react';
import { Skeleton } from 'primereact/skeleton'; 
import 'primereact/resources/primereact.min.css';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users";
import { DropdownService, IDropdownOption } from '../../../../../Share/Services/DropdownService';
import { NFADocService, ISavedDocument } from '../../../../../Share/Services/NFADocService';
import { NFAService, INFAFormData } from '../../../../../Share/Services/NFAService';
import { useNavigate, useParams } from 'react-router-dom';
import "../../../../../Share/styles/global.scss";

interface IViewNFAProps {
  context: WebPartContext;
  sp: SPFI;
}

const getInitials = (name: string) => {
  if (!name || name === 'Pending') return 'UN';
  return name.split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();
};

const UserAvatar: React.FC<{ email?: string; initials: string; className: string }> = ({ email, initials, className }) => {
  const [imgFailed, setImgFailed] = useState(false);
  const photoUrl = email ? `/_layouts/15/userphoto.aspx?size=S&accountname=${email}` : null;

  if (photoUrl && !imgFailed) {
    return (
      <img 
        src={photoUrl} 
        alt={initials}
        onError={() => setImgFailed(true)}
        className={`${className} object-cover rounded-full`}
      />
    );
  }
  return (
    <div className={`${className} flex items-center justify-center font-bold`}>
      {initials}
    </div>
  );
};

const ViewNFA: React.FC<IViewNFAProps> = ({ context, sp }) => {
  const { id } = useParams();
  const navigate = useNavigate();

  const hasInitialized = useRef(false);
  const [isLoading, setIsLoading] = useState(true);
  const [form, setForm] = useState<INFAFormData | null>(null);
  const [savedDocuments, setSavedDocuments] = useState<ISavedDocument[]>([]);
  const [users, setUsers] = useState<{label: string, value: string, email?: string}[]>([]);
  const [history, setHistory] = useState<any[]>([]);
  const [comments, setComments] = useState<any[]>([]);

  const currentStepRef = useRef<HTMLDivElement>(null);
  const stepperContainerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (hasInitialized.current) return;

    const loadData = async () => {
      console.log("ViewNFA: Starting to load data for ID:", id);
      const nfaService = new NFAService(context);
      const docService = new NFADocService(context);

      try {
        if (!id) { console.error("ViewNFA: No ID found in URL"); navigate('/'); return; }

        const nfaData = await nfaService.getNFA(parseInt(id));
        console.log("ViewNFA: NFA Data found:", !!nfaData);

        if (nfaData) {
          const [docs, siteUsers, historyData] = await Promise.all([
            docService.getDocumentsByUniqId(nfaData.uniqId),
            sp.web.siteUsers.filter("PrincipalType eq 1").select("Id", "Title", "Email")()
              .then(res => res.map(u => ({ label: u.Title, value: String(u.Id), email: u.Email }))),
            nfaService.getHistory(parseInt(id))
          ]);

           setSavedDocuments(docs);
          setUsers(siteUsers);
          
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
              const match = cleanLine.match(/^\[(.*?)\]\s+(.*?)\s+\((.*?)\):\s+(.*)$/);
              if (match) {
                parsed.push({
                  id: Math.random(),
                  date: match[1],
                  user: match[2],
                  action: match[3],
                  remarks: match[4]
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

          let finalHistory = historyData;
          if (!finalHistory || finalHistory.length === 0) {
            if (nfaData.workflowHistory) {
              finalHistory = parseWorkflowHistory(nfaData.workflowHistory);
            } else if (nfaData.commentHistory) {
              finalHistory = parseCommentHistory(nfaData.commentHistory);
            }
          }
          setHistory(finalHistory.reverse());

          let finalComments: any[] = [];
          if (nfaData.commentHistory) {
            finalComments = parseCommentHistory(nfaData.commentHistory);
          }
          setComments(finalComments);

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

  // Auto scroll to current approval step on load
  useEffect(() => {
    if (!isLoading && currentStepRef.current && stepperContainerRef.current) {
      const timer = setTimeout(() => {
        if (currentStepRef.current && stepperContainerRef.current) {
          const container = stepperContainerRef.current;
          const element = currentStepRef.current;
          const containerWidth = container.offsetWidth;
          const elementOffset = element.offsetLeft;
          const elementWidth = element.offsetWidth;
          
          container.scrollTo({
            left: elementOffset - (containerWidth / 2) + (elementWidth / 2),
            behavior: 'smooth'
          });
        }
      }, 300);
      return () => clearTimeout(timer);
    }
  }, [isLoading]);

  const workflowSteps = useMemo(() => {
    if (!form) return [];
    const steps = [{ label: 'Initiator', id: form.initiatorId, type: 'initiator' }];
    
    form.approvers.forEach((appId, idx) => {
      if (appId) {
        steps.push({ label: `Level ${idx + 1}`, id: Number(appId), type: 'approver' });
      }
    });
    return steps;
  }, [form]);

  if (isLoading || !form) {
    return (
      <div className="pageContainer animate-pulse">
        <div className="flex items-center gap-3 mb-6">
          <Skeleton shape="circle" size="2.5rem" />
          <div className="flex-1">
            <Skeleton width="6rem" height="0.75rem" className="mb-2" />
            <Skeleton width="18rem" height="1.5rem" />
          </div>
        </div>
        <div className="flex flex-col lg:flex-row gap-6">
          <div className="w-full lg:w-[70%] space-y-6">
            <Skeleton height="150px" className="rounded-2xl" />
            <Skeleton height="250px" className="rounded-2xl" />
            <Skeleton height="200px" className="rounded-2xl" />
          </div>
          <div className="w-full lg:w-[30%] space-y-6">
            <Skeleton height="180px" className="rounded-2xl" />
            <Skeleton height="180px" className="rounded-2xl" />
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="pageContainer">
      
      {/* Approval Chain Card */}
      <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm py-3 px-6 mb-6">
        <div 
          ref={stepperContainerRef}
          className="relative flex items-center justify-start gap-1 overflow-x-auto pb-1.5 pt-0.5 no-scrollbar"
        >
          {workflowSteps.map((step, idx) => {
            const user = users.find(u => u.value === String(step.id));
            const isCompleted = idx < form.level;
            const isCurrent = idx === form.level;
            const isInitiator = idx === 0;
            const userLabel = user?.label || 'Pending';
            const initials = getInitials(userLabel);
            
            let borderClass = 'border-slate-200 bg-white';
            let avatarClass = 'bg-slate-100 text-slate-400';
            let dotClass = 'bg-slate-300';
            let labelColor = 'text-slate-400';
            
            if (isInitiator) {
              borderClass = 'border-blue-500 bg-blue-50/10 shadow-sm shadow-blue-50';
              avatarClass = 'bg-blue-500 text-white';
              dotClass = 'bg-blue-500';
              labelColor = 'text-blue-600 font-extrabold';
            } else if (isCompleted) {
              borderClass = 'border-emerald-500 bg-emerald-50/10 shadow-sm shadow-emerald-50';
              avatarClass = 'bg-emerald-500 text-white';
              dotClass = 'bg-emerald-500';
              labelColor = 'text-emerald-600 font-extrabold';
            } else if (isCurrent) {
              borderClass = 'border-amber-500 bg-amber-50/10 shadow-lg shadow-amber-100';
              avatarClass = 'bg-amber-500 text-white animate-pulse';
              dotClass = 'bg-amber-500';
              labelColor = 'text-amber-600 font-extrabold';
            }
            
            return (
              <React.Fragment key={idx}>
                <div 
                  ref={isCurrent ? currentStepRef : undefined}
                  className="flex flex-col items-center shrink-0"
                >
                  <span className={`text-[9px] font-extrabold uppercase tracking-wider mb-1 ${labelColor}`}>
                    {step.label}
                  </span>
                  
                  <div className={`w-[145px] h-[40px] rounded-xl border flex flex-row items-center gap-2 px-2 py-0.5 transition-all duration-350 ${borderClass}`}>
                    <UserAvatar 
                      email={user?.email} 
                      initials={initials} 
                      className={`w-6 h-6 rounded-full shadow-inner shrink-0 ${avatarClass}`} 
                    />
                    <span className="text-[10px] font-extrabold text-slate-700 text-left leading-tight line-clamp-2 flex-1">
                      {userLabel}
                    </span>
                  </div>
                  
                  <div className={`w-2 h-2 rounded-full mt-1.5 ${dotClass}`} />
                </div>
                
                {idx < workflowSteps.length - 1 && (
                  <div className="flex items-center justify-center text-slate-500 shrink-0 mt-4">
                    <ChevronRight size={16} strokeWidth={3} />
                  </div>
                )}
              </React.Fragment>
            );
          })}
        </div>
      </div>

      {/* ── 70% / 30% Columns Wrapper ── */}
      <div className="flex flex-col lg:flex-row gap-6 items-start">
        
        {/* ── Left Column (70% Width) ── */}
        <div className="w-full lg:w-[70%] space-y-6">

          {/* Basic Information */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-6">
            <div className="flex items-center gap-2 mb-4 border-b border-slate-100 pb-3">
              <div className="w-1.5 h-4 bg-blue-600 rounded-full" />
              <span className="font-bold text-slate-800 text-xs uppercase tracking-wider">Basic Information</span>
            </div>
            
            <div className="flex flex-col md:flex-row md:flex-wrap gap-4">
              <div className="flex-1 min-w-[220px] flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Project</span>
                <span className="text-sm font-semibold text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100">{form.project || '-'}</span>
              </div>
              <div className="flex-1 min-w-[220px] flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Sub-Entity</span>
                <span className="text-sm font-semibold text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100">
                  {form.projectSubEntity && form.projectSubEntity.length > 0 ? form.projectSubEntity.join(', ') : '-'}
                </span>
              </div>
              
              <div className="w-full md:hidden" />
              
              <div className="flex-1 min-w-[220px] flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Function</span>
                <span className="text-sm font-semibold text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100">{form.originatorDept || '-'}</span>
              </div>
              <div className="flex-1 min-w-[220px] flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Date</span>
                <span className="text-sm font-semibold text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100">{form.date || '-'}</span>
              </div>
              
              <div className="w-full md:hidden" />
              
              <div className="flex-1 min-w-[220px] flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Workflow</span>
                <span className="text-sm font-semibold text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 uppercase">{form.workflow || '-'}</span>
              </div>
              <div className="flex-1 min-w-[220px] flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">DOA Matrix</span>
                <span className="text-sm font-semibold text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100">{form.doaMatrix || '-'}</span>
              </div>
            </div>
          </div>

          {/* NFA Details */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-6">
            <div className="flex items-center gap-2 mb-4 border-b border-slate-100 pb-3">
              <div className="w-1.5 h-4 bg-blue-600 rounded-full" />
              <span className="font-bold text-slate-800 text-xs uppercase tracking-wider">NFA Details</span>
            </div>
            
            <div className="flex flex-col gap-4">
              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Subject</span>
                <span className="text-sm font-semibold text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{form.referenceSubject || '-'}</span>
              </div>
              
              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Introduction</span>
                <div className="text-sm text-slate-700 bg-slate-50/40 p-4 rounded-xl border border-slate-100 prose prose-sm max-w-none">
                  <div dangerouslySetInnerHTML={{ __html: form.introduction || '-' }} />
                </div>
              </div>

              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Chronology</span>
                <div className="text-sm text-slate-700 bg-slate-50/40 p-4 rounded-xl border border-slate-100 prose prose-sm max-w-none">
                  <div dangerouslySetInnerHTML={{ __html: form.chronology || '-' }} />
                </div>
              </div>

              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Commercial Proposals</span>
                <div className="text-sm text-slate-700 bg-slate-50/40 p-4 rounded-xl border border-slate-100 prose prose-sm max-w-none">
                  <div dangerouslySetInnerHTML={{ __html: form.commercialProposal || '-' }} />
                </div>
              </div>

              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Payment Terms</span>
                <div className="text-sm text-slate-700 bg-slate-50/40 p-4 rounded-xl border border-slate-100 prose prose-sm max-w-none">
                  <div dangerouslySetInnerHTML={{ __html: form.paymentTerms || '-' }} />
                </div>
              </div>
            </div>
          </div>

          {/* Financial Details */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-6">
            <div className="flex items-center gap-2 mb-4 border-b border-slate-100 pb-3">
              <div className="w-1.5 h-4 bg-blue-600 rounded-full" />
              <span className="font-bold text-slate-800 text-xs uppercase tracking-wider">Financial Details</span>
            </div>
            
            <div className="space-y-5">
              <div className="flex flex-col md:flex-row gap-4">
                <div className="flex-1 flex flex-col gap-1 p-3.5 bg-emerald-50/30 rounded-xl border border-emerald-100/50">
                  <span className="text-[9px] font-extrabold text-emerald-600 uppercase tracking-wider mb-0.5">Awarded Cost</span>
                  <span className="text-base font-black text-emerald-700">₹ {Number(form.packageAwardedCost || 0).toLocaleString('en-IN')}</span>
                </div>
                <div className="flex-1 flex flex-col gap-1 p-3.5 bg-blue-50/30 rounded-xl border border-blue-100/50">
                  <span className="text-[9px] font-extrabold text-blue-600 uppercase tracking-wider mb-0.5">Budget</span>
                  <span className="text-base font-black text-blue-700">₹ {Number(form.budget || 0).toLocaleString('en-IN')}</span>
                </div>
                <div className="flex-1 flex flex-col gap-1 p-3.5 bg-purple-50/30 rounded-xl border border-purple-100/50">
                  <span className="text-[9px] font-extrabold text-purple-600 uppercase tracking-wider mb-0.5">Capex No</span>
                  <span className="text-sm font-extrabold text-purple-700 truncate">{form.capexSerialNo || '-'}</span>
                </div>
              </div>
              
              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Remarks</span>
                <div className="text-xs text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{form.budgetRemarks || '-'}</div>
              </div>
            </div>
          </div>

          {/* Recommendation */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-6">
            <div className="flex items-center gap-2 mb-4 border-b border-slate-100 pb-3">
              <div className="w-1.5 h-4 bg-blue-600 rounded-full" />
              <span className="font-bold text-slate-800 text-xs uppercase tracking-wider">Recommendation</span>
            </div>
            
            <div className="flex flex-col gap-4">
              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Justification</span>
                <div className="text-sm text-slate-700 bg-slate-50/40 p-4 rounded-xl border border-slate-100 prose prose-sm max-w-none">
                  <div dangerouslySetInnerHTML={{ __html: form.recommendation || '-' }} />
                </div>
              </div>

              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Contractor Name</span>
                <span className="text-sm font-bold text-slate-800 bg-blue-50/30 p-3.5 rounded-xl border border-blue-100/50">{form.contractorName || '-'}</span>
              </div>
            </div>
          </div>

          {/* Contract Terms */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-6">
            <div className="flex items-center gap-2 mb-4 border-b border-slate-100 pb-3">
              <div className="w-1.5 h-4 bg-blue-600 rounded-full" />
              <span className="font-bold text-slate-800 text-xs uppercase tracking-wider">Contract Terms</span>
            </div>
            
            <div className="flex flex-col gap-4">
              <div className="flex flex-col md:flex-row gap-4">
                <div className="flex-1 flex flex-col gap-1">
                  <span className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">LD</span>
                  <div className="text-xs text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{form.liquidatedDamage || '-'}</div>
                </div>
                <div className="flex-1 flex flex-col gap-1">
                  <span className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">CPG</span>
                  <div className="text-xs text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{form.contractPerf || '-'}</div>
                </div>
              </div>
              <div className="flex flex-col md:flex-row gap-4">
                <div className="flex-1 flex flex-col gap-1">
                  <span className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">DLP</span>
                  <div className="text-xs text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{form.defectLiability || '-'}</div>
                </div>
                <div className="flex-1 flex flex-col gap-1">
                  <span className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">Work Completion</span>
                  <div className="text-xs text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{form.workCompletion || '-'}</div>
                </div>
              </div>
            </div>
          </div>

          {/* Annexures & Notes */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-6">
            <div className="flex items-center gap-2 mb-4 border-b border-slate-100 pb-3">
              <div className="w-1.5 h-4 bg-blue-600 rounded-full" />
              <span className="font-bold text-slate-800 text-xs uppercase tracking-wider">Annexures & Notes</span>
            </div>
            
            <div className="flex flex-col gap-4">
              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Annexures</span>
                <div className="text-xs text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{form.annexure || '-'}</div>
              </div>
              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Note 1</span>
                <div className="text-xs text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{form.note1 || '-'}</div>
              </div>
              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Note 2</span>
                <div className="text-xs text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{form.note2 || '-'}</div>
              </div>
              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Note 3</span>
                <div className="text-xs text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{form.note3 || '-'}</div>
              </div>
              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Note 4</span>
                <div className="text-xs text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{form.note4 || '-'}</div>
              </div>
            </div>
          </div>

        </div>

        {/* ── Right Column (30% Width) ── */}
        <div className="w-full lg:w-[30%] space-y-6">
          
          {/* Attachments */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-5">
            <div className="flex items-center justify-between mb-4 border-b border-slate-100 pb-2.5">
              <div className="flex items-center gap-2">
                <Paperclip size={15} className="text-emerald-500 shrink-0"/> 
                <span className="font-bold text-slate-700 text-[11px] uppercase tracking-wider">Attachments</span>
              </div>
              <span className="bg-emerald-50 text-emerald-600 text-[10px] font-extrabold px-2 py-0.5 rounded-full border border-emerald-100">
                {savedDocuments.length}
              </span>
            </div>
            {savedDocuments.length > 0 ? (
              <div className="space-y-2.5">
                {savedDocuments.map(doc => {
                  const ext = doc.fileName.split('.').pop()?.toLowerCase();
                  let iconColor = 'text-slate-500 bg-slate-50 border-slate-100';
                  if (ext === 'pdf') iconColor = 'text-rose-500 bg-rose-50 border-rose-100';
                  else if (['xlsx', 'xls', 'csv'].indexOf(ext || '') > -1) iconColor = 'text-emerald-500 bg-emerald-50 border-emerald-100';
                  else if (['docx', 'doc'].indexOf(ext || '') > -1) iconColor = 'text-blue-500 bg-blue-50 border-blue-100';
                  
                  return (
                    <div key={doc.id} className="group flex items-center justify-between p-2.5 bg-slate-50/50 hover:bg-emerald-50/10 rounded-xl border border-slate-100 transition-all">
                      <div className="flex items-center gap-2.5 min-w-0">
                        <div className={`p-1.5 rounded-lg border transition-colors ${iconColor}`}>
                          <FileText size={14} />
                        </div>
                        <span className="text-[11px] font-bold text-slate-700 truncate" title={doc.fileName}>{doc.fileName}</span>
                      </div>
                      <a 
                        href={doc.serverRelativeUrl} 
                        target="_blank" 
                        rel="noopener noreferrer" 
                        className="text-[10px] text-emerald-600 font-extrabold px-2.5 py-1 bg-white hover:bg-emerald-600 hover:text-white rounded-lg border border-slate-150 hover:border-emerald-600 shadow-sm transition-all"
                      >
                        View
                      </a>
                    </div>
                  );
                })}
              </div>
            ) : (
              <div className="flex flex-col items-center justify-center py-6 border border-dashed border-slate-200 rounded-xl">
                <Paperclip size={20} className="text-slate-300 mb-1.5" />
                <span className="text-[10px] text-slate-400 font-bold italic">No attachments found</span>
              </div>
            )}
          </div>

          {/* Workflow History */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-5">
            <div className="flex items-center gap-2 mb-4 border-b border-slate-100 pb-2.5">
              <History size={15} className="text-blue-500 shrink-0"/> 
              <span className="font-bold text-slate-700 text-[11px] uppercase tracking-wider">Workflow History</span>
            </div>
            
            <div className="space-y-4">
              {history.length > 0 ? history.map((item, idx) => {
                const isApproved = item.action === 'Approved';
                const isRejected = item.action === 'Rejected';
                const isQuery = item.action === 'Query';
                
                let badgeClass = 'bg-blue-50 text-blue-700 border-blue-100';
                let dotClass = 'bg-blue-500';
                
                if (isApproved) {
                  badgeClass = 'bg-emerald-50 text-emerald-700 border-emerald-100';
                  dotClass = 'bg-emerald-500';
                } else if (isRejected) {
                  badgeClass = 'bg-rose-50 text-rose-700 border-rose-100';
                  dotClass = 'bg-rose-500';
                } else if (isQuery) {
                  badgeClass = 'bg-amber-50 text-amber-700 border-amber-100';
                  dotClass = 'bg-amber-500';
                }
                
                return (
                  <div key={item.id || idx} className="relative pl-7 pb-1 last:pb-0">
                    {idx < history.length - 1 && (
                      <div className="absolute left-[7px] top-5 bottom-0 w-[1.5px] bg-slate-100" />
                    )}
                    <div className={`absolute left-0 top-1 w-3.5 h-3.5 rounded-full border border-white flex items-center justify-center z-10 ${dotClass} shadow-sm`} />
                    <div className="flex flex-col min-w-0">
                      <p className="text-[11px] text-slate-600 font-medium leading-relaxed">
                        {item.rawText || `${item.user} (${item.action})${item.remarks ? `: ${item.remarks}` : ""}`}
                      </p>
                    </div>
                  </div>
                );
              }) : (
                <div className="text-center py-4 text-slate-400 italic text-[11px]">
                  No audit history found.
                </div>
              )}
            </div>
          </div>

          {/* Comments History */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-5">
            <div className="flex items-center gap-2 mb-4 border-b border-slate-100 pb-2.5">
              <MessageSquare size={15} className="text-indigo-500 shrink-0"/> 
              <span className="font-bold text-slate-700 text-[11px] uppercase tracking-wider">Comments History</span>
            </div>
            
            <div className="space-y-4">
              {comments.filter(item => item.remarks && item.remarks !== "No remarks" && item.remarks.trim() !== "").length > 0 ? (
                comments.filter(item => item.remarks && item.remarks !== "No remarks" && item.remarks.trim() !== "").map((item, idx) => (
                  <div key={item.id || idx} className="flex flex-col gap-1.5 p-3 bg-slate-50/60 hover:bg-slate-50 rounded-xl border border-slate-100 transition-colors">
                    <div className="flex items-center justify-between gap-2">
                      <span className="text-[10px] font-extrabold text-slate-600">{item.user}</span>
                      <span className="text-[9px] text-slate-400">{item.date.split(',')[0]}</span>
                    </div>
                    <div className="p-2 bg-white rounded-lg border border-slate-100 relative shadow-sm">
                      <div className="absolute -top-1 left-3 w-2.5 h-2.5 bg-white border-t border-l border-slate-100 rotate-45" />
                      <p className="text-[10.5px] text-slate-500 leading-normal italic relative z-10">
                        "{item.remarks}"
                      </p>
                    </div>
                  </div>
                ))
              ) : (
                <div className="text-center py-6 text-slate-400 italic text-[11px]">
                  No comments or remarks recorded.
                </div>
              )}
            </div>
          </div>

        </div>

      </div>

      {/* Back to Home Button */}
      <div className="flex justify-center gap-4 py-8">
        <button 
          onClick={() => navigate('/')} 
          className="px-6 h-[40px] bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl font-extrabold text-xs uppercase tracking-wider transition-all border border-slate-300 active:scale-98 shadow-sm"
        >
          Back to Home
        </button>
      </div>

    </div>
  );
};

export default ViewNFA;
