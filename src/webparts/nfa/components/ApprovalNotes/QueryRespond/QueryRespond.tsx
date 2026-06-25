import * as React from 'react';
import { useState, useEffect, useMemo, useRef } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import { 
  Building2, FileText, Wrench, DollarSign, Paperclip, 
  Settings, Star, ChevronDown, ArrowLeft, History, 
  MessageSquare, User, CheckCircle2, Clock, Circle,
  ChevronRight, Trash2, Upload
} from 'lucide-react';
import { Skeleton } from 'primereact/skeleton';
import 'primereact/resources/primereact.min.css';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users";
import { NFAService, INFAFormData } from '../../../../../Share/Services/NFAService';
import { NFADocService, ISavedDocument } from '../../../../../Share/Services/NFADocService';
import { ApprovalService } from '../../../../../Share/Services/ApprovalService';
import { CommonService } from '../../../../../Share/Services/CommonService';
import MessagePopup from '../../../../../Share/Components/MessagePopup';
import "../../../../../Share/styles/global.scss";
import { forceScrollToTop } from '../../../../../Share/utils';

interface IQueryRespondProps {
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

const QueryRespond: React.FC<IQueryRespondProps> = ({ context, sp }) => {
  const { id } = useParams();
  const navigate = useNavigate();
  const currentUserId = context.pageContext.legacyPageContext.userId;

  const [isLoading, setIsLoading] = useState(true);
  const [isActionLoading, setIsActionLoading] = useState(false);
  const [nfa, setNfa] = useState<INFAFormData | null>(null);
  const [docs, setDocs] = useState<ISavedDocument[]>([]);
  const [users, setUsers] = useState<{label: string, value: string, email?: string}[]>([]);
  const [history, setHistory] = useState<any[]>([]);
  const [comments, setComments] = useState<any[]>([]);
  const [queryHistory, setQueryHistory] = useState<any[]>([]);
  const [activeQuery, setActiveQuery] = useState<any | null>(null);
  
  // Action State
  const [remarks, setRemarks] = useState('');
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  
  const currentStepRef = useRef<HTMLDivElement>(null);
  const stepperContainerRef = useRef<HTMLDivElement>(null);

  // Popup state
  const [popupConfig, setPopupConfig] = useState<{
    show: boolean;
    type: 'success' | 'error' | 'warning' | 'info';
    title: string;
    messages: string[];
    onSuccessRedirect?: boolean;
  }>({
    show: false,
    type: 'info',
    title: '',
    messages: [],
    onSuccessRedirect: false
  });

  const handlePopupClose = () => {
    const redirect = popupConfig.onSuccessRedirect;
    setPopupConfig(prev => ({ ...prev, show: false }));
    if (redirect) {
      navigate('/inbox');
    }
  };

  const handleRespond = async () => {
    if (!id || !nfa) return;
    
    if (!remarks.trim()) {
      setPopupConfig({
        show: true,
        type: 'warning',
        title: 'Response Comments Required',
        messages: ['Please type your response comments in the text box.'],
        onSuccessRedirect: false
      });
      return;
    }
    
    setIsActionLoading(true);
    const service = new ApprovalService(context, sp);
    const itemId = parseInt(id);
    
    try {
      const result = await service.respondToQuery(itemId, remarks, currentUserId, selectedFiles);
      
      setPopupConfig({
        show: true,
        type: result.success ? 'success' : 'error',
        title: result.success ? 'Success' : 'Error',
        messages: [result.message],
        onSuccessRedirect: result.success
      });
    } catch (err) {
      console.error("Error during NFA Query Respond:", err);
      setPopupConfig({
        show: true,
        type: 'error',
        title: 'Submit Error',
        messages: ['An error occurred while attempting to submit the response.'],
        onSuccessRedirect: false
      });
    } finally {
      setIsActionLoading(false);
    }
  };

  useEffect(() => {
    const loadData = async () => {
      if (!id) return;
      const nfaService = new NFAService(context);
      const docService = new NFADocService(context);
      
      try {
        const commonService = new CommonService(context, sp);
        
        // 1. Fetch NFA Data first to verify status
        const nfaData = await nfaService.getNFA(parseInt(id));

        if (!nfaData) {
          console.warn("QueryRespond: NFA not found. Redirecting to home.");
          navigate('/', { replace: true });
          return;
        }

        // If the NFA status is not Query or not pending at current user, redirect to approver page
        if ((nfaData.status || "").toLowerCase() !== 'query' || nfaData.pendingAtId !== currentUserId) {
          console.warn("QueryRespond: Item is not under query for current user. Redirecting to approver-page.");
          navigate(`/approver-page/${id}`, { replace: true });
          return;
        }

        setNfa(nfaData);

        // 2. Fetch pending query from QueryMgt for query console display (with robust fallback)
        let activeQueryResult = await sp.web.lists.getByTitle("QueryMgt").items
          .filter(`WFID eq '${id}' and QueryToId eq ${currentUserId} and Status eq 'Pending'`)
          .orderBy("ID", false)
          .select("Id", "Query", "QueriedOn", "QueryBy/Title")
          .expand("QueryBy")();

        // Fallback 1: Any pending query for this NFA
        if (activeQueryResult.length === 0) {
          activeQueryResult = await sp.web.lists.getByTitle("QueryMgt").items
            .filter(`WFID eq '${id}' and Status eq 'Pending'`)
            .orderBy("ID", false)
            .select("Id", "Query", "QueriedOn", "QueryBy/Title")
            .expand("QueryBy")();
        }

        // Fallback 2: Any query for this NFA (latest first)
        if (activeQueryResult.length === 0) {
          activeQueryResult = await sp.web.lists.getByTitle("QueryMgt").items
            .filter(`WFID eq '${id}'`)
            .orderBy("ID", false)
            .select("Id", "Query", "QueriedOn", "QueryBy/Title")
            .expand("QueryBy")();
        }

        if (activeQueryResult.length > 0) {
          setActiveQuery(activeQueryResult[0]);
        } else {
          setActiveQuery({
            Query: "No pending query details found in QueryMgt list.",
            QueriedOn: "N/A",
            QueryBy: { Title: "Unknown Approver" }
          });
        }

        const [siteUsers, historyData, qHistory] = await Promise.all([
          sp.web.siteUsers.filter("PrincipalType eq 1").select("Id", "Title", "Email")()
            .then(users => users.map(u => ({ label: u.Title, value: String(u.Id), email: u.Email }))),
          nfaService.getHistory(parseInt(id)),
          commonService.getQueryHistory(id)
        ]);

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

          const attachedDocs = await docService.getDocumentsByUniqId(nfaData.uniqId);
          setDocs(attachedDocs);
          setQueryHistory(qHistory);
      } catch (err) {
        console.error("QueryRespond: Data loading error:", err);
      } finally {
        setIsLoading(false);
      }
    };
    void loadData();
  }, [id]);

  // Auto scroll to current approval step on load
  useEffect(() => {
    if (!isLoading && nfa && currentStepRef.current && stepperContainerRef.current) {
      const timer = setTimeout(() => {
        const container = stepperContainerRef.current;
        const currentStep = currentStepRef.current;
        if (container && currentStep) {
          const containerWidth = container.clientWidth;
          const stepOffsetLeft = currentStep.offsetLeft;
          const stepWidth = currentStep.clientWidth;
          const targetScrollLeft = stepOffsetLeft - (containerWidth / 2) + (stepWidth / 2);
          container.scrollTo({
            left: targetScrollLeft,
            behavior: 'smooth'
          });
        }
      }, 300);
      return () => clearTimeout(timer);
    }
  }, [isLoading, nfa]);

  // --- Workflow Stepper Logic ---
  const workflowSteps = useMemo(() => {
    if (!nfa) return [];
    const steps = [{ label: 'Initiator', id: nfa.initiatorId, type: 'initiator' }];
    
    nfa.approvers.forEach((appId, idx) => {
      if (appId) {
        steps.push({ label: `Level ${idx + 1}`, id: Number(appId), type: 'approver' });
      }
    });
    return steps;
  }, [nfa]);

  const queryDocuments = useMemo(() => {
    const docsList: any[] = [];
    queryHistory.forEach(item => {
      if (item.AttachmentFiles && item.AttachmentFiles.length > 0) {
        item.AttachmentFiles.forEach((file: any) => {
          docsList.push({
            fileName: file.FileName,
            serverRelativeUrl: file.ServerRelativeUrl,
            queryBy: item.QueryBy ? item.QueryBy.Title : "Unknown"
          });
        });
      }
    });
    return docsList;
  }, [queryHistory]);

  if (isLoading || !nfa) {
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
            const isCompleted = idx < nfa.level;
            const isCurrent = idx === nfa.level;
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

      {/* Columns Wrapper */}
      <div className="flex flex-col lg:flex-row gap-6 items-start">
        
        {/* Left Column (70% Width) */}
        <div className="w-full lg:w-[70%] space-y-6">

          {/* Basic Information */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-4">
            <div className="flex items-center gap-2 mb-3 border-b border-slate-100 pb-2">
              <div className="w-1.5 h-4 bg-blue-600 rounded-full" />
              <span className="font-bold text-slate-800 text-[10.5px] uppercase tracking-wider">
                Request No: {nfa.requestNo || '-'} (Level {nfa.level || '1'})
              </span>
            </div>
            
            <div className="flex flex-col gap-2">
              {/* Row 1: Project (35%) | Sub-Entity (30%) | Date (25%) */}
              <div className="flex flex-col md:flex-row gap-3">
                <div className="w-full md:w-[35%] flex flex-col gap-1">
                  <span className="text-[9px] font-medium text-slate-400 uppercase tracking-wider">Project</span>
                  <span className="text-xs font-normal text-slate-700 bg-slate-50/60 p-1.5 px-2.5 rounded-lg border border-slate-100">{nfa.project || '-'}</span>
                </div>
                <div className="w-full md:w-[30%] flex flex-col gap-1">
                  <span className="text-[9px] font-medium text-slate-400 uppercase tracking-wider">Sub-Entity</span>
                  <span className="text-xs font-normal text-slate-700 bg-slate-50/60 p-1.5 px-2.5 rounded-lg border border-slate-100">
                    {nfa.projectSubEntity && nfa.projectSubEntity.length > 0 ? nfa.projectSubEntity.join(', ') : '-'}
                  </span>
                </div>
                <div className="w-full md:w-[25%] flex flex-col gap-1">
                  <span className="text-[9px] font-medium text-slate-400 uppercase tracking-wider">Date</span>
                  <span className="text-xs font-normal text-slate-700 bg-slate-50/60 p-1.5 px-2.5 rounded-lg border border-slate-100">{nfa.date || '-'}</span>
                </div>
              </div>

              {/* Row 2: Function (35%) | DOA Matrix (30%) | Workflow (25%) */}
              <div className="flex flex-col md:flex-row gap-3">
                <div className="w-full md:w-[35%] flex flex-col gap-1">
                  <span className="text-[9px] font-medium text-slate-400 uppercase tracking-wider">Function</span>
                  <span className="text-xs font-normal text-slate-700 bg-slate-50/60 p-1.5 px-2.5 rounded-lg border border-slate-100">{nfa.originatorDept || '-'}</span>
                </div>
                <div className="w-full md:w-[30%] flex flex-col gap-1">
                  <span className="text-[9px] font-medium text-slate-400 uppercase tracking-wider">DOA Matrix</span>
                  <span className="text-xs font-normal text-slate-700 bg-slate-50/60 p-1.5 px-2.5 rounded-lg border border-slate-100">{nfa.doaMatrix || '-'}</span>
                </div>
                <div className="w-full md:w-[25%] flex flex-col gap-1">
                  <span className="text-[9px] font-medium text-slate-400 uppercase tracking-wider">Workflow</span>
                  <span className="text-xs font-normal text-slate-700 bg-slate-50/60 p-1.5 px-2.5 rounded-lg border border-slate-100 uppercase">{nfa.workflow || '-'}</span>
                </div>
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
                <span className="text-sm font-semibold text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{nfa.referenceSubject || '-'}</span>
              </div>
              
              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Introduction</span>
                <div className="text-sm text-slate-700 bg-slate-50/40 p-4 rounded-xl border border-slate-100 richTextContent max-w-none">
                  <div dangerouslySetInnerHTML={{ __html: nfa.introduction || '-' }} />
                </div>
              </div>

              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Chronology</span>
                <div className="text-sm text-slate-700 bg-slate-50/40 p-4 rounded-xl border border-slate-100 richTextContent max-w-none">
                  <div dangerouslySetInnerHTML={{ __html: nfa.chronology || '-' }} />
                </div>
              </div>

              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Commercial Proposals</span>
                <div className="text-sm text-slate-700 bg-slate-50/40 p-4 rounded-xl border border-slate-100 richTextContent max-w-none">
                  <div dangerouslySetInnerHTML={{ __html: nfa.commercialProposal || '-' }} />
                </div>
              </div>

              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Payment Terms</span>
                <div className="text-sm text-slate-700 bg-slate-50/40 p-4 rounded-xl border border-slate-100 richTextContent max-w-none">
                  <div dangerouslySetInnerHTML={{ __html: nfa.paymentTerms || '-' }} />
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
                  <span className="text-base font-black text-emerald-700">₹ {Number(nfa.packageAwardedCost || 0).toLocaleString('en-IN')}</span>
                </div>
                <div className="flex-1 flex flex-col gap-1 p-3.5 bg-blue-50/30 rounded-xl border border-blue-100/50">
                  <span className="text-[9px] font-extrabold text-blue-600 uppercase tracking-wider mb-0.5">Budget</span>
                  <span className="text-base font-black text-blue-700">₹ {Number(nfa.budget || 0).toLocaleString('en-IN')}</span>
                </div>
                <div className="flex-1 flex flex-col gap-1 p-3.5 bg-purple-50/30 rounded-xl border border-purple-100/50">
                  <span className="text-[9px] font-extrabold text-purple-600 uppercase tracking-wider mb-0.5">Capex No</span>
                  <span className="text-sm font-extrabold text-purple-700 truncate">{nfa.capexSerialNo || '-'}</span>
                </div>
              </div>
              
              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Remarks</span>
                <div className="text-xs text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{nfa.budgetRemarks || '-'}</div>
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
                <div className="text-sm text-slate-700 bg-slate-50/40 p-4 rounded-xl border border-slate-100 richTextContent max-w-none">
                  <div dangerouslySetInnerHTML={{ __html: nfa.recommendation || '-' }} />
                </div>
              </div>

              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Contractor Name</span>
                <span className="text-sm font-bold text-slate-800 bg-blue-50/30 p-3.5 rounded-xl border border-blue-100/50">{nfa.contractorName || '-'}</span>
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
                  <div className="text-xs text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{nfa.liquidatedDamage || '-'}</div>
                </div>
                <div className="flex-1 flex flex-col gap-1">
                  <span className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">CPG</span>
                  <div className="text-xs text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{nfa.contractPerf || '-'}</div>
                </div>
              </div>
              <div className="flex flex-col md:flex-row gap-4">
                <div className="flex-1 flex flex-col gap-1">
                  <span className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">DLP</span>
                  <div className="text-xs text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{nfa.defectLiability || '-'}</div>
                </div>
                <div className="flex-1 flex flex-col gap-1">
                  <span className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">Work Completion</span>
                  <div className="text-xs text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{nfa.workCompletion || '-'}</div>
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
                <div className="text-xs text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{nfa.annexure || '-'}</div>
              </div>
              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Note 1</span>
                <div className="text-xs text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{nfa.note1 || '-'}</div>
              </div>
              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Note 2</span>
                <div className="text-xs text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{nfa.note2 || '-'}</div>
              </div>
              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Note 3</span>
                <div className="text-xs text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{nfa.note3 || '-'}</div>
              </div>
              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Note 4</span>
                <div className="text-xs text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 whitespace-pre-wrap">{nfa.note4 || '-'}</div>
              </div>
            </div>
          </div>

        </div>

        {/* Right Column (30% Width) */}
        <div className="w-full lg:w-[30%] space-y-6">
          
          {/* Attachments */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-4">
            <div className="flex items-center justify-between mb-3 border-b border-slate-100 pb-2">
              <div className="flex items-center gap-2">
                <Paperclip size={14} className="text-emerald-500 shrink-0"/> 
                <span className="font-bold text-slate-700 text-[10.5px] uppercase tracking-wider">Attachments</span>
              </div>
              <span className="bg-emerald-50 text-emerald-600 text-[9.5px] font-extrabold px-1.5 py-0.5 rounded-full border border-emerald-100">
                {docs.length}
              </span>
            </div>
            {docs.length > 0 ? (
              <div className="space-y-1.5 max-h-[160px] overflow-y-auto pr-1">
                {docs.map(doc => {
                  const ext = doc.fileName.split('.').pop()?.toLowerCase();
                  let iconColor = 'text-slate-500 bg-slate-50 border-slate-100';
                  if (ext === 'pdf') iconColor = 'text-rose-500 bg-rose-50 border-rose-100';
                  else if (['xlsx', 'xls', 'csv'].indexOf(ext || '') > -1) iconColor = 'text-emerald-500 bg-emerald-50 border-emerald-100';
                  else if (['docx', 'doc'].indexOf(ext || '') > -1) iconColor = 'text-blue-500 bg-blue-50 border-blue-100';
                  
                  return (
                    <a 
                      key={doc.id}
                      href={doc.serverRelativeUrl} 
                      target="_blank" 
                      rel="noopener noreferrer" 
                      className="group flex items-center gap-2 p-1.5 bg-slate-50/50 hover:bg-emerald-50/15 rounded-lg border border-slate-100 hover:border-emerald-150 transition-all min-w-0"
                    >
                      <div className={`p-1 rounded border transition-colors shrink-0 ${iconColor}`}>
                        <FileText size={12} />
                      </div>
                      <span className="text-[10.5px] font-bold text-slate-700 truncate group-hover:text-emerald-600 transition-colors" title={doc.fileName}>
                        {doc.fileName}
                      </span>
                    </a>
                  );
                })}
              </div>
            ) : (
              <div className="flex flex-col items-center justify-center py-4 border border-dashed border-slate-200 rounded-xl">
                <Paperclip size={18} className="text-slate-300 mb-1" />
                <span className="text-[9.5px] text-slate-400 font-bold italic">No attachments found</span>
              </div>
            )}
          </div>

          {/* Query Documents (Show only if available) */}
          {queryDocuments.length > 0 && (
            <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-4">
              <div className="flex items-center justify-between mb-3 border-b border-slate-100 pb-2">
                <div className="flex items-center gap-2">
                  <Paperclip size={14} className="text-amber-500 shrink-0"/> 
                  <span className="font-bold text-slate-700 text-[10.5px] uppercase tracking-wider">Query Documents</span>
                </div>
                <span className="bg-amber-50 text-amber-600 text-[9.5px] font-extrabold px-1.5 py-0.5 rounded-full border border-amber-100">
                  {queryDocuments.length}
                </span>
              </div>
              <div className="space-y-1.5 max-h-[160px] overflow-y-auto pr-1">
                {queryDocuments.map((doc, idx) => {
                  const ext = doc.fileName.split('.').pop()?.toLowerCase();
                  let iconColor = 'text-slate-500 bg-slate-50 border-slate-100';
                  if (ext === 'pdf') iconColor = 'text-rose-500 bg-rose-50 border-rose-100';
                  else if (['xlsx', 'xls', 'csv'].indexOf(ext || '') > -1) iconColor = 'text-emerald-500 bg-emerald-50 border-emerald-100';
                  else if (['docx', 'doc'].indexOf(ext || '') > -1) iconColor = 'text-blue-500 bg-blue-50 border-blue-100';
                  
                  return (
                    <a 
                      key={idx}
                      href={doc.serverRelativeUrl} 
                      target="_blank" 
                      rel="noopener noreferrer" 
                      className="group flex items-center gap-2 p-1.5 bg-slate-50/50 hover:bg-amber-50/15 rounded-lg border border-slate-100 hover:border-amber-150 transition-all min-w-0"
                    >
                      <div className={`p-1 rounded border transition-colors shrink-0 ${iconColor}`}>
                        <FileText size={12} />
                      </div>
                      <span className="text-[10.5px] font-bold text-slate-700 truncate group-hover:text-amber-600 transition-colors" title={doc.fileName}>
                        {doc.fileName}
                      </span>
                    </a>
                  );
                })}
              </div>
            </div>
          )}

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
          {comments.filter(item => item.remarks && item.remarks !== "No remarks" && item.remarks.trim() !== "").length > 0 && (
            <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-4">
              <div className="flex items-center gap-2 mb-3 border-b border-slate-100 pb-2">
                <MessageSquare size={14} className="text-indigo-500 shrink-0"/> 
                <span className="font-bold text-slate-700 text-[10.5px] uppercase tracking-wider">Comments History</span>
              </div>
              
              <div className="space-y-3">
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
            </div>
          )}

          {/* Query Remarks History */}
          {queryHistory.length > 0 && (
            <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-4">
              <div className="flex items-center gap-2 mb-3 border-b border-slate-100 pb-2">
                <MessageSquare size={14} className="text-amber-500 shrink-0"/> 
                <span className="font-bold text-slate-700 text-[10.5px] uppercase tracking-wider">Query Remarks History</span>
              </div>
              
              <div className="space-y-3">
                {queryHistory.map((item, idx) => (
                  <div key={item.id || idx} className="flex flex-col gap-2 p-2.5 bg-slate-50/60 hover:bg-slate-50 rounded-xl border border-slate-100 transition-colors">
                    {/* The Query Part */}
                    <div className="flex flex-col gap-1">
                      <div className="flex items-center justify-between gap-2">
                        <span className="text-[10px] font-medium text-slate-600 text-left">Query: {item.Query || ""}</span>
                      </div>
                      <div className="p-2 bg-white rounded-lg border border-slate-100 relative shadow-sm">
                        <div className="absolute -top-1 left-3 w-2 h-2 bg-white border-t border-l border-slate-100 rotate-45" />
                        <p className="text-[10px] text-slate-500 leading-normal relative z-10">
                          Remarks Added by {item.QueryBy ? item.QueryBy.Title : "Unknown"} At Level {item.WFLevel || "1"} on {item.QueriedOn || ""}
                        </p>
                      </div>
                    </div>

                    {/* The Response Part */}
                    {item.Response && item.Response.trim() !== "" && (
                      <div className="flex flex-col gap-1 mt-1.5 border-t border-slate-100 pt-1.5">
                        <div className="flex items-center justify-between gap-2">
                          <span className="text-[10px] font-medium text-slate-600 text-left">Query Response: {item.Response}</span>
                        </div>
                        <div className="p-2 bg-white rounded-lg border border-slate-100 relative shadow-sm">
                          <div className="absolute -top-1 left-3 w-2 h-2 bg-white border-t border-l border-slate-100 rotate-45" />
                          <p className="text-[10px] text-slate-500 leading-normal relative z-10">
                            Query Responded by {item.QueryTo ? item.QueryTo.Title : "Unknown"} At Level {item.WFLevel || "1"} on {item.RespondedOn || ""}
                          </p>
                        </div>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </div>
          )}

        </div>

      </div>

      {/* Query Console (Full Page Width below Left and Right columns) */}
      <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-6 space-y-5 mt-6">
        <div className="flex items-center gap-2 mb-2 border-b border-slate-100 pb-3">
          <div className="w-1.5 h-4 bg-amber-500 rounded-full" />
          <span className="font-bold text-slate-800 text-xs uppercase tracking-wider">Query Console</span>
        </div>
        
        {/* Query details block */}
        {activeQuery && (
          <div className="p-4 bg-amber-50/40 border border-amber-200/60 rounded-xl space-y-2">
            <div className="flex flex-wrap gap-4 text-xs">
              <div className="flex items-center gap-2">
                <span className="font-extrabold text-slate-500 uppercase tracking-wider text-[10px]">Query By:</span>
                <span className="font-bold text-amber-700 bg-amber-50 px-2 py-0.5 rounded border border-amber-100">
                  {activeQuery.QueryBy ? activeQuery.QueryBy.Title : "Unknown"}
                </span>
              </div>
              <div className="flex items-center gap-2">
                <span className="font-extrabold text-slate-500 uppercase tracking-wider text-[10px]">Query Date:</span>
                <span className="font-bold text-amber-700 bg-amber-50 px-2 py-0.5 rounded border border-amber-100">
                  {activeQuery.QueriedOn || "N/A"}
                </span>
              </div>
            </div>
            <div className="flex flex-col gap-1">
              <span className="font-extrabold text-slate-500 uppercase tracking-wider text-[10px]">Query Text:</span>
              <p className="text-xs text-slate-700 italic bg-white p-3 rounded-lg border border-slate-100 shadow-inner leading-relaxed">
                "{activeQuery.Query}"
              </p>
            </div>
          </div>
        )}

        <div className="flex flex-col md:flex-row gap-6 items-stretch">
          
          {/* Response text box (60%) */}
          <div className="w-full md:w-[60%] flex flex-col">
            <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1.5 block shrink-0">
              Response Remarks
            </label>
            <div className="relative flex-1">
              <textarea 
                className="w-full bg-white border border-slate-200 rounded-xl px-4 py-3 text-xs focus:outline-none focus:ring-2 focus:ring-blue-500/15 focus:border-blue-500 transition-all resize-none h-[110px] leading-normal"
                placeholder="Type your response to the query here..."
                value={remarks}
                onChange={(e) => setRemarks(e.target.value)}
              />
              <MessageSquare size={14} className="absolute right-4 top-4.5 text-slate-300 group-focus-within:text-blue-500 transition-colors" />
            </div>
          </div>
          
          {/* Direct File Attachment (40%) */}
          <div className="w-full md:w-[40%] flex flex-col justify-start">
            <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1.5 block">
              Attach Documents (Direct to Query)
            </label>
            
            <div className="flex flex-col gap-2 p-4 bg-slate-50/50 border border-slate-100 rounded-xl min-h-[110px]">
              <label className="flex items-center justify-center gap-2 px-4 py-3 bg-white border border-dashed border-slate-300 hover:border-blue-500 rounded-xl cursor-pointer hover:bg-blue-50/10 transition-all text-xs font-semibold text-slate-600">
                <Upload size={14} className="text-slate-400" />
                <span>Choose Files...</span>
                <input 
                  type="file"
                  multiple
                  className="hidden"
                  onChange={(e) => {
                    if (e.target.files && e.target.files.length > 0) {
                      const filesArr = Array.from(e.target.files);
                      setSelectedFiles(prev => [...prev, ...filesArr]);
                    }
                  }}
                />
              </label>

              {selectedFiles.length > 0 ? (
                <div className="space-y-1.5 max-h-[80px] overflow-y-auto pr-1">
                  {selectedFiles.map((file, idx) => (
                    <div key={idx} className="flex items-center justify-between p-1.5 bg-white border border-slate-100 rounded-lg text-[10px]">
                      <span className="font-bold text-slate-700 truncate max-w-[180px]">{file.name}</span>
                      <button 
                        type="button" 
                        className="text-red-500 hover:text-red-750 font-bold px-1"
                        onClick={() => setSelectedFiles(prev => prev.filter((_, i) => i !== idx))}
                      >
                        ✕
                      </button>
                    </div>
                  ))}
                </div>
              ) : (
                <div className="text-center py-3 text-slate-400 text-[10px] italic">
                  No files selected
                </div>
              )}
            </div>
          </div>
          
        </div>

        {/* Action Buttons Row */}
        <div className="flex flex-wrap gap-2.5 items-center justify-center pt-4 border-t border-slate-100">
          <button 
            onClick={handleRespond} 
            disabled={isActionLoading} 
            className="px-6 h-[38px] bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 text-white rounded-xl font-extrabold text-[11px] uppercase tracking-wider transition-all shadow-md shadow-emerald-100 hover:shadow-lg active:scale-[0.98] disabled:opacity-50"
          >
            {isActionLoading ? 'Submitting...' : 'Respond'}
          </button>
          <button 
            onClick={() => {
              forceScrollToTop();
              navigate('/inbox');
            }} 
            disabled={isActionLoading} 
            className="px-6 h-[38px] bg-gradient-to-r from-slate-600 to-slate-700 hover:from-slate-700 hover:to-slate-800 text-white rounded-xl font-extrabold text-[11px] uppercase tracking-wider transition-all shadow-md shadow-slate-100 hover:shadow-lg active:scale-[0.98] disabled:opacity-50 border border-slate-350"
          >
            Cancel
          </button>
        </div>
      </div>

      <MessagePopup 
        isOpen={popupConfig.show}
        type={popupConfig.type}
        title={popupConfig.title}
        messages={popupConfig.messages}
        onClose={handlePopupClose}
      />
    </div>
  );
};

export default QueryRespond;
