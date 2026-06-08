import * as React from 'react';
import { useState, useEffect, useMemo, useRef } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import { 
  Building2, FileText, Wrench, DollarSign, Paperclip, 
  Settings, Star, ChevronDown, ArrowLeft, History, 
  MessageSquare, User, CheckCircle2, Clock, Circle,
  ChevronRight
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
import MessagePopup from '../../../../../Share/Components/MessagePopup';
import { InputField } from '../../../../../Share/Components/InputField';
import { TextAreaField } from '../../../../../Share/Components/TextAreaField';
import { RichTextEditor } from '../../../../../Share/Components/RichTextEditor';
import DropdownField from '../../../../../Share/Components/DropdownField';
import CommonButton from '../../../../../Share/Components/CommonButton';
import "../../../../../Share/styles/global.scss";

interface IApproverPageProps {
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

const ApproverPage: React.FC<IApproverPageProps> = ({ context, sp }) => {
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
  
  // Action State
  const [remarks, setRemarks] = useState('');
  const [selectedUser, setSelectedUser] = useState('');
  const [userSearch, setUserSearch] = useState('');
  const [showSuggestions, setShowSuggestions] = useState(false);
  const suggestionsRef = useRef<HTMLDivElement>(null);
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
      navigate('/');
    }
  };

  const handleAction = async (actionType: 'Approve' | 'Reject' | 'Query' | 'Delegate' | 'Return') => {
    if (!id || !nfa) return;
    
    if ((actionType === 'Reject' || actionType === 'Query' || actionType === 'Return') && !remarks.trim()) {
      setPopupConfig({
        show: true,
        type: 'warning',
        title: 'Remarks Required',
        messages: [`Remarks / Comments are required to ${actionType.toLowerCase()} this NFA.`],
        onSuccessRedirect: false
      });
      return;
    }
    
    if ((actionType === 'Query' || actionType === 'Delegate') && !selectedUser) {
      setPopupConfig({
        show: true,
        type: 'warning',
        title: 'User Selection Required',
        messages: [`Please select a user to ${actionType.toLowerCase()} to.`],
        onSuccessRedirect: false
      });
      return;
    }

    setIsActionLoading(true);
    const service = new ApprovalService(context, sp);
    const subject = nfa.referenceSubject || "NFA Action Note";
    const itemId = parseInt(id);
    
    let result = { success: false, message: "" };

    try {
      if (actionType === 'Approve') {
        result = await service.approveNFA(itemId, nfa.level, nfa.approvers, remarks, currentUserId, subject);
      } else if (actionType === 'Reject') {
        result = await service.rejectNFA(itemId, remarks, currentUserId, subject);
      } else if (actionType === 'Query') {
        result = await service.queryNFA(itemId, remarks, parseInt(selectedUser), currentUserId, subject);
      } else if (actionType === 'Delegate') {
        result = await service.delegateNFA(itemId, remarks, parseInt(selectedUser), nfa.level, currentUserId, subject);
      } else if (actionType === 'Return') {
        const initiatorId = nfa.initiatorId || currentUserId;
        result = await service.returnNFA(itemId, remarks, initiatorId, currentUserId, subject);
      }
      
      setPopupConfig({
        show: true,
        type: result.success ? 'success' : 'error',
        title: result.success ? 'Success' : 'Error',
        messages: [result.message],
        onSuccessRedirect: result.success
      });
    } catch (err) {
      console.error(`Error during NFA ${actionType}:`, err);
      setPopupConfig({
        show: true,
        type: 'error',
        title: 'Action Error',
        messages: [`An error occurred while attempting to ${actionType.toLowerCase()} the NFA.`],
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
        const [nfaData, siteUsers, historyData] = await Promise.all([
          nfaService.getNFA(parseInt(id)),
          sp.web.siteUsers.filter("PrincipalType eq 1").select("Id", "Title", "Email")()
            .then(users => users.map(u => ({ label: u.Title, value: String(u.Id), email: u.Email }))),
          nfaService.getHistory(parseInt(id))
        ]);

        if (nfaData) {
          setNfa(nfaData);
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

          const attachedDocs = await docService.getDocumentsByUniqId(nfaData.uniqId);
          setDocs(attachedDocs);
        }
      } catch (err) {
        console.error("ApproverPage: Data loading error:", err);
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

  // Handle click outside of suggestions to close the list
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (suggestionsRef.current && !suggestionsRef.current.contains(event.target as Node)) {
        setShowSuggestions(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

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

  const selectedUserObj = useMemo(() => {
    return users.find(u => u.value === selectedUser);
  }, [selectedUser, users]);

  const filteredUsers = useMemo(() => {
    if (!userSearch) return users.slice(0, 5);
    return users.filter(u => u.label.toLowerCase().includes(userSearch.toLowerCase())).slice(0, 8);
  }, [userSearch, users]);

  const isCurrentApprover = nfa?.pendingAtId === currentUserId;

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
            
            // Color configuration mapping the user's reference image
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
                  {/* Above Card Label */}
                  <span className={`text-[9px] font-extrabold uppercase tracking-wider mb-1 ${labelColor}`}>
                    {step.label}
                  </span>
                  
                  {/* Card Box - Compact Horizontal Layout */}
                  <div className={`w-[145px] h-[40px] rounded-xl border flex flex-row items-center gap-2 px-2 py-0.5 transition-all duration-350 ${borderClass}`}>
                    {/* Avatar circle */}
                    <UserAvatar 
                      email={user?.email} 
                      initials={initials} 
                      className={`w-6 h-6 rounded-full shadow-inner shrink-0 ${avatarClass}`} 
                    />
                    {/* Name */}
                    <span className="text-[10px] font-extrabold text-slate-700 text-left leading-tight line-clamp-2 flex-1">
                      {userLabel}
                    </span>
                  </div>
                  
                  {/* Below Card Dot */}
                  <div className={`w-2 h-2 rounded-full mt-1.5 ${dotClass}`} />
                </div>
                
                {/* Chevron Connector */}
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
                <span className="text-sm font-semibold text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100">{nfa.project || '-'}</span>
              </div>
              <div className="flex-1 min-w-[220px] flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Sub-Entity</span>
                <span className="text-sm font-semibold text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100">
                  {nfa.projectSubEntity && nfa.projectSubEntity.length > 0 ? nfa.projectSubEntity.join(', ') : '-'}
                </span>
              </div>
              
              <div className="w-full md:hidden" />
              
              <div className="flex-1 min-w-[220px] flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Function</span>
                <span className="text-sm font-semibold text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100">{nfa.originatorDept || '-'}</span>
              </div>
              <div className="flex-1 min-w-[220px] flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Date</span>
                <span className="text-sm font-semibold text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100">{nfa.date || '-'}</span>
              </div>
              
              <div className="w-full md:hidden" />
              
              <div className="flex-1 min-w-[220px] flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Workflow</span>
                <span className="text-sm font-semibold text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100 uppercase">{nfa.workflow || '-'}</span>
              </div>
              <div className="flex-1 min-w-[220px] flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">DOA Matrix</span>
                <span className="text-sm font-semibold text-slate-700 bg-slate-50/60 p-3.5 rounded-xl border border-slate-100">{nfa.doaMatrix || '-'}</span>
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
                <div className="text-sm text-slate-700 bg-slate-50/40 p-4 rounded-xl border border-slate-100 prose prose-sm max-w-none">
                  <div dangerouslySetInnerHTML={{ __html: nfa.introduction || '-' }} />
                </div>
              </div>

              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Chronology</span>
                <div className="text-sm text-slate-700 bg-slate-50/40 p-4 rounded-xl border border-slate-100 prose prose-sm max-w-none">
                  <div dangerouslySetInnerHTML={{ __html: nfa.chronology || '-' }} />
                </div>
              </div>

              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Commercial Proposals</span>
                <div className="text-sm text-slate-700 bg-slate-50/40 p-4 rounded-xl border border-slate-100 prose prose-sm max-w-none">
                  <div dangerouslySetInnerHTML={{ __html: nfa.commercialProposal || '-' }} />
                </div>
              </div>

              <div className="flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Payment Terms</span>
                <div className="text-sm text-slate-700 bg-slate-50/40 p-4 rounded-xl border border-slate-100 prose prose-sm max-w-none">
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
                <div className="text-sm text-slate-700 bg-slate-50/40 p-4 rounded-xl border border-slate-100 prose prose-sm max-w-none">
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
                {docs.length}
              </span>
            </div>
            {docs.length > 0 ? (
              <div className="space-y-2.5">
                {docs.map(doc => {
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

      {/* Remarks & Action Console (Full Page Width below Left and Right columns) */}
      {isCurrentApprover && (
        <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-6 space-y-5 mt-6">
          <div className="flex items-center gap-2 mb-2 border-b border-slate-100 pb-3">
            <div className="w-1.5 h-4 bg-blue-600 rounded-full" />
            <span className="font-bold text-slate-800 text-xs uppercase tracking-wider">Action Console</span>
          </div>
          
          <div className="flex flex-col md:flex-row gap-6 items-stretch">
            
            {/* Remarks Multiline Text Area (Left Side - 60% Width) */}
            <div className="w-full md:w-[60%] flex flex-col">
              <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1.5 block shrink-0">
                Remarks / Comments
              </label>
              <div className="relative flex-1">
                <textarea 
                  className="w-full bg-white border border-slate-200 rounded-xl px-4 py-3 text-xs focus:outline-none focus:ring-2 focus:ring-blue-500/15 focus:border-blue-500 transition-all resize-none h-[110px] leading-normal"
                  placeholder="Type your approval/rejection comments here..."
                  value={remarks}
                  onChange={(e) => setRemarks(e.target.value)}
                />
                <MessageSquare size={14} className="absolute right-4 top-4.5 text-slate-300 group-focus-within:text-blue-500 transition-colors" />
              </div>
            </div>
            
            {/* Delegate/Query User (Right Side - 40% Width) */}
            <div className="w-full md:w-[40%] flex flex-col justify-start">
              
              {/* Custom People Picker Search Input */}
              <div className="relative" ref={suggestionsRef}>
                <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1.5 block">
                  Delegate / Query User
                </label>
                
                {selectedUserObj ? (
                  /* Selected User Pill */
                  <div className="flex items-center justify-between bg-blue-50 border border-blue-200 rounded-xl p-2.5 h-[42px]">
                    <div className="flex items-center gap-2 min-w-0">
                      <UserAvatar 
                        email={selectedUserObj.email} 
                        initials={getInitials(selectedUserObj.label)} 
                        className="w-6 h-6 rounded-full bg-blue-600 text-white text-[10px] shrink-0 shadow-inner" 
                      />
                      <span className="text-xs font-bold text-slate-700 truncate">{selectedUserObj.label}</span>
                    </div>
                    <button 
                      type="button" 
                      onClick={() => {
                        setSelectedUser('');
                        setUserSearch('');
                      }}
                      className="text-slate-400 hover:text-slate-650 font-bold text-xs p-1"
                    >
                      ✕
                    </button>
                  </div>
                ) : (
                  /* Search Input */
                  <div className="relative">
                    <input 
                      type="text"
                      className="w-full bg-white border border-slate-200 rounded-xl pl-8 pr-4 py-2.5 text-xs focus:outline-none focus:ring-2 focus:ring-blue-500/15 focus:border-blue-500 transition-all h-[42px]"
                      placeholder="Search people by name..."
                      value={userSearch}
                      onChange={(e) => {
                        setUserSearch(e.target.value);
                        setShowSuggestions(true);
                      }}
                      onFocus={() => setShowSuggestions(true)}
                    />
                    <User size={13} className="absolute left-3 top-[14px] text-slate-400" />
                  </div>
                )}

                {/* Suggestions Dropdown */}
                {showSuggestions && !selectedUser && (
                  <div className="absolute left-0 right-0 mt-1 bg-white border border-slate-200 rounded-xl shadow-lg z-50 max-h-[180px] overflow-y-auto p-1.5">
                    {filteredUsers.length > 0 ? (
                      filteredUsers.map(u => (
                        <button
                          key={u.value}
                          type="button"
                          className="w-full flex items-center gap-2 p-2 hover:bg-slate-50 rounded-lg text-left transition-colors"
                          onClick={() => {
                            setSelectedUser(u.value);
                            setShowSuggestions(false);
                          }}
                        >
                          <UserAvatar 
                            email={u.email} 
                            initials={getInitials(u.label)} 
                            className="w-6 h-6 rounded-full bg-slate-100 text-slate-600 text-[9px] shrink-0 shadow-inner" 
                          />
                          <span className="text-xs font-bold text-slate-700">{u.label}</span>
                        </button>
                      ))
                    ) : (
                      <div className="text-center py-3 text-slate-400 text-xs italic">
                        No matches found
                      </div>
                    )}
                  </div>
                )}
              </div>
              
            </div>
            
          </div>

          {/* Action Buttons Row */}
          <div className="flex flex-wrap gap-2.5 items-center justify-center pt-4 border-t border-slate-100">
            <button 
              onClick={() => handleAction('Approve')} 
              disabled={isActionLoading} 
              className="px-5 h-[38px] bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 text-white rounded-xl font-extrabold text-[11px] uppercase tracking-wider transition-all shadow-md shadow-emerald-100 hover:shadow-lg hover:shadow-emerald-200 active:scale-[0.98] disabled:opacity-50"
            >
              Approve
            </button>
            <button 
              onClick={() => handleAction('Reject')} 
              disabled={isActionLoading} 
              className="px-5 h-[38px] bg-gradient-to-r from-rose-600 to-red-600 hover:from-rose-700 hover:to-red-700 text-white rounded-xl font-extrabold text-[11px] uppercase tracking-wider transition-all shadow-md shadow-rose-100 hover:shadow-lg hover:shadow-rose-200 active:scale-[0.98] disabled:opacity-50"
            >
              Reject
            </button>
            <button 
              onClick={() => handleAction('Query')} 
              disabled={isActionLoading} 
              className="px-3.5 h-[38px] bg-gradient-to-r from-amber-500 to-orange-500 hover:from-amber-600 hover:to-orange-600 text-white rounded-xl font-extrabold text-[11px] uppercase tracking-wider transition-all shadow-md shadow-amber-100 hover:shadow-lg hover:shadow-amber-200 active:scale-[0.98] disabled:opacity-50 border-0"
            >
              Query
            </button>
            <button 
              onClick={() => handleAction('Delegate')} 
              disabled={isActionLoading} 
              className="px-3.5 h-[38px] bg-gradient-to-r from-indigo-500 to-blue-600 hover:from-indigo-600 hover:to-blue-700 text-white rounded-xl font-extrabold text-[11px] uppercase tracking-wider transition-all shadow-md shadow-indigo-100 hover:shadow-lg hover:shadow-indigo-200 active:scale-[0.98] disabled:opacity-50 border-0"
            >
              Delegate
            </button>
            <button 
              onClick={() => handleAction('Return')} 
              disabled={isActionLoading} 
              className="px-3.5 h-[38px] bg-gradient-to-r from-slate-600 to-slate-700 hover:from-slate-700 hover:to-slate-800 text-white rounded-xl font-extrabold text-[11px] uppercase tracking-wider transition-all shadow-md shadow-slate-100 hover:shadow-lg active:scale-[0.98] disabled:opacity-50 border border-slate-300"
            >
              Return
            </button>
          </div>
        </div>
      )}

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

export default ApproverPage;
