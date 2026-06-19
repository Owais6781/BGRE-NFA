import * as React from 'react';
import { useState, useEffect, useMemo, useRef } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import { 
  Building2, FileText, Wrench, DollarSign, Paperclip, 
  Settings, Star, ChevronDown, ArrowLeft, History, 
  MessageSquare, User, CheckCircle2, Clock, Circle,
  ChevronRight, Check, X, HelpCircle, UserCheck, CornerUpLeft, Download
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

const formatFileSize = (bytes?: number) => {
  if (!bytes) return '0 KB';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
};

const splitHistoryTextAndDate = (item: any) => {
  const raw = item.rawText || "";
  const date = item.date || "";
  
  let cleanText = raw.replace(/<[^>]*>/g, "").trim();
  
  if (date) {
    const dateIndex = cleanText.lastIndexOf(date);
    if (dateIndex !== -1) {
      const textBeforeDate = cleanText.substring(0, dateIndex).trim();
      if (textBeforeDate.endsWith("on")) {
        cleanText = textBeforeDate.substring(0, textBeforeDate.length - 2).trim();
      } else {
        cleanText = textBeforeDate;
      }
    }
  }
  
  cleanText = cleanText.replace(/\.$/, "").trim();
  
  return {
    text: cleanText,
    date: date || "System Log"
  };
};

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

const renderHistoryText = (text: string) => {
  const statusRegex = /(Pending|Approved?|Rejected?|Initiated|Queried|Query|Returned?|Delegated?|Resolved|Submitted)/i;
  const parts = text.split(statusRegex);
  
  return (
    <span className="leading-relaxed">
      {parts.map((part, idx) => {
        if (statusRegex.test(part)) {
          return (
            <span key={idx} className="text-[13px] font-bold text-slate-800">
              {part}
            </span>
          );
        }
        return (
          <span key={idx} className="text-[11px] font-medium text-slate-500">
            {part}
          </span>
        );
      })}
    </span>
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
  const [queryHistory, setQueryHistory] = useState<any[]>([]);
  
  // Action State
  const [remarks, setRemarks] = useState('');
  const [selectedUser, setSelectedUser] = useState('');
  const [userSearch, setUserSearch] = useState('');
  const [showSuggestions, setShowSuggestions] = useState(false);
  const suggestionsRef = useRef<HTMLDivElement>(null);
  const currentStepRef = useRef<HTMLDivElement>(null);
  const stepperContainerRef = useRef<HTMLDivElement>(null);

  // Tab states
  const [activeTab, setActiveTab] = useState<'details' | 'financial' | 'recommendation' | 'terms' | 'notes'>('details');
  const [activeDocTab, setActiveDocTab] = useState<'attachments' | 'query'>('attachments');
  const [activeCommentTab, setActiveCommentTab] = useState<'comments' | 'query'>('comments');
  const [isIntroExpanded, setIsIntroExpanded] = useState(false);
  const [isChronologyExpanded, setIsChronologyExpanded] = useState(false);
  const [isCommercialExpanded, setIsCommercialExpanded] = useState(false);
  const [isHistoryExpanded, setIsHistoryExpanded] = useState(false);

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
      
      if (result.success && (actionType === 'Approve' || actionType === 'Reject')) {
        try {
          const itemObj = await sp.web.lists.getByTitle("DigitalApprovalNotes").items.getById(itemId).select("WorkflowHistoryHTML")();
          const currentHTML = itemObj.WorkflowHistoryHTML || "";

          const approverName = context.pageContext.user.displayName || "Approver";
          const actionName = actionType === 'Approve' ? 'Approved' : 'Rejected';
          
          const pad = (num: number) => num < 10 ? '0' + num : String(num);
          const now = new Date();
          const day = pad(now.getDate());
          const month = pad(now.getMonth() + 1);
          const year = now.getFullYear();
          let hours = now.getHours();
          const minutes = pad(now.getMinutes());
          const ampm = hours >= 12 ? 'PM' : 'AM';
          hours = hours % 12;
          hours = hours ? hours : 12;
          const dateTimeVal = `${day}/${month}/${year} ${hours}:${minutes} ${ampm}`;

          const newRowHTML = `<tr><td>(${nfa.level} Approver)</td><td></td><td>${approverName}</td><td>${remarks || ""}</td><td>${actionName}</td><td>${dateTimeVal}</td></tr>`;
          
          let updatedHTML = "";
          const cleanedCurrent = currentHTML.trim();
          if (!cleanedCurrent) {
            updatedHTML = `<table><tbody>${newRowHTML}</tbody></table>`;
          } else if (!cleanedCurrent.toLowerCase().includes("<table")) {
            updatedHTML = `<table><tbody>${cleanedCurrent}${newRowHTML}</tbody></table>`;
          } else {
            const lowerHTML = cleanedCurrent.toLowerCase();
            const lastTbodyIndex = lowerHTML.lastIndexOf("</tbody>");
            if (lastTbodyIndex !== -1) {
              updatedHTML = cleanedCurrent.substring(0, lastTbodyIndex) + newRowHTML + cleanedCurrent.substring(lastTbodyIndex);
            } else {
              const lastTableIndex = lowerHTML.lastIndexOf("</table>");
              if (lastTableIndex !== -1) {
                updatedHTML = cleanedCurrent.substring(0, lastTableIndex) + `<tbody>${newRowHTML}</tbody>` + cleanedCurrent.substring(lastTableIndex);
              } else {
                updatedHTML = cleanedCurrent + `<table><tbody>${newRowHTML}</tbody></table>`;
              }
            }
          }

          await sp.web.lists.getByTitle("DigitalApprovalNotes").items.getById(itemId).update({
            WorkflowHistoryHTML: updatedHTML
          });
        } catch (htmlErr) {
          console.error("ApproverPage: Error updating WorkflowHistoryHTML:", htmlErr);
        }
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
        const commonService = new CommonService(context, sp);
        const [nfaData, siteUsers, historyData, qHistory] = await Promise.all([
          nfaService.getNFA(parseInt(id)),
          sp.web.siteUsers.filter("PrincipalType eq 1").select("Id", "Title", "Email")()
            .then(users => users.map(u => ({ label: u.Title, value: String(u.Id), email: u.Email }))),
          nfaService.getHistory(parseInt(id)),
          commonService.getQueryHistory(id)
        ]);

        if (nfaData) {
          // If status is Query and pending user is current user, redirect to query response page
          if ((nfaData.status || "").toLowerCase() === 'query' && nfaData.pendingAtId === currentUserId) {
            navigate(`/query-respond/${id}`, { replace: true });
            return;
          }
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
                  <div className="flex items-center justify-center text-slate-500 shrink-0 mt-4 mx-0.5">
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

          {/* Request Details Card */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-5">
            <div className="flex items-center justify-between mb-4 border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2">
                <div className="w-1.5 h-4 bg-blue-600 rounded-full" />
                <span className="font-bold text-slate-800 text-[11px] uppercase tracking-wider">Request Details</span>
                <span className="bg-emerald-50 text-emerald-600 text-[9px] font-extrabold px-2 py-0.5 rounded-md border border-emerald-100 uppercase ml-2">
                  Level {nfa.level || '1'}
                </span>
              </div>
              <span className="text-blue-600 font-bold text-[13px]">
                Request No. {nfa.requestNo || '-'}
              </span>
            </div>
            
            {/* 2-Column Details Grid */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-x-8 gap-y-4" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px 32px' }}>
              {/* Row 1: Project (Full Width) */}
              <div className="flex flex-col pb-2.5 border-b border-slate-100 col-span-2" style={{ gridColumn: 'span 2 / span 2' }}>
                <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Project</span>
                <span className="text-[14px] text-slate-700 font-semibold mt-1">{nfa.project || '-'}</span>
              </div>
              
              {/* Row 2: Sub-Entity | Date */}
              <div className="flex flex-col pb-2.5 border-b border-slate-100 col-span-1">
                <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Sub-Entity</span>
                <span className="text-[14px] text-slate-700 font-semibold mt-1">
                  {nfa.projectSubEntity && nfa.projectSubEntity.length > 0 ? nfa.projectSubEntity.join(', ') : '-'}
                </span>
              </div>
              <div className="flex flex-col pb-2.5 border-b border-slate-100 col-span-1">
                <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Date</span>
                <span className="text-[14px] text-slate-700 font-semibold mt-1">{nfa.date || '-'}</span>
              </div>

              {/* Row 3: Function (Full Width) */}
              <div className="flex flex-col pb-2.5 border-b border-slate-100 col-span-2" style={{ gridColumn: 'span 2 / span 2' }}>
                <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Function</span>
                <span className="text-[14px] text-slate-700 font-semibold mt-1">{nfa.originatorDept || '-'}</span>
              </div>

              {/* Row 4: Data Matrix | Workflow */}
              <div className="flex flex-col pb-2.5 border-b border-slate-100 col-span-1">
                <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Data Matrix</span>
                <span className="text-[14px] text-slate-700 font-semibold mt-1">{nfa.doaMatrix || '-'}</span>
              </div>
              <div className="flex flex-col pb-2.5 border-b border-slate-100 col-span-1">
                <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Workflow</span>
                <span className="text-[14px] text-slate-700 font-semibold mt-1 uppercase">{nfa.workflow || '-'}</span>
              </div>
            </div>
          </div>

          {/* Financial Summary Card */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-5 space-y-4">
            <div className="flex items-center gap-2 border-b border-slate-100 pb-3">
              <div className="w-1.5 h-4 bg-blue-600 rounded-full" />
              <span className="font-bold text-slate-800 text-[11px] uppercase tracking-wider">Financial Summary</span>
            </div>
            
            <div style={{ display: 'flex', flexDirection: 'row', gap: '12px', width: '100%', alignItems: 'stretch' }}>
              <div className="p-2 bg-blue-50/20 border border-blue-100/50 rounded-xl" style={{ flex: '1 1 0%', minWidth: '0px' }}>
                <span style={{ fontSize: '9px', fontWeight: 700, color: '#2563eb', textTransform: 'uppercase', letterSpacing: '0.05em', display: 'block', lineHeight: '1.2' }}>Budget</span>
                <div style={{ fontSize: '13px', fontWeight: 900, color: '#1d4ed8', marginTop: '2px', lineHeight: '1.2' }}>₹ {Number(nfa.budget || 0).toLocaleString('en-IN')}</div>
              </div>
              <div className="p-2 bg-emerald-50/20 border border-emerald-100/50 rounded-xl" style={{ flex: '1 1 0%', minWidth: '0px' }}>
                <span style={{ fontSize: '9px', fontWeight: 700, color: '#059669', textTransform: 'uppercase', letterSpacing: '0.05em', display: 'block', lineHeight: '1.2' }}>Awarded Cost</span>
                <div style={{ fontSize: '13px', fontWeight: 900, color: '#047857', marginTop: '2px', lineHeight: '1.2' }}>₹ {Number(nfa.packageAwardedCost || 0).toLocaleString('en-IN')}</div>
              </div>
              <div className="p-2 bg-purple-50/20 border border-purple-100/50 rounded-xl" style={{ flex: '1 1 0%', minWidth: '0px' }}>
                <span style={{ fontSize: '9px', fontWeight: 700, color: '#7e22ce', textTransform: 'uppercase', letterSpacing: '0.05em', display: 'block', lineHeight: '1.2' }}>Capex Serial No</span>
                <div style={{ fontSize: '13px', fontWeight: 800, color: '#6b21a8', marginTop: '2px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', lineHeight: '1.2' }} title={nfa.capexSerialNo}>{nfa.capexSerialNo || '-'}</div>
              </div>
            </div>
            {nfa.budgetRemarks && nfa.budgetRemarks.trim() !== "" && (
              <div className="p-2.5 bg-slate-50/50 border border-slate-100 rounded-xl">
                <span className="text-[8px] font-bold text-slate-400 uppercase tracking-wider block">Remarks</span>
                <div className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-0.5 whitespace-pre-wrap">{nfa.budgetRemarks}</div>
              </div>
            )}
          </div>

          {/* Details Card with Tabs */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-6">
            
            {/* Horizontal Tabs Header */}
            <div className="flex flex-wrap border-b border-slate-200 gap-x-6 mb-5 overflow-x-hidden" style={{ marginBottom: '20px' }}>
              <button 
                type="button"
                onClick={() => setActiveTab('details')}
                className={`font-bold transition-all flex items-center gap-2 pb-2.5 border-b-2 -mb-[2px] ${
                  activeTab === 'details' 
                    ? 'border-teal-800 text-teal-800 font-extrabold' 
                    : 'border-transparent text-slate-500 hover:text-slate-700 hover:border-b-slate-300/50'
                }`}
                style={{ fontSize: '14px' }}
              >
                <FileText size={18} className="text-blue-500 shrink-0" />
                NFA DETAILS
              </button>
              <button 
                type="button"
                onClick={() => setActiveTab('recommendation')}
                className={`font-bold transition-all flex items-center gap-2 pb-2.5 border-b-2 -mb-[2px] ${
                  activeTab === 'recommendation' 
                    ? 'border-teal-800 text-teal-800 font-extrabold' 
                    : 'border-transparent text-slate-500 hover:text-slate-700 hover:border-b-slate-300/50'
                }`}
                style={{ fontSize: '14px' }}
              >
                <Star size={18} className="text-amber-500 shrink-0" />
                RECOMMENDATION
              </button>
              <button 
                type="button"
                onClick={() => setActiveTab('terms')}
                className={`font-bold transition-all flex items-center gap-2 pb-2.5 border-b-2 -mb-[2px] ${
                  activeTab === 'terms' 
                    ? 'border-teal-800 text-teal-800 font-extrabold' 
                    : 'border-transparent text-slate-500 hover:text-slate-700 hover:border-b-slate-300/50'
                }`}
                style={{ fontSize: '14px' }}
              >
                <Wrench size={18} className="text-purple-500 shrink-0" />
                CONTRACT TERMS
              </button>
              <button 
                type="button"
                onClick={() => setActiveTab('notes')}
                className={`font-bold transition-all flex items-center gap-2 pb-2.5 border-b-2 -mb-[2px] ${
                  activeTab === 'notes' 
                    ? 'border-teal-800 text-teal-800 font-extrabold' 
                    : 'border-transparent text-slate-500 hover:text-slate-700 hover:border-b-slate-300/50'
                }`}
                style={{ fontSize: '14px' }}
              >
                <Paperclip size={18} className="text-rose-500 shrink-0" />
                ANNEXURES & NOTES
              </button>
            </div>

            {/* Tab Contents */}
            {activeTab === 'details' && (
              <div className="space-y-6">
                <div>
                  <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Subject</span>
                  <h2 className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1">{nfa.referenceSubject || '-'}</h2>
                </div>
                
                <div className="border-t border-slate-100 pt-4">
                  <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Introduction / Preamble / Background</span>
                  <div 
                    className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1.5 richTextContent transition-all duration-300"
                    style={!isIntroExpanded ? { display: '-webkit-box', WebkitLineClamp: 4, WebkitBoxOrient: 'vertical', overflow: 'hidden' } : undefined}
                  >
                    <div dangerouslySetInnerHTML={{ __html: nfa.introduction || '-' }} />
                  </div>
                  {nfa.introduction && nfa.introduction.replace(/<[^>]*>/g, "").length > 250 && (
                    <button 
                      type="button"
                      onClick={() => setIsIntroExpanded(!isIntroExpanded)}
                      className="text-[10px] font-bold text-blue-600 hover:text-blue-800 mt-1 cursor-pointer block"
                    >
                      {isIntroExpanded ? 'View less' : 'View more'}
                    </button>
                  )}
                </div>

                <div className="border-t border-slate-100 pt-4">
                  <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Chronology</span>
                  <div 
                    className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1 richTextContent transition-all duration-300"
                    style={!isChronologyExpanded ? { display: '-webkit-box', WebkitLineClamp: 4, WebkitBoxOrient: 'vertical', overflow: 'hidden' } : undefined}
                  >
                    <div dangerouslySetInnerHTML={{ __html: nfa.chronology || '-' }} />
                  </div>
                  {nfa.chronology && nfa.chronology.replace(/<[^>]*>/g, "").length > 250 && (
                    <button 
                      type="button"
                      onClick={() => setIsChronologyExpanded(!isChronologyExpanded)}
                      className="text-[10px] font-bold text-blue-600 hover:text-blue-800 mt-1 cursor-pointer block"
                    >
                      {isChronologyExpanded ? 'View less' : 'View more'}
                    </button>
                  )}
                </div>

                <div className="border-t border-slate-100 pt-4">
                  <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Commercial Proposals</span>
                  <div 
                    className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1 richTextContent transition-all duration-300"
                    style={!isCommercialExpanded ? { display: '-webkit-box', WebkitLineClamp: 4, WebkitBoxOrient: 'vertical', overflow: 'hidden' } : undefined}
                  >
                    <div dangerouslySetInnerHTML={{ __html: nfa.commercialProposal || '-' }} />
                  </div>
                  {nfa.commercialProposal && nfa.commercialProposal.replace(/<[^>]*>/g, "").length > 250 && (
                    <button 
                      type="button"
                      onClick={() => setIsCommercialExpanded(!isCommercialExpanded)}
                      className="text-[10px] font-bold text-blue-600 hover:text-blue-800 mt-1 cursor-pointer block"
                    >
                      {isCommercialExpanded ? 'View less' : 'View more'}
                    </button>
                  )}
                </div>

                <div className="border-t border-slate-100 pt-4">
                  <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Payment Terms</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1 richTextContent">
                    <div dangerouslySetInnerHTML={{ __html: nfa.paymentTerms || '-' }} />
                  </div>
                </div>
              </div>
            )}

            {activeTab === 'recommendation' && (
              <div className="space-y-4">
                <div className="flex flex-col md:flex-row gap-4">
                  <div className="flex-1 p-3.5 bg-blue-50/20 border border-blue-100/50 rounded-xl">
                    <span className="text-[9px] font-bold text-blue-600 uppercase tracking-wider">Contractor Name</span>
                    <div className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-0.5">{nfa.contractorName || '-'}</div>
                  </div>
                </div>
                
                <div>
                  <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Justification / Recommendation</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1.5 richTextContent">
                    <div dangerouslySetInnerHTML={{ __html: nfa.recommendation || '-' }} />
                  </div>
                </div>
              </div>
            )}

            {activeTab === 'terms' && (
              <div className="space-y-4">
                <div className="p-3 bg-slate-50/40 border border-slate-100 rounded-xl">
                  <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">LD</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1.5 whitespace-pre-wrap">{nfa.liquidatedDamage || '-'}</div>
                </div>
                <div className="p-3 bg-slate-50/40 border border-slate-100 rounded-xl">
                  <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">CPG</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1.5 whitespace-pre-wrap">{nfa.contractPerf || '-'}</div>
                </div>
                <div className="p-3 bg-slate-50/40 border border-slate-100 rounded-xl">
                  <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">DLP</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1.5 whitespace-pre-wrap">{nfa.defectLiability || '-'}</div>
                </div>
                <div className="p-3 bg-slate-50/40 border border-slate-100 rounded-xl">
                  <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Work Completion</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1.5 whitespace-pre-wrap">{nfa.workCompletion || '-'}</div>
                </div>
              </div>
            )}

            {activeTab === 'notes' && (
              <div className="space-y-4">
                <div>
                  <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Annexures</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed bg-slate-50/40 p-3 rounded-xl border border-slate-100 mt-1 whitespace-pre-wrap">{nfa.annexure || '-'}</div>
                </div>
                <div>
                  <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Note 1</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed bg-slate-50/40 p-3 rounded-xl border border-slate-100 mt-1 whitespace-pre-wrap">{nfa.note1 || '-'}</div>
                </div>
                <div>
                  <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Note 2</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed bg-slate-50/40 p-3 rounded-xl border border-slate-100 mt-1 whitespace-pre-wrap">{nfa.note2 || '-'}</div>
                </div>
                <div>
                  <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Note 3</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed bg-slate-50/40 p-3 rounded-xl border border-slate-100 mt-1 whitespace-pre-wrap">{nfa.note3 || '-'}</div>
                </div>
                <div>
                  <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Note 4</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed bg-slate-50/40 p-3 rounded-xl border border-slate-100 mt-1 whitespace-pre-wrap">{nfa.note4 || '-'}</div>
                </div>
              </div>
            )}

          </div>

        </div>

        {/* ── Right Column (30% Width) ── */}
        <div className="w-full lg:w-[30%] space-y-6">

          {/* Documents Card with Tabs */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-4">
            
            {/* Documents Tabs */}
            <div className="flex flex-wrap border-b border-slate-200 mb-3 gap-4 overflow-x-hidden">
              <button 
                type="button"
                onClick={() => setActiveDocTab('attachments')}
                className={`pb-2 font-bold transition-all flex items-center gap-1.5 border-b-2 -mb-[2px] uppercase tracking-wider text-[10.5px] ${
                  activeDocTab === 'attachments' 
                    ? 'border-teal-800 text-slate-700 font-extrabold' 
                    : 'border-transparent text-slate-400 hover:text-slate-600'
                }`}
              >
                <FileText size={14} className={activeDocTab === 'attachments' ? 'text-slate-700' : 'text-slate-400'} />
                DOCUMENTS ({docs.length})
              </button>
              <button 
                type="button"
                onClick={() => setActiveDocTab('query')}
                className={`pb-2 font-bold transition-all flex items-center gap-1.5 border-b-2 -mb-[2px] uppercase tracking-wider text-[10.5px] ${
                  activeDocTab === 'query' 
                    ? 'border-teal-800 text-slate-700 font-extrabold' 
                    : 'border-transparent text-slate-400 hover:text-slate-600'
                }`}
              >
                <HelpCircle size={14} className={activeDocTab === 'query' ? 'text-slate-700' : 'text-slate-400'} />
                QUERY DOCUMENTS ({queryDocuments.length})
              </button>
            </div>

            {/* Documents Content */}
            {activeDocTab === 'attachments' ? (
              docs.length > 0 ? (
                <div className="space-y-2 max-h-[300px] overflow-y-auto pr-1">
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
                        className="group flex items-center justify-between p-2 bg-slate-50/50 hover:bg-slate-50 hover:border-slate-300 rounded-lg border border-slate-100 transition-all min-w-0 cursor-pointer"
                      >
                        <div className="flex items-center gap-2 min-w-0">
                          <div className={`p-1.5 rounded border shrink-0 ${iconColor}`}>
                            <FileText size={12} />
                          </div>
                          <div className="flex flex-col min-w-0">
                            <span className="text-[10px] font-bold text-slate-700 truncate" title={doc.fileName}>
                              {doc.fileName}
                            </span>
                            <span className="text-[9px] text-slate-400 font-semibold mt-0.5">
                              {formatFileSize(doc.fileSize)}
                            </span>
                          </div>
                        </div>
                      </a>
                    );
                  })}
                </div>
              ) : (
                <div className="flex flex-col items-center justify-center py-6 border border-dashed border-slate-200 rounded-xl">
                  <Paperclip size={18} className="text-slate-300 mb-1" />
                  <span className="text-[9.5px] text-slate-400 font-bold italic">No attachments found</span>
                </div>
              )
            ) : (
              queryDocuments.length > 0 ? (
                <div className="space-y-2 max-h-[300px] overflow-y-auto pr-1">
                  {queryDocuments.map((doc, idx) => {
                    const ext = doc.fileName.split('.').pop()?.toLowerCase();
                    let iconColor = 'text-slate-500 bg-slate-50 border-slate-100';
                    if (ext === 'pdf') iconColor = 'text-rose-500 bg-rose-50 border-rose-100';
                    else if (['xlsx', 'xls', 'csv'].indexOf(ext || '') > -1) iconColor = 'text-emerald-500 bg-emerald-50 border-emerald-100';
                    else if (['docx', 'doc'].indexOf(ext || '') > -1) iconColor = 'text-blue-500 bg-blue-50 border-blue-100';
                    
                    return (
                      <a 
                        key={doc.id || idx}
                        href={doc.serverRelativeUrl} 
                        target="_blank" 
                        rel="noopener noreferrer" 
                        className="group flex items-center justify-between p-2 bg-slate-50/50 hover:bg-slate-50 hover:border-slate-300 rounded-lg border border-slate-100 transition-all min-w-0 cursor-pointer"
                      >
                        <div className="flex items-center gap-2 min-w-0">
                          <div className={`p-1.5 rounded border shrink-0 ${iconColor}`}>
                            <FileText size={12} />
                          </div>
                          <div className="flex flex-col min-w-0">
                            <span className="text-[10px] font-bold text-slate-700 truncate" title={doc.fileName}>
                              {doc.fileName}
                            </span>
                            <span className="text-[9px] text-slate-400 font-semibold mt-0.5">
                              Query File
                            </span>
                          </div>
                        </div>
                      </a>
                    );
                  })}
                </div>
              ) : (
                <div className="flex flex-col items-center justify-center py-6 border border-dashed border-slate-200 rounded-xl">
                  <Paperclip size={18} className="text-slate-300 mb-1" />
                  <span className="text-[9.5px] text-slate-400 font-bold italic">No query documents found</span>
                </div>
              )
            )}

            {docs.length > 0 && (
              <div className="mt-3 pt-2 border-t border-slate-50 text-center">
                <span className="text-[10px] font-bold text-blue-600 hover:text-blue-800 cursor-pointer">
                  View all documents
                </span>
              </div>
            )}
          </div>

          {/* Workflow History Card */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-4">
            <div className="flex items-center gap-2 mb-6 border-b border-slate-100 pb-2">
              <History size={16} className="text-blue-600 shrink-0"/> 
              <span className="font-bold text-slate-700 text-[10.5px] uppercase tracking-wider">Workflow History</span>
            </div>
            
            <div className="space-y-0.5">
              {history.length > 0 ? (
                (() => {
                  const displayedHistory = isHistoryExpanded ? history : history.slice(0, 3);
                  return (
                    <>
                      {displayedHistory.map((item, idx) => {
                        const rawText = item.rawText || "";
                        const lower = rawText.toLowerCase();
                        
                        let dotColor = "bg-blue-500";
                        if (lower.includes("approved") || lower.includes("initiated") || lower.includes("resolved") || lower.includes("submitted")) {
                          dotColor = "bg-emerald-500";
                        } else if (lower.includes("rejected")) {
                          dotColor = "bg-rose-500";
                        } else if (lower.includes("query")) {
                          dotColor = "bg-amber-500";
                        }
                        
                        return (
                          <div key={item.id || idx} className="relative pl-7 pb-1 last:pb-0">
                            {idx < displayedHistory.length - 1 && (
                              <div className="absolute left-[7px] top-4 bottom-0 w-[1.5px]" style={{ backgroundColor: '#cbd5e1' }} />
                            )}
                            <div className={`absolute left-0 top-1.5 w-3.5 h-3.5 rounded-full border-2 border-white flex items-center justify-center z-10 ${dotColor} shadow-sm`} />
                            
                            <div className="flex flex-col min-w-0">
                              <div className="leading-relaxed">
                                {renderHistoryText(rawText)}
                              </div>
                            </div>
                          </div>
                        );
                      })}
                      
                      {history.length > 3 && (
                        <div className="pt-1 text-left">
                          <button
                            type="button"
                            onClick={() => setIsHistoryExpanded(!isHistoryExpanded)}
                            className="text-[10px] font-bold text-blue-600 hover:text-blue-800 transition-colors uppercase tracking-wider cursor-pointer"
                          >
                            {isHistoryExpanded ? "Less...." : "More...."}
                          </button>
                        </div>
                      )}
                    </>
                  );
                })()
              ) : (
                <div className="text-center py-4 text-slate-400 italic text-[10px] font-bold">
                  No workflow history found.
                </div>
              )}
            </div>
          </div>

          {/* Unified Comments History Card */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-4">
            
            {/* Tab selection */}
            <div className="flex flex-wrap border-b border-slate-200 mb-3 gap-4 overflow-x-hidden">
              <button 
                type="button"
                onClick={() => setActiveCommentTab('comments')}
                className={`pb-2 font-bold transition-all flex items-center gap-1.5 border-b-2 -mb-[2px] uppercase tracking-wider text-[10.5px] ${
                  activeCommentTab === 'comments' 
                    ? 'border-teal-800 text-slate-700 font-extrabold' 
                    : 'border-transparent text-slate-400 hover:text-slate-600'
                }`}
              >
                <MessageSquare size={14} className={activeCommentTab === 'comments' ? 'text-slate-700' : 'text-slate-400'} />
                COMMENTS HISTORY
              </button>
              <button 
                type="button"
                onClick={() => setActiveCommentTab('query')}
                className={`pb-2 font-bold transition-all flex items-center gap-1.5 border-b-2 -mb-[2px] uppercase tracking-wider text-[10.5px] ${
                  activeCommentTab === 'query' 
                    ? 'border-teal-800 text-slate-700 font-extrabold' 
                    : 'border-transparent text-slate-400 hover:text-slate-600'
                }`}
              >
                <HelpCircle size={14} className={activeCommentTab === 'query' ? 'text-slate-700' : 'text-slate-400'} />
                QUERY REMARKS
              </button>
            </div>

            {/* Comments content */}
            {activeCommentTab === 'comments' ? (
              comments.filter(item => item.remarks && item.remarks !== "No remarks" && item.remarks.trim() !== "").length > 0 ? (
                <div className="space-y-3 max-h-[300px] overflow-y-auto pr-1">
                  {comments.filter(item => item.remarks && item.remarks !== "No remarks" && item.remarks.trim() !== "").map((item, idx) => (
                    <div key={item.id || idx} className="bg-slate-50/40 p-3.5 rounded-2xl border border-slate-100/80 flex flex-col gap-2.5 shadow-sm">
                      <p className="text-[13px] text-slate-700 font-medium leading-relaxed">
                        <span className="font-semibold text-slate-800">Comment:</span> {item.remarks || ""}
                      </p>
                      <div className="p-2.5 bg-white border border-slate-100 rounded-xl shadow-sm">
                        <span className="text-[11px] text-slate-500 font-medium leading-relaxed">
                          Remarks Added by <span className="font-semibold text-slate-700">{item.user}</span> at Level <span className="font-semibold text-slate-700">{item.level || "1"}</span> on <span className="font-semibold text-slate-700">{item.date}</span>
                        </span>
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <div className="text-center py-6 text-slate-400 italic text-[13px] font-bold">
                  No comment history found
                </div>
              )
            ) : (
              queryHistory.length > 0 ? (
                <div className="space-y-3 max-h-[300px] overflow-y-auto pr-1">
                  {queryHistory.map((item, idx) => (
                    <div key={item.id || idx} className="bg-slate-50/40 p-3.5 rounded-2xl border border-slate-100/80 flex flex-col gap-2.5 shadow-sm">
                      
                      {/* Query section */}
                      <div className="flex flex-col gap-1 min-w-0">
                        <p className="text-[13px] text-amber-700 leading-relaxed">
                          <span className="font-semibold text-amber-800">Query:</span> {item.Query || ""}
                        </p>
                        <div className="p-2 bg-white border border-slate-100 rounded-xl shadow-sm mt-1">
                          <span className="text-[11px] text-slate-500 font-medium">
                            Queried by <span className="font-semibold text-slate-700">{item.QueryBy ? item.QueryBy.Title : "Unknown"}</span> at Level <span className="font-semibold text-slate-700">{item.WFLevel || "1"}</span> on <span className="font-semibold text-slate-700">{item.QueriedOn || ""}</span>
                          </span>
                        </div>
                      </div>

                      {/* Response section */}
                      {item.Response && item.Response.trim() !== "" && (
                        <div className="flex flex-col gap-1 pl-3 border-l-2 border-emerald-300 mt-1 min-w-0">
                          <p className="text-[13px] text-emerald-700 leading-relaxed">
                            <span className="font-semibold text-emerald-800">Response:</span> {item.Response}
                          </p>
                          <div className="p-2 bg-white border border-slate-100 rounded-xl shadow-sm mt-1">
                            <span className="text-[11px] text-slate-500 font-medium">
                              Responded by <span className="font-semibold text-slate-700">{item.QueryTo ? item.QueryTo.Title : "Unknown"}</span> on <span className="font-semibold text-slate-700">{item.RespondedOn || ""}</span>
                            </span>
                          </div>
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              ) : (
                <div className="text-center py-6 text-slate-400 italic text-[13px] font-bold">
                  No query remarks history found
                </div>
              )
            )}

            {(comments.filter(item => item.remarks && item.remarks !== "No remarks" && item.remarks.trim() !== "").length > 0 || queryHistory.length > 0) && (
              <div className="mt-4 pt-2 border-t border-slate-50 text-center">
                <span className="text-[10px] font-bold text-blue-600 hover:text-blue-800 cursor-pointer">
                  View all comments
                </span>
              </div>
            )}
          </div>

        </div>

      </div>

      {/* Remarks & Action Console (Level 4) */}
      {isCurrentApprover && (
        <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-6 space-y-5 mt-6">
          <div className="flex items-center gap-2 mb-2 border-b border-slate-100 pb-3">
            <div className="w-1.5 h-4 bg-blue-600 rounded-full" />
            <span className="font-bold text-slate-800 text-xs uppercase tracking-wider">
              Action Console (Level {nfa.level || '4'})
            </span>
          </div>
          
          <div className="flex flex-col md:flex-row gap-6 items-stretch">
            
            {/* Remarks TextArea (Left 60%) */}
            <div className="w-full md:w-[60%] flex flex-col">
              <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-2 block shrink-0">
                Remarks / Comments <span className="text-rose-500">*</span>
              </label>
              <div className="relative flex-1">
                <textarea 
                  className="w-full bg-white border border-slate-200 rounded-xl px-4 py-3 text-[13px] focus:outline-none focus:ring-2 focus:ring-blue-500/15 focus:border-blue-500 transition-all resize-none h-[110px] leading-relaxed shadow-sm"
                  placeholder="Type your approval / rejection comments here..."
                  value={remarks}
                  onChange={(e) => setRemarks(e.target.value)}
                />
              </div>
            </div>
            
            {/* Delegate/Query User (Right 40%) */}
            <div className="w-full md:w-[40%] flex flex-col justify-start">
              
              {/* Custom People Picker */}
              <div className="relative" ref={suggestionsRef}>
                <label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-2 block">
                  Delegate / Query User
                </label>
                
                {selectedUserObj ? (
                  /* Selected User Pill */
                  <div className="flex items-center justify-between bg-blue-50/50 border border-blue-200/80 rounded-xl p-2.5 h-[42px] shadow-sm">
                    <div className="flex items-center gap-2 min-w-0">
                      <UserAvatar 
                        email={selectedUserObj.email} 
                        initials={getInitials(selectedUserObj.label)} 
                        className="w-6 h-6 rounded-full bg-blue-600 text-white text-[10px] shrink-0 shadow-inner" 
                      />
                      <span className="text-[13px] font-bold text-slate-700 truncate">{selectedUserObj.label}</span>
                    </div>
                    <button 
                      type="button" 
                      onClick={() => {
                        setSelectedUser('');
                        setUserSearch('');
                      }}
                      className="text-slate-400 hover:text-slate-600 font-bold text-[13px] p-1"
                    >
                      ✕
                    </button>
                  </div>
                ) : (
                  /* Search input */
                  <div className="relative">
                    <input 
                      type="text"
                      className="w-full bg-white border border-slate-200 rounded-xl pl-8 pr-4 py-2.5 text-[13px] focus:outline-none focus:ring-2 focus:ring-blue-500/15 focus:border-blue-500 transition-all h-[42px] shadow-sm"
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

                {/* Dropdown Suggestions */}
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
                          <span className="text-[13px] font-bold text-slate-700">{u.label}</span>
                        </button>
                      ))
                    ) : (
                      <div className="text-center py-3 text-slate-400 text-[13px] italic">
                        No matches found
                      </div>
                    )}
                  </div>
                )}
              </div>
              
              <span className="text-[9.5px] text-slate-400 font-semibold mt-2.5 block leading-none pl-0.5">
                Delegated user will be notified to take action.
              </span>
            </div>
            
          </div>

          {/* Centered Action Buttons Bar (Using inline background colors to avoid Tailwind purging issues) */}
          <div className="flex flex-wrap gap-3 items-center justify-center pt-5 border-t border-slate-100">
            <button 
              type="button"
              onClick={() => handleAction('Approve')} 
              disabled={isActionLoading} 
              style={{ backgroundColor: '#00b294', color: '#ffffff' }}
              className="px-5 h-[38px] hover:opacity-90 rounded-lg font-bold text-xs uppercase tracking-wider transition-all flex items-center gap-2 shadow-sm disabled:opacity-50 active:scale-[0.98]"
            >
              <Check size={14} strokeWidth={2.5} /> Approve
            </button>
            <button 
              type="button"
              onClick={() => handleAction('Reject')} 
              disabled={isActionLoading} 
              style={{ backgroundColor: '#e81123', color: '#ffffff' }}
              className="px-5 h-[38px] hover:opacity-90 rounded-lg font-bold text-xs uppercase tracking-wider transition-all flex items-center gap-2 shadow-sm disabled:opacity-50 active:scale-[0.98]"
            >
              <X size={14} strokeWidth={2.5} /> Reject
            </button>
            <button 
              type="button"
              onClick={() => handleAction('Query')} 
              disabled={isActionLoading} 
              style={{ backgroundColor: '#ff8c00', color: '#ffffff' }}
              className="px-5 h-[38px] hover:opacity-90 rounded-lg font-bold text-xs uppercase tracking-wider transition-all flex items-center gap-2 shadow-sm disabled:opacity-50 active:scale-[0.98]"
            >
              <HelpCircle size={14} strokeWidth={2.5} /> Query
            </button>
            <button 
              type="button"
              onClick={() => handleAction('Delegate')} 
              disabled={isActionLoading} 
              style={{ backgroundColor: '#0078d4', color: '#ffffff' }}
              className="px-5 h-[38px] hover:opacity-90 rounded-lg font-bold text-xs uppercase tracking-wider transition-all flex items-center gap-2 shadow-sm disabled:opacity-50 active:scale-[0.98]"
            >
              <UserCheck size={14} strokeWidth={2.5} /> Delegate
            </button>
            <button 
              type="button"
              onClick={() => handleAction('Return')} 
              disabled={isActionLoading} 
              style={{ backgroundColor: '#ffffff', color: '#475569', borderColor: '#cbd5e1' }}
              className="px-5 h-[38px] hover:bg-slate-50 rounded-lg font-bold text-xs uppercase tracking-wider border transition-all flex items-center gap-2 shadow-sm disabled:opacity-50 active:scale-[0.98]"
            >
              <CornerUpLeft size={14} strokeWidth={2.5} /> Return
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
