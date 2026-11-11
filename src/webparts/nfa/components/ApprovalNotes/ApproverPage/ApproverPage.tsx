import * as React from 'react';
import { useState, useEffect, useMemo, useRef } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import {
  Building2, FileText, Wrench, DollarSign, Paperclip,
  Settings, Star, ChevronDown, ArrowLeft, History,
  MessageSquare, User, CheckCircle2, Clock, Circle,
  ChevronRight, Check, X, HelpCircle, UserCheck, CornerUpLeft, Download,
  MapPin, Calendar, Hourglass, Layers, GitBranch, BadgeDollarSign, Target, Tag,
  ClipboardList, ExternalLink
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
import { forceScrollToTop, stripHtml } from '../../../../../Share/utils';
import { downloadNfaPdf } from '../../../../../Share/utils/nfaPdfGenerator';
import { downloadFilesAsZip } from '../../../../../Share/utils/zipUtils';

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
    <span className="leading-relaxed text-[13px] text-slate-600 font-medium">
      {parts.map((part, idx) => {
        if (statusRegex.test(part)) {
          return (
            <span key={idx} className="font-bold text-slate-800">
              {part}
            </span>
          );
        }
        return (
          <span key={idx}>
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
  const [users, setUsers] = useState<{ label: string, value: string, email?: string }[]>([]);
  const [history, setHistory] = useState<any[]>([]);
  const [comments, setComments] = useState<any[]>([]);
  const [queryHistory, setQueryHistory] = useState<any[]>([]);
  const [isGeneratingPdf, setIsGeneratingPdf] = useState<boolean>(false);
  const [isZippingDocs, setIsZippingDocs] = useState<boolean>(false);
  const [dashboardApprover, setDashboardApprover] = useState<{
    currentApprover: string;
    currentApproverId: string | number | null;
  } | null>(null);

  const handleDownloadPdf = async () => {
    if (!nfa) return;
    try {
      setIsGeneratingPdf(true);
      await downloadNfaPdf({ form: nfa, users, history, comments });
    } catch (err) {
      console.error('Error generating PDF:', err);
      alert('Failed to generate PDF. Please try again.');
    } finally {
      setIsGeneratingPdf(false);
    }
  };

  const handleDownloadZip = async () => {
    const isAttachmentTab = activeDocTab === 'attachments';
    const targetFiles = isAttachmentTab ? docs : queryDocuments;
    if (!targetFiles || targetFiles.length === 0) return;

    try {
      setIsZippingDocs(true);
      const docType = isAttachmentTab ? 'Attachments' : 'Query_Documents';
      const prefix = nfa?.requestNo ? `${nfa.requestNo}_${docType}` : `NFA_${docType}`;
      await downloadFilesAsZip(targetFiles, `${prefix}.zip`);
    } catch (err) {
      console.error('Error downloading files as ZIP:', err);
      alert('Failed to download files as ZIP. Please try again.');
    } finally {
      setIsZippingDocs(false);
    }
  };

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
  const [isDocsExpanded, setIsDocsExpanded] = useState(false);

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
        const [nfaData, siteUsers, historyData, qHistory, dashItems] = await Promise.all([
          nfaService.getNFA(parseInt(id)),
          sp.web.siteUsers.filter("PrincipalType eq 1").select("Id", "Title", "Email")()
            .then(users => users.map(u => ({ label: u.Title, value: String(u.Id), email: u.Email }))),
          nfaService.getHistory(parseInt(id)),
          commonService.getQueryHistory(id),
          sp.web.lists.getByTitle("DashboardList").items
            .filter(`WFID eq '${id}'`)
            .top(1)()
            .catch(async () => {
              try {
                return await sp.web.lists.getByTitle("DashboardList").items
                  .filter(`WFID eq ${id}`)
                  .top(1)();
              } catch {
                return [];
              }
            })
        ]);

        let dashItem: any = null;
        if (Array.isArray(dashItems) && dashItems.length > 0) {
          dashItem = dashItems[0];
        }

        if (dashItem) {
          // Check CurrentApprover and CurrentApproverID columns from DashboardList
          const cApproverId = dashItem.CurrentApproverID ?? dashItem.CurrentApproverId ?? dashItem.CurrentApprovalID ?? (dashItem.CurrentApprover && typeof dashItem.CurrentApprover === 'object' ? dashItem.CurrentApprover.Id : null);
          const cApprover = typeof dashItem.CurrentApprover === 'string'
            ? dashItem.CurrentApprover
            : (dashItem.CurrentApprover?.Title || dashItem.CurrentApprovalName || "");

          setDashboardApprover({
            currentApprover: cApprover,
            currentApproverId: cApproverId ?? null
          });
        }

        if (nfaData) {
          // If status is Query and pending user is current user, redirect to query response page
          const pendingApproverId = dashItem
            ? (dashItem.CurrentApproverID ?? dashItem.CurrentApproverId ?? dashItem.CurrentApprovalID ?? nfaData.pendingAtId)
            : nfaData.pendingAtId;
          const currentStatus = (dashItem?.Status || nfaData.status || "").toLowerCase();
          if (currentStatus === 'query' && String(pendingApproverId).trim() === String(currentUserId).trim()) {
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

          const attachedDocs = await docService.getDocumentsByUniqId(nfaData.uniqId, parseInt(id), nfaData.requestNo);
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

  // Check CurrentApproverID from DashboardList and compare with current logged-in user's ID
  const isCurrentApprover = useMemo(() => {
    if (dashboardApprover !== null) {
      const approverId = dashboardApprover.currentApproverId;
      if (approverId === null || approverId === undefined || approverId === "") {
        return false;
      }
      return String(approverId).trim() === String(currentUserId).trim();
    }
    // Fallback to nfa.pendingAtId if DashboardList record is not loaded yet
    return nfa?.pendingAtId === currentUserId;
  }, [dashboardApprover, currentUserId, nfa?.pendingAtId]);

  const currentApproverName = useMemo(() => {
    if (dashboardApprover?.currentApprover) {
      return dashboardApprover.currentApprover;
    }
    if (!nfa) return "Unknown";
    const user = users.find(u => u.value === String(dashboardApprover?.currentApproverId || nfa.pendingAtId));
    return user ? user.label : "Unknown";
  }, [dashboardApprover, nfa, users]);

  const receivedOnDate = useMemo(() => {
    if (!nfa) return "Unknown";
    if (history && history.length > 0) {
      const latestAction = history[0];
      if (latestAction && latestAction.date) {
        return latestAction.date;
      }
    }
    return nfa.date || "Unknown";
  }, [history, nfa]);

  const ageingText = useMemo(() => {
    if (!receivedOnDate || receivedOnDate === "Unknown") return "0 Hours";

    try {
      const match = receivedOnDate.match(/(\d{1,2})[/-](\d{1,2})[/-](\d{4})(?:\s+(\d{1,2}):(\d{2})\s*(AM|PM)?)?/i);
      if (!match) {
        const parsed = Date.parse(receivedOnDate);
        if (!isNaN(parsed)) {
          const diffMs = new Date().getTime() - parsed;
          if (diffMs <= 0) return "0 Hours";
          const diffHrs = Math.floor(diffMs / (1000 * 60 * 60));
          const days = Math.floor(diffHrs / 24);
          const remainingHours = diffHrs % 24;
          return days > 0 ? `${days} Day${days > 1 ? 's' : ''} ${remainingHours} Hour${remainingHours > 1 ? 's' : ''}` : `${remainingHours} Hour${remainingHours > 1 ? 's' : ''}`;
        }
        return "0 Hours";
      }

      const day = parseInt(match[1]);
      const month = parseInt(match[2]) - 1;
      const year = parseInt(match[3]);
      let parsedHours = match[4] ? parseInt(match[4]) : 0;
      const minutes = match[5] ? parseInt(match[5]) : 0;
      const ampm = match[6];

      if (ampm) {
        if (ampm.toUpperCase() === 'PM' && parsedHours < 12) parsedHours += 12;
        if (ampm.toUpperCase() === 'AM' && parsedHours === 12) parsedHours = 0;
      }

      const receivedDate = new Date(year, month, day, parsedHours, minutes);
      const diffMs = new Date().getTime() - receivedDate.getTime();
      if (diffMs <= 0) return "0 Hours";
      const diffHrs = Math.floor(diffMs / (1000 * 60 * 60));
      const days = Math.floor(diffHrs / 24);
      const remainingHours = diffHrs % 24;
      return days > 0 ? `${days} Day${days > 1 ? 's' : ''} ${remainingHours} Hour${remainingHours > 1 ? 's' : ''}` : `${remainingHours} Hour${remainingHours > 1 ? 's' : ''}`;
    } catch (err) {
      console.error("Error calculating ageing:", err);
      return "0 Hours";
    }
  }, [receivedOnDate]);

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
      <style>{`
        .no-scrollbar::-webkit-scrollbar {
          height: 4px !important;
        }
        .no-scrollbar::-webkit-scrollbar-track {
          background: #f8fafc !important;
          border-radius: 9999px !important;
        }
        .no-scrollbar::-webkit-scrollbar-thumb {
          background: #cbd5e1 !important;
          border-radius: 9999px !important;
        }
        .no-scrollbar::-webkit-scrollbar-thumb:hover {
          background: #94a3b8 !important;
        }
        .no-scrollbar::-webkit-scrollbar-button {
          display: none !important;
          width: 0 !important;
          height: 0 !important;
        }
        .no-scrollbar {
          scrollbar-width: thin !important;
          scrollbar-color: #cbd5e1 #f8fafc !important;
          -ms-overflow-style: auto !important;
        }
      `}</style>

      {/* Top Action Bar */}
      {/* <div className="flex items-center justify-between gap-4 bg-white rounded-xl border border-slate-200/80 shadow-sm py-2.5 px-4 mb-3.5">
        <div className="flex items-center gap-2.5">
          <span className="text-xs font-extrabold text-[#4A081E] tracking-tight">
            {nfa.requestNo || 'Note for Approval'}
          </span>
          <span className="text-[10px] font-extrabold uppercase px-2 py-0.5 rounded-full bg-slate-100 text-slate-600 border border-slate-200">
            Level {nfa.level || '1'}
          </span>
        </div>
        <div className="flex items-center gap-2">
          <button
            type="button"
            onClick={handleDownloadPdf}
            disabled={isGeneratingPdf}
            className="flex items-center gap-1.5 px-3.5 py-1.5 bg-[#4A081E] hover:bg-[#380617] text-white rounded-lg font-bold text-xs shadow-sm transition-all active:scale-95 disabled:opacity-50 cursor-pointer"
            title="Download PDF"
          >
            <Download size={14} />
            <span>{isGeneratingPdf ? 'Generating PDF...' : 'Download PDF'}</span>
          </button>
          <button
            type="button"
            onClick={() => {
              forceScrollToTop();
              navigate('/inbox');
            }}
            className="px-3.5 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg font-bold text-xs transition-all border border-slate-300 active:scale-95 cursor-pointer"
          >
            Back
          </button>
        </div>
      </div> */}

      {/* Approval Chain Card */}
      <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm py-4 px-6 mb-6">
        <div
          ref={stepperContainerRef}
          className="relative flex items-center justify-start gap-1 overflow-x-auto pb-2 pt-0.5 no-scrollbar"
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
            let labelColor = 'text-slate-400';
            let statusBadge: React.ReactNode = null;

            if (isCurrent) {
              borderClass = 'border-amber-500 bg-amber-50/10 shadow-lg shadow-amber-100';
              avatarClass = 'bg-amber-500 text-white animate-pulse';
              labelColor = 'text-amber-600 font-extrabold';
              statusBadge = (
                <div className="w-4 h-4 rounded-full bg-amber-500 text-white flex items-center justify-center shrink-0 shadow-sm shadow-amber-100 animate-pulse">
                  <Clock size={10} strokeWidth={3} />
                </div>
              );
            } else if (isInitiator) {
              borderClass = 'border-slate-200 bg-white';
              avatarClass = 'bg-blue-500 text-white';
              labelColor = 'text-blue-600 font-extrabold';
              statusBadge = (
                <div className="w-4 h-4 rounded-full bg-emerald-500 text-white flex items-center justify-center shrink-0 shadow-sm shadow-emerald-100">
                  <Check size={10} strokeWidth={4} />
                </div>
              );
            } else if (isCompleted) {
              borderClass = 'border-slate-200 bg-white';
              avatarClass = 'bg-emerald-500 text-white';
              labelColor = 'text-emerald-600 font-extrabold';
              statusBadge = (
                <div className="w-4 h-4 rounded-full bg-emerald-500 text-white flex items-center justify-center shrink-0 shadow-sm shadow-emerald-100">
                  <Check size={10} strokeWidth={4} />
                </div>
              );
            } else {
              borderClass = 'border-slate-200 bg-white';
              avatarClass = 'bg-slate-100 text-slate-400';
              labelColor = 'text-slate-400';
              statusBadge = (
                <div className="w-4 h-4 rounded-full bg-slate-50 border border-slate-200 text-slate-400 flex items-center justify-center shrink-0">
                  <Clock size={10} strokeWidth={2.5} />
                </div>
              );
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

                  <div className={`w-[155px] h-[40px] rounded-xl border flex flex-row items-center gap-2 px-2 py-0.5 transition-all duration-350 ${borderClass}`}>
                    <UserAvatar
                      email={user?.email}
                      initials={initials}
                      className={`w-6 h-6 rounded-full shadow-inner shrink-0 ${avatarClass}`}
                    />
                    <span className="text-[10px] font-extrabold text-slate-700 text-left leading-tight line-clamp-2 flex-1">
                      {userLabel}
                    </span>
                    {statusBadge}
                  </div>
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
            {/* Card Header — blue icon box like reference image */}
            <div className="flex items-center justify-between mb-4 border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2.5">
                <div className="w-7 h-7 bg-blue-50 border border-blue-100 rounded-lg flex items-center justify-center shrink-0">
                  <ClipboardList size={15} className="text-blue-600" />
                </div>
                <span className="font-bold text-slate-800 text-[14px]">Request Overview</span>
                <span className="bg-blue-50 text-blue-600 text-[11px] font-extrabold px-2 py-0.5 rounded-md border border-blue-100 uppercase ml-1">
                  Level {nfa.level || '1'}
                </span>
              </div>
              <span className="text-[12.5px] font-bold text-slate-500">
                {nfa.requestNo || '-'}
              </span>
            </div>

            {/* Strict 2-Column Details Grid — all fields in pairs, matching reference */}
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px 32px' }}>

              {/* Col 1: Project */}
              <div className="flex items-start gap-2.5">
                <div className="w-6 h-6 bg-blue-50 rounded-md flex items-center justify-center shrink-0 mt-0.5">
                  <Building2 size={13} className="text-blue-500" />
                </div>
                <div className="flex flex-col min-w-0">
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Entity</span>
                  <span className="text-[13px] text-slate-700 font-semibold mt-0.5">{nfa.project || '-'}</span>
                </div>
              </div>

              {/* Col 2: Sub-Entity */}
              <div className="flex items-start gap-2.5">
                <div className="w-6 h-6 bg-teal-50 rounded-md flex items-center justify-center shrink-0 mt-0.5">
                  <MapPin size={13} className="text-teal-500" />
                </div>
                <div className="flex flex-col min-w-0">
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Sub-Entity</span>
                  <span className="text-[13px] text-slate-700 font-semibold mt-0.5">
                    {nfa.projectSubEntity && nfa.projectSubEntity.length > 0 ? nfa.projectSubEntity.join(', ') : '-'}
                  </span>
                </div>
              </div>

              {/* Col 1: Function */}
              <div className="flex items-start gap-2.5">
                <div className="w-6 h-6 bg-purple-50 rounded-md flex items-center justify-center shrink-0 mt-0.5">
                  <Layers size={13} className="text-purple-500" />
                </div>
                <div className="flex flex-col min-w-0">
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Function</span>
                  <span className="text-[13px] text-slate-700 font-semibold mt-0.5">{nfa.originatorDept || '-'}</span>
                </div>
              </div>

              {/* Col 2: Date */}
              <div className="flex items-start gap-2.5">
                <div className="w-6 h-6 bg-indigo-50 rounded-md flex items-center justify-center shrink-0 mt-0.5">
                  <Calendar size={13} className="text-indigo-500" />
                </div>
                <div className="flex flex-col min-w-0">
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Date</span>
                  <span className="text-[13px] text-slate-700 font-semibold mt-0.5">{nfa.date || '-'}</span>
                </div>
              </div>

              {/* Col 1: Data Matrix */}
              <div className="flex items-start gap-2.5">
                <div className="w-6 h-6 bg-amber-50 rounded-md flex items-center justify-center shrink-0 mt-0.5">
                  <Layers size={13} className="text-amber-500" />
                </div>
                <div className="flex flex-col min-w-0">
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">DOA Matrix</span>
                  <span className="text-[13px] text-slate-700 font-semibold mt-0.5">{nfa.doaMatrix || '-'}</span>
                </div>
              </div>

              {/* Col 2: Workflow */}
              <div className="flex items-start gap-2.5">
                <div className="w-6 h-6 bg-emerald-50 rounded-md flex items-center justify-center shrink-0 mt-0.5">
                  <GitBranch size={13} className="text-emerald-500" />
                </div>
                <div className="flex flex-col min-w-0">
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Workflow</span>
                  <span className="text-[13px] text-slate-700 font-semibold mt-0.5 uppercase">{nfa.workflow || '-'}</span>
                </div>
              </div>

            </div>
          </div>

          {/* Financial Summary Card */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-5 space-y-4">
            <div className="flex items-center gap-2.5 border-b border-slate-100 pb-3">
              <div className="w-7 h-7 bg-emerald-50 border border-emerald-100 rounded-lg flex items-center justify-center shrink-0">
                <DollarSign size={15} className="text-emerald-600" />
              </div>
              <span className="font-bold text-slate-800 text-[14px]">Financial Summary</span>
            </div>

            {/* Budget, Awarded, Capex — forced single row */}
            <div style={{ display: 'flex', flexDirection: 'row', gap: '10px', width: '100%' }}>
              {/* Packeage Awarded Cost */}
              <div style={{ flex: '1 1 0%', minWidth: 0, padding: '10px', background: 'rgba(236,253,245,0.6)', border: '1px solid #a7f3d0', borderRadius: '12px', display: 'flex', flexDirection: 'column', gap: '6px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <div style={{ width: '18px', height: '18px', background: '#d1fae5', borderRadius: '6px', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                    <Target size={10} className="text-emerald-600" />
                  </div>
                  <span style={{ fontSize: '11px', fontWeight: 700, color: '#059669', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Packeage Awarded Cost</span>
                </div>
                <div style={{ fontSize: '14.5px', fontWeight: 800, color: '#047857' }}>₹ {Number(nfa.packageAwardedCost || 0).toLocaleString('en-IN')}</div>
              </div>

              {/* Budget */}
              <div style={{ flex: '1 1 0%', minWidth: 0, padding: '10px', background: 'rgba(239,246,255,0.6)', border: '1px solid #bfdbfe', borderRadius: '12px', display: 'flex', flexDirection: 'column', gap: '6px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <div style={{ width: '18px', height: '18px', background: '#dbeafe', borderRadius: '6px', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                    <BadgeDollarSign size={10} className="text-blue-600" />
                  </div>
                  <span style={{ fontSize: '11px', fontWeight: 700, color: '#2563eb', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Budget</span>
                </div>
                <div style={{ fontSize: '14.5px', fontWeight: 800, color: '#1d4ed8' }}>₹ {Number(nfa.budget || 0).toLocaleString('en-IN')}</div>
              </div>



              {/*Budgeted/ Non Budgeted */}
              <div style={{ flex: '1 1 0%', minWidth: 0, padding: '10px', background: 'rgba(245,243,255,0.6)', border: '1px solid #ddd6fe', borderRadius: '12px', display: 'flex', flexDirection: 'column', gap: '6px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <div style={{ width: '18px', height: '18px', background: '#ede9fe', borderRadius: '6px', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                    <Tag size={10} className="text-purple-600" />
                  </div>
                  <span style={{ fontSize: '11px', fontWeight: 700, color: '#7c3aed', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Budgeted/ Non Budgeted</span>
                </div>
                <div style={{ fontSize: '13.5px', fontWeight: 700, color: '#6d28d9', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={nfa.budgetedNonBudgeted || ''}>{nfa.budgetedNonBudgeted || '-'}</div>
              </div>
            </div>

            {nfa.budgetRemarks && nfa.budgetRemarks.trim() !== "" && (
              <div className="p-3 bg-slate-50 border border-slate-100 rounded-xl">
                <div className="flex items-center gap-1.5 mb-1">
                  <MessageSquare size={11} className="text-slate-400" />
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Remarks</span>
                </div>
                <div className="text-[13.5px] text-slate-600 font-medium leading-relaxed whitespace-pre-wrap">{nfa.budgetRemarks}</div>
              </div>
            )}
          </div>

          {/* Details Card with Tabs */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-5">

            {/* Horizontal Tabs Header — matching reference image style */}
            <div className="flex flex-wrap border-b border-slate-200 gap-x-6 gap-y-2.5 mb-5">
              <button
                type="button"
                onClick={() => setActiveTab('details')}
                className={`font-bold transition-all flex items-center gap-2 px-3.5 pb-2.5 border-b-2 -mb-[2px] text-[13px] ${activeTab === 'details'
                  ? 'border-blue-600 text-blue-700 font-extrabold'
                  : 'border-transparent text-slate-500 hover:text-slate-700'
                  }`}
              >
                <Paperclip size={15} className={activeTab === 'details' ? 'text-blue-600' : 'text-slate-400'} />
                NFA Details
              </button>
              <button
                type="button"
                onClick={() => setActiveTab('terms')}
                className={`font-bold transition-all flex items-center gap-2 px-3.5 pb-2.5 border-b-2 -mb-[2px] text-[13px] ${activeTab === 'terms'
                  ? 'border-blue-600 text-blue-700 font-extrabold'
                  : 'border-transparent text-slate-500 hover:text-slate-700'
                  }`}
              >
                <Settings size={15} className={activeTab === 'terms' ? 'text-blue-600' : 'text-slate-400'} />
                Contract Terms
              </button>
              <button
                type="button"
                onClick={() => setActiveTab('recommendation')}
                className={`font-bold transition-all flex items-center gap-2 px-3.5 pb-2.5 border-b-2 -mb-[2px] text-[13px] ${activeTab === 'recommendation'
                  ? 'border-blue-600 text-blue-700 font-extrabold'
                  : 'border-transparent text-slate-500 hover:text-slate-700'
                  }`}
              >
                <Star size={15} className={activeTab === 'recommendation' ? 'text-blue-600' : 'text-slate-400'} />
                Recommendation
              </button>

              <button
                type="button"
                onClick={() => setActiveTab('notes')}
                className={`font-bold transition-all flex items-center gap-2 px-3.5 pb-2.5 border-b-2 -mb-[2px] text-[13px] ${activeTab === 'notes'
                  ? 'border-blue-600 text-blue-700 font-extrabold'
                  : 'border-transparent text-slate-500 hover:text-slate-700'
                  }`}
              >
                <Paperclip size={15} className={activeTab === 'notes' ? 'text-blue-600' : 'text-slate-400'} />
                Annexures & Notes
              </button>
            </div>

            {/* Tab Contents */}
            {activeTab === 'details' && (
              <div className="space-y-6">
                <div>
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Subject</span>
                  <h2 className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1">{stripHtml(nfa.referenceSubject) || '-'}</h2>
                </div>

                <div className="border-t border-slate-100 pt-4">
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Introduction / Preamble / Background</span>
                  <div
                    className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1.5 richTextContent transition-all duration-300"
                    style={!isIntroExpanded ? { display: '-webkit-box', WebkitLineClamp: 7, WebkitBoxOrient: 'vertical', overflow: 'hidden' } : undefined}
                  >
                    <div dangerouslySetInnerHTML={{ __html: nfa.introduction || '-' }} />
                  </div>
                  {nfa.introduction && nfa.introduction.replace(/<[^>]*>/g, "").length > 450 && (
                    <button
                      type="button"
                      onClick={() => setIsIntroExpanded(!isIntroExpanded)}
                      className="text-[12px] font-bold text-blue-600 hover:text-blue-800 mt-1 cursor-pointer block"
                    >
                      {isIntroExpanded ? 'View less' : 'View more'}
                    </button>
                  )}
                </div>
                <div className="flex flex-row flex-nowrap items-stretch gap-3 w-full border-t border-slate-100 pt-4">

                  {/* Complete Indent / PR Received */}
                  <div className="flex flex-col flex-1 min-w-0 p-3 bg-slate-50/40 border border-slate-100 rounded-xl">
                    <span className="block text-[11px] font-bold text-slate-400 uppercase tracking-wider leading-5">
                      Complete Indent/PR received from & Date
                    </span>
                    <div className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1.5 whitespace-pre-wrap break-words">
                      {nfa.completeIndentPRReceivedFromDate
                        ? (() => {
                          const [year, month, day] =
                            nfa.completeIndentPRReceivedFromDate.split("T")[0].split("-");
                          return `${day.slice(-1)}/${month}/${year}`;
                        })()
                        : "-"}
                    </div>
                  </div>

                  {/* BOQ / Specifications Revisions */}
                  <div className="flex flex-col flex-1 min-w-0 p-3 bg-slate-50/40 border border-slate-100 rounded-xl">
                    <span className="block text-[11px] font-bold text-slate-400 uppercase tracking-wider leading-5">
                      BOQ/Specifications Revisions by & Date
                    </span>
                    <div className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1.5 whitespace-pre-wrap break-words">
                      {nfa.boqSpecificationsRevisionsByDate
                        ? (() => {
                          const [year, month, day] =
                            nfa.boqSpecificationsRevisionsByDate.split("T")[0].split("-");
                          return `${day.slice(-1)}/${month}/${year}`;
                        })()
                        : "-"}
                    </div>
                  </div>

                  {/* RFQ Ref No */}
                  <div className="flex flex-col flex-1 min-w-0 p-3 bg-slate-50/40 border border-slate-100 rounded-xl">
                    <span className="block text-[11px] font-bold text-slate-400 uppercase tracking-wider leading-5">
                      RFQ Ref No
                    </span>
                    <div className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1.5 whitespace-pre-wrap break-words">
                      {nfa.RFQRefNo || "-"}
                    </div>
                  </div>

                </div>

                <div className="flex flex-col gap-1">
                  <span className="text-[9.5px] font-bold text-slate-400 uppercase tracking-wider">Remarks</span>
                  <span className="text-xs font-semibold text-slate-700 bg-slate-50/60 p-2.5 rounded-lg border border-slate-100 whitespace-pre-wrap">{stripHtml(nfa.remark) || '-'}</span>
                </div>

                <div className="border-t border-slate-100 pt-4">
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Chronology</span>
                  <div
                    className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1 richTextContent transition-all duration-300"
                    style={!isChronologyExpanded ? { display: '-webkit-box', WebkitLineClamp: 7, WebkitBoxOrient: 'vertical', overflow: 'hidden' } : undefined}
                  >
                    <div dangerouslySetInnerHTML={{ __html: nfa.chronology || '-' }} />
                  </div>
                  {nfa.chronology && nfa.chronology.replace(/<[^>]*>/g, "").length > 450 && (
                    <button
                      type="button"
                      onClick={() => setIsChronologyExpanded(!isChronologyExpanded)}
                      className="text-[12px] font-bold text-blue-600 hover:text-blue-800 mt-1 cursor-pointer block"
                    >
                      {isChronologyExpanded ? 'View less' : 'View more'}
                    </button>
                  )}
                </div>

                <div className="border-t border-slate-100 pt-4">
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Commercial Proposals</span>
                  <div
                    className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1 richTextContent transition-all duration-300"
                    style={!isCommercialExpanded ? { display: '-webkit-box', WebkitLineClamp: 7, WebkitBoxOrient: 'vertical', overflow: 'hidden' } : undefined}
                  >
                    <div dangerouslySetInnerHTML={{ __html: nfa.commercialProposal || '-' }} />
                  </div>
                  {nfa.commercialProposal && nfa.commercialProposal.replace(/<[^>]*>/g, "").length > 450 && (
                    <button
                      type="button"
                      onClick={() => setIsCommercialExpanded(!isCommercialExpanded)}
                      className="text-[12px] font-bold text-blue-600 hover:text-blue-800 mt-1 cursor-pointer block"
                    >
                      {isCommercialExpanded ? 'View less' : 'View more'}
                    </button>
                  )}
                </div>

                <div className="border-t border-slate-100 pt-4">
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Payment Terms</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1 richTextContent">
                    <div dangerouslySetInnerHTML={{ __html: nfa.paymentTerms || '-' }} />
                  </div>
                </div>
              </div>
            )}

            {activeTab === 'recommendation' && (
              <div className="space-y-4">
                <div>
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Recommendation with Justification</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1.5 richTextContent">
                    <div dangerouslySetInnerHTML={{ __html: nfa.recommendation || '-' }} />
                  </div>
                </div>
                <div className="flex flex-col md:flex-row gap-4">
                  <div className="flex-1 p-3.5 bg-blue-50/20 border border-blue-100/50 rounded-xl">
                    <span className="text-[11px] font-bold text-blue-600 uppercase tracking-wider">Contractor Name</span>
                    <div className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-0.5">{nfa.contractorName || '-'}</div>
                  </div>
                </div>


              </div>
            )}

            {activeTab === 'terms' && (
              <div className="space-y-4">
                <div className="p-3 bg-slate-50/40 border border-slate-100 rounded-xl">
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Liquidated Damage</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1.5 whitespace-pre-wrap">{nfa.liquidatedDamage || '-'}</div>
                </div>
                <div className="p-3 bg-slate-50/40 border border-slate-100 rounded-xl">
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Contract Performance Gurantee</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1.5 whitespace-pre-wrap">{nfa.contractPerf || '-'}</div>
                </div>
                <div className="p-3 bg-slate-50/40 border border-slate-100 rounded-xl">
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Defect Liability Period</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1.5 whitespace-pre-wrap">{nfa.defectLiability || '-'}</div>
                </div>
                <div className="p-3 bg-slate-50/40 border border-slate-100 rounded-xl">
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Work Completion</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed mt-1.5 whitespace-pre-wrap">{nfa.workCompletion || '-'}</div>
                </div>
              </div>
            )}

            {activeTab === 'notes' && (
              <div className="space-y-4">
                <div>
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Annexure/ Enclosur(if any)</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed bg-slate-50/40 p-3 rounded-xl border border-slate-100 mt-1 whitespace-pre-wrap">{nfa.annexure || '-'}</div>
                </div>
                <div>
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Note 1</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed bg-slate-50/40 p-3 rounded-xl border border-slate-100 mt-1 whitespace-pre-wrap">{nfa.note1 || '-'}</div>
                </div>
                <div>
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Note 2</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed bg-slate-50/40 p-3 rounded-xl border border-slate-100 mt-1 whitespace-pre-wrap">{nfa.note2 || '-'}</div>
                </div>
                <div>
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Note 3</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed bg-slate-50/40 p-3 rounded-xl border border-slate-100 mt-1 whitespace-pre-wrap">{nfa.note3 || '-'}</div>
                </div>
                <div>
                  <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Note 4</span>
                  <div className="text-[13px] text-slate-700 font-semibold leading-relaxed bg-slate-50/40 p-3 rounded-xl border border-slate-100 mt-1 whitespace-pre-wrap">{nfa.note4 || '-'}</div>
                </div>
              </div>
            )}

          </div>

        </div>

        {/* ── Right Column (30% Width) ── */}
        <div className="w-full lg:w-[30%] space-y-6">

          {/* Documents Card — reference style */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-4">

            {/* Header / Tabs in one row */}
            <div className="flex items-center justify-between gap-1.5 pb-2 border-b border-slate-100 mb-3 flex-wrap sm:flex-nowrap">
              <div className="flex items-center gap-2 min-w-0">
                <button
                  type="button"
                  className={`flex items-center gap-1 pb-1.5 border-b-2 -mb-[10px] transition-all text-[11px] uppercase tracking-wider shrink-0 ${activeDocTab === 'attachments'
                    ? 'border-blue-600 text-blue-700 font-extrabold'
                    : 'border-transparent text-slate-500 hover:text-slate-700 font-bold'
                    }`}
                  onClick={() => setActiveDocTab('attachments')}
                >
                  <FileText size={13} className={activeDocTab === 'attachments' ? 'text-blue-600' : 'text-slate-400'} />
                  <span className="text-[11px] font-semibold">Docs ({docs.length})</span>
                </button>

                <button
                  type="button"
                  className={`flex items-center gap-1 pb-1.5 border-b-2 -mb-[10px] transition-all text-[11px] uppercase tracking-wider shrink-0 ${activeDocTab === 'query'
                    ? 'border-blue-600 text-blue-700 font-extrabold'
                    : 'border-transparent text-slate-500 hover:text-slate-700 font-bold'
                    }`}
                  onClick={() => setActiveDocTab('query')}
                >
                  <HelpCircle size={13} className={activeDocTab === 'query' ? 'text-blue-600' : 'text-slate-400'} />
                  <span className="text-[11px] font-semibold">Query Docs ({queryDocuments.length})</span>
                </button>
              </div>

              {/* Download ZIP button */}
              {((activeDocTab === 'attachments' && docs.length > 0) || (activeDocTab === 'query' && queryDocuments.length > 0)) && (
                <button
                  type="button"
                  onClick={handleDownloadZip}
                  disabled={isZippingDocs}
                  className="flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-bold bg-blue-50 text-blue-700 border border-blue-200 hover:bg-blue-100 transition-colors cursor-pointer disabled:opacity-50 shrink-0 whitespace-nowrap ml-auto"
                  title="Download files in this tab as ZIP"
                >
                  <Download size={11} className={isZippingDocs ? "animate-spin" : ""} />
                  <span>{isZippingDocs ? "Zipping..." : "Download ZIP"}</span>
                </button>
              )}
            </div>

            {/* Content area */}
            <div className="pt-3">
              {activeDocTab === 'attachments' ? (
                docs.length > 0 ? (
                  <>
                    <div className="relative">
                      <div className={`space-y-2 overflow-y-auto pr-1 transition-all ${isDocsExpanded ? 'max-h-[500px]' : 'max-h-[210px]'
                        }`}>
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
                              data-interception="off"
                              className="group flex items-center justify-between gap-2 p-1.5 px-2 bg-slate-50/50 hover:bg-emerald-50/15 rounded-lg border border-slate-100 hover:border-emerald-150 transition-all min-w-0"
                            >
                              <div className="flex items-center gap-2 min-w-0 overflow-hidden"><div className={`p-1 rounded border transition-colors shrink-0 ${iconColor}`}>
                                <FileText size={13} />
                              </div>
                                <span className="text-[13px] font-semibold text-slate-700 truncate group-hover:text-emerald-600 transition-colors" title={doc.fileName}>{doc.fileName}</span></div><span className="shrink-0 text-[10px] font-bold px-2 py-0.5 rounded-md bg-slate-100 text-slate-700 border border-slate-200 group-hover:bg-emerald-50 group-hover:text-emerald-700 group-hover:border-emerald-200 transition-colors">{doc.documentType || 'Attachment'}</span><span style={{ display: 'none' }}>
                                {doc.fileName}
                              </span>
                            </a>
                          );
                        })}
                      </div>
                      {!isDocsExpanded && docs.length > 3 && (
                        <div className="absolute bottom-0 left-0 right-0 h-8 bg-gradient-to-t from-white to-transparent pointer-events-none" />
                      )}
                    </div>
                    <div className="mt-3 text-center">
                      <button
                        type="button"
                        onClick={() => setIsDocsExpanded(!isDocsExpanded)}
                        className="text-[12px] font-bold text-blue-600 hover:text-blue-800 transition-colors"
                      >
                        {isDocsExpanded ? 'Show Less ↑' : `View All (${docs.length}) ↓`}
                      </button>
                    </div>
                  </>
                ) : (
                  <div className="flex flex-col items-center justify-center py-8 text-center bg-slate-50/20 border border-dashed border-slate-100 rounded-xl">
                    <div className="w-10 h-10 bg-slate-50 rounded-full flex items-center justify-center mb-2">
                      <Paperclip size={18} className="text-slate-300" />
                    </div>
                    <span className="text-[13px] text-slate-400 font-semibold">No attachments found</span>
                    <span className="text-[12px] text-slate-300 mt-0.5">Files will appear here once uploaded</span>
                  </div>
                )
              ) : (
                queryDocuments.length > 0 ? (
                  <div className="space-y-2 max-h-[210px] overflow-y-auto pr-1">
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
                          data-interception="off"
                          className="group flex items-center gap-2 p-1.5 bg-slate-50/50 hover:bg-amber-50/15 rounded-lg border border-slate-100 hover:border-amber-150 transition-all min-w-0"
                        >
                          <div className={`p-1 rounded border transition-colors shrink-0 ${iconColor}`}>
                            <FileText size={13} />
                          </div>
                          <span className="text-[13px] font-semibold text-slate-700 truncate group-hover:text-amber-600 transition-colors" title={doc.fileName}>
                            {doc.fileName}
                          </span>
                        </a>
                      );
                    })}
                  </div>
                ) : (
                  <div className="flex flex-col items-center justify-center py-8 text-center bg-slate-50/20 border border-dashed border-slate-100 rounded-xl">
                    <div className="w-10 h-10 bg-slate-50 rounded-full flex items-center justify-center mb-2">
                      <HelpCircle size={18} className="text-slate-300" />
                    </div>
                    <span className="text-[13px] text-slate-400 font-semibold">No query documents found</span>
                  </div>
                )
              )}
            </div>
          </div>

          {/* Workflow History Card — matching reference image */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-5">
            <div className="flex items-center gap-2 mb-4 border-b border-slate-100 pb-2.5">
              <History size={15} className="text-blue-500 shrink-0" />
              <span className="font-bold text-slate-700 text-[12px] uppercase tracking-wider">Workflow History</span>
            </div>

            <div className="space-y-4">
              {history.length > 0 ? (
                (() => {
                  const displayedHistory = isHistoryExpanded ? history : history.slice(0, 4);
                  return (
                    <>
                      {displayedHistory.map((item, idx) => {
                        const rawText = item.rawText || "";
                        const lower = rawText.toLowerCase();

                        let dotClass = 'bg-blue-500';
                        if (lower.includes("approved") || lower.includes("resolved") || lower.includes("submitted")) {
                          dotClass = 'bg-emerald-500';
                        } else if (lower.includes("rejected")) {
                          dotClass = 'bg-rose-500';
                        } else if (lower.includes("query")) {
                          dotClass = 'bg-amber-500';
                        }

                        let descText = rawText;
                        let dateText = item.date || "";

                        if (dateText) {
                          const dateStr = dateText.trim();
                          const lastOnIndex = rawText.lastIndexOf(" on " + dateStr);
                          if (lastOnIndex !== -1) {
                            descText = rawText.substring(0, lastOnIndex).trim();
                          } else {
                            const dateIndex = rawText.lastIndexOf(dateStr);
                            if (dateIndex !== -1) {
                              descText = rawText.substring(0, dateIndex).replace(/\s+on\s*$/, "").trim();
                            }
                          }
                        } else {
                          const dateRegex = /\s+on\s+(\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM)?)$/i;
                          const match = rawText.match(dateRegex);
                          if (match) {
                            descText = rawText.substring(0, match.index).trim();
                            dateText = match[1].trim();
                          }
                        }

                        return (
                          <div key={item.id || idx} className="relative pl-7 pb-1 last:pb-0">
                            {idx < displayedHistory.length - 1 && (
                              <div className="absolute left-[7px] top-5 bottom-0 w-[1.5px] bg-slate-100" />
                            )}
                            <div className={`absolute left-0 top-1 w-3.5 h-3.5 rounded-full border border-white flex items-center justify-center z-10 ${dotClass} shadow-sm`} />

                            <div className="flex flex-col min-w-0">
                              <p className="text-[13px] text-slate-600 font-medium leading-relaxed">
                                {renderHistoryText(descText)}
                              </p>
                              {dateText && (
                                <span className="text-[11px] text-slate-400 mt-0.5 font-normal">
                                  {dateText}
                                </span>
                              )}
                            </div>
                          </div>
                        );
                      })}

                      {history.length > 4 && (
                        <div className="pt-3 text-center">
                          <button
                            type="button"
                            onClick={() => setIsHistoryExpanded(!isHistoryExpanded)}
                            className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full border border-slate-200 text-[12px] font-bold text-blue-600 hover:bg-blue-50 hover:border-blue-300 transition-all uppercase tracking-wider"
                          >
                            {isHistoryExpanded ? 'SHOW LESS' : `VIEW MORE (${history.length - 4})`}
                          </button>
                        </div>
                      )}
                    </>
                  );
                })()
              ) : (
                <div className="text-center py-4 text-slate-400 italic text-[13px]">
                  No audit history found.
                </div>
              )}
            </div>
          </div>

          {/* Comments & Query History Card — matching reference image */}
          <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-5">
            <div className="flex items-center gap-2 mb-3 border-b border-slate-100 pb-2">
              <MessageSquare size={14} className="text-indigo-500 shrink-0" />
              <span className="font-bold text-slate-700 text-[12px] uppercase tracking-wider">Comments History</span>
            </div>

            {/* Subtle Pill-style tabs below the main card line header */}
            <div className="flex items-center gap-3 mb-4 pt-1">
              <button
                type="button"
                className={`px-3 py-1 rounded-full text-[12px] font-bold transition-all border ${activeCommentTab === 'comments'
                  ? 'bg-blue-50 text-blue-700 border-blue-200 shadow-sm'
                  : 'bg-slate-50 text-slate-500 border-slate-200/80 hover:bg-slate-100'
                  }`}
                onClick={() => setActiveCommentTab('comments')}
              >
                Comments ({comments.filter(item => item.remarks && item.remarks !== "No remarks" && item.remarks.trim() !== "").length})
              </button>

              <button
                type="button"
                className={`px-3 py-1 rounded-full text-[12px] font-bold transition-all border ${activeCommentTab === 'query'
                  ? 'bg-blue-50 text-blue-700 border-blue-200 shadow-sm'
                  : 'bg-slate-50 text-slate-500 border-slate-200/80 hover:bg-slate-100'
                  }`}
                onClick={() => setActiveCommentTab('query')}
              >
                Query Remarks ({queryHistory.length})
              </button>
            </div>

            {/* Content area */}
            <div className="pt-1">
              {activeCommentTab === 'comments' ? (
                comments.filter(item => item.remarks && item.remarks !== "No remarks" && item.remarks.trim() !== "").length > 0 ? (
                  <div className="space-y-3 max-h-[300px] overflow-y-auto pr-1 animate-fade-in">
                    {comments.filter(item => item.remarks && item.remarks !== "No remarks" && item.remarks.trim() !== "").map((item, idx) => {
                      return (
                        <div key={item.id || idx} className="flex flex-col gap-1 p-2.5 bg-slate-50/60 hover:bg-slate-50 rounded-xl border border-slate-100 transition-colors">
                          <div className="flex items-center justify-between gap-2">
                            <span className="text-[13px] font-semibold text-slate-600 text-left">Comment: {item.remarks || ""}</span>
                          </div>
                          <div className="p-2 bg-white rounded-lg border border-slate-100 relative shadow-sm">
                            <div className="absolute -top-1 left-3 w-2.5 h-2.5 bg-white border-t border-l border-slate-100 rotate-45" />
                            <p className="text-[11.5px] text-slate-500 leading-normal relative z-10">
                              Remarks Added by {item.user} at Level {item.level || "1"} on {item.date}
                            </p>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                ) : (
                  <div className="text-center py-4 text-slate-400 italic text-[13px]">
                    No comments found.
                  </div>
                )
              ) : (
                queryHistory.length > 0 ? (
                  <div className="space-y-3 max-h-[300px] overflow-y-auto pr-1 animate-fade-in">
                    {queryHistory.map((item, idx) => (
                      <div key={item.id || idx} className="flex flex-col gap-2 p-2.5 bg-slate-50/60 hover:bg-slate-50 rounded-xl border border-slate-100 transition-colors">
                        {/* Query Part */}
                        <div className="flex flex-col gap-1">
                          <div className="flex items-center justify-between gap-2">
                            <span className="text-[13px] font-semibold text-slate-600 text-left">Query: {item.Query || ""}</span>
                          </div>
                          <div className="p-2 bg-white rounded-lg border border-slate-100 relative shadow-sm">
                            <div className="absolute -top-1 left-3 w-2 h-2 bg-white border-t border-l border-slate-100 rotate-45" />
                            <p className="text-[11.5px] text-slate-500 leading-normal relative z-10">
                              Remarks Added by {item.QueryBy ? item.QueryBy.Title : "Unknown"} At Level {item.WFLevel || "1"} on {item.QueriedOn || ""}
                            </p>
                          </div>
                        </div>
                        {/* Response Part */}
                        {item.Response && item.Response.trim() !== "" && (
                          <div className="flex flex-col gap-1 mt-1.5 border-t border-slate-100 pt-1.5">
                            <div className="flex items-center justify-between gap-2">
                              <span className="text-[13px] font-semibold text-slate-600 text-left">Query Response: {item.Response}</span>
                            </div>
                            <div className="p-2 bg-white rounded-lg border border-slate-100 relative shadow-sm">
                              <div className="absolute -top-1 left-3 w-2 h-2 bg-white border-t border-l border-slate-100 rotate-45" />
                              <p className="text-[11.5px] text-slate-500 leading-normal relative z-10">
                                Query Responded by {item.QueryTo ? item.QueryTo.Title : "Unknown"} At Level {item.WFLevel || "1"} on {item.RespondedOn || ""}
                              </p>
                            </div>
                          </div>
                        )}
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="text-center py-4 text-slate-400 italic text-[13px]">
                    No query remarks found.
                  </div>
                )
              )}
            </div>
          </div>
        </div>
      </div>

      {/* Remarks & Action Console (Level 4) */}
      {isCurrentApprover && (
        <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm p-5 space-y-5 mt-6">
          <div className="flex items-center gap-2 mb-2 border-b border-slate-100 pb-3">
            <div className="w-1.5 h-4 bg-blue-600 rounded-full" />
            <span className="font-bold text-slate-800 text-[14px] uppercase tracking-wider">
              Action Console (Level {nfa.level || '4'})
            </span>
          </div>

          <div className="flex flex-col md:flex-row gap-6 items-stretch">

            {/* Remarks TextArea (Left 60%) */}
            <div className="w-full md:w-[60%] flex flex-col">
              <label className="text-[11px] font-bold text-slate-400 uppercase tracking-wider mb-2 block shrink-0">
                Remarks / Comments <span className="text-rose-500">*</span>
              </label>
              <div className="relative flex-1">
                <textarea
                  className="w-full bg-white border border-slate-200 rounded-xl px-4 py-3 text-[13.5px] focus:outline-none focus:ring-2 focus:ring-blue-500/15 focus:border-blue-500 transition-all resize-none h-[110px] leading-relaxed shadow-sm"
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
                <label className="text-[11px] font-bold text-slate-400 uppercase tracking-wider mb-2 block">
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
                      <span className="text-[13px] font-semibold text-slate-700 truncate">{selectedUserObj.label}</span>
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
                          <span className="text-[13px] font-semibold text-slate-700">{u.label}</span>
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

              <span className="text-[11px] text-slate-400 font-semibold mt-2.5 block leading-none pl-0.5">
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
            {/* <button
              type="button"
              onClick={() => handleAction('Delegate')}
              disabled={isActionLoading}
              style={{ backgroundColor: '#0078d4', color: '#ffffff' }}
              className="px-5 h-[38px] hover:opacity-90 rounded-lg font-bold text-xs uppercase tracking-wider transition-all flex items-center gap-2 shadow-sm disabled:opacity-50 active:scale-[0.98]"
            >
              <UserCheck size={14} strokeWidth={2.5} /> Delegate
            </button> */}
            <button
              type="button"
              onClick={handleDownloadPdf}
              disabled={isGeneratingPdf}
              style={{
                backgroundColor: isGeneratingPdf ? "#4A081E" : "#4A081E",
                color: "#FFFFFF",
              }}
              className="px-5 h-[38px] bg-[#4A081E] hover:bg-[#380617] text-white rounded-lg font-bold text-xs uppercase tracking-wider transition-all flex items-center gap-2 shadow-sm disabled:opacity-50 active:scale-[0.98] cursor-pointer"
            >
              <Download size={14} strokeWidth={2.5} />
              <span>{isGeneratingPdf ? 'Generating...' : 'Download PDF'}</span>
            </button>
            <button
              type="button"
              onClick={() => {
                forceScrollToTop();
                navigate('/inbox');
              }}
              style={{ backgroundColor: '#ffffff', color: '#475569', borderColor: '#cbd5e1' }}
              className="px-5 h-[38px] hover:bg-slate-50 rounded-lg font-bold text-xs uppercase tracking-wider border transition-all flex items-center gap-2 shadow-sm active:scale-[0.98]"
            >
              <ArrowLeft size={14} strokeWidth={2.5} /> Cancel
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
