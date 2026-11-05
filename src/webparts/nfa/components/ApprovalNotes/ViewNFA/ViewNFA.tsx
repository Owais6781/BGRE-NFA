import * as React from 'react';
import { useState, useEffect, useRef, useMemo } from 'react';
import {
  Building2, FileText, Wrench, DollarSign, Paperclip, Settings, Star, ChevronDown, ArrowLeft,
  History, MessageSquare, ChevronRight, Download
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
import { CommonService } from '../../../../../Share/Services/CommonService';
import { useNavigate, useParams } from 'react-router-dom';
import "../../../../../Share/styles/global.scss";
import ExcelDataViewer from '../../../../../Share/Components/ExcelDataViewer';
import { forceScrollToTop, stripHtml } from '../../../../../Share/utils';
import { downloadNfaPdf } from '../../../../../Share/utils/nfaPdfGenerator';

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
  const [users, setUsers] = useState<{ label: string, value: string, email?: string }[]>([]);
  const [history, setHistory] = useState<any[]>([]);
  const [comments, setComments] = useState<any[]>([]);
  const [queryHistory, setQueryHistory] = useState<any[]>([]);
  const [isGeneratingPdf, setIsGeneratingPdf] = useState<boolean>(false);

  const handleDownloadPdf = async () => {
    if (!form) return;
    try {
      setIsGeneratingPdf(true);
      await downloadNfaPdf({ form, users, history, comments });
    } catch (err) {
      console.error('Error generating PDF:', err);
      alert('Failed to generate PDF. Please try again.');
    } finally {
      setIsGeneratingPdf(false);
    }
  };

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
          if ((nfaData.status || "").toLowerCase() === 'draft') {
            navigate(`/new-nfa/${id}`, { replace: true });
            return;
          }
          const commonService = new CommonService(context, sp);
          const [docs, siteUsers, historyData, qHistory] = await Promise.all([
            docService.getDocumentsByUniqId(nfaData.uniqId),
            sp.web.siteUsers.filter("PrincipalType eq 1").select("Id", "Title", "Email")()
              .then(res => res.map(u => ({ label: u.Title, value: String(u.Id), email: u.Email }))),
            nfaService.getHistory(parseInt(id)),
            commonService.getQueryHistory(id)
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
          setQueryHistory(qHistory);

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

      {/* Top Action Bar */}
      {/* <div className="flex items-center justify-between gap-4 bg-white rounded-xl border border-slate-200/80 shadow-sm py-2.5 px-4 mb-3.5">
        <div className="flex items-center gap-2.5">
          <span className="text-xs font-extrabold text-[#4A081E] tracking-tight">
            {form.requestNo || 'Note for Approval'}
          </span>
          <span className="text-[10px] font-extrabold uppercase px-2 py-0.5 rounded-full bg-slate-100 text-slate-600 border border-slate-200">
            Level {form.level || '1'}
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
              navigate(-1);
            }}
            className="px-3.5 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg font-bold text-xs transition-all border border-slate-300 active:scale-95 cursor-pointer"
          >
            Back
          </button>
        </div>
      </div> */}

      {/* Approval Chain Card */}
      <div className="bg-white rounded-xl border border-slate-200/80 shadow-sm py-2 px-4 mb-3.5">
        <div
          ref={stepperContainerRef}
          className="relative flex items-center justify-start gap-1 overflow-x-auto pb-1 pt-0.5 no-scrollbar"
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
                  <span className={`text-[9px] font-extrabold uppercase tracking-wider mb-0.5 ${labelColor}`}>
                    {step.label}
                  </span>

                  <div className={`w-[140px] h-[38px] rounded-lg border flex flex-row items-center gap-1.5 px-2 py-0.5 transition-all duration-350 ${borderClass}`}>
                    <UserAvatar
                      email={user?.email}
                      initials={initials}
                      className={`w-6 h-6 rounded-full shadow-inner shrink-0 ${avatarClass}`}
                    />
                    <span className="text-[10px] font-extrabold text-slate-700 text-left leading-tight line-clamp-2 flex-1">
                      {userLabel}
                    </span>
                  </div>

                  <div className={`w-1.5 h-1.5 rounded-full mt-1 ${dotClass}`} />
                </div>

                {idx < workflowSteps.length - 1 && (
                  <div className="flex items-center justify-center text-slate-400 shrink-0 mt-3">
                    <ChevronRight size={14} strokeWidth={2.5} />
                  </div>
                )}
              </React.Fragment>
            );
          })}
        </div>
      </div>

      {/* ── MAIN DATA (100% FULL WIDTH - COMPACT VERTICAL SPACING) ── */}
      <div className="w-full space-y-3.5">

        {/* Basic Information */}
        <div className="bg-white rounded-xl border border-slate-200/80 shadow-sm p-4">
          <div className="flex items-center gap-2 mb-2.5 border-b border-slate-100 pb-2">
            <div className="w-1.5 h-3.5 bg-blue-600 rounded-full" />
            <span className="font-bold text-slate-800 text-[11px] uppercase tracking-wider">
              Request No: {form.requestNo || '-'} (Level {form.level || '1'})
            </span>
          </div>

          <div className="flex flex-col gap-2">
            {/* Row 1: Project | Sub-Entity | Date */}
            <div className="flex flex-col md:flex-row gap-3">
              <div className="w-full md:w-[35%] flex flex-col gap-0.5">
                <span className="text-[9px] font-semibold text-slate-400 uppercase tracking-wider">Entity</span>
                <span className="text-xs font-medium text-slate-700 bg-slate-50/60 p-1.5 px-2.5 rounded-lg border border-slate-100">{form.project || '-'}</span>
              </div>
              <div className="w-full md:w-[35%] flex flex-col gap-0.5">
                <span className="text-[9px] font-semibold text-slate-400 uppercase tracking-wider">Sub-Entity</span>
                <span className="text-xs font-medium text-slate-700 bg-slate-50/60 p-1.5 px-2.5 rounded-lg border border-slate-100">
                  {form.projectSubEntity && form.projectSubEntity.length > 0 ? form.projectSubEntity.join(', ') : '-'}
                </span>
              </div>
              <div className="w-full md:w-[30%] flex flex-col gap-0.5">
                <span className="text-[9px] font-semibold text-slate-400 uppercase tracking-wider">Date</span>
                <span className="text-xs font-medium text-slate-700 bg-slate-50/60 p-1.5 px-2.5 rounded-lg border border-slate-100">{form.date || '-'}</span>
              </div>
            </div>

            {/* Row 2: Function | DOA Matrix | Workflow */}
            <div className="flex flex-col md:flex-row gap-3">
              <div className="w-full md:w-[35%] flex flex-col gap-0.5">
                <span className="text-[9px] font-semibold text-slate-400 uppercase tracking-wider">Function</span>
                <span className="text-xs font-medium text-slate-700 bg-slate-50/60 p-1.5 px-2.5 rounded-lg border border-slate-100">{form.originatorDept || '-'}</span>
              </div>
              <div className="w-full md:w-[35%] flex flex-col gap-0.5">
                <span className="text-[9px] font-semibold text-slate-400 uppercase tracking-wider">DOA Matrix</span>
                <span className="text-xs font-medium text-slate-700 bg-slate-50/60 p-1.5 px-2.5 rounded-lg border border-slate-100">{form.doaMatrix || '-'}</span>
              </div>
              <div className="w-full md:w-[30%] flex flex-col gap-0.5">
                <span className="text-[9px] font-semibold text-slate-400 uppercase tracking-wider">Workflow</span>
                <span className="text-xs font-medium text-slate-700 bg-slate-50/60 p-1.5 px-2.5 rounded-lg border border-slate-100 uppercase">{form.workflow || '-'}</span>
              </div>
            </div>
          </div>
        </div>

        {/* NFA Details (Main Data & Excel Viewer) */}
        <div className="bg-white rounded-xl border border-slate-200/80 shadow-sm p-4">
          <div className="flex items-center gap-2 mb-2.5 border-b border-slate-100 pb-2">
            <div className="w-1.5 h-3.5 bg-blue-600 rounded-full" />
            <span className="font-bold text-slate-800 text-[11px] uppercase tracking-wider">NFA Details</span>
          </div>

          <div className="flex flex-col gap-3">
            <div className="flex flex-col gap-1">
              <span className="text-[9.5px] font-bold text-slate-400 uppercase tracking-wider">Refrence / Subject</span>
              <span className="text-xs font-semibold text-slate-700 bg-slate-50/60 p-2.5 rounded-lg border border-slate-100 whitespace-pre-wrap">{stripHtml(form.referenceSubject) || '-'}</span>
            </div>

            <div className="flex flex-col gap-1">
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Introduction / Preamble / Background</span>
              <ExcelDataViewer html={form.introduction} label="Introduction" />
            </div>


            <div className="flex flex-nowrap w-full  gap-3 items-start">
              <div className="flex flex-col gap-1 flex-1 min-w-0 ">
                <span className="text-[9.5px] font-bold text-slate-400 uppercase tracking-wider">
                  Complete Indent/PR received from & Date
                </span>
                <span className="text-xs font-semibold text-slate-700 bg-slate-50/60 p-2.5 rounded-lg border border-slate-100 whitespace-pre-wrap">
                  {form.completeIndentPRReceivedFromDate
                    ? (() => {
                      const [year, month, day] =
                        form.completeIndentPRReceivedFromDate.split("T")[0].split("-");
                      return `${day.slice(-1)}/${month}/${year}`;
                    })()
                    : "-"}
                </span>
              </div>

              {/* BOQ / Specifications Revisions */}
              <div className="flex flex-col gap-1 flex-1 min-w-0">
                <span className="text-[9.5px] font-bold text-slate-400 uppercase tracking-wider">
                  BOQ/Specifications Revisions by & Date
                </span>
                <span className="text-xs font-semibold text-slate-700 bg-slate-50/60 p-2.5 rounded-lg border border-slate-100 whitespace-pre-wrap">
                  {form.boqSpecificationsRevisionsByDate
                    ? (() => {
                      const [year, month, day] =
                        form.boqSpecificationsRevisionsByDate.split("T")[0].split("-");
                      return `${day.slice(-1)}/${month}/${year}`;
                    })()
                    : "-"}
                </span>
              </div>

              {/* RFQ Reference Number */}
              <div className="flex flex-col gap-1 flex-1 min-w-0">
                <span className="text-[9.5px] font-bold text-slate-400 uppercase tracking-wider">
                  RFQ Ref No
                </span>
                <span className="text-xs font-semibold text-slate-700 bg-slate-50/60 p-2.5 rounded-lg border border-slate-100 whitespace-pre-wrap">
                  {form.RFQRefNo || "-"}
                </span>
              </div>

            </div>
            <div className="flex flex-col gap-1">
              <span className="text-[9.5px] font-bold text-slate-400 uppercase tracking-wider">Remarks</span>
              <span className="text-xs font-semibold text-slate-700 bg-slate-50/60 p-2.5 rounded-lg border border-slate-100 whitespace-pre-wrap">{stripHtml(form.remark) || '-'}</span>
            </div>


            <div className="flex flex-col gap-1">
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Chronology</span>
              <ExcelDataViewer html={form.chronology} label="Chronology" />
            </div>

            <div className="flex flex-col gap-1">
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Commercial Proposals</span>
              <ExcelDataViewer html={form.commercialProposal} label="Commercial Proposals" />
            </div>

            <div className="flex flex-col gap-1">
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Payment Terms</span>
              <ExcelDataViewer html={form.paymentTerms} label="Payment Terms" />
            </div>
          </div>
        </div>

        {/* Financial Details */}
        <div className="bg-white rounded-xl border border-slate-200/80 shadow-sm p-4">
          <div className="flex items-center gap-2 mb-2.5 border-b border-slate-100 pb-2">
            <div className="w-1.5 h-3.5 bg-blue-600 rounded-full" />
            <span className="font-bold text-slate-800 text-[11px] uppercase tracking-wider">Financial Details</span>
          </div>

          <div className="space-y-3">
            <div className="flex flex-col md:flex-row gap-3">
              <div className="flex-1 flex flex-col gap-0.5 p-2.5 bg-emerald-50/30 rounded-lg border border-emerald-100/50">
                <span className="text-[9px] font-extrabold text-emerald-600 uppercase tracking-wider">Packeage Awarded Cost</span>
                <span className="text-sm font-black text-emerald-700">₹ {Number(form.packageAwardedCost || 0).toLocaleString('en-IN')}</span>
              </div>
              <div className="flex-1 flex flex-col gap-0.5 p-2.5 bg-blue-50/30 rounded-lg border border-blue-100/50">
                <span className="text-[9px] font-extrabold text-blue-600 uppercase tracking-wider">Budget</span>
                <span className="text-sm font-black text-blue-700">₹ {Number(form.budget || 0).toLocaleString('en-IN')}</span>
              </div>
              <div className="flex-1 flex flex-col gap-0.5 p-2.5 bg-purple-50/30 rounded-lg border border-purple-100/50">
                <span className="text-[9px] font-extrabold text-purple-600 uppercase tracking-wider">Budgeted/ Non Budgeted</span>
                <span className="text-xs font-extrabold text-purple-700 truncate">{form.budgetedNonBudgeted || '-'}</span>
              </div>
            </div>

            <div className="flex flex-col gap-0.5">
              <span className="text-[9.5px] font-bold text-slate-400 uppercase tracking-wider">Remarks</span>
              <div className="text-xs text-slate-700 bg-slate-50/60 p-2 rounded-lg border border-slate-100 whitespace-pre-wrap">{form.budgetRemarks || '-'}</div>
            </div>
          </div>
        </div>

        {/* Recommendation */}
        <div className="bg-white rounded-xl border border-slate-200/80 shadow-sm p-4">
          <div className="flex items-center gap-2 mb-2.5 border-b border-slate-100 pb-2">
            <div className="w-1.5 h-3.5 bg-blue-600 rounded-full" />
            <span className="font-bold text-slate-800 text-[11px] uppercase tracking-wider">Recommendation</span>
          </div>

          <div className="flex flex-col gap-2.5">
            <div className="flex flex-col gap-1">
              <span className="text-[9.5px] font-bold text-slate-400 uppercase tracking-wider">Recommendation with Justification</span>
              <ExcelDataViewer html={form.recommendation} label="Justification" />
            </div>

            <div className="flex flex-col gap-0.5">
              <span className="text-[9.5px] font-bold text-slate-400 uppercase tracking-wider">Contractor Name</span>
              <span className="text-xs font-bold text-slate-800 bg-blue-50/30 p-2.5 rounded-lg border border-blue-100/50">{form.contractorName || '-'}</span>
            </div>
          </div>
        </div>

        {/* Contract Terms (Compact 4-Column Row) */}
        <div className="bg-white rounded-xl border border-slate-200/80 shadow-sm p-4">
          <div className="flex items-center gap-2 mb-2.5 border-b border-slate-100 pb-2">
            <div className="w-1.5 h-3.5 bg-blue-600 rounded-full" />
            <span className="font-bold text-slate-800 text-[11px] uppercase tracking-wider">Contract Terms</span>
          </div>

          {/* <div className="grid grid-cols-2 md:grid-cols-4 gap-2.5"> */}
          <div className="grid grid-cols-4 gap-2.5">
            <div className="flex flex-col gap-0.5">
              <span className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">Liquidated Damage</span>
              <div className="text-xs text-slate-700 bg-slate-50/60 p-2 rounded-lg border border-slate-100 whitespace-pre-wrap">{form.liquidatedDamage || '-'}</div>
            </div>
            <div className="flex flex-col gap-0.5">
              <span className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">Contract Performance Gurantee</span>
              <div className="text-xs text-slate-700 bg-slate-50/60 p-2 rounded-lg border border-slate-100 whitespace-pre-wrap">{form.contractPerf || '-'}</div>
            </div>
            <div className="flex flex-col gap-0.5">
              <span className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">Defect Liability Period</span>
              <div className="text-xs text-slate-700 bg-slate-50/60 p-2 rounded-lg border border-slate-100 whitespace-pre-wrap">{form.defectLiability || '-'}</div>
            </div>
            <div className="flex flex-col gap-0.5">
              <span className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">Work Completion</span>
              <div className="text-xs text-slate-700 bg-slate-50/60 p-2 rounded-lg border border-slate-100 whitespace-pre-wrap">{form.workCompletion || '-'}</div>
            </div>
          </div>
        </div>

        {/* Annexures & Notes (Compact 2-Column Row for Notes) */}
        <div className="bg-white rounded-xl border border-slate-200/80 shadow-sm p-4">
          <div className="flex items-center gap-2 mb-2.5 border-b border-slate-100 pb-2">
            <div className="w-1.5 h-3.5 bg-blue-600 rounded-full" />
            <span className="font-bold text-slate-800 text-[11px] uppercase tracking-wider">Annexures & Notes</span>
          </div>

          <div className="flex flex-col gap-2.5">
            <div className="flex flex-col gap-0.5">
              <span className="text-[9.5px] font-bold text-slate-400 uppercase tracking-wider">Annexure/ Enclosur(if any)</span>
              <div className="text-xs text-slate-700 bg-slate-50/60 p-2 rounded-lg border border-slate-100 whitespace-pre-wrap">{form.annexure || '-'}</div>
            </div>

            {/* Note 1 & Note 2 in 2 columns */}
            <div className="flex flex-col md:flex-row gap-2.5">
              <div className="flex-1 flex flex-col gap-0.5">
                <span className="text-[9.5px] font-bold text-slate-400 uppercase tracking-wider">Note 1</span>
                <div className="text-xs text-slate-700 bg-slate-50/60 p-2 rounded-lg border border-slate-100 whitespace-pre-wrap">{form.note1 || '-'}</div>
              </div>
              <div className="flex-1 flex flex-col gap-0.5">
                <span className="text-[9.5px] font-bold text-slate-400 uppercase tracking-wider">Note 2</span>
                <div className="text-xs text-slate-700 bg-slate-50/60 p-2 rounded-lg border border-slate-100 whitespace-pre-wrap">{form.note2 || '-'}</div>
              </div>
            </div>

            {/* Note 3 & Note 4 in 2 columns */}
            <div className="flex flex-col md:flex-row gap-2.5">
              <div className="flex-1 flex flex-col gap-0.5">
                <span className="text-[9.5px] font-bold text-slate-400 uppercase tracking-wider">Note 3</span>
                <div className="text-xs text-slate-700 bg-slate-50/60 p-2 rounded-lg border border-slate-100 whitespace-pre-wrap">{form.note3 || '-'}</div>
              </div>
              <div className="flex-1 flex flex-col gap-0.5">
                <span className="text-[9.5px] font-bold text-slate-400 uppercase tracking-wider">Note 4</span>
                <div className="text-xs text-slate-700 bg-slate-50/60 p-2 rounded-lg border border-slate-100 whitespace-pre-wrap">{form.note4 || '-'}</div>
              </div>
            </div>
          </div>
        </div>

      </div>

      {/* ── SECONDARY INFORMATION & AUDIT HISTORY (EXACT SAME ROW) ── */}
      <div className="mt-5 pt-3.5 border-t border-slate-200/80">
        <div className="flex items-center gap-2 mb-3">
          <div className="w-1.5 h-3.5 bg-emerald-600 rounded-full" />
          <span className="font-bold text-slate-800 text-xs uppercase tracking-wider">
            Supporting Information & Audit History
          </span>
        </div>

        <div
          className="w-full"
          style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))',
            gap: '16px',
            alignItems: 'start'
          }}
        >

          {/* Left Column: Documents (Attachments & Query Documents) */}
          <div className="flex flex-col gap-3.5 w-full">

            {/* Attachments */}
            <div className="bg-white rounded-xl border border-slate-200/80 shadow-sm p-4">
              <div className="flex items-center justify-between mb-2.5 border-b border-slate-100 pb-2">
                <div className="flex items-center gap-2">
                  <Paperclip size={14} className="text-emerald-500 shrink-0" />
                  <span className="font-bold text-slate-700 text-[11px] uppercase tracking-wider">Attachments</span>
                </div>
                <span className="bg-emerald-50 text-emerald-600 text-[9px] font-extrabold px-2 py-0.5 rounded-full border border-emerald-100">
                  {savedDocuments.length}
                </span>
              </div>
              {savedDocuments.length > 0 ? (
                <div className="space-y-1.5 max-h-[190px] overflow-y-auto pr-1">
                  {savedDocuments.map(doc => {
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
                        <div className="flex items-center gap-2 min-w-0 overflow-hidden">
                          <div className={`p-1 rounded border transition-colors shrink-0 ${iconColor}`}>
                            <FileText size={12} />
                          </div>
                          <span className="text-[10.5px] font-bold text-slate-700 truncate group-hover:text-emerald-600 transition-colors" title={doc.fileName}>
                            {doc.fileName}
                          </span>
                        </div>
                        <span className="shrink-0 text-[9.5px] font-bold px-2 py-0.5 rounded-md bg-slate-100 text-slate-700 border border-slate-200 group-hover:bg-emerald-50 group-hover:text-emerald-700 group-hover:border-emerald-200 transition-colors">
                          {doc.documentType || 'Attachment'}
                        </span>

                      </a>
                    );
                  })}
                </div>
              ) : (
                <div className="flex flex-col items-center justify-center py-4 border border-dashed border-slate-200 rounded-lg">
                  <Paperclip size={18} className="text-slate-300 mb-1" />
                  <span className="text-[10px] text-slate-400 font-semibold italic">No attachments found</span>
                </div>
              )}
            </div>

            {/* Query Documents (Show only if available) */}
            {queryDocuments.length > 0 && (
              <div className="bg-white rounded-xl border border-slate-200/80 shadow-sm p-4">
                <div className="flex items-center justify-between mb-2.5 border-b border-slate-100 pb-2">
                  <div className="flex items-center gap-2">
                    <Paperclip size={14} className="text-amber-500 shrink-0" />
                    <span className="font-bold text-slate-700 text-[11px] uppercase tracking-wider">Query Documents</span>
                  </div>
                  <span className="bg-amber-50 text-amber-600 text-[9px] font-extrabold px-2 py-0.5 rounded-full border border-amber-100">
                    {queryDocuments.length}
                  </span>
                </div>
                <div className="space-y-1.5 max-h-[190px] overflow-y-auto pr-1">
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
                        data-interception="off"
                        className="group flex items-center gap-2 p-1.5 px-2 bg-slate-50/50 hover:bg-amber-50/15 rounded-lg border border-slate-100 hover:border-amber-150 transition-all min-w-0"
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

          </div>

          {/* Right Column: Workflow History & Comments */}
          <div className="flex flex-col gap-3.5 w-full">

            {/* Workflow History */}
            <div className="bg-white rounded-xl border border-slate-200/80 shadow-sm p-4">
              <div className="flex items-center gap-2 mb-2.5 border-b border-slate-100 pb-2">
                <History size={15} className="text-blue-500 shrink-0" />
                <span className="font-bold text-slate-700 text-[11px] uppercase tracking-wider">Workflow History</span>
              </div>

              <div className="space-y-2 max-h-[190px] overflow-y-auto pr-1">
                {history.length > 0 ? history.map((item, idx) => {
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
                    <div key={item.id || idx} className="relative pl-6 pb-0.5 last:pb-0">
                      {idx < history.length - 1 && (
                        <div className="absolute left-[6px] top-4 bottom-0 w-[1.5px] bg-slate-100" />
                      )}
                      <div className={`absolute left-0 top-1 w-3 h-3 rounded-full border border-white flex items-center justify-center z-10 ${dotClass} shadow-sm`} />
                      <div className="flex flex-col min-w-0">
                        <p className="text-[10.5px] text-slate-600 font-medium leading-relaxed">
                          {item.rawText || `${item.user} (${item.action})${item.remarks ? `: ${item.remarks}` : ""}`}
                        </p>
                      </div>
                    </div>
                  );
                }) : (
                  <div className="text-center py-3 text-slate-400 italic text-[10.5px]">
                    No audit history found.
                  </div>
                )}
              </div>
            </div>

            {/* Comments History */}
            {comments.filter(item => item.remarks && item.remarks !== "No remarks" && item.remarks.trim() !== "").length > 0 && (
              <div className="bg-white rounded-xl border border-slate-200/80 shadow-sm p-4">
                <div className="flex items-center gap-2 mb-2.5 border-b border-slate-100 pb-2">
                  <MessageSquare size={14} className="text-indigo-500 shrink-0" />
                  <span className="font-bold text-slate-700 text-[11px] uppercase tracking-wider">Comments History</span>
                </div>

                <div className="space-y-2 max-h-[190px] overflow-y-auto pr-1">
                  {comments.filter(item => item.remarks && item.remarks !== "No remarks" && item.remarks.trim() !== "").map((item, idx) => (
                    <div key={item.id || idx} className="flex flex-col gap-1 p-2.5 bg-slate-50/60 hover:bg-slate-50 rounded-lg border border-slate-100 transition-colors">
                      <div className="flex items-center justify-between gap-2">
                        <span className="text-[10px] font-medium text-slate-700 text-left">Comment: {item.remarks || ""}</span>
                      </div>
                      <div className="p-2 bg-white rounded-md border border-slate-100 relative shadow-sm">
                        <div className="absolute -top-1 left-3 w-2 h-2 bg-white border-t border-l border-slate-100 rotate-45" />
                        <p className="text-[9.5px] text-slate-500 leading-normal relative z-10">
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
              <div className="bg-white rounded-xl border border-slate-200/80 shadow-sm p-4">
                <div className="flex items-center gap-2 mb-2.5 border-b border-slate-100 pb-2">
                  <MessageSquare size={14} className="text-amber-500 shrink-0" />
                  <span className="font-bold text-slate-700 text-[11px] uppercase tracking-wider">Query Remarks History</span>
                </div>

                <div className="space-y-2 max-h-[190px] overflow-y-auto pr-1">
                  {queryHistory.map((item, idx) => (
                    <div key={item.id || idx} className="flex flex-col gap-1.5 p-2.5 bg-slate-50/60 hover:bg-slate-50 rounded-lg border border-slate-100 transition-colors">
                      {/* The Query Part */}
                      <div className="flex flex-col gap-0.5">
                        <div className="flex items-center justify-between gap-2">
                          <span className="text-[10px] font-medium text-slate-700 text-left">Query: {item.Query || ""}</span>
                        </div>
                        <div className="p-2 bg-white rounded-md border border-slate-100 relative shadow-sm">
                          <div className="absolute -top-1 left-3 w-2 h-2 bg-white border-t border-l border-slate-100 rotate-45" />
                          <p className="text-[9.5px] text-slate-500 leading-normal relative z-10">
                            Remarks Added by {item.QueryBy ? item.QueryBy.Title : "Unknown"} At Level {item.WFLevel || "1"} on {item.QueriedOn || ""}
                          </p>
                        </div>
                      </div>

                      {/* The Response Part */}
                      {item.Response && item.Response.trim() !== "" && (
                        <div className="flex flex-col gap-0.5 mt-1 border-t border-slate-100 pt-1">
                          <div className="flex items-center justify-between gap-2">
                            <span className="text-[10px] font-medium text-slate-700 text-left">Query Response: {item.Response}</span>
                          </div>
                          <div className="p-2 bg-white rounded-md border border-slate-100 relative shadow-sm">
                            <div className="absolute -top-1 left-3 w-2 h-2 bg-white border-t border-l border-slate-100 rotate-45" />
                            <p className="text-[9.5px] text-slate-500 leading-normal relative z-10">
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
      </div>

      {/* Bottom Action Buttons */}
      <div className="flex justify-center items-center gap-4 py-4">

        <button
          type="button"
          onClick={handleDownloadPdf}
          disabled={isGeneratingPdf}
          style={{
            backgroundColor: isGeneratingPdf ? "#4A081E" : "#4A081E",
            color: "#FFFFFF",
          }}
          className="flex items-center gap-2 px-6 h-[38px] hover:!bg-[#380617] rounded-lg font-extrabold text-xs uppercase tracking-wider transition-all shadow-md active:scale-98 disabled:opacity-50 cursor-pointer"
        >
          <Download size={15} />
          <span>
            {isGeneratingPdf ? "Generating PDF..." : "Download PDF"}
          </span>
        </button>
        <button
          onClick={() => {
            forceScrollToTop();
            navigate(-1);
          }}
          className="px-6 h-[38px] bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg font-extrabold text-xs uppercase tracking-wider transition-all border border-slate-300 active:scale-98 shadow-sm"
        >
          Back
        </button>
      </div>

    </div>
  );
};

export default ViewNFA;
