import * as React from 'react';
import { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { ChevronUp, ChevronDown } from 'lucide-react';
import ApprovalDetailsFilter from './ApprovalDetailsFilter';
import ApprovalDetailsTable  from './ApprovalDetailsTable';
import "../../../../../Share/styles/global.scss";
import PageTitle from '../../../../../Share/Components/PageTitle';
import { forceScrollToTop, stripHtml } from '../../../../../Share/utils';

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { DropdownService, IDropdownOption } from '../../../../../Share/Services/DropdownService';

interface ApprovalRecord {
  id: number;
  nfa: string;
  subject: string;
  function: string;
  createdBy: string;
  createdOn: string;
  status: string;
  approvedOn: string;
  pendingAt: string;
  pendingFrom: string;
  project: string;
  packageCost: string;
}

interface IApprovalDetailsProps {
  context: WebPartContext;
  sp: SPFI;
}

const ApprovalDetailsView: React.FC<IApprovalDetailsProps> = ({ context, sp }) => {
  const navigate = useNavigate();
  const [tableData,   setTableData]   = useState<ApprovalRecord[]>([]);
  const [filters,     setFilters]     = useState<any>(null);
  const [isCollapsed, setIsCollapsed] = useState(false);
  const [projectOptions, setProjectOptions] = useState<IDropdownOption[]>([]);
  const [functionOptions, setFunctionOptions] = useState<IDropdownOption[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const loadData = async () => {
      try {
        setLoading(true);
        const dropdownService = new DropdownService(context);

        // 1. Get current logged-in user
        const currentUser = await sp.web.currentUser();
        const currentUserId = currentUser.Id;
        const currentUserName = (currentUser.Title || "").trim().toLowerCase();

        // 2. Fetch master dropdowns with fallback
        let projects: IDropdownOption[] = [];
        let funcs: IDropdownOption[] = [];
        try {
          projects = await dropdownService.getProjects();
        } catch (pErr) {
          console.warn("ApprovalDetailsView: Failed to load projects from ProjectMaster:", pErr);
        }
        if (!projects || projects.length === 0) {
          try {
            const pItems = await sp.web.lists.getByTitle("Project Master").items
              .select("Id", "ProjectName", "Title")
              .top(500)();
            projects = pItems.map((p: any) => ({
              label: p.ProjectName || p.Title,
              value: String(p.Id || p.ProjectName || p.Title)
            }));
          } catch { }
        }

        try {
          funcs = await dropdownService.getFunctionMaster();
        } catch (fErr) {
          console.warn("ApprovalDetailsView: Failed to load functions from FunctionMaster:", fErr);
        }

        // 3. Query WFLogList for note IDs where user acted or was assigned as approver
        const userWfLogIds = new Set<string>();
        try {
          const logItems = await sp.web.lists.getByTitle("WFLogList").items
            .filter(`AuthorId eq ${currentUserId} or PPLPendingAtId eq ${currentUserId}`)
            .select("WFID")
            .top(5000)();
          logItems.forEach((l: any) => {
            if (l.WFID) userWfLogIds.add(String(l.WFID));
          });
        } catch (logErr) {
          console.warn("ApprovalDetailsView: Failed to query WFLogList:", logErr);
        }

        // 4. Fetch dashboard items
        let dashboardItems: any[] = [];
        const dashMap: Record<string, any> = {};
        try {
          dashboardItems = await sp.web.lists.getByTitle("DashboardList").items
            .select("WFID", "PendingFrom", "CurrentApproverId", "CurrentApprovalID", "CurrentApprovalName", "Status")
            .top(5000)();
          dashboardItems.forEach((d: any) => {
            if (d.WFID) {
              dashMap[d.WFID] = d;
            }
          });
        } catch (dashErr) {
          console.warn("ApprovalDetailsView: Error fetching DashboardList:", dashErr);
        }

        // 5. Fetch approval notes from DigitalApprovalNotes including all 12 approver levels
        let nfaItems: any[] = [];
        try {
          const levels = Array.from({ length: 12 }, (_, i) => `ApprovalLevel${i + 1}Id`);
          nfaItems = await sp.web.lists.getByTitle("DigitalApprovalNotes").items
            .select(
              "Id", "Title", "SubjectMatter", "Dept", "FunctionName", "Project",
              "RequestedOn", "Status", "ApprovalAmount", "Author/Title", "Author/Id",
              "PendingAt", "Modified", "Created", "PendingFrom", "WFLevel",
              ...levels
            )
            .expand("Author")
            .top(5000)();
        } catch (nfaErr) {
          console.error("ApprovalDetailsView: Error fetching DigitalApprovalNotes:", nfaErr);
        }

        // 6. Filter notes assigned to the currently logged-in user
        const userNfaItems = nfaItems.filter((item: any) => {
          // Check if user is configured in any ApprovalLevel1..12
          for (let i = 1; i <= 12; i++) {
            const lvlVal = item[`ApprovalLevel${i}Id`];
            if (lvlVal && (lvlVal === currentUserId || String(lvlVal) === String(currentUserId))) {
              return true;
            }
          }

          // Check if PendingAt text matches current user's name
          const pendingAt = (typeof item.PendingAt === 'string' ? item.PendingAt : (item.PendingAt?.Title || "")).trim().toLowerCase();
          if (pendingAt && (pendingAt === currentUserName || pendingAt.includes(currentUserName))) {
            return true;
          }

          // Check DashboardList for active assignment
          const dash = dashMap[item.Id.toString()];
          if (dash) {
            if (dash.CurrentApproverId === currentUserId || String(dash.CurrentApprovalID) === String(currentUserId)) {
              return true;
            }
            const dashApprover = (dash.CurrentApprovalName || "").trim().toLowerCase();
            if (dashApprover && (dashApprover === currentUserName || dashApprover.includes(currentUserName))) {
              return true;
            }
          }

          // Check if recorded in WFLogList for this user
          if (userWfLogIds.has(String(item.Id))) {
            return true;
          }

          return false;
        });

        // 7. Fallback dropdown options from user-assigned notes if master lists were empty
        if ((!projects || projects.length === 0) && userNfaItems.length > 0) {
          const uniqueProjects = Array.from(new Set(userNfaItems.map((it: any) => it.Project).filter(Boolean)));
          projects = uniqueProjects.map((p: any) => ({ label: String(p), value: String(p) }));
        }

        if ((!funcs || funcs.length === 0) && userNfaItems.length > 0) {
          const uniqueDepts = Array.from(new Set(userNfaItems.map((it: any) => it.Dept || it.FunctionName).filter(Boolean)));
          funcs = uniqueDepts.map((d: any) => ({ label: String(d), value: String(d) }));
        }

        setProjectOptions(projects || []);
        setFunctionOptions(funcs || []);

        // 8. Map to ApprovalRecord
        const mapped: ApprovalRecord[] = userNfaItems.map((item: any) => {
          const dash = dashMap[item.Id.toString()];
          let approvedOnVal = "";
          if (item.Status === "Approved" || item.Status === "Completed" || item.Status === "Rejected") {
            approvedOnVal = item.Modified || "";
          }

          let pendingAtVal = "";
          if (typeof item.PendingAt === 'object' && item.PendingAt) {
            pendingAtVal = item.PendingAt.Title || "";
          } else if (typeof item.PendingAt === 'string') {
            pendingAtVal = item.PendingAt;
          }
          if (!pendingAtVal && dash && dash.CurrentApprovalName) {
            pendingAtVal = dash.CurrentApprovalName;
          }

          const pendingFromVal = (dash && dash.PendingFrom) ? dash.PendingFrom : (item.PendingFrom || "");

          return {
            id: item.Id,
            nfa: item.Title || "",
            subject: stripHtml(item.SubjectMatter || item.Title || ""),
            function: item.Dept || item.FunctionName || "",
            createdBy: item.Author ? item.Author.Title : "",
            createdOn: item.RequestedOn || item.Created || "",
            status: item.Status || "",
            approvedOn: approvedOnVal,
            pendingAt: pendingAtVal,
            pendingFrom: pendingFromVal,
            project: item.Project || "",
            packageCost: item.ApprovalAmount || "0.00"
          };
        });

        setTableData(mapped);
      } catch (err) {
        console.error("ApprovalDetailsView: Error loading page data:", err);
      } finally {
        setLoading(false);
      }
    };
    loadData();
  }, [context]);

  useEffect(() => {
    forceScrollToTop();
  }, [loading]);

  // --- Filtering Logic ---
  const displayData = React.useMemo(() => {
    if (!filters) return tableData;

    return tableData.filter(item => {
      // Date Filter
      if (filters.fromDate || filters.toDate) {
        const parseDate = (dStr: string) => {
          if (!dStr) return null;
          if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(dStr)) {
            const parts = dStr.split(/[/\s]/);
            return new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
          }
          const d = new Date(dStr);
          return isNaN(d.getTime()) ? null : d;
        };

        const itemDate = parseDate(item.createdOn);
        if (itemDate) {
          itemDate.setHours(0, 0, 0, 0);

          if (filters.fromDate) {
            const from = new Date(filters.fromDate);
            from.setHours(0, 0, 0, 0);
            if (itemDate < from) return false;
          }
          if (filters.toDate) {
            const to = new Date(filters.toDate);
            to.setHours(0, 0, 0, 0);
            if (itemDate > to) return false;
          }
        }
      }

      // Dropdown Filters
      if (filters.status && item.status !== filters.status) return false;

      if (filters.function) {
        const funcObj = functionOptions.find(f => f.value === filters.function || f.label === filters.function);
        const funcName = funcObj ? funcObj.label : filters.function;
        if (item.function !== funcName) return false;
      }

      if (filters.project) {
        const projObj = projectOptions.find(p => p.value === filters.project || p.label === filters.project);
        const projName = projObj ? projObj.label : filters.project;
        if (item.project !== projName) return false;
      }

      return true;
    });
  }, [tableData, filters, projectOptions, functionOptions]);

  const handleSearch = (f: any) => {
    setFilters(f);
  };

  return (
    <div className="pageContainer">

      <PageTitle 
        title="Detailed View for Approver" 
        subtitle="Comprehensive list and details of notes assigned to you." 
        action={
          <button className="iconBtn" onClick={() => setIsCollapsed(c => !c)} title={isCollapsed ? 'Expand' : 'Collapse'}>
            {isCollapsed ? <ChevronDown size={16} /> : <ChevronUp size={16} />}
          </button>
        }
      />

      {!isCollapsed && (
        <>
          <ApprovalDetailsFilter 
            onSearch={handleSearch} 
            onClear={() => setFilters(null)} 
            projectOptions={projectOptions}
            functionOptions={functionOptions}
          />
          {loading ? (
            <div className="flex justify-center items-center p-8 bg-white rounded-2xl border border-slate-200 mt-3">
              <span className="text-slate-500 font-medium">Loading approval details...</span>
            </div>
          ) : (
            <ApprovalDetailsTable data={displayData} onRecordClick={(id) => {
              const rec = displayData.find(d => d.id === id);
              if ((rec?.status || "").toLowerCase() === 'draft') {
                navigate(`/new-nfa/${id}`);
              } else {
                navigate(`/view-nfa/${id}`);
              }
            }} />
          )}
        </>
      )}

    </div>
  );
};

export default ApprovalDetailsView;
