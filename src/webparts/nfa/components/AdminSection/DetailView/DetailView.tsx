import * as React from 'react';
import { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { ChevronUp, ChevronDown } from 'lucide-react';
import DetailViewFilter from './DetailViewFilter';
import DetailViewTable, { DetailViewRecord }  from './DetailViewTable';
import PageTitle from '../../../../../Share/Components/PageTitle';
import "../../../../../Share/styles/global.scss";
import { forceScrollToTop, stripHtml } from '../../../../../Share/utils';

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { DropdownService, IDropdownOption } from '../../../../../Share/Services/DropdownService';

interface IDetailViewProps {
  context?: WebPartContext;
  sp?: SPFI;
}

const DetailView: React.FC<IDetailViewProps> = ({ context, sp }) => {
  const navigate = useNavigate();
  const [tableData,   setTableData]   = useState<DetailViewRecord[]>([]);
  const [filters,     setFilters]     = useState<any>(null);
  const [isCollapsed, setIsCollapsed] = useState(false);
  const [projectOptions, setProjectOptions] = useState<IDropdownOption[]>([]);
  const [functionOptions, setFunctionOptions] = useState<IDropdownOption[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const loadData = async () => {
      if (!context || !sp) {
        setLoading(false);
        return;
      }

      try {
        setLoading(true);
        const dropdownService = new DropdownService(context);

        // 1. Fetch master dropdowns with fallback
        let projects: IDropdownOption[] = [];
        let funcs: IDropdownOption[] = [];
        try {
          projects = await dropdownService.getProjects();
        } catch (pErr) {
          console.warn("DetailView: Failed to load projects from ProjectMaster:", pErr);
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
          console.warn("DetailView: Failed to load functions from FunctionMaster:", fErr);
        }

        // 2. Fetch dashboard items
        let dashboardItems: any[] = [];
        const dashMap: Record<string, any> = {};
        try {
          dashboardItems = await sp.web.lists.getByTitle("DashboardList").items
            .select("WFID", "PendingFrom", "CurrentApprovalName", "Status")
            .top(5000)();
          dashboardItems.forEach((d: any) => {
            if (d.WFID) {
              dashMap[d.WFID] = d;
            }
          });
        } catch (dashErr) {
          console.warn("DetailView: Error fetching DashboardList:", dashErr);
        }

        // 3. Fetch all approval records from DigitalApprovalNotes
        let nfaItems: any[] = [];
        try {
          nfaItems = await sp.web.lists.getByTitle("DigitalApprovalNotes").items
            .select(
              "Id", "Title", "SubjectMatter", "Dept", "FunctionName", "Project",
              "RequestedOn", "Status", "ApprovalAmount", "Author/Title", "Author/Id",
              "PendingAt", "Modified", "Created", "PendingFrom", "WFLevel",
              "Approvedrejectedby", "Approvedrejectdatetime"
            )
            .expand("Author")
            .top(5000)();
        } catch (nfaErr) {
          console.error("DetailView: Error fetching DigitalApprovalNotes:", nfaErr);
        }

        // 4. Fallback dropdown options from distinct values in nfaItems if master lists were empty
        if ((!projects || projects.length === 0) && nfaItems.length > 0) {
          const uniqueProjects = Array.from(new Set(nfaItems.map((it: any) => it.Project).filter(Boolean)));
          projects = uniqueProjects.map((p: any) => ({ label: String(p), value: String(p) }));
        }

        if ((!funcs || funcs.length === 0) && nfaItems.length > 0) {
          const uniqueDepts = Array.from(new Set(nfaItems.map((it: any) => it.Dept || it.FunctionName).filter(Boolean)));
          funcs = uniqueDepts.map((d: any) => ({ label: String(d), value: String(d) }));
        }

        setProjectOptions(projects || []);
        setFunctionOptions(funcs || []);

        // 5. Map all records to DetailViewRecord
        const mapped: DetailViewRecord[] = nfaItems.map((item: any) => {
          const dash = dashMap[item.Id.toString()];
          let approvedOnVal = "";
          if (item.Approvedrejectdatetime) {
            approvedOnVal = item.Approvedrejectdatetime;
          } else if (item.Status === "Approved" || item.Status === "Completed" || item.Status === "Rejected") {
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
          const levelNum = parseInt(item.WFLevel || "1", 10);

          return {
            id: item.Id,
            nfa: item.Title || "",
            subject: stripHtml(item.SubjectMatter || item.Title || ""),
            function: item.Dept || item.FunctionName || "",
            createdBy: item.Author ? item.Author.Title : "",
            createdOn: item.RequestedOn || item.Created || "",
            status: item.Status || "",
            approvedRejOn: approvedOnVal,
            pendingAt: pendingAtVal,
            pendingFrom: pendingFromVal,
            level: isNaN(levelNum) ? 1 : levelNum,
            project: item.Project || "",
            packageCost: item.ApprovalAmount || "0.00"
          };
        });

        setTableData(mapped);
      } catch (err) {
        console.error("DetailView: Error loading page data:", err);
      } finally {
        setLoading(false);
      }
    };

    loadData();
  }, [context, sp]);

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
            const parts = dStr.split(/[\/\s]/);
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
        title="Detailed View" 
        subtitle="Detailed administrative overview of all approval records." 
        action={
          <button className="iconBtn" onClick={() => setIsCollapsed(c => !c)} title={isCollapsed ? 'Expand' : 'Collapse'}>
            {isCollapsed ? <ChevronDown size={16} /> : <ChevronUp size={16} />}
          </button>
        }
      />

      {!isCollapsed && (
        <>
          <DetailViewFilter 
            onSearch={handleSearch} 
            onClear={() => setFilters(null)} 
            projectOptions={projectOptions}
            functionOptions={functionOptions}
          />
          {loading ? (
            <div className="flex justify-center items-center p-8 bg-white rounded-2xl border border-slate-200 mt-3">
              <span className="text-slate-500 font-medium">Loading detailed records...</span>
            </div>
          ) : (
            <DetailViewTable  data={displayData} onRecordClick={(id) => navigate(`/view-nfa/${id}`)} />
          )}
        </>
      )}

    </div>
  );
};

export default DetailView;
