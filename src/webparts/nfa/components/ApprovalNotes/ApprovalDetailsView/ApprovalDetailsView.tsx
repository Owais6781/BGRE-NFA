import * as React from 'react';
import { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { ChevronUp, ChevronDown } from 'lucide-react';
import ApprovalDetailsFilter from './ApprovalDetailsFilter';
import ApprovalDetailsTable  from './ApprovalDetailsTable';
import "../../../../../Share/styles/global.scss";

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
        
        const [projects, funcs, nfaItems, dashboardItems] = await Promise.all([
          dropdownService.getProjects(),
          dropdownService.getFunctionMaster(),
          sp.web.lists.getByTitle("DigitalApprovalNotes").items
            .select("Id", "Title", "SubjectMatter", "Dept", "Project", "RequestedOn", "Status", "ApprovalAmount", "Author/Title", "PendingAt/Title", "Modified", "Created")
            .expand("Author", "PendingAt")
            .top(5000)(),
          sp.web.lists.getByTitle("DashboardList").items
            .select("WFID", "PendingFrom")
            .top(5000)()
        ]);

        setProjectOptions(projects);
        setFunctionOptions(funcs);

        const dashMap: Record<string, any> = {};
        dashboardItems.forEach((d: any) => {
          if (d.WFID) {
            dashMap[d.WFID] = d;
          }
        });

        const mapped: ApprovalRecord[] = nfaItems.map((item: any) => {
          const dash = dashMap[item.Id.toString()];
          let approvedOnVal = "";
          if (item.Status === "Approved" || item.Status === "Rejected") {
            approvedOnVal = item.Modified || "";
          }

          return {
            id: item.Id,
            nfa: item.Title || "",
            subject: item.SubjectMatter || item.Title || "",
            function: item.Dept || "",
            createdBy: item.Author ? item.Author.Title : "",
            createdOn: item.RequestedOn || item.Created || "",
            status: item.Status || "",
            approvedOn: approvedOnVal,
            pendingAt: item.PendingAt ? item.PendingAt.Title : "",
            pendingFrom: dash ? (dash.PendingFrom || "") : "",
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

  // --- Filtering Logic ---
  const displayData = React.useMemo(() => {
    if (!filters) return tableData;

    return tableData.filter(item => {
      // Date Filter
      if (filters.fromDate || filters.toDate) {
        const itemDate = new Date(item.createdOn);
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

      // Dropdown Filters
      if (filters.status && item.status !== filters.status) return false;

      if (filters.function) {
        const funcObj = functionOptions.find(f => f.value === filters.function);
        const funcName = funcObj ? funcObj.label : filters.function;
        if (item.function !== funcName) return false;
      }

      if (filters.project) {
        const projObj = projectOptions.find(p => p.value === filters.project);
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

      <div className="sectionTitleBar flex items-center justify-between">
        <span className="sectionTitle">Detailed View for Approver</span>
        <button className="iconBtn" onClick={() => setIsCollapsed(c => !c)} title={isCollapsed ? 'Expand' : 'Collapse'}>
          {isCollapsed ? <ChevronDown size={16} /> : <ChevronUp size={16} />}
        </button>
      </div>

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
            <ApprovalDetailsTable  data={displayData} onRecordClick={(id) => navigate(`/view-nfa/${id}`)} />
          )}
        </>
      )}

    </div>
  );
};

export default ApprovalDetailsView;
