import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import HomeFilter from './HomeFilter';
import HomeTable  from './HomeTable';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { DropdownService, IDropdownOption } from '../../../../../Share/Services/DropdownService';
import { NFAService } from '../../../../../Share/Services/NFAService';
import { useNavigate } from 'react-router-dom';
import { Clock, CheckCircle2, XCircle, FileEdit, FileText } from 'lucide-react';
import "../../../../../Share/styles/global.scss";

interface IHomeProps {
  context: WebPartContext;
  sp: SPFI;
}

const Home: React.FC<IHomeProps> = ({ context, sp }) => {
  const [tableData, setTableData] = useState<any[]>([]);
  const [projectOptions, setProjectOptions] = useState<IDropdownOption[]>([]);
  const [functionOptions, setFunctionOptions] = useState<IDropdownOption[]>([]);
  const [filters, setFilters] = useState<any>(null);
  const [loading, setLoading] = useState(true);
  const navigate = useNavigate();

  useEffect(() => {
    const loadData = async () => {
      const dropdownService = new DropdownService(context);
      const nfaService = new NFAService(context);
      try {
        setLoading(true);
        const [projects, funcs, nfas] = await Promise.all([
          dropdownService.getProjects(),
          dropdownService.getFunctionMaster(),
          nfaService.getNFAList()
        ]);
        
        setProjectOptions(projects);
        setFunctionOptions(funcs);
        setTableData(nfas);
      } catch (err) {
        console.error("Failed to load Home data:", err);
      } finally {
        setLoading(false);
      }
    };
    loadData();
  }, []);

  // --- Filtering Logic ---
  const displayData = useMemo(() => {
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

      // Dropdown Filters (Empty value "" means "All")
      if (filters.status && item.status !== filters.status) return false;
      
      if (filters.project) {
        // Since project in item might be label but projectVal is ID, we check label
        const projObj = projectOptions.find(p => p.value === filters.project);
        const projName = projObj ? projObj.label : filters.project;
        if (item.project !== projName) return false;
      }

      if (filters.function) {
        const funcObj = functionOptions.find(f => f.value === filters.function);
        const funcName = funcObj ? funcObj.label : filters.function;
        if (item.function !== funcName) return false;
      }

      return true;
    });
  }, [tableData, filters, projectOptions, functionOptions]);

  // --- Calculate Stats ---
  const stats = useMemo(() => {
    return {
      total:    displayData.length,
      pending:  displayData.filter(item => (item.status || "").toLowerCase() === 'pending').length,
      approved: displayData.filter(item => (item.status || "").toLowerCase() === 'approved').length,
      rejected: displayData.filter(item => (item.status || "").toLowerCase() === 'rejected').length,
      draft:    displayData.filter(item => (item.status || "").toLowerCase() === 'draft').length,
    };
  }, [displayData]);

  const handleSearch = (newFilters: any) => {
    setFilters(newFilters);
  };

  const SummaryCard = ({ icon: Icon, label, count, iconBgClass, iconColorClass }: any) => (
    <div className="flex-1 min-w-[150px] h-24 bg-white rounded-2xl px-5 shadow-sm border border-[#AF8941]/40 hover:shadow-md hover:border-[#AF8941]/70 transition-all duration-300 group flex items-center gap-4">
      <div className={`p-3 rounded-full flex items-center justify-center ${iconBgClass}`}>
        <Icon size={24} className={iconColorClass} />
      </div>
      <div className="flex flex-col">
        <span className="text-slate-400 text-xs font-bold uppercase tracking-widest mb-1">{label}</span>
        <span className="text-3xl font-extrabold text-slate-800 leading-none">{count}</span>
      </div>
    </div>
  );

  return (
    <div className="pageContainer">

      <div className="sectionTitleBar flex items-center justify-between mb-4">
        <span className="sectionTitle">Home</span>
      </div>

      {/* ── Summary Stats ── */}
      <div className="flex flex-wrap lg:flex-nowrap gap-4" style={{ marginBottom: '16px' }}>
        <SummaryCard icon={FileText}     label="Total Requests" count={stats.total}    iconBgClass="bg-blue-500 bg-opacity-10"    iconColorClass="text-blue-500" />
        <SummaryCard icon={Clock}        label="Pending"        count={stats.pending}  iconBgClass="bg-amber-500 bg-opacity-10"   iconColorClass="text-amber-500" />
        <SummaryCard icon={CheckCircle2} label="Approved"       count={stats.approved} iconBgClass="bg-emerald-500 bg-opacity-10" iconColorClass="text-emerald-500" />
        <SummaryCard icon={XCircle}      label="Rejected"       count={stats.rejected} iconBgClass="bg-rose-500 bg-opacity-10"    iconColorClass="text-rose-500" />
        <SummaryCard icon={FileEdit}     label="Draft"          count={stats.draft}    iconBgClass="bg-slate-500 bg-opacity-10"   iconColorClass="text-slate-500" />
      </div>

      {/* ── Filter ── */}
      <HomeFilter 
        onSearch={handleSearch} 
        onClear={() => setFilters(null)}
        projectOptions={projectOptions} 
        functionOptions={functionOptions} 
      />

      {/* ── Table ── */}
      <HomeTable
        data={displayData}
        loading={loading}
        onRowClick={(record) => {
           const status = (record.status || "").toLowerCase();
           if (status === 'draft') {
             navigate(`/new-nfa/${record.id}`);
           } else if (status === 'pending') {
             navigate(`/approver-page/${record.id}`);
           } else {
             navigate(`/view-nfa/${record.id}`);
           }
        }}
      />

    </div>
  );
};

export default Home;
