import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import HomeFilter from './HomeFilter';
import HomeTable  from './HomeTable';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { DropdownService, IDropdownOption } from '../../../../../Share/Services/DropdownService';
import { NFAService } from '../../../../../Share/Services/NFAService';
import { useNavigate } from 'react-router-dom';
import { Clock, CheckCircle2, XCircle, FileEdit, Files } from 'lucide-react';
import "../../../../../Share/styles/global.scss";

interface IHomeProps {
  context: WebPartContext;
  sp: SPFI;
}

const Home: React.FC<IHomeProps> = ({ context, sp }) => {
  const [tableData, setTableData] = useState<any[]>([]);
  const [projectOptions, setProjectOptions] = useState<IDropdownOption[]>([]);
  const [functionOptions, setFunctionOptions] = useState<IDropdownOption[]>([]);
  const navigate = useNavigate();

  useEffect(() => {
    const loadData = async () => {
      const dropdownService = new DropdownService(context);
      const nfaService = new NFAService(context);
      try {
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
      }
    };
    loadData();
  }, []);

  // --- Calculate Stats ---
  const stats = useMemo(() => {
    return {
      total:    tableData.length,
      pending:  tableData.filter(item => (item.status || "").toLowerCase() === 'pending').length,
      approved: tableData.filter(item => (item.status || "").toLowerCase() === 'approved').length,
      rejected: tableData.filter(item => (item.status || "").toLowerCase() === 'rejected').length,
      draft:    tableData.filter(item => (item.status || "").toLowerCase() === 'draft').length,
    };
  }, [tableData]);

  const handleSearch = (filters: any) => {
    console.log('Search filters:', filters);
  };

  const SummaryCard = ({ icon: Icon, label, count, iconBgClass, iconColorClass }: any) => (
    <div className="flex-1 min-w-[180px] max-w-[250px] h-24 bg-white rounded-2xl px-5 shadow-sm border border-[#AF8941]/40 hover:shadow-md hover:border-[#AF8941]/70 transition-all duration-300 group flex items-center gap-4">
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

      <div className="flex items-center justify-between mb-4">
        <span className="sectionTitle">Home</span>
      </div>

      {/* ── Summary Stats ── */}
      <div className="flex flex-wrap gap-4 mb-6">
        <SummaryCard icon={Files}        label="Total Requests" count={stats.total}    iconBgClass="bg-blue-500 bg-opacity-10"    iconColorClass="text-blue-500" />
        <SummaryCard icon={Clock}        label="Pending"        count={stats.pending}  iconBgClass="bg-amber-500 bg-opacity-10"   iconColorClass="text-amber-500" />
        <SummaryCard icon={CheckCircle2} label="Approved"       count={stats.approved} iconBgClass="bg-emerald-500 bg-opacity-10" iconColorClass="text-emerald-500" />
        <SummaryCard icon={XCircle}      label="Rejected"       count={stats.rejected} iconBgClass="bg-rose-500 bg-opacity-10"    iconColorClass="text-rose-500" />
        <SummaryCard icon={FileEdit}     label="Draft"          count={stats.draft}    iconBgClass="bg-slate-500 bg-opacity-10"   iconColorClass="text-slate-500" />
      </div>

      {/* ── Filter ── */}
      <HomeFilter 
        onSearch={handleSearch} 
        projectOptions={projectOptions} 
        functionOptions={functionOptions} 
      />

      {/* ── Table ── */}
      <HomeTable
        data={tableData}
        onRequestClick={(record) => {
           const status = (record.status || "").toLowerCase();
           if (status === 'draft') {
             navigate(`/new-nfa/${record.id}`);
           } else {
             navigate(`/view-nfa/${record.id}`);
           }
        }}
      />

    </div>
  );
};

export default Home;
