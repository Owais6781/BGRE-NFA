import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import HomeFilter from './HomeFilter';
import HomeTable  from './HomeTable';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { DropdownService, IDropdownOption } from '../../../../../Share/Services/DropdownService';
import { NFAService } from '../../../../../Share/Services/NFAService';
import { useNavigate } from 'react-router-dom';
import { Clock, CheckCircle2, XCircle, FileEdit } from 'lucide-react';
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
      pending:  tableData.filter(item => (item.status || "").toLowerCase() === 'pending').length,
      approved: tableData.filter(item => (item.status || "").toLowerCase() === 'approved').length,
      rejected: tableData.filter(item => (item.status || "").toLowerCase() === 'rejected').length,
      draft:    tableData.filter(item => (item.status || "").toLowerCase() === 'draft').length,
    };
  }, [tableData]);

  const handleSearch = (filters: any) => {
    console.log('Search filters:', filters);
  };

  const SummaryCard = ({ icon: Icon, label, count, colorClass }: any) => (
    <div className="flex-1 min-w-[160px] bg-white rounded-xl border border-slate-200 p-4 shadow-sm flex items-center gap-4 transition-transform hover:scale-[1.02]">
      <div className={`p-3 rounded-lg ${colorClass} bg-opacity-10`}>
        <Icon size={20} className={colorClass.replace('bg-', 'text-')} />
      </div>
      <div>
        <p className="text-[11px] font-semibold text-slate-500 uppercase tracking-wider">{label}</p>
        <p className="text-xl font-bold text-slate-800">{count}</p>
      </div>
    </div>
  );

  return (
    <div className="pageContainer">

      <div className="sectionTitleBar flex items-center justify-between">
        <span className="sectionTitle">Home</span>
      </div>

      {/* ── Summary Stats ── */}
      <div className="flex flex-wrap gap-4 mb-6">
        <SummaryCard icon={Clock}        label="Pending"  count={stats.pending}  colorClass="bg-amber-500" />
        <SummaryCard icon={CheckCircle2} label="Approved" count={stats.approved} colorClass="bg-emerald-500" />
        <SummaryCard icon={XCircle}      label="Rejected" count={stats.rejected} colorClass="bg-rose-500" />
        <SummaryCard icon={FileEdit}     label="Draft"    count={stats.draft}    colorClass="bg-slate-500" />
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
