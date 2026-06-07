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

  const today = useMemo(() => new Date().toISOString().split('T')[0], []);
  const minus11 = useMemo(() => new Date(Date.now() - 11 * 24 * 60 * 60 * 1000).toISOString().split('T')[0], []);

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

  const activeFiltersList = useMemo(() => {
    if (!filters) return [];
    const list = [];
    if (filters.fromDate && filters.fromDate !== minus11) {
      list.push({ key: 'fromDate', label: `From: ${filters.fromDate}` });
    }
    if (filters.toDate && filters.toDate !== today) {
      list.push({ key: 'toDate', label: `To: ${filters.toDate}` });
    }
    if (filters.status) {
      list.push({ key: 'status', label: `Status: ${filters.status}` });
    }
    if (filters.project) {
      const proj = projectOptions.find(p => p.value === filters.project);
      list.push({ key: 'project', label: `Project: ${proj ? proj.label : filters.project}` });
    }
    if (filters.function) {
      const func = functionOptions.find(f => f.value === filters.function);
      list.push({ key: 'function', label: `Function: ${func ? func.label : filters.function}` });
    }
    return list;
  }, [filters, projectOptions, functionOptions, today, minus11]);

  const removeFilterKey = (key: string) => {
    if (!filters) return;
    const newFilters = { ...filters };
    if (key === 'fromDate') {
      newFilters.fromDate = minus11;
    } else if (key === 'toDate') {
      newFilters.toDate = today;
    } else {
      delete newFilters[key];
    }

    const hasActive = 
      newFilters.fromDate !== minus11 ||
      newFilters.toDate !== today ||
      (newFilters.status && newFilters.status !== '') ||
      (newFilters.project && newFilters.project !== '') ||
      (newFilters.function && newFilters.function !== '');

    if (!hasActive) {
      setFilters(null);
    } else {
      setFilters(newFilters);
    }
  };

  const handleCardClick = (status: string) => {
    const currentStatus = filters?.status || '';
    let targetStatus = '';
    if (status !== 'total') {
      targetStatus = currentStatus === status ? '' : status;
    }

    const baseFilters = filters || {
      fromDate: minus11,
      toDate: today,
      function: '',
      project: ''
    };

    const newFilters = {
      ...baseFilters,
      status: targetStatus
    };

    const hasActive = 
      newFilters.fromDate !== minus11 ||
      newFilters.toDate !== today ||
      newFilters.status !== '' ||
      newFilters.project !== '' ||
      newFilters.function !== '';

    if (!hasActive) {
      setFilters(null);
    } else {
      setFilters(newFilters);
    }
  };

  const handleSearch = (newFilters: any) => {
    setFilters(newFilters);
  };

  const SummaryCard = ({ icon: Icon, label, count, iconBgClass, iconColorClass, isActive, activeGlowClass, onClick }: any) => (
    <div 
      className={`flex-1 min-w-[150px] h-24 bg-white rounded-2xl px-5 shadow-sm border transition-all duration-300 group flex items-center gap-4 cursor-pointer select-none ${
        isActive 
          ? `border-[#AF8941] scale-[1.03] shadow-md ${activeGlowClass}` 
          : 'border-slate-200 hover:shadow-md hover:border-[#AF8941]/50 hover:-translate-y-0.5'
      }`}
      onClick={onClick}
    >
      <div className={`p-3 rounded-full flex items-center justify-center transition-transform duration-300 group-hover:scale-110 ${iconBgClass}`}>
        <Icon size={24} className={iconColorClass} />
      </div>
      <div className="flex flex-col">
        <span className="text-slate-400 text-[10px] font-bold uppercase tracking-widest mb-1">{label}</span>
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
        <SummaryCard 
          icon={FileText}     
          label="Total Requests" 
          count={stats.total}    
          iconBgClass="bg-blue-500 bg-opacity-10"    
          iconColorClass="text-blue-500" 
          isActive={!filters?.status}
          activeGlowClass="shadow-blue-50"
          onClick={() => handleCardClick('total')}
        />
        <SummaryCard 
          icon={Clock}        
          label="Pending"        
          count={stats.pending}  
          iconBgClass="bg-amber-500 bg-opacity-10"   
          iconColorClass="text-amber-500" 
          isActive={filters?.status === 'Pending'}
          activeGlowClass="shadow-amber-100"
          onClick={() => handleCardClick('Pending')}
        />
        <SummaryCard 
          icon={CheckCircle2} 
          label="Approved"       
          count={stats.approved} 
          iconBgClass="bg-emerald-500 bg-opacity-10" 
          iconColorClass="text-emerald-500" 
          isActive={filters?.status === 'Approved'}
          activeGlowClass="shadow-emerald-100"
          onClick={() => handleCardClick('Approved')}
        />
        <SummaryCard 
          icon={XCircle}      
          label="Rejected"       
          count={stats.rejected} 
          iconBgClass="bg-rose-500 bg-opacity-10"    
          iconColorClass="text-rose-500" 
          isActive={filters?.status === 'Rejected'}
          activeGlowClass="shadow-rose-100"
          onClick={() => handleCardClick('Rejected')}
        />
        <SummaryCard 
          icon={FileEdit}     
          label="Draft"          
          count={stats.draft}    
          iconBgClass="bg-slate-500 bg-opacity-10"   
          iconColorClass="text-slate-500" 
          isActive={filters?.status === 'Draft'}
          activeGlowClass="shadow-slate-100"
          onClick={() => handleCardClick('Draft')}
        />
      </div>

      {/* ── Active Filter Pills ── */}
      {activeFiltersList.length > 0 && (
        <div className="flex flex-wrap items-center gap-2 mb-4 mt-2 px-1">
          <span className="text-[10px] font-extrabold text-slate-400 uppercase tracking-widest mr-1">Active Filters:</span>
          {activeFiltersList.map(f => (
            <div key={f.key} className="flex items-center gap-1.5 px-3 py-1 rounded-full bg-slate-200/50 border border-slate-200/80 text-[11px] font-bold text-slate-700 transition-colors hover:bg-slate-200">
              <span>{f.label}</span>
              <button 
                className="text-slate-400 hover:text-red-500 font-bold ml-1 transition-colors text-xs"
                onClick={() => removeFilterKey(f.key)}
              >
                ✕
              </button>
            </div>
          ))}
          <button 
            className="text-[11px] font-extrabold text-red-500 hover:text-red-700 ml-2 transition-colors underline uppercase tracking-wider"
            onClick={() => setFilters(null)}
          >
            Clear All
          </button>
        </div>
      )}

      {/* ── Filter ── */}
      <HomeFilter 
        filters={filters}
        onSearch={handleSearch} 
        onClear={() => setFilters(null)}
        projectOptions={projectOptions} 
        functionOptions={functionOptions} 
      />

      {/* ── Table ── */}
      <HomeTable
        data={displayData}
        loading={loading}
        onClearFilters={() => setFilters(null)}
        onRowClick={(record) => {
           const status = (record.status || "").toLowerCase();
           if (status === 'draft') {
             navigate(`/new-nfa/${record.id}`);
           } else if (status === 'pending') {
             navigate(`/view-nfa/${record.id}`);
           } else {
             navigate(`/view-nfa/${record.id}`);
           }
        }}
      />

    </div>
  );
};

export default Home;
