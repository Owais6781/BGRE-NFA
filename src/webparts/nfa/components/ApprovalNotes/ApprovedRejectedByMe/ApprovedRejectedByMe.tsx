import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import { ChevronUp, ChevronDown, CheckCircle2, XCircle, FileText, Building2, Wrench } from 'lucide-react';
import { CommonTable } from '../../../../../Share/Components/CommonTable';
import { Pagination }  from '../../../../../Share/Components/Pagination';
import TableToolbar    from '../../../../../Share/Components/TableToolbar';
import ExportToExcel   from '../../../../../Share/Components/ExportToExcel';
import { SearchInput } from '../../../../../Share/Components/SearchInput';
import { useNavigate } from 'react-router-dom';
import { Skeleton }    from 'primereact/skeleton';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { CommonService, ApprovedRejectedRecord } from '../../../../../Share/Services/CommonService';
import { DropdownService, IDropdownOption } from '../../../../../Share/Services/DropdownService';
import ApprovedRejectedFilter from './ApprovedRejectedFilter';
import "../../../../../Share/styles/global.scss";

const PAGE_SIZE_OPTIONS = [10, 25, 50, 100];

const exportColumns = [
  { key: 'requestNo',            label: 'Request No.'                },
  { key: 'function',             label: 'Function'                   },
  { key: 'project',              label: 'Project'                    },
  { key: 'createdBy',            label: 'Created By'                 },
  { key: 'createdOn',            label: 'Created On'                 },
  { key: 'status',               label: 'Status'                     },
  { key: 'level',                label: 'Level'                      },
  { key: 'approvedRejectedBy',   label: 'Approved/Rejected By'       },
  { key: 'approvedRejectedDate', label: 'Approved/Rejected Datetime' },
];

interface IApprovedRejectedProps {
  context: WebPartContext;
  sp: SPFI;
}

const ApprovedRejectedByMe: React.FC<IApprovedRejectedProps> = ({ context, sp }) => {
  const [data, setData] = useState<ApprovedRejectedRecord[]>([]);
  const [projectOptions, setProjectOptions] = useState<IDropdownOption[]>([]);
  const [functionOptions, setFunctionOptions] = useState<IDropdownOption[]>([]);
  const [filters, setFilters] = useState<any>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [pageSize, setPageSize] = useState(10);
  const [search, setSearch] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [isCollapsed, setIsCollapsed] = useState(false);
  const navigate = useNavigate();

  const today = useMemo(() => new Date().toISOString().split('T')[0], []);
  const minus11 = useMemo(() => new Date(Date.now() - 11 * 24 * 60 * 60 * 1000).toISOString().split('T')[0], []);

  useEffect(() => {
    const loadData = async () => {
      setLoading(true);
      setError('');
      try {
        const commonService = new CommonService(context, sp);
        const dropdownService = new DropdownService(context);
        
        const [projects, funcs, records] = await Promise.all([
          dropdownService.getProjects(),
          dropdownService.getFunctionMaster(),
          commonService.getApprovedRejectedByMe()
        ]);

        setProjectOptions(projects);
        setFunctionOptions(funcs);
        setData(records);
      } catch (err) {
        console.error("ApprovedRejectedByMe: Failed to load data:", err);
        setError("Failed to retrieve approved/rejected requests. Please try again later.");
      } finally {
        setLoading(false);
      }
    };
    loadData();
  }, []);

  // --- Filtering Logic ---
  const displayData = useMemo(() => {
    let result = data;

    // Search bar filter (local table toolbar search)
    if (search) {
      result = result.filter((row) =>
        Object.keys(row).some((k) => String((row as any)[k]).toLowerCase().includes(search.toLowerCase()))
      );
    }

    if (!filters) return result;

    return result.filter(item => {
      // Date Filter
      if (filters.fromDate || filters.toDate) {
        const dateStr = item.createdOnRaw || item.createdOn;
        if (dateStr) {
          const itemDate = new Date(dateStr);
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

      // Status Filter
      if (filters.status && item.status !== filters.status) return false;

      // Project Filter
      if (filters.project) {
        const projObj = projectOptions.find(p => p.value === filters.project);
        const projName = projObj ? projObj.label : filters.project;
        if (item.project !== projName) return false;
      }

      // Function Filter
      if (filters.function) {
        const funcObj = functionOptions.find(f => f.value === filters.function);
        const funcName = funcObj ? funcObj.label : filters.function;
        if (item.function !== funcName) return false;
      }

      return true;
    });
  }, [data, search, filters, projectOptions, functionOptions]);

  // --- Stats Calculation ---
  const stats = useMemo(() => {
    const total = displayData.length;
    const approved = displayData.filter(item => item.status === 'Approved').length;
    const rejected = displayData.filter(item => item.status === 'Rejected').length;
    const projects = new Set(displayData.map(item => item.project).filter(p => p)).size;
    const functions = new Set(displayData.map(item => item.function).filter(f => f)).size;

    let totalMs = 0;
    let count = 0;
    displayData.forEach(item => {
      if (item.createdOnRaw && item.approvedRejectedDateRaw) {
        const initDate = new Date(item.createdOnRaw);
        const actionDate = new Date(item.approvedRejectedDateRaw);
        if (!isNaN(initDate.getTime()) && !isNaN(actionDate.getTime())) {
          const diff = actionDate.getTime() - initDate.getTime();
          if (diff >= 0) {
            totalMs += diff;
            count++;
          }
        }
      }
    });

    let avgTimeStr = "N/A";
    if (count > 0) {
      const avgMs = totalMs / count;
      const avgHours = avgMs / (1000 * 60 * 60);
      if (avgHours < 24) {
        avgTimeStr = `${Math.round(avgHours)}h`;
      } else {
        avgTimeStr = `${(avgHours / 24).toFixed(1)}d`;
      }
    }

    return { total, approved, rejected, projects, functions, avgTimeStr };
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
    setCurrentPage(1);
  };

  const handleClear = () => {
    setFilters(null);
    setCurrentPage(1);
  };

  const handlePageSizeChange = (size: number) => {
    setPageSize(size);
    setCurrentPage(1);
  };

  const handleTableSearch = (val: string) => {
    setSearch(val);
    setCurrentPage(1);
  };

  const columns = [
    { key: 'requestNo', label: 'Request No.',
      render: (val: string, row: ApprovedRejectedRecord) => (
        <button 
          className="tableLinkBtn" 
          onClick={(e) => { 
            e.stopPropagation(); 
            navigate(`/view-nfa/${row.id}`); 
          }}
        >
          {val}
        </button>
      ) 
    },
    { key: 'function',  label: 'Function'   },
    { key: 'project',   label: 'Project'    },
    { key: 'createdBy', label: 'Created By' },
    { key: 'createdOn', label: 'Created On' },
    { key: 'status',    label: 'Status',
      render: (val: string) => (
        <span className={`statusBadge ${val === 'Approved' ? 'statusApproved' : val === 'Rejected' ? 'statusRejected' : 'statusPending'}`}>{val}</span>
      ) },
    { key: 'level',                label: 'Level'                     },
    { key: 'approvedRejectedBy',   label: 'Approved/Rejected By'      },
    { key: 'approvedRejectedDate', label: 'Approved/Rejected Datetime' },
  ];

  const totalPages = Math.max(1, Math.ceil(displayData.length / pageSize));
  const paginated = displayData.slice((currentPage - 1) * pageSize, currentPage * pageSize);

  const SummaryCard = ({ icon: Icon, label, count, iconBgClass, iconColorClass, isActive, activeGlowClass, onClick }: any) => (
    <div 
      className={`flex-1 min-w-[150px] h-24 bg-white rounded-2xl px-5 shadow-sm border transition-all duration-300 group flex items-center gap-4 select-none ${
        onClick ? 'cursor-pointer hover:shadow-md hover:border-[#AF8941]/50 hover:-translate-y-0.5' : 'cursor-default'
      } ${
        isActive 
          ? `border-[#AF8941] scale-[1.03] shadow-md ${activeGlowClass}` 
          : 'border-slate-200'
      }`}
      onClick={onClick || undefined}
    >
      <div className={`p-3 rounded-full flex items-center justify-center transition-transform duration-300 ${onClick ? 'group-hover:scale-110' : ''} ${iconBgClass}`}>
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
      <div className="contentCard">

        <div className="sectionTitleBar flex items-center justify-between mb-4">
          <span className="sectionTitle">Approved / Rejected by Me</span>
          <button className="iconBtn" onClick={() => setIsCollapsed(c => !c)} title={isCollapsed ? 'Expand' : 'Collapse'}>
            {isCollapsed ? <ChevronDown size={16} /> : <ChevronUp size={16} />}
          </button>
        </div>

        {!isCollapsed && (
          <>
            {error && <div className="alertError">{error}</div>}

            {/* Stats Cards */}
            <div className="flex flex-wrap lg:flex-nowrap gap-4 mb-4">
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
                icon={Building2}     
                label="Projects"          
                count={stats.projects}    
                iconBgClass="bg-cyan-500 bg-opacity-10"   
                iconColorClass="text-cyan-500" 
                isActive={false}
                activeGlowClass="shadow-cyan-100"
                onClick={null}
              />
              <SummaryCard 
                icon={Wrench}     
                label="Functions"          
                count={stats.functions}    
                iconBgClass="bg-purple-500 bg-opacity-10"   
                iconColorClass="text-purple-500" 
                isActive={false}
                activeGlowClass="shadow-purple-100"
                onClick={null}
              />
            </div>

            {/* Active Filters Pills */}
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
                  onClick={handleClear}
                >
                  Clear All
                </button>
              </div>
            )}

            {/* Filter Panel */}
            <div className="mb-4">
              <ApprovedRejectedFilter 
                filters={filters}
                onSearch={handleSearch} 
                onClear={handleClear}
                projectOptions={projectOptions} 
                functionOptions={functionOptions} 
              />
            </div>

            <TableToolbar
              pageSize={pageSize} pageSizeOptions={PAGE_SIZE_OPTIONS} onPageSizeChange={handlePageSizeChange}
              search={search} onSearch={handleTableSearch}
              right={
                <div className="flex items-center gap-3">
                  <ExportToExcel data={displayData} columns={exportColumns} fileName="Approved_Rejected_By_Me" sheetName="Approved Rejected" />
                  <SearchInput value={search} onChange={handleTableSearch} />
                </div>
              }
            />

            {loading ? (
              <div className="w-full mt-3 border border-slate-200 rounded-xl overflow-hidden">
                {/* Header Skeleton */}
                <div className="flex bg-slate-50 border-b border-slate-200 p-4">
                  {columns.map((col) => (
                    <div key={col.key} className="flex-1">
                      <Skeleton width="60%" height="1.2rem" />
                    </div>
                  ))}
                </div>
                {/* Rows Skeleton */}
                {Array.from({ length: pageSize || 5 }).map((_, i) => (
                  <div key={i} className="flex border-b border-slate-100 p-4 items-center">
                    {columns.map((col) => (
                      <div key={col.key} className="flex-1">
                        <Skeleton width="80%" height="1rem" />
                      </div>
                    ))}
                  </div>
                ))}
              </div>
            ) : (
              <CommonTable
                columns={columns}
                data={paginated}
                emptyMessage="No approved/rejected records found"
                stickyFirstColumn={true}
              />
            )}

            <Pagination
              page={currentPage} total={totalPages} pageSize={pageSize}
              totalRecords={displayData.length} pageSizeOptions={PAGE_SIZE_OPTIONS}
              onPageChange={setCurrentPage} onPageSizeChange={handlePageSizeChange}
            />
          </>
        )}

      </div>
    </div>
  );
};

export default ApprovedRejectedByMe;
