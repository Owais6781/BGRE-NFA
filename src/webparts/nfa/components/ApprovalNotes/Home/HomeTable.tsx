import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import { CommonTable } from '../../../../../Share/Components/CommonTable';
import { Pagination }  from '../../../../../Share/Components/Pagination';
import TableToolbar    from '../../../../../Share/Components/TableToolbar';
import ExportToExcel   from '../../../../../Share/Components/ExportToExcel';
import { SearchInput } from '../../../../../Share/Components/SearchInput';
import { formatDateFriendly } from '../../../../../Share/utils/dateUtils';
import "../../../../../Share/styles/global.scss";

interface NFARecord {
  id: number;
  uniqId?: string;
  requestNo: string; 
  function: string;
  contractorName: string;
  project: string;
  createdOn: string; // CDate
  status: string; 
  Approvedrejectedby: string; 
  Approvedrejectdatetime: string; 
  pendingAt: string;
  level: number;
  PendingFrom: string;
  packageAwardedCost: string;
   doaMatrix: string;
}

interface HomeTableProps {
  data: NFARecord[];
  onRowClick?: (record: NFARecord) => void;
  loading?: boolean;
  onClearFilters?: () => void;
}

const PAGE_SIZE_OPTIONS = [10, 25, 50, 100, 500];

const exportColumns = [
  { key: 'requestNo', label: 'Request No.' },
  { key: 'function',  label: 'Function'    },
  { key: 'contractorName',  label: 'Contractor Name'    },
  { key: 'project',   label: 'Entity'     },
  { key: 'createdOn', label: 'CDate'       },
  { key: 'status',    label: 'Status'      },
  { key: 'Approvedrejectedby',    label: 'Approved Reject By'      },
  { key: 'Approvedrejectdatetime',    label: 'Approved Reject Date'      },
  { key: 'pendingAt', label: 'Pending At'  },
  { key: 'level',     label: 'Level'       },
   { key: 'PendingFrom',     label: 'Pending From'       },
   { key: 'doaMatrix',     label: 'DOA Clause'},
   { key: 'packageAwardedCost',     label: 'Amount'},
];


const HomeTable: React.FC<HomeTableProps> = ({ data, onRowClick, loading = false, onClearFilters }) => {
  const [pageSize,     setPageSize]    = useState(10);
  const [search,       setSearch]      = useState('');
  const [currentPage,  setCurrentPage] = useState(1);

  const columns = [
    {
      key: 'requestNo', label: 'Request No.',
      render: (val: string, row: NFARecord) => (
        <button className="tableLinkBtn"
        onClick={(e) => {
        e.preventDefault();
        e.stopPropagation();

        if (row.id) {
          onRowClick?.(row);
        }
      }}
         >
          {val}
        </button>
      ),
    },
    { key: 'function', label: 'Function' },
     { key: 'contractorName', label: 'Contractor Name' },
        { 
      key: 'createdOn', 
      label: 'Created On',
      render: (val: string) => formatDateFriendly(val)
    },
    {
      key: 'status', label: 'Status',
      render: (val: string) => {
        const clsMap: Record<string, { bg: string, text: string, border: string }> = {
          'Approved':  { bg: 'bg-emerald-50/60', text: 'text-emerald-700', border: 'border-emerald-500/30' },
          'Completed': { bg: 'bg-emerald-50/60', text: 'text-emerald-700', border: 'border-emerald-500/30' },
          'Rejected': { bg: 'bg-rose-50/60', text: 'text-rose-700', border: 'border-rose-500/30' },
          'Draft':    { bg: 'bg-slate-50/60', text: 'text-slate-700', border: 'border-slate-500/30' },
          'Pending':  { bg: 'bg-amber-50/60', text: 'text-amber-700', border: 'border-amber-500/30' }
        };
        const style = clsMap[val] || { bg: 'bg-blue-50/60', text: 'text-blue-700', border: 'border-blue-500/30' };
        
        return (
          <span className={`inline-flex items-center px-2.5 py-1 rounded-full text-xs font-semibold border ${style.bg} ${style.text} ${style.border}`}>
            <span className="w-1.5 h-1.5 rounded-full mr-1.5 shrink-0 bg-current opacity-75"></span>
            {val}
          </span>
        );
      }
    },
      { key: 'Approvedrejectedby',  label: 'Approved Reject By'  },
        { key: 'Approvedrejectdatetime',  label: 'Approved Reject Date'  },


   
 
    
    { 
      key: 'pendingAt', 
      label: 'Pending At',
      render: (val: string, row: NFARecord) => {
        const statusLower = (row.status || "").toLowerCase();
        const isActivelyPending = (statusLower === "pending" || statusLower === "query") && !row.Approvedrejectedby;
        if (!isActivelyPending || !val) return <span className="text-slate-400 font-normal italic">—</span>;
        
        const parts = val.trim().split(/\s+/);
        const initials = parts.map(p => p[0]).join('').substring(0, 2).toUpperCase();
        
        const colors = [
          'bg-blue-500/10 text-blue-600 border-blue-200',
          'bg-emerald-500/10 text-emerald-600 border-emerald-200',
          'bg-indigo-500/10 text-indigo-600 border-indigo-200',
          'bg-purple-500/10 text-purple-600 border-purple-200',
          'bg-amber-500/10 text-amber-600 border-amber-200',
          'bg-rose-500/10 text-rose-600 border-rose-200'
        ];
        const idx = Math.abs(val.split('').reduce((acc, char) => acc + char.charCodeAt(0), 0)) % colors.length;
        const colorClass = colors[idx];
        
        return (
          <div className="flex items-center gap-2">
            <div className={`w-7 h-7 rounded-full flex items-center justify-center font-bold text-xs shrink-0 border ${colorClass}`}>
              {initials}
            </div>
            <span className="font-semibold text-slate-800 text-sm truncate max-w-[140px]">{val}</span>
          </div>
        );
      }
    },
    { 
      key: 'PendingFrom',     
      label: 'Pending From',
      render: (val: string, row: NFARecord) => {
        const statusLower = (row.status || "").toLowerCase();
        const isActivelyPending = (statusLower === "pending" || statusLower === "query") && !row.Approvedrejectedby;
        if (!isActivelyPending || !val) return <span className="text-slate-400 font-normal italic">—</span>;
        return <span>{val}</span>;
      }
    },
    { key: 'level',     label: 'Level'      },
     { key: 'doaMatrix',     label: 'DOA Clause'},
     { key: 'project',  label: 'Entity'  },
      { key: 'packageAwardedCost',  label: 'Amount'  },
  ];

  useEffect(() => {
    setCurrentPage(1);
  }, [data]);

  const filtered = useMemo(() => {
    if (!search || !search.trim()) return data;
    const query = search.trim().toLowerCase();
    return data.filter((row) =>
      Object.keys(row).some((k) => {
        const val = (row as any)[k];
        return val !== null && val !== undefined && String(val).toLowerCase().includes(query);
      })
    );
  }, [data, search]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
  const safePage   = Math.min(Math.max(1, currentPage), totalPages);
  const paginated  = useMemo(() => {
    const start = (safePage - 1) * pageSize;
    return filtered.slice(start, start + pageSize);
  }, [filtered, safePage, pageSize]);

  const handleSearch         = (val: string) => { setSearch(val);     setCurrentPage(1); };
  const handlePageSizeChange = (size: number) => { setPageSize(size); setCurrentPage(1); };

  const handleResetAll = () => {
    setSearch('');
    onClearFilters?.();
  };

  return (
    <div className="contentCard">

      <TableToolbar
        pageSize={pageSize} pageSizeOptions={PAGE_SIZE_OPTIONS} onPageSizeChange={handlePageSizeChange}
        search={search} onSearch={handleSearch}
        right={
          <div className="flex items-center gap-3">
            <ExportToExcel data={data} columns={exportColumns} fileName="NFA_Home" sheetName="NFA Data" />
            <SearchInput value={search} onChange={handleSearch} />
          </div>
        }
      />

      {loading ? (
        <div className="tableWrapper">
          <table className="commonTable">
            <thead>
              <tr>
                {columns.map((col, index) => (
                  <th 
                    key={col.key} 
                    className={`tableHeader stickyHeader ${index === 0 ? 'stickyCol' : ''}`}
                  >
                    {col.label}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {Array(5).fill(0).map((_, i) => (
                <tr key={i} className="tableRow">
                  {columns.map((col, index) => (
                    <td 
                      key={col.key} 
                      className={`tableCell py-4 ${index === 0 ? 'stickyCol' : ''}`}
                    >
                      <div className="h-4 bg-slate-200/60 rounded animate-pulse w-full max-w-[120px] min-w-[50px]"></div>
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : filtered.length === 0 ? (
        <div className="flex flex-col items-center justify-center py-16 px-4 text-center bg-white rounded-2xl border border-dashed border-slate-200">
          <div className="w-14 h-14 rounded-full bg-slate-50 flex items-center justify-center mb-4 text-slate-400 border border-slate-100">
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
              <path strokeLinecap="round" strokeLinejoin="round" d="M20.25 7.5l-.625 10.632a2.25 2.25 0 01-2.247 2.118H6.622a2.25 2.25 0 01-2.247-2.118L3.75 7.5m8.25 3v6.75m0 0l-3-3m3 3l3-3M3.375 7.5h17.25c.621 0 1.125-.504 1.125-1.125v-1.5c0-.621-.504-1.125-1.125-1.125H3.375c-.621 0-1.125.504-1.125 1.125v1.5c0 .621.504 1.125 1.125 1.125z" />
            </svg>
          </div>
          <h3 className="text-slate-700 font-bold text-sm mb-1">No NFA Records Found</h3>
          <p className="text-slate-400 text-xs max-w-xs mb-4 leading-normal">We couldn't find any Note for Approval records matching your active filters or search query.</p>
          <button 
            onClick={handleResetAll}
            className="btn btnPrimary px-4 py-2 text-xs font-semibold rounded-xl"
            style={{ padding: '8px 16px' }}
          >
            Reset All Filters
          </button>
        </div>
      ) : (
        <CommonTable
          columns={columns}
          data={paginated}
          emptyMessage="No NFA records found"
          stickyFirstColumn={true}
          loading={false}
        />
      )}

      <Pagination
        page={safePage} total={totalPages} pageSize={pageSize}
        totalRecords={filtered.length} pageSizeOptions={PAGE_SIZE_OPTIONS}
        onPageChange={setCurrentPage} onPageSizeChange={handlePageSizeChange}
      />
    </div>
  );
};

export default HomeTable;
