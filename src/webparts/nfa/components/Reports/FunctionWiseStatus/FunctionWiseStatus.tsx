import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import { SlidersHorizontal } from 'lucide-react';
import { CommonTable } from '../../../../../Share/Components/CommonTable';
import { Pagination }  from '../../../../../Share/Components/Pagination';
import TableToolbar    from '../../../../../Share/Components/TableToolbar';
import ExportToExcel   from '../../../../../Share/Components/ExportToExcel';
import { SearchInput } from '../../../../../Share/Components/SearchInput';
import { Skeleton }    from 'primereact/skeleton';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { ReportService, IFunctionStatusRecord } from '../../../../../Share/Services/ReportService';
import { DropdownService, IDropdownOption } from '../../../../../Share/Services/DropdownService';
import FunctionWiseStatusFilter from './FunctionWiseStatusFilter';
import "../../../../../Share/styles/global.scss";

const COLUMNS = [
  { key: 'functionName', label: 'Function Name' },
  { key: 'approved',     label: 'Approved'       },
  { key: 'pending',      label: 'Pending'        },
  { key: 'query',        label: 'Query'          },
  { key: 'delegated',    label: 'Delegated'      },
  { key: 'rejected',     label: 'Rejected'       },
  { key: 'grandTotal',   label: 'Grand Total'    },
];

const PAGE_SIZE_OPTIONS = [10, 25, 50, 100];

interface IFunctionWiseStatusProps {
  context: WebPartContext;
  sp: SPFI;
}

const FunctionWiseStatus: React.FC<IFunctionWiseStatusProps> = ({ context, sp }) => {
  const [data, setData] = useState<IFunctionStatusRecord[]>([]);
  const [functionOptions, setFunctionOptions] = useState<IDropdownOption[]>([]);
  const [filters, setFilters] = useState<any>(null);
  const [loading, setLoading] = useState(false);
  const [pageSize, setPageSize] = useState(10);
  const [search, setSearch] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [isCollapsed, setIsCollapsed] = useState(false);

  useEffect(() => {
    const initPage = async () => {
      setLoading(true);
      try {
        const dropdownService = new DropdownService(context);
        const funcs = await dropdownService.getFunctionMaster();
        setFunctionOptions(funcs);

        // Load default report data on page mount (last 21 days)
        const todayStr = new Date().toISOString().split('T')[0];
        const minus21Str = new Date(Date.now() - 21 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];

        const reportService = new ReportService(sp);
        const records = await reportService.getFunctionWiseStatus(minus21Str, todayStr, "");
        setData(records);
        setFilters({ fromDate: minus21Str, toDate: todayStr, functionName: "" });
      } catch (err) {
        console.error("FunctionWiseStatus: Error initializing page data:", err);
      } finally {
        setLoading(false);
      }
    };
    initPage();
  }, [context]);

  const handleSearch = async (filterVals: any) => {
    setLoading(true);
    setFilters(filterVals);
    try {
      const reportService = new ReportService(sp);
      const records = await reportService.getFunctionWiseStatus(
        filterVals.fromDate,
        filterVals.toDate,
        filterVals.functionName
      );
      setData(records);
      setCurrentPage(1);
    } catch (err) {
      console.error("FunctionWiseStatus: Error loading report data:", err);
    } finally {
      setLoading(false);
    }
  };

  const handleReset = () => {
    setData([]);
    setFilters(null);
    setSearch('');
    setCurrentPage(1);
  };

  const displayData = useMemo(() => {
    if (!search) return data;
    return data.filter(row => 
      row.functionName.toLowerCase().includes(search.toLowerCase())
    );
  }, [data, search]);

  const totalPages = Math.ceil(displayData.length / pageSize) || 1;

  const paginated = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return displayData.slice(start, start + pageSize);
  }, [displayData, currentPage, pageSize]);

  return (
    <div className="pageContainer">
      <div className="sectionTitleBar flex items-center justify-between">
        <span className="sectionTitle">Function Wise Status Report</span>
      </div>

      {/* Filter Card */}
      <div className="mb-4">
        <FunctionWiseStatusFilter 
          onSearch={handleSearch} 
          onReset={handleReset} 
          functionOptions={functionOptions}
        />
      </div>

      <div className="contentCard">
        <TableToolbar
          pageSize={pageSize}
          pageSizeOptions={PAGE_SIZE_OPTIONS}
          onPageSizeChange={(size) => { setPageSize(size); setCurrentPage(1); }}
          right={
            <div className="flex items-center gap-3">
              <SearchInput value={search} onChange={(val) => { setSearch(val); setCurrentPage(1); }} />
              {displayData.length > 0 && (
                <ExportToExcel 
                  data={displayData} 
                  columns={COLUMNS} 
                  fileName="Function_Wise_Status_Report" 
                  sheetName="Status Report" 
                />
              )}
            </div>
          }
        />

        {loading ? (
          <div className="w-full mt-3 border border-slate-200 rounded-xl overflow-hidden">
            {/* Header Skeleton */}
            <div className="flex bg-slate-50 border-b border-slate-200 p-4">
              {COLUMNS.map((col) => (
                <div key={col.key} className="flex-1">
                  <Skeleton width="60%" height="1.2rem" />
                </div>
              ))}
            </div>
            {/* Rows Skeleton */}
            {Array.from({ length: pageSize || 5 }).map((_, i) => (
              <div key={i} className="flex border-b border-slate-100 p-4 items-center">
                {COLUMNS.map((col) => (
                  <div key={col.key} className="flex-1">
                    <Skeleton width="80%" height="1rem" />
                  </div>
                ))}
              </div>
            ))}
          </div>
        ) : (
          <CommonTable
            columns={COLUMNS}
            data={paginated}
            emptyMessage={filters ? "No records found matching filters" : "Please click on Search to load the report"}
            stickyFirstColumn={true}
          />
        )}

        {displayData.length > 0 && (
          <Pagination
            page={currentPage}
            total={totalPages}
            pageSize={pageSize}
            totalRecords={displayData.length}
            pageSizeOptions={PAGE_SIZE_OPTIONS}
            onPageChange={setCurrentPage}
            onPageSizeChange={(size) => { setPageSize(size); setCurrentPage(1); }}
          />
        )}
      </div>
    </div>
  );
};

export default FunctionWiseStatus;
