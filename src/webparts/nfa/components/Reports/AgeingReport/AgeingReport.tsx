import * as React from 'react';
import { useState, useMemo } from 'react';
import { CommonTable } from '../../../../../Share/Components/CommonTable';
import { Pagination }  from '../../../../../Share/Components/Pagination';
import TableToolbar    from '../../../../../Share/Components/TableToolbar';
import ExportToExcel   from '../../../../../Share/Components/ExportToExcel';
import { SearchInput } from '../../../../../Share/Components/SearchInput';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { ReportService, IAgeingStatusRecord } from '../../../../../Share/Services/ReportService';
import AgeingReportFilter from './AgeingReportFilter';
import PageTitle from '../../../../../Share/Components/PageTitle';
import "../../../../../Share/styles/global.scss";

const COLUMNS = [
  { key: 'pendingAt',               label: 'Pending At'                   },
  { key: 'approvalNotes',           label: 'Approval Notes'               },
  { key: 'status',                  label: 'Status'                       },
  { key: 'function',                label: 'Function'                     },
  { key: 'amount',                  label: 'Amount'                       },
  { key: 'ageingFromCreated',       label: 'Ageing (from Created Date)'   },
  { key: 'ageingFromLastApprover',  label: 'Ageing (from Last Approver)'  },
];

const PAGE_SIZE_OPTIONS = [10, 25, 50, 100];

interface IAgeingReportProps {
  context: WebPartContext;
  sp: SPFI;
}

const AgeingReport: React.FC<IAgeingReportProps> = ({ context, sp }) => {
  const [data, setData] = useState<IAgeingStatusRecord[]>([]);
  const [filters, setFilters] = useState<any>(null);
  const [loading, setLoading] = useState(false);
  const [pageSize, setPageSize] = useState(10);
  const [search, setSearch] = useState('');
  const [currentPage, setCurrentPage] = useState(1);

  const handleSearch = async (filterVals: any) => {
    setLoading(true);
    setFilters(filterVals);
    try {
      const reportService = new ReportService(sp);
      const records = await reportService.getAgeingReport(
        filterVals.fromDate,
        filterVals.toDate
      );
      setData(records);
      setCurrentPage(1);
    } catch (err) {
      console.error("AgeingReport: Error loading ageing report data:", err);
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
      Object.keys(row).some(k => 
        String((row as any)[k]).toLowerCase().includes(search.toLowerCase())
      )
    );
  }, [data, search]);

  const totalPages = Math.ceil(displayData.length / pageSize) || 1;

  const paginated = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return displayData.slice(start, start + pageSize);
  }, [displayData, currentPage, pageSize]);

  return (
    <div className="pageContainer">
      <PageTitle 
        title="Ageing Report" 
        subtitle="Monitor ageing brackets for open approval requests." 
      />

      {/* Filter Card */}
      <div className="mb-4">
        <AgeingReportFilter 
          onSearch={handleSearch} 
          onReset={handleReset} 
        />
      </div>

      <div className="contentCard">
        <TableToolbar
          pageSize={pageSize}
          pageSizeOptions={PAGE_SIZE_OPTIONS}
          onPageSizeChange={(size) => { setPageSize(size); setCurrentPage(1); }}
          search={search}
          onSearch={(val) => { setSearch(val); setCurrentPage(1); }}
          right={
            <div className="flex items-center gap-3">
              {data.length > 0 && (
                <ExportToExcel 
                  data={displayData} 
                  columns={COLUMNS} 
                  fileName="Ageing_Report" 
                  sheetName="Ageing Report" 
                />
              )}
              <SearchInput value={search} onChange={(val) => { setSearch(val); setCurrentPage(1); }} />
            </div>
          }
        />

        <CommonTable
          columns={COLUMNS}
          data={paginated}
          loading={loading}
          emptyMessage={filters ? "No records found matching filters" : "Please click on Search to load the report"}
          stickyFirstColumn={true}
        />

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

export default AgeingReport;
