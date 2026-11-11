import * as React from 'react';
import { useState, useMemo, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
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
  { key: 'pendingAt',               label: 'Pending At'                                           },
  { key: 'approvalNotes',           label: 'Approval Notes'                                       },
  { key: 'status',                  label: 'Status'                                               },
  { key: 'function',                label: 'Function'                                             },
  { key: 'amount',                  label: 'Amount'                                               },
  { key: 'ageingFromCreated',       label: 'Ageing (from Created Date)',   align: 'center' as const },
  { key: 'ageingFromLastApprover',  label: 'Ageing (from Last Approver)',  align: 'center' as const },
];

const PAGE_SIZE_OPTIONS = [10, 25, 50, 100];

interface IAgeingReportProps {
  context: WebPartContext;
  sp: SPFI;
}

const AgeingReport: React.FC<IAgeingReportProps> = ({ context, sp }) => {
  const navigate = useNavigate();
  const [data, setData] = useState<IAgeingStatusRecord[]>([]);
  const [filters, setFilters] = useState<any>(null);
  const [loading, setLoading] = useState(false);
  const [pageSize, setPageSize] = useState(10);
  const [search, setSearch] = useState('');
  const [currentPage, setCurrentPage] = useState(1);

  const handleRecordClick = async (row: IAgeingStatusRecord) => {
    if (row.id) {
      navigate(`/view-nfa/${row.id}`);
      return;
    }
    if (row.approvalNotes && sp) {
      try {
        const found = await sp.web.lists.getByTitle("DigitalApprovalNotes").items
          .filter(`Title eq '${row.approvalNotes.trim()}'`)
          .select("Id")
          .top(1)();
        if (found && found.length > 0) {
          navigate(`/view-nfa/${found[0].Id}`);
          return;
        }
      } catch (err) {
        console.error("AgeingReport: Error finding NFA by Title:", err);
      }
    }
  };

  const columns = useMemo(() => [
    { key: 'pendingAt',               label: 'Pending At'                                           },
    {
      key: 'approvalNotes',
      label: 'Approval Notes',
      render: (val: string, row: IAgeingStatusRecord) => (
        <button
          type="button"
          className="tableLinkBtn"
          onClick={(e) => {
            e.preventDefault();
            e.stopPropagation();
            void handleRecordClick(row);
          }}
        >
          {val || '—'}
        </button>
      ),
    },
    { key: 'status',                  label: 'Status'                                               },
    { key: 'function',                label: 'Function'                                             },
    { key: 'amount',                  label: 'Amount'                                               },
    { key: 'ageingFromCreated',       label: 'Ageing (from Created Date)',   align: 'center' as const },
    { key: 'ageingFromLastApprover',  label: 'Ageing (from Last Approver)',  align: 'center' as const },
  ], [navigate, sp]);

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

  useEffect(() => {
    if (sp) {
      const today = new Date().toISOString().split('T')[0];
      const minus21 = new Date(Date.now() - 21 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
      void handleSearch({ fromDate: minus21, toDate: today });
    }
  }, [sp]);

  const handleReset = () => {
    setSearch('');
    setCurrentPage(1);
    const today = new Date().toISOString().split('T')[0];
    const minus21 = new Date(Date.now() - 21 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
    void handleSearch({ fromDate: minus21, toDate: today });
  };

  const displayData = useMemo(() => {
    if (!search) return data;
    return data.filter(row => 
      Object.keys(row).some(k => 
        String((row as any)[k]).toLowerCase().includes(search.toLowerCase())
      )
    );
  }, [data, search]);

  const totalPages = Math.max(1, Math.ceil(displayData.length / pageSize));
  const safePage   = Math.min(Math.max(1, currentPage), totalPages);

  const paginated = useMemo(() => {
    const start = (safePage - 1) * pageSize;
    return displayData.slice(start, start + pageSize);
  }, [displayData, safePage, pageSize]);

  const totals = useMemo(() => {
    let sumAmount = 0;
    let sumAgeingFromCreated = 0;
    let sumAgeingFromLastApprover = 0;

    displayData.forEach(row => {
      const cleanAmt = String(row.amount || '').replace(/,/g, '').trim();
      const amtNum = parseFloat(cleanAmt);
      if (!isNaN(amtNum)) {
        sumAmount += amtNum;
      }

      const ageingCreatedNum = Number(row.ageingFromCreated);
      if (!isNaN(ageingCreatedNum)) {
        sumAgeingFromCreated += ageingCreatedNum;
      }

      const ageingLastNum = Number(row.ageingFromLastApprover);
      if (!isNaN(ageingLastNum)) {
        sumAgeingFromLastApprover += ageingLastNum;
      }
    });

    return {
      amount: sumAmount,
      ageingFromCreated: sumAgeingFromCreated,
      ageingFromLastApprover: sumAgeingFromLastApprover
    };
  }, [displayData]);

  const tableFooter = useMemo(() => {
    if (displayData.length === 0) return null;
    return (
      <tr className="dayWiseGrandTotal font-bold">
        <td className="tableCell"></td>
        <td className="tableCell"></td>
        <td className="tableCell"></td>
        <td className="tableCell font-bold text-[#1e3a5f]" style={{ fontWeight: 700 }}>
          Total:
        </td>
        <td className="tableCell font-bold text-[#1e3a5f]" style={{ fontWeight: 700 }}>
          {totals.amount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
        </td>
        <td className="tableCell font-bold text-[#1e3a5f] text-center" style={{ fontWeight: 700, textAlign: 'center' }}>
          {totals.ageingFromCreated}
        </td>
        <td className="tableCell font-bold text-[#1e3a5f] text-center" style={{ fontWeight: 700, textAlign: 'center' }}>
          {totals.ageingFromLastApprover}
        </td>
      </tr>
    );
  }, [displayData.length, totals]);

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
          columns={columns}
          data={paginated}
          loading={loading}
          emptyMessage={filters ? "No records found matching filters" : "Please click on Search to load the report"}
          stickyFirstColumn={true}
          footer={tableFooter}
        />

        {displayData.length > 0 && (
          <Pagination
            page={safePage}
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
