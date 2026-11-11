import * as React from 'react';
import { useState, useEffect } from 'react';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import ApproverWisePendencyFilter, { FilterValues } from './ApproverWisePendencyFilter';
import ExportToExcel from '../../../../../Share/Components/ExportToExcel';
import PageTitle from '../../../../../Share/Components/PageTitle';
import { ReportService, IApproverPendencyRecord } from '../../../../../Share/Services/ReportService';
import "../../../../../Share/styles/global.scss";

const COLUMNS = [
  { key: 'approverName',    label: 'Approver Name'      },
  { key: 'noOfPendingTask', label: 'No of Pending Task' },
  { key: 'status',          label: 'Status'             },
];

interface IApproverWisePendencyProps {
  context?: WebPartContext;
  sp?: SPFI;
}

const ApproverWisePendency: React.FC<IApproverWisePendencyProps> = ({ context, sp }) => {
  const [data, setData] = useState<IApproverPendencyRecord[]>([]);
  const [loading, setLoading] = useState(false);
  const [hasSearched, setHasSearched] = useState(false);

  const totalTasks = data.reduce((sum, row) => sum + row.noOfPendingTask, 0);

  const loadData = async (fromDate: string, toDate: string, status: string) => {
    if (!sp) return;
    setLoading(true);
    setHasSearched(true);
    try {
      const reportService = new ReportService(sp);
      const records = await reportService.getApproverWisePendency({
        fromDate,
        toDate,
        status,
      });
      setData(records);
    } catch (err) {
      console.error("ApproverWisePendency: Error loading data:", err);
      setData([]);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (sp) {
      const tomorrow = new Date(Date.now() + 1 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
      const minus10 = new Date(Date.now() - 10 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
      loadData(minus10, tomorrow, 'All');
    }
  }, [sp]);

  const handleSearch = (filters: FilterValues) => {
    loadData(filters.fromDate, filters.toDate, filters.status);
  };

  const handleReset = () => {
    setData([]);
    setHasSearched(false);
  };

  return (
    <div className="pageContainer">
      <PageTitle 
        title="Approver Wise Pendency Report" 
        subtitle="Track pending approvals and turnaround times across approvers." 
      />

      <ApproverWisePendencyFilter onSearch={handleSearch} onReset={handleReset} />

      <div className="contentCard">
        {data.length > 0 && (
          <div className="flex justify-end mb-3">
            <ExportToExcel 
              data={data} 
              columns={COLUMNS} 
              fileName="Approvers_Wise_Pendency_Report" 
              sheetName="Pendency Report" 
            />
          </div>
        )}

        <div className="tableWrapper">
          <table className="commonTable" id="tblDataMain">
            <thead>
              <tr>
                {COLUMNS.map((col) => (
                  <th 
                    key={col.key} 
                    className="reportTableHeader" 
                    style={{ textAlign: col.key === 'approverName' ? 'left' : 'center' }}
                  >
                    {col.label}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {loading ? (
                <tr>
                  <td colSpan={3} className="reportTableCell text-center py-6" style={{ textAlign: 'center' }}>
                    <span className="inline-flex items-center gap-2 text-[#94a3b8]">
                      <span className="inline-block w-4 h-4 border-2 border-[#3b82f6] border-t-transparent rounded-full animate-spin" />
                      Loading…
                    </span>
                  </td>
                </tr>
              ) : data.length === 0 ? (
                <tr>
                  <td colSpan={3} className="reportTableCell text-center py-8" style={{ color: '#94a3b8', textAlign: 'center' }}>
                    {hasSearched ? "Please refine your search to Get Pendency report" : "Please click on Get Pendency to load the report"}
                  </td>
                </tr>
              ) : (
                <>
                  {data.map((row, i) => (
                    <tr key={i} className={i % 2 === 0 ? 'dayWiseRowEven' : 'dayWiseRowOdd'}>
                      <td className="reportTableCell font-medium text-[#1e3a5f]">{row.approverName}</td>
                      <td className="reportTableCell text-center" style={{ textAlign: 'center' }}>{row.noOfPendingTask}</td>
                      <td className="reportTableCell text-center" style={{ textAlign: 'center' }}>{row.status}</td>
                    </tr>
                  ))}
                  <tr className="dayWiseGrandTotal font-bold">
                    <td className="reportTableCell font-semibold">Total :</td>
                    <td className="reportTableCell text-center font-bold text-[#1e3a5f]" style={{ textAlign: 'center' }}>
                      {totalTasks}
                    </td>
                    <td className="reportTableCell"></td>
                  </tr>
                </>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

export default ApproverWisePendency;
