import * as React from 'react';
import { useState, useEffect } from 'react';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import DayWiseAgeingFilter from './DayWiseAgeingFilter';
import PageTitle from '../../../../../Share/Components/PageTitle';
import ExportToExcel from '../../../../../Share/Components/ExportToExcel';
import { ReportService, IDayWiseAgeingRecord } from '../../../../../Share/Services/ReportService';
import "../../../../../Share/styles/global.scss";

const COLUMNS = [
  { key: 'pendingDays', label: 'Pending Days' },
  { key: 'count', label: 'Count of Pending Approval Notes' },
];

interface IDayWiseAgeingProps {
  context?: WebPartContext;
  sp?: SPFI;
}

const DayWiseAgeing: React.FC<IDayWiseAgeingProps> = ({ context, sp }) => {
  const [data, setData] = useState<IDayWiseAgeingRecord[]>([]);
  const [loading, setLoading] = useState(false);
  const [hasSearched, setHasSearched] = useState(false);

  const grandTotal = data.reduce((sum, row) => sum + row.count, 0);

  const loadData = async (fromDate: string, toDate: string) => {
    if (!sp) return;
    setLoading(true);
    setHasSearched(true);
    try {
      const reportService = new ReportService(sp);
      const records = await reportService.getDayWiseAgeingReport(fromDate, toDate);
      setData(records);
    } catch (err) {
      console.error("DayWiseAgeing: Error loading data:", err);
      setData([]);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (sp) {
      const today = new Date().toISOString().split('T')[0];
      const minus16 = new Date(Date.now() - 16 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
      loadData(minus16, today);
    }
  }, [sp]);

  const handleSearch = (filters: { fromDate: string; toDate: string }) => {
    loadData(filters.fromDate, filters.toDate);
  };

  const handleReset = () => {
    setData([]);
    setHasSearched(false);
  };

  return (
    <div className="pageContainer">
      <PageTitle 
        title="Day Wise Ageing Report" 
        subtitle="Daily ageing analytics and duration breakdown for pending notes." 
      />

      <DayWiseAgeingFilter onSearch={handleSearch} onReset={handleReset} />

      <div className="contentCard">
        {data.length > 0 && (
          <div className="flex justify-end mb-3">
            <ExportToExcel 
              data={data} 
              columns={COLUMNS} 
              fileName="DayWise_Aging_Report" 
              sheetName="Day Wise Ageing" 
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
                    className="reportTableHeader text-center" 
                    style={{ textAlign: 'center' }}
                  >
                    {col.label}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {loading ? (
                <tr>
                  <td colSpan={2} className="reportTableCell text-center py-6" style={{ textAlign: 'center' }}>
                    <span className="inline-flex items-center gap-2 text-[#94a3b8]">
                      <span className="inline-block w-4 h-4 border-2 border-[#3b82f6] border-t-transparent rounded-full animate-spin" />
                      Loading…
                    </span>
                  </td>
                </tr>
              ) : data.length === 0 ? (
                <tr>
                  <td colSpan={2} className="reportTableCell text-center py-8" style={{ color: '#94a3b8', textAlign: 'center' }}>
                    {hasSearched ? "No Day wise ageing report found matching filters" : "Please click on Search to load the report"}
                  </td>
                </tr>
              ) : (
                <>
                  {data.map((row, i) => (
                    <tr key={i} className={i % 2 === 0 ? 'dayWiseRowEven' : 'dayWiseRowOdd'}>
                      <td className="reportTableCell text-center font-medium text-[#1e3a5f]" style={{ textAlign: 'center' }}>
                        {row.pendingDays}
                      </td>
                      <td className="reportTableCell text-center" style={{ textAlign: 'center' }}>
                        {row.count}
                      </td>
                    </tr>
                  ))}
                  <tr className="dayWiseGrandTotal font-bold">
                    <td className="reportTableCell text-center font-semibold" style={{ textAlign: 'center' }}>
                      Grand Total
                    </td>
                    <td className="reportTableCell text-center font-bold text-[#1e3a5f]" style={{ textAlign: 'center' }}>
                      {grandTotal}
                    </td>
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

export default DayWiseAgeing;
