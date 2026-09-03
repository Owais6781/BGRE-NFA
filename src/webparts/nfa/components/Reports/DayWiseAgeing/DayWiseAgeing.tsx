import * as React          from 'react';
import { useState }        from 'react';
import DayWiseAgeingFilter from './DayWiseAgeingFilter';
import PageTitle          from '../../../../../Share/Components/PageTitle';
import "../../../../../Share/styles/global.scss";

const COLUMNS = [
  { key: 'pendingDays', label: 'Pending Days'                     },
  { key: 'count',       label: 'Count of Pending Approval Notes'  },
];

const DayWiseAgeing: React.FC = () => {
  const [data,    setData]    = useState<{ pendingDays: number; count: number }[]>([]);
  const [loading, setLoading] = useState(false);

  const grandTotal = data.reduce((sum, row) => sum + row.count, 0);

  const handleSearch = (filters: any) => {
    setLoading(true);
    setTimeout(() => { setData([]); setLoading(false); }, 500);
  };

  return (
    <div className="pageContainer">
      <PageTitle 
        title="Day Wise Ageing Report" 
        subtitle="Daily ageing analytics and duration breakdown for pending notes." 
      />

      <DayWiseAgeingFilter onSearch={handleSearch} onReset={() => setData([])} />

      <div className="contentCard">
        <div className="tableWrapper">
          <table className="commonTable">
            <thead>
              <tr>
                {COLUMNS.map((col) => (
                  <th key={col.key} className="reportTableHeader text-center">{col.label}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {loading ? (
                <tr>
                  <td colSpan={2} className="reportTableCell text-center py-6">
                    <span className="inline-flex items-center gap-2 text-[#94a3b8]">
                      <span className="inline-block w-4 h-4 border-2 border-[#3b82f6] border-t-transparent rounded-full animate-spin" />
                      Loading…
                    </span>
                  </td>
                </tr>
              ) : data.length === 0 ? (
                <tr>
                  <td colSpan={2} className="reportTableCell text-center py-8" style={{ color: '#94a3b8' }}>
                    Please click on Search to load the report
                  </td>
                </tr>
              ) : (
                <>
                  {data.map((row, i) => (
                    <tr key={i} className={i % 2 === 0 ? 'dayWiseRowEven' : 'dayWiseRowOdd'}>
                      <td className="reportTableCell text-center">{row.pendingDays}</td>
                      <td className="reportTableCell text-center">{row.count}</td>
                    </tr>
                  ))}
                  <tr className="dayWiseGrandTotal">
                    <td className="reportTableCell text-right font-semibold" style={{ paddingRight: 16 }}>Grand Total:</td>
                    <td className="reportTableCell text-center font-semibold">{grandTotal}</td>
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
