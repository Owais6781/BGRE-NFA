import * as React       from 'react';
import { useState }     from 'react';
import AgeingReportFilter from './AgeingReportFilter';
import ReportTable        from '../../../../../Share/Components/ReportTable';
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

const AgeingReport: React.FC = () => {
  const [data,    setData]    = useState<any[]>([]);
  const [loading, setLoading] = useState(false);

  const handleSearch = (filters: any) => {
    setLoading(true);
    setTimeout(() => { setData([]); setLoading(false); }, 500);
  };

  return (
    <div className="pageContainer">
      <div className="sectionTitleBar flex items-center justify-between">
        <span className="sectionTitle">Ageing Report</span>
      </div>

      <AgeingReportFilter onSearch={handleSearch} onReset={() => setData([])} />

      <div className="contentCard">
        <ReportTable columns={COLUMNS} data={data} loading={loading}
          emptyMessage="Please click on Search to load the report" />
      </div>
    </div>
  );
};

export default AgeingReport;
