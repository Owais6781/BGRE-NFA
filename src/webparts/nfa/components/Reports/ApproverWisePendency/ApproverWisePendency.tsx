import * as React                 from 'react';
import { useState }               from 'react';
import ApproverWisePendencyFilter from './ApproverWisePendencyFilter';
import ReportTable                from '../../../../../Share/Components/ReportTable';
import "../../../../../Share/styles/global.scss";

const COLUMNS = [
  { key: 'approverName',    label: 'Approver Name'      },
  { key: 'noOfPendingTask', label: 'No of Pending Task' },
  { key: 'status',          label: 'Status'             },
];

const ApproverWisePendency: React.FC = () => {
  const [data,    setData]    = useState<any[]>([]);
  const [loading, setLoading] = useState(false);

  const handleSearch = (filters: any) => {
    setLoading(true);
    setTimeout(() => { setData([]); setLoading(false); }, 500);
  };

  return (
    <div className="pageContainer">
      <div className="sectionTitleBar flex items-center justify-between">
        <span className="sectionTitle">Approver Wise Pendency Report</span>
      </div>

      <ApproverWisePendencyFilter onSearch={handleSearch} onReset={() => setData([])} />

      <div className="contentCard">
        <ReportTable columns={COLUMNS} data={data} loading={loading}
          emptyMessage="Please click on Get Pendency to load the report" />
      </div>
    </div>
  );
};

export default ApproverWisePendency;
