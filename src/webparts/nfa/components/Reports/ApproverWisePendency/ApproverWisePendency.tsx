import * as React                 from 'react';
import { useState }               from 'react';
import ApproverWisePendencyFilter from './ApproverWisePendencyFilter';
import ReportTable                from '../../../../../Share/Components/ReportTable';
import PageTitle                  from '../../../../../Share/Components/PageTitle';
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
      <PageTitle 
        title="Approver Wise Pendency Report" 
        subtitle="Track pending approvals and turnaround times across approvers." 
      />

      <ApproverWisePendencyFilter onSearch={handleSearch} onReset={() => setData([])} />

      <div className="contentCard">
        <ReportTable columns={COLUMNS} data={data} loading={loading}
          emptyMessage="Please click on Get Pendency to load the report" />
      </div>
    </div>
  );
};

export default ApproverWisePendency;
