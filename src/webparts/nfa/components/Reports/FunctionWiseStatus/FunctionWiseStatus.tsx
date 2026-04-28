import * as React              from 'react';
import { useState }            from 'react';
import FunctionWiseStatusFilter from './FunctionWiseStatusFilter';
import ReportTable              from '../../../../../Share/Components/ReportTable';
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

const FunctionWiseStatus: React.FC = () => {
  const [data,    setData]    = useState<any[]>([]);
  const [loading, setLoading] = useState(false);

  const handleSearch = (filters: any) => {
    setLoading(true);
    setTimeout(() => { setData([]); setLoading(false); }, 500);
  };

  return (
    <div className="pageContainer">
      <div className="sectionTitleBar flex items-center justify-between">
        <span className="sectionTitle">Function Wise Status Report</span>
      </div>

      <FunctionWiseStatusFilter onSearch={handleSearch} onReset={() => setData([])} />

      <div className="contentCard">
        <ReportTable columns={COLUMNS} data={data} loading={loading}
          emptyMessage="Please click on Search to load the report" />
      </div>
    </div>
  );
};

export default FunctionWiseStatus;
