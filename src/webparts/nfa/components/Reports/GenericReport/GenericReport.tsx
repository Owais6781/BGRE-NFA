import * as React          from 'react';
import { useState }        from 'react';
import GenericReportFilter from './GenericReportFilter';
import ReportTable         from '../../../../../Share/Components/ReportTable';
import ExportToExcel       from '../../../../../Share/Components/ExportToExcel';
import TableToolbar        from '../../../../../Share/Components/TableToolbar';
import PageTitle           from '../../../../../Share/Components/PageTitle';
import "../../../../../Share/styles/global.scss";

const COLUMNS = [
  { key: 'nfaNo',          label: 'NFA No.'             },
  { key: 'subjectMatter',  label: 'Subject / Matter'    },
  { key: 'function',       label: 'Function'             },
  { key: 'project',        label: 'Project'              },
  { key: 'contractor',     label: 'Contractor / Vendor'  },
  { key: 'packageAwarded', label: 'Package Awarded Cost' },
  { key: 'budgetedAmount', label: 'Budgeted Amount'      },
  { key: 'status',         label: 'Status'               },
  { key: 'createdBy',      label: 'Created By'           },
  { key: 'createdOn',      label: 'Created On'           },
];

const GenericReport: React.FC = () => {
  const [data,    setData]    = useState<any[]>([]);
  const [loading, setLoading] = useState(false);

  const handleSearch = (filters: any) => {
    setLoading(true);
    setTimeout(() => { setData([]); setLoading(false); }, 500);
  };

  return (
    <div className="pageContainer">
      <PageTitle 
        title="NFA Generic Report" 
        subtitle="Custom filterable report for Note for Approval records." 
      />

      <GenericReportFilter onSearch={handleSearch} onReset={() => setData([])} />

      <div className="contentCard">
        <TableToolbar
          right={<ExportToExcel data={data} columns={COLUMNS} fileName="NFA_Generic_Report" sheetName="Generic Report" />}
        />
        <ReportTable columns={COLUMNS} data={data} loading={loading}
          emptyMessage="Please click on Search to load the report" />
      </div>
    </div>
  );
};

export default GenericReport;
