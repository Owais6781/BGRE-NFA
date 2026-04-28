import * as React from 'react';
import { useState }    from 'react';
import { Inbox as InboxIcon, MessageSquare } from 'lucide-react';
import { CommonTable } from '../../../../../Share/Components/CommonTable';
import { Pagination }  from '../../../../../Share/Components/Pagination';
import TableToolbar    from '../../../../../Share/Components/TableToolbar';
import ExportToExcel   from '../../../../../Share/Components/ExportToExcel';
import "../../../../../Share/styles/global.scss";

interface ApproverInboxRecord {
  pendingTasks: string; subject: string; function: string;
  initiator: string; createdDate: string; pendingFrom: string;
  agingDays: number; level: number; project: string; amount: string;
}
interface QueryInboxRecord {
  pendingTasks: string; subject: string; function: string;
  queryBy: string; queryOn: string; pendingFrom: string;
  agingDays: number; level: number; project: string; amount: string;
}

const APPROVER_DATA: ApproverInboxRecord[] = [
  { pendingTasks: 'LPHRL/GYM/2026/Apr/6037', subject: 'test data',  function: 'GYM', initiator: 'NFA CORP', createdDate: '9/4/2026 10:55 AM',  pendingFrom: '9/4/2026 10:55 AM',  agingDays: 10, level: 1, project: 'LPHRL-Leela Palaces Hotels & Resorts Limited', amount: '1,000.00'        },
  { pendingTasks: 'LPHRL/GYM/2026/Feb/5576', subject: '',           function: 'GYM', initiator: 'NFA CORP', createdDate: '17/2/2026 11:17 AM', pendingFrom: '26/2/2026 3:35 PM',  agingDays: 52, level: 2, project: 'LPHRL-Leela Palaces Hotels & Resorts Limited', amount: '5,000.00'        },
  { pendingTasks: 'LPHRL/GYM/2026/Feb/5582', subject: 'Test Data',  function: 'GYM', initiator: 'NFA CORP', createdDate: '18/2/2026 11:55 AM', pendingFrom: '8/4/2026 10:46 PM',  agingDays: 11, level: 2, project: 'LPHRL-Leela Palaces Hotels & Resorts Limited', amount: '1,000.00'        },
  { pendingTasks: 'LPHRL/IT/2026/Apr/6025',  subject: 'Test data',  function: 'IT',  initiator: 'NFA CORP', createdDate: '8/4/2026 12:29 AM',  pendingFrom: '8/4/2026 10:28 PM',  agingDays: 11, level: 4, project: 'LPHRL-Leela Palaces Hotels & Resorts Limited', amount: '12,34,56,789.00' },
];
const QUERY_DATA: QueryInboxRecord[] = [];
const PAGE_SIZE_OPTIONS = [10, 25, 50, 100];

interface InboxTableProps<T> {
  title: string; icon: React.ReactNode; accentColor: string;
  data: T[]; columns: any[];
  exportColumns: { key: string; label: string }[];
  fileName: string;
}

function InboxTable<T>({ title, icon, accentColor, data, columns, exportColumns, fileName }: InboxTableProps<T>) {
  const [pageSize,    setPageSize]    = useState(10);
  const [search,      setSearch]      = useState('');
  const [currentPage, setCurrentPage] = useState(1);

  const filtered    = (data as any[]).filter((row) =>
    Object.keys(row).some((k) => String(row[k]).toLowerCase().includes(search.toLowerCase()))
  );
  const totalPages  = Math.max(1, Math.ceil(filtered.length / pageSize));
  const paginated   = filtered.slice((currentPage - 1) * pageSize, currentPage * pageSize);

  const handleSearch         = (val: string) => { setSearch(val);     setCurrentPage(1); };
  const handlePageSizeChange = (size: number) => { setPageSize(size); setCurrentPage(1); };

  return (
    <div className="contentCard mb-5">

      {/* Section header */}
      <div className="flex items-center justify-between mb-4 pb-3" style={{ borderBottom: '2px solid #e2e8f0' }}>
        <div className="flex items-center gap-2">
          <div className="flex items-center justify-center w-8 h-8 rounded-lg" style={{ background: accentColor + '20' }}>
            <span style={{ color: accentColor }}>{icon}</span>
          </div>
          <span className="sectionTitle">{title}</span>
          <span className="sidebarBadge" style={{ background: accentColor }}>{data.length}</span>
        </div>
        <ExportToExcel data={data as any[]} columns={exportColumns} fileName={fileName} sheetName={title} />
      </div>

      <TableToolbar
        pageSize={pageSize} pageSizeOptions={PAGE_SIZE_OPTIONS} onPageSizeChange={handlePageSizeChange}
        search={search} onSearch={handleSearch}
      />

      <CommonTable columns={columns} data={paginated} emptyMessage={`No ${title} records`} />

      <Pagination
        page={currentPage} total={totalPages} pageSize={pageSize}
        totalRecords={filtered.length} pageSizeOptions={PAGE_SIZE_OPTIONS}
        onPageChange={setCurrentPage} onPageSizeChange={handlePageSizeChange}
      />
    </div>
  );
}

const Inbox: React.FC = () => {
  const handleClick = (val: string) => console.log('Clicked:', val);

  const approverColumns = [
    { key: 'pendingTasks', label: 'Pending Tasks',
      render: (val: string) => <button className="tableLinkBtn" onClick={() => handleClick(val)}>{val}</button> },
    { key: 'subject',     label: 'Subject'     },
    { key: 'function',    label: 'Function'    },
    { key: 'initiator',   label: 'Initiator'   },
    { key: 'createdDate', label: 'Created Date' },
    { key: 'pendingFrom', label: 'Pending From' },
    { key: 'agingDays',   label: 'Aging (Days)' },
    { key: 'level',       label: 'Level'        },
    { key: 'project',     label: 'Project'      },
    { key: 'amount',      label: 'Amount'       },
  ];

  const queryColumns = [
    { key: 'pendingTasks', label: 'Pending Tasks',
      render: (val: string) => <button className="tableLinkBtn" onClick={() => handleClick(val)}>{val}</button> },
    { key: 'subject',     label: 'Subject'     },
    { key: 'function',    label: 'Function'    },
    { key: 'queryBy',     label: 'Query By'    },
    { key: 'queryOn',     label: 'Query On'    },
    { key: 'pendingFrom', label: 'Pending From' },
    { key: 'agingDays',   label: 'Aging (Days)' },
    { key: 'level',       label: 'Level'        },
    { key: 'project',     label: 'Project'      },
    { key: 'amount',      label: 'Amount'       },
  ];

  return (
    <div className="pageContainer">

      <div className="sectionTitleBar flex items-center justify-between">
        <span className="sectionTitle">My Inbox</span>
      </div>

      <InboxTable
        title="Approver Inbox"
        icon={<InboxIcon size={16} />}
        accentColor="#3b82f6"
        data={APPROVER_DATA}
        columns={approverColumns}
        exportColumns={approverColumns.map(c => ({ key: c.key, label: c.label }))}
        fileName="Approver_Inbox"
      />

      <InboxTable
        title="Query Inbox"
        icon={<MessageSquare size={16} />}
        accentColor="#0891b2"
        data={QUERY_DATA}
        columns={queryColumns}
        exportColumns={queryColumns.map(c => ({ key: c.key, label: c.label }))}
        fileName="Query_Inbox"
      />

    </div>
  );
};

export default Inbox;
