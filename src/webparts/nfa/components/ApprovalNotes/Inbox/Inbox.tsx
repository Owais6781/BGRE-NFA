import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import { Inbox as InboxIcon, MessageSquare } from 'lucide-react';
import { useNavigate } from 'react-router-dom';
import { CommonTable } from '../../../../../Share/Components/CommonTable';
import { Pagination }  from '../../../../../Share/Components/Pagination';
import TableToolbar    from '../../../../../Share/Components/TableToolbar';
import ExportToExcel   from '../../../../../Share/Components/ExportToExcel';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { NFAService } from '../../../../../Share/Services/NFAService';
import "../../../../../Share/styles/global.scss";

interface IInboxProps {
  context: WebPartContext;
  sp: SPFI;
}

const PAGE_SIZE_OPTIONS = [5, 10, 25, 50, 100];

interface InboxTableProps<T> {
  title: string; 
  icon: React.ReactNode; 
  accentColor: string;
  data: T[]; 
  columns: any[];
  exportColumns: { key: string; label: string }[];
  fileName: string;
  loading: boolean;
  onRowClick: (row: T) => void;
}

function InboxTable<T>({ title, icon, accentColor, data, columns, exportColumns, fileName, loading, onRowClick }: InboxTableProps<T>) {
  const [pageSize,    setPageSize]    = useState(5);
  const [search,      setSearch]      = useState('');
  const [currentPage, setCurrentPage] = useState(1);

  const filtered = useMemo(() => {
    return (data as any[]).filter((row) =>
      Object.keys(row).some((k) => String(row[k]).toLowerCase().includes(search.toLowerCase()))
    );
  }, [data, search]);

  const totalPages  = Math.max(1, Math.ceil(filtered.length / pageSize));
  const paginated   = useMemo(() => {
    return filtered.slice((currentPage - 1) * pageSize, currentPage * pageSize);
  }, [filtered, currentPage, pageSize]);

  const handleSearch         = (val: string) => { setSearch(val);     setCurrentPage(1); };
  const handlePageSizeChange = (size: number) => { setPageSize(size); setCurrentPage(1); };

  return (
    <div className="contentCard mb-8">
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

      <CommonTable
        columns={columns}
        data={paginated}
        loading={loading}
        emptyMessage={`No ${title} records`}
        stickyFirstColumn={true}
        onRowClick={onRowClick}
      />

      <Pagination
        page={currentPage} total={totalPages} pageSize={pageSize}
        totalRecords={filtered.length} pageSizeOptions={PAGE_SIZE_OPTIONS}
        onPageChange={setCurrentPage} onPageSizeChange={handlePageSizeChange}
      />
    </div>
  );
}

const Inbox: React.FC<IInboxProps> = ({ context, sp }) => {
  const navigate = useNavigate();
  const [inboxData, setInboxData] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const fetchInbox = async () => {
      const service = new NFAService(context);
      try {
        const data = await service.getInboxData();
        setInboxData(data);
      } catch (err) {
        console.error("Failed to load inbox data:", err);
      } finally {
        setLoading(false);
      }
    };
    fetchInbox();
  }, []);

  const pendingInbox = useMemo(() => inboxData.filter(item => (item.status || "").toLowerCase() === 'pending'), [inboxData]);
  const queryInbox = useMemo(() => inboxData.filter(item => (item.status || "").toLowerCase() === 'query'), [inboxData]);

  const approverColumns = [
    { key: 'requestNo', label: 'Pending Tasks',
      render: (val: string) => <span className="tableLinkBtn">{val}</span> },
    { key: 'function',    label: 'Function'    },
    { key: 'project',     label: 'Project'      },
    { key: 'initiator',   label: 'Initiator'   },
    { key: 'createdDate', label: 'Created Date' },
    { key: 'agingDays',   label: 'Aging (Days)' },
    { key: 'level',       label: 'Level'        },
    { key: 'amount',      label: 'Amount'       },
  ];

  const queryColumns = [
    { key: 'requestNo', label: 'Pending Tasks',
      render: (val: string) => <span className="tableLinkBtn">{val}</span> },
    { key: 'function',    label: 'Function'    },
    { key: 'project',     label: 'Project'      },
    { key: 'initiator',   label: 'Query By'    },
    { key: 'createdDate', label: 'Query On'    },
    { key: 'agingDays',   label: 'Aging (Days)' },
    { key: 'level',       label: 'Level'        },
    { key: 'amount',      label: 'Amount'       },
  ];

  const handleRecordClick = (row: any) => {
    console.log("Inbox: Navigating for ID:", row.id);
    if (row.id) {
      if ((row.status || "").toLowerCase() === 'query') {
        navigate(`/query-respond/${row.id}`);
      } else {
        navigate(`/approver-page/${row.id}`);
      }
    } else {
      console.warn("Inbox: Row ID is missing", row);
    }
  };

  return (
    <div className="pageContainer">
      <div className="sectionTitleBar flex items-center justify-between">
        <span className="sectionTitle">My Inbox</span>
      </div>

      <InboxTable
        title="Approver Inbox"
        icon={<InboxIcon size={16} />}
        accentColor="#3b82f6"
        data={pendingInbox}
        loading={loading}
        columns={approverColumns}
        exportColumns={approverColumns.map(c => ({ key: c.key, label: c.label }))}
        fileName="Approver_Inbox"
        onRowClick={handleRecordClick}
      />

      <InboxTable
        title="Query Inbox"
        icon={<MessageSquare size={16} />}
        accentColor="#0891b2"
        data={queryInbox}
        loading={loading}
        columns={queryColumns}
        exportColumns={queryColumns.map(c => ({ key: c.key, label: c.label }))}
        fileName="Query_Inbox"
        onRowClick={handleRecordClick}
      />
    </div>
  );
};

export default Inbox;
