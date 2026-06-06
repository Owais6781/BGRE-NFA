import * as React from 'react';
import { useState } from 'react';
import { ChevronUp, ChevronDown } from 'lucide-react';
import { CommonTable } from '../../../../../Share/Components/CommonTable';
import { Pagination }  from '../../../../../Share/Components/Pagination';
import TableToolbar    from '../../../../../Share/Components/TableToolbar';
import ExportToExcel   from '../../../../../Share/Components/ExportToExcel';
import "../../../../../Share/styles/global.scss";

interface ApprovedRejectedRecord {
  requestNo: string; function: string; project: string;
  createdBy: string; createdOn: string; status: string;
  level: number; approvedRejectedBy: string; approvedRejectedDate: string;
}

const PAGE_SIZE_OPTIONS = [10, 25, 50, 100];

const exportColumns = [
  { key: 'requestNo',            label: 'Request No.'                },
  { key: 'function',             label: 'Function'                   },
  { key: 'project',              label: 'Project'                    },
  { key: 'createdBy',            label: 'Created By'                 },
  { key: 'createdOn',            label: 'Created On'                 },
  { key: 'status',               label: 'Status'                     },
  { key: 'level',                label: 'Level'                      },
  { key: 'approvedRejectedBy',   label: 'Approved/Rejected By'       },
  { key: 'approvedRejectedDate', label: 'Approved/Rejected Datetime' },
];

const ApprovedRejectedByMe: React.FC = () => {
  const [data,        setData]        = useState<ApprovedRejectedRecord[]>([]);
  const [loading,     setLoading]     = useState(false);
  const [error,       setError]       = useState('');
  const [pageSize,    setPageSize]    = useState(10);
  const [search,      setSearch]      = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [isCollapsed, setIsCollapsed] = useState(false);

  const columns = [
    { key: 'requestNo', label: 'Request No.',
      render: (val: string) => <button className="tableLinkBtn" onClick={(e) => { e.stopPropagation(); console.log(val); }}>{val}</button> },
    { key: 'function',  label: 'Function'   },
    { key: 'project',   label: 'Project'    },
    { key: 'createdBy', label: 'Created By' },
    { key: 'createdOn', label: 'Created On' },
    { key: 'status',    label: 'Status',
      render: (val: string) => (
        <span className={`statusBadge ${val === 'Approved' ? 'statusApproved' : val === 'Rejected' ? 'statusRejected' : 'statusPending'}`}>{val}</span>
      ) },
    { key: 'level',                label: 'Level'                     },
    { key: 'approvedRejectedBy',   label: 'Approved/Rejected By'      },
    { key: 'approvedRejectedDate', label: 'Approved/Rejected Datetime' },
  ];

  const filtered   = data.filter((row) =>
    Object.keys(row).some((k) => String((row as any)[k]).toLowerCase().includes(search.toLowerCase()))
  );
  const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
  const paginated  = filtered.slice((currentPage - 1) * pageSize, currentPage * pageSize);

  const handleSearch         = (val: string) => { setSearch(val);     setCurrentPage(1); };
  const handlePageSizeChange = (size: number) => { setPageSize(size); setCurrentPage(1); };

  return (
    <div className="pageContainer">
      <div className="contentCard">

        <div className="sectionTitleBar flex items-center justify-between">
          <span className="sectionTitle">Approved / Rejected by Me</span>
          <button className="iconBtn" onClick={() => setIsCollapsed(c => !c)} title={isCollapsed ? 'Expand' : 'Collapse'}>
            {isCollapsed ? <ChevronDown size={16} /> : <ChevronUp size={16} />}
          </button>
        </div>

        {!isCollapsed && (
          <>
            {error && <div className="alertError">{error}</div>}

            <div className="flex justify-end mb-3">
              <ExportToExcel data={data} columns={exportColumns} fileName="Approved_Rejected_By_Me" sheetName="Approved Rejected" />
            </div>

            <TableToolbar
              pageSize={pageSize} pageSizeOptions={PAGE_SIZE_OPTIONS} onPageSizeChange={handlePageSizeChange}
              search={search} onSearch={handleSearch}
            />

            <CommonTable
              columns={columns}
              data={paginated}
              loading={loading}
              emptyMessage="No approved/rejected records found"
              stickyFirstColumn={true}
            />

            <Pagination
              page={currentPage} total={totalPages} pageSize={pageSize}
              totalRecords={filtered.length} pageSizeOptions={PAGE_SIZE_OPTIONS}
              onPageChange={setCurrentPage} onPageSizeChange={handlePageSizeChange}
            />
          </>
        )}

      </div>
    </div>
  );
};

export default ApprovedRejectedByMe;
