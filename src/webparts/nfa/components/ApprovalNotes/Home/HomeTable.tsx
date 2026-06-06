import * as React from 'react';
import { useState }    from 'react';
import { CommonTable } from '../../../../../Share/Components/CommonTable';
import { Pagination }  from '../../../../../Share/Components/Pagination';
import TableToolbar    from '../../../../../Share/Components/TableToolbar';
import ExportToExcel   from '../../../../../Share/Components/ExportToExcel';
import { SearchInput } from '../../../../../Share/Components/SearchInput';
import "../../../../../Share/styles/global.scss";

interface NFARecord {
  id: number;
  requestNo: string; 
  function: string; 
  project: string;
  createdOn: string; // CDate
  status: string; 
  pendingAt: string;
  level: number;
}

interface HomeTableProps {
  data: NFARecord[];
  onRowClick?: (record: NFARecord) => void;
  loading?: boolean;
}

const PAGE_SIZE_OPTIONS = [10, 25, 50, 100, 500];

const exportColumns = [
  { key: 'requestNo', label: 'Request No.' },
  { key: 'function',  label: 'Function'    },
  { key: 'project',   label: 'Project'     },
  { key: 'createdOn', label: 'CDate'       },
  { key: 'status',    label: 'Status'      },
  { key: 'pendingAt', label: 'Pending At'  },
  { key: 'level',     label: 'Level'       },
];

const HomeTable: React.FC<HomeTableProps> = ({ data, onRowClick, loading = false }) => {
  const [pageSize,     setPageSize]    = useState(10);
  const [search,       setSearch]      = useState('');
  const [currentPage,  setCurrentPage] = useState(1);

  const columns = [
    {
      key: 'requestNo', label: 'Title / Request No.',
      render: (val: string, row: NFARecord) => (
        <button className="tableLinkBtn" onClick={(e) => { e.stopPropagation(); onRowClick?.(row); }}>
          {val}
        </button>
      ),
    },
    { key: 'function', label: 'Function' },
    { key: 'project',  label: 'Project'  },
    { key: 'createdOn', label: 'CDate'   },
    {
      key: 'status', label: 'Status',
      render: (val: string) => (
        <span className={`statusBadge ${val === 'Approved' ? 'statusApproved' : val === 'Rejected' ? 'statusRejected' : 'statusPending'}`}>
          {val}
        </span>
      ),
    },
    { key: 'pendingAt', label: 'Pending At' },
    { key: 'level',     label: 'Level'      },
  ];

  const filtered    = data.filter((row) =>
    Object.keys(row).some((k) => String((row as any)[k]).toLowerCase().includes(search.toLowerCase()))
  );
  const totalPages  = Math.max(1, Math.ceil(filtered.length / pageSize));
  const paginated   = filtered.slice((currentPage - 1) * pageSize, currentPage * pageSize);

  const handleSearch         = (val: string) => { setSearch(val);     setCurrentPage(1); };
  const handlePageSizeChange = (size: number) => { setPageSize(size); setCurrentPage(1); };

  return (
    <div className="contentCard">

      <TableToolbar
        pageSize={pageSize} pageSizeOptions={PAGE_SIZE_OPTIONS} onPageSizeChange={handlePageSizeChange}
        search={search} onSearch={handleSearch}
        right={
          <div className="flex items-center gap-3">
            <ExportToExcel data={data} columns={exportColumns} fileName="NFA_Home" sheetName="NFA Data" />
            <SearchInput value={search} onChange={handleSearch} />
          </div>
        }
      />

      <CommonTable
        columns={columns}
        data={paginated}
        emptyMessage="No NFA records found"
        stickyFirstColumn={true}
        onRowClick={onRowClick}
        loading={loading}
      />

      <Pagination
        page={currentPage} total={totalPages} pageSize={pageSize}
        totalRecords={filtered.length} pageSizeOptions={PAGE_SIZE_OPTIONS}
        onPageChange={setCurrentPage} onPageSizeChange={handlePageSizeChange}
      />
    </div>
  );
};

export default HomeTable;
