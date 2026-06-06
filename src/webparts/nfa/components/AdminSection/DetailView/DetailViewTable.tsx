import * as React from 'react';
import { useState }      from 'react';
import { CommonTable }   from '../../../../../Share/Components/CommonTable';
import { SearchInput }   from '../../../../../Share/Components/SearchInput';
import { Pagination }    from '../../../../../Share/Components/Pagination';
import ExportToExcel     from '../../../../../Share/Components/ExportToExcel';
import "../../../../../Share/styles/global.scss";

interface DetailViewRecord {
  nfa:              string;
  subject:          string;
  function:         string;
  createdBy:        string;
  createdOn:        string;
  status:           string;
  approvedRejOn:    string;
  pendingAt:        string;
  pendingFrom:      string;
  level:            number;
  project:          string;
  packageCost:      string;
}

interface DetailViewTableProps {
  data:           DetailViewRecord[];
  onRecordClick?: (nfa: string) => void;
}

const PAGE_SIZE_OPTIONS = [10, 25, 50, 100];

const exportColumns = [
  { key: 'nfa',           label: 'NFA'                  },
  { key: 'subject',       label: 'Subject'               },
  { key: 'function',      label: 'Function'              },
  { key: 'createdBy',     label: 'Created By'            },
  { key: 'createdOn',     label: 'Created On'            },
  { key: 'status',        label: 'Status'                },
  { key: 'approvedRejOn', label: 'Approved Rejected On'  },
  { key: 'pendingAt',     label: 'Pending At'            },
  { key: 'pendingFrom',   label: 'Pending From'          },
  { key: 'level',         label: 'Level'                 },
  { key: 'project',       label: 'Project'               },
  { key: 'packageCost',   label: 'Package Cost'          },
];

const DetailViewTable: React.FC<DetailViewTableProps> = ({
  data,
  onRecordClick,
}) => {
  const [pageSize,    setPageSize]    = useState(10);
  const [search,      setSearch]      = useState('');
  const [currentPage, setCurrentPage] = useState(1);

  // ── Column definitions ──
  const columns = [
    {
      key: 'nfa',
      label: 'NFA',
      render: (val: string) => (
        <button
          className="tableLinkBtn"
          onClick={(e) => { e.stopPropagation(); onRecordClick?.(val); }}
        >
          {val}
        </button>
      ),
    },
    { key: 'subject',   label: 'Subject'    },
    { key: 'function',  label: 'Function'   },
    { key: 'createdBy', label: 'Created By' },
    { key: 'createdOn', label: 'Created On' },
    {
      key: 'status',
      label: 'Status',
      render: (val: string) => (
        <span className={`statusBadge ${
          val === 'Approved' ? 'statusApproved' :
          val === 'Rejected' ? 'statusRejected' :
          'statusPending'
        }`}>
          {val}
        </span>
      ),
    },
    {
      key: 'approvedRejOn',
      label: 'Approved Rejected On',
    },
    { key: 'pendingAt',   label: 'Pending At'   },
    { key: 'pendingFrom', label: 'Pending From' },
    { key: 'level',       label: 'Level'        },
    { key: 'project',     label: 'Project'      },
    { key: 'packageCost', label: 'Package Cost' },
  ];

  // ── Filter ──
  const filtered = data.filter((row) =>
    Object.keys(row).some((key) =>
      String((row as any)[key]).toLowerCase().includes(search.toLowerCase())
    )
  );

  // ── Paginate ──
  const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
  const paginated  = filtered.slice(
    (currentPage - 1) * pageSize,
    currentPage * pageSize
  );

  const handleSearch         = (val: string) => { setSearch(val);     setCurrentPage(1); };
  const handlePageSizeChange = (size: number) => { setPageSize(size); setCurrentPage(1); };

  return (
    <div className="contentCard mt-3">

      {/* ── Export ── */}
      <div className="flex justify-end mb-2">
        <ExportToExcel
          data={data}
          columns={exportColumns}
          fileName="Detail_View"
          sheetName="Detail View"
        />
      </div>

      {/* ── Show entries + Search ── */}
      <div className="flex items-center justify-between flex-wrap gap-3 mb-3">
        <div className="flex items-center gap-2">
          <span className="formLabel" style={{ marginBottom: 0 }}>Show</span>
          <select
            value={pageSize}
            onChange={(e) => handlePageSizeChange(Number(e.target.value))}
            className="formInput formSelect"
            style={{ width: 70, padding: '4px 28px 4px 8px' }}
          >
            {PAGE_SIZE_OPTIONS.map((n) => (
              <option key={n} value={n}>{n}</option>
            ))}
          </select>
          <span className="formLabel" style={{ marginBottom: 0 }}>entries</span>
        </div>
        <SearchInput value={search} onChange={handleSearch} />
      </div>

      {/* ── Table ── */}
      <CommonTable
        columns={columns}
        data={paginated}
        emptyMessage="No records found"
        stickyFirstColumn={true}
      />

      {/* ── Pagination ── */}
      <Pagination
        page={currentPage}
        total={totalPages}
        pageSize={pageSize}
        totalRecords={filtered.length}
        pageSizeOptions={PAGE_SIZE_OPTIONS}
        onPageChange={setCurrentPage}
        onPageSizeChange={handlePageSizeChange}
      />

    </div>
  );
};

export default DetailViewTable;