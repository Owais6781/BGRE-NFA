import * as React from 'react';
import { useState } from 'react';
import { ChevronUp, ChevronDown } from 'lucide-react';
import ApprovalDetailsFilter from './ApprovalDetailsFilter';
import ApprovalDetailsTable  from './ApprovalDetailsTable';
import "../../../../../Share/styles/global.scss";

interface ApprovalRecord {
  nfa: string; subject: string; function: string;
  createdBy: string; createdOn: string; status: string;
  approvedOn: string; pendingAt: string; pendingFrom: string;
  project: string; packageCost: string;
}

const ApprovalDetailsView: React.FC = () => {
  const [tableData,   setTableData]   = useState<ApprovalRecord[]>([]);
  const [filters,     setFilters]     = useState<any>(null);
  const [isCollapsed, setIsCollapsed] = useState(false);

  // --- Filtering Logic ---
  const displayData = React.useMemo(() => {
    if (!filters) return tableData;

    return tableData.filter(item => {
      // Date Filter
      if (filters.fromDate || filters.toDate) {
        const itemDate = new Date(item.createdOn);
        itemDate.setHours(0, 0, 0, 0);

        if (filters.fromDate) {
          const from = new Date(filters.fromDate);
          from.setHours(0, 0, 0, 0);
          if (itemDate < from) return false;
        }
        if (filters.toDate) {
          const to = new Date(filters.toDate);
          to.setHours(0, 0, 0, 0);
          if (itemDate > to) return false;
        }
      }

      // Dropdown Filters
      if (filters.status && item.status !== filters.status) return false;
      if (filters.function && item.function !== filters.function) return false;
      if (filters.project && item.project !== filters.project) return false;

      return true;
    });
  }, [tableData, filters]);

  const handleSearch = (f: any) => {
    setFilters(f);
  };

  return (
    <div className="pageContainer">

      <div className="sectionTitleBar flex items-center justify-between">
        <span className="sectionTitle">Detailed View for Approver</span>
        <button className="iconBtn" onClick={() => setIsCollapsed(c => !c)} title={isCollapsed ? 'Expand' : 'Collapse'}>
          {isCollapsed ? <ChevronDown size={16} /> : <ChevronUp size={16} />}
        </button>
      </div>

      {!isCollapsed && (
        <>
          <ApprovalDetailsFilter onSearch={handleSearch} onClear={() => setFilters(null)} />
          <ApprovalDetailsTable  data={displayData} onRecordClick={(nfa) => console.log('Clicked:', nfa)} />
        </>
      )}

    </div>
  );
};

export default ApprovalDetailsView;
