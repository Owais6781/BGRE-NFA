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
  const [isCollapsed, setIsCollapsed] = useState(false);

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
          <ApprovalDetailsFilter onSearch={(f) => { console.log(f); setTableData([]); }} />
          <ApprovalDetailsTable  data={tableData} onRecordClick={(nfa) => console.log('Clicked:', nfa)} />
        </>
      )}

    </div>
  );
};

export default ApprovalDetailsView;
