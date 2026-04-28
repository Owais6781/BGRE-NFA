import * as React from 'react';
import { useState } from 'react';
import { ChevronUp, ChevronDown } from 'lucide-react';
import DetailViewFilter from './DetailViewFilter';
import DetailViewTable  from './DetailViewTable';
import "../../../../../Share/styles/global.scss";

interface DetailViewRecord {
  nfa: string; subject: string; function: string;
  createdBy: string; createdOn: string; status: string;
  approvedRejOn: string; pendingAt: string; pendingFrom: string;
  level: number; project: string; packageCost: string;
}

const DetailView: React.FC = () => {
  const [tableData,   setTableData]   = useState<DetailViewRecord[]>([]);
  const [isCollapsed, setIsCollapsed] = useState(false);

  return (
    <div className="pageContainer">

      <div className="sectionTitleBar flex items-center justify-between">
        <span className="sectionTitle">Detailed View</span>
        <button className="iconBtn" onClick={() => setIsCollapsed(c => !c)} title={isCollapsed ? 'Expand' : 'Collapse'}>
          {isCollapsed ? <ChevronDown size={16} /> : <ChevronUp size={16} />}
        </button>
      </div>

      {!isCollapsed && (
        <>
          <DetailViewFilter onSearch={(f) => { console.log(f); setTableData([]); }} />
          <DetailViewTable  data={tableData} onRecordClick={(nfa) => console.log('Clicked:', nfa)} />
        </>
      )}

    </div>
  );
};

export default DetailView;
