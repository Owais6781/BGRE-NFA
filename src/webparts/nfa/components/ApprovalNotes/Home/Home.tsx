import * as React from 'react';
import { useState } from 'react';
import HomeFilter from './HomeFilter';
import HomeTable  from './HomeTable';
import "../../../../../Share/styles/global.scss";

const SAMPLE_DATA = [
  {
    requestNo: 'LPHRL/IT/2026/Apr/6035', function: 'IT', contractorName: 'Test data',
    createdOn: '09/04/2026 1:21 AM', status: 'Approved', approvedRejBy: 'NFA CORP',
    approvedRejDate: '09/04/2026 2:17 AM', pendingAt: 'NFA CORP',
    pendingFrom: '09/04/2026 2:13 AM', level: 4, project: 'LPH Lee Pal Hot Res Lim',
  },
];

const Home: React.FC = () => {
  const [tableData, setTableData] = useState(SAMPLE_DATA);

  const handleSearch = (filters: any) => {
    console.log('Search filters:', filters);
    setTableData(SAMPLE_DATA);
  };

  return (
    <div className="pageContainer">

      <div className="sectionTitleBar flex items-center justify-between">
        <span className="sectionTitle">Home</span>
      </div>

      {/* ── Filter ── */}
      <HomeFilter onSearch={handleSearch} />

      {/* ── Table ── */}
      <HomeTable
        data={tableData}
        onRequestClick={(no) => console.log('Clicked:', no)}
      />

    </div>
  );
};

export default Home;
