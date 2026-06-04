import * as React from 'react';
import { useState, useEffect } from 'react';
import HomeFilter from './HomeFilter';
import HomeTable  from './HomeTable';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { DropdownService, IDropdownOption } from '../../../../../Share/Services/DropdownService';
import { NFAService } from '../../../../../Share/Services/NFAService';
import { useNavigate } from 'react-router-dom';
import "../../../../../Share/styles/global.scss";

const SAMPLE_DATA = [
  {
    requestNo: 'LPHRL/IT/2026/Apr/6035', function: 'IT', contractorName: 'Test data',
    createdOn: '09/04/2026 1:21 AM', status: 'Approved', approvedRejBy: 'NFA CORP',
    approvedRejDate: '09/04/2026 2:17 AM', pendingAt: 'NFA CORP',
    pendingFrom: '09/04/2026 2:13 AM', level: 4, project: 'LPH Lee Pal Hot Res Lim',
  },
];

interface IHomeProps {
  context: WebPartContext;
  sp: SPFI;
}

const Home: React.FC<IHomeProps> = ({ context, sp }) => {
  const [tableData, setTableData] = useState<any[]>([]);
  const [projectOptions, setProjectOptions] = useState<IDropdownOption[]>([]);
  const [functionOptions, setFunctionOptions] = useState<IDropdownOption[]>([]);
  const navigate = useNavigate();

  useEffect(() => {
    const loadData = async () => {
      const dropdownService = new DropdownService(context);
      const nfaService = new NFAService(context);
      try {
        console.log("Home Component: Starting to load data...");
        const [projects, funcs, nfas] = await Promise.all([
          dropdownService.getProjects(),
          dropdownService.getFunctionMaster(),
          nfaService.getNFAList()
        ]);
        
        console.log("Home Component: Projects loaded:", projects.length);
        console.log("Home Component: Functions loaded:", funcs.length);
        console.log("Home Component: NFA Records loaded:", nfas.length);
        
        setProjectOptions(projects);
        setFunctionOptions(funcs);
        setTableData(nfas);
      } catch (err) {
        console.error("Failed to load Home data:", err);
      }
    };
    loadData();
  }, []);

  const handleSearch = (filters: any) => {
    console.log('Search filters:', filters);
    // In a real app, we'd filter tableData here based on dates, etc.
  };

  return (
    <div className="pageContainer">

      <div className="sectionTitleBar flex items-center justify-between">
        <span className="sectionTitle">Home</span>
      </div>

      {/* ── Filter ── */}
      <HomeFilter 
        onSearch={handleSearch} 
        projectOptions={projectOptions} 
        functionOptions={functionOptions} 
      />

      {/* ── Table ── */}
      <HomeTable
        data={tableData}
        onRequestClick={(id) => navigate(`/new-nfa/${id}`)}
      />

    </div>
  );
};

export default Home;
