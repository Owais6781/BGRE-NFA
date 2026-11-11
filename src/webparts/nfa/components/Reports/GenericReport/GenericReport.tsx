import * as React from 'react';
import { useState, useEffect } from 'react';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import GenericReportFilter, { FilterValues } from './GenericReportFilter';
import ReportTable from '../../../../../Share/Components/ReportTable';
import ExportToExcel from '../../../../../Share/Components/ExportToExcel';
import TableToolbar from '../../../../../Share/Components/TableToolbar';
import PageTitle from '../../../../../Share/Components/PageTitle';
import { ReportService, IGenericReportRecord } from '../../../../../Share/Services/ReportService';
import { DropdownService, IDropdownOption } from '../../../../../Share/Services/DropdownService';
import "../../../../../Share/styles/global.scss";

const EXPORT_COLUMNS = [
  { key: 'nfaNo',          label: 'NFA No.' },
  { key: 'subjectMatter',  label: 'Subject / Matter' },
  { key: 'function',       label: 'Function' },
  { key: 'project',        label: 'Entity' },
  { key: 'contractor',     label: 'Contractor / Vendor' },
  { key: 'packageAwarded', label: 'Package Awarded Cost' },
  { key: 'budgetedAmount', label: 'Budgeted Amount' },
  { key: 'status',         label: 'Status' },
  { key: 'createdBy',      label: 'Created By' },
  { key: 'createdOn',      label: 'Created On' },
];

const TABLE_COLUMNS = [
  { key: 'nfaNoDisplay',   label: 'NFA No.' },
  { key: 'subjectDisplay', label: 'Subject / Matter' },
  { key: 'function',       label: 'Function' },
  { key: 'project',        label: 'Entity' },
  { key: 'contractor',     label: 'Contractor / Vendor' },
  { key: 'packageAwarded', label: 'Package Awarded Cost' },
  { key: 'budgetedAmount', label: 'Budgeted Amount' },
  { key: 'status',         label: 'Status' },
  { key: 'createdBy',      label: 'Created By' },
  { key: 'createdOn',      label: 'Created On' },
];

interface IGenericReportProps {
  context?: WebPartContext;
  sp?: SPFI;
}

const SubjectMatterCell: React.FC<{ text: string }> = ({ text }) => {
  const [expanded, setExpanded] = useState(false);
  if (!text || text.length <= 50) return <span>{text || '—'}</span>;
  return (
    <span>
      {expanded ? text : `${text.substring(0, 50)}... `}
      <button 
        type="button" 
        onClick={() => setExpanded(!expanded)} 
        className="text-[#000080] font-bold underline ml-1 text-xs hover:text-[#2563eb]"
      >
        {expanded ? "Show less" : "Read more..."}
      </button>
    </span>
  );
};

const GenericReport: React.FC<IGenericReportProps> = ({ context, sp }) => {
  const [data, setData] = useState<IGenericReportRecord[]>([]);
  const [loading, setLoading] = useState(false);
  const [hasSearched, setHasSearched] = useState(false);

  const [functionOptions, setFunctionOptions] = useState<IDropdownOption[]>([]);
  const [projectOptions, setProjectOptions] = useState<IDropdownOption[]>([]);
  const [contractorOptions, setContractorOptions] = useState<{ label: string; value: string }[]>([]);

  useEffect(() => {
    if (!context || !sp) return;

    const loadDropdowns = async () => {
      try {
        const dropdownService = new DropdownService(context);
        const reportService = new ReportService(sp);

        const [funcs, projs, contractors] = await Promise.all([
          dropdownService.getFunctionMaster(),
          reportService.getProjectList(),
          reportService.getContractorList(),
        ]);

        setFunctionOptions(funcs);
        setProjectOptions(projs);
        setContractorOptions(contractors.map(c => ({ label: c, value: c })));
      } catch (err) {
        console.error("GenericReport: Error loading dropdown options:", err);
      }
    };

    loadDropdowns();
  }, [context, sp]);

  const loadReport = async (filters: FilterValues) => {
    if (!sp) return;
    setLoading(true);
    setHasSearched(true);
    try {
      const reportService = new ReportService(sp);
      
      const funcObj = functionOptions.find(f => f.value === filters.function);
      const funcName = funcObj ? funcObj.label : filters.function;

      const projObj = projectOptions.find(p => p.value === filters.project);
      const projName = projObj ? projObj.label : filters.project;

      const records = await reportService.getGenericReport({
        fromDate: filters.fromDate,
        toDate: filters.toDate,
        functionId: filters.function !== '0' ? filters.function : undefined,
        functionName: filters.function !== '0' ? funcName : undefined,
        project: filters.project !== '0' ? projName : undefined,
        projectId: filters.project !== '0' ? filters.project : undefined,
        contractor: filters.contractor !== '0' ? filters.contractor : undefined,
        packageCostFrom: filters.packageCostFrom || undefined,
        packageCostTo: filters.packageCostTo || undefined,
      });

      setData(records);
    } catch (err) {
      console.error("GenericReport: Error loading generic report data:", err);
      setData([]);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (sp) {
      const tomorrow = new Date(Date.now() + 1 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
      const minus15 = new Date(Date.now() - 15 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
      loadReport({
        fromDate: minus15,
        toDate: tomorrow,
        function: '0',
        project: '0',
        contractor: '0',
        packageCostFrom: '',
        packageCostTo: ''
      });
    }
  }, [sp]);

  const handleSearch = (filters: FilterValues) => {
    loadReport(filters);
  };

  const handleReset = () => {
    setData([]);
    setHasSearched(false);
  };

  const displayData = React.useMemo(() => {
    return data.map(item => {
      const isApprovedOrCompleted = item.status === "Approved" || item.status === "Completed";
      const viewUrl = isApprovedOrCompleted
        ? `#/view-nfa/${item.id}`
        : `#/view-nfa/${item.id}`;

      return {
        ...item,
        nfaNoDisplay: (
          <a href={viewUrl} className="font-semibold text-[#1e3a5f] hover:underline">
            {item.nfaNo || '—'}
          </a>
        ),
        subjectDisplay: <SubjectMatterCell text={item.subjectMatter} />
      };
    });
  }, [data]);

  return (
    <div className="pageContainer">
      <PageTitle 
        title="NFA Generic Report" 
        subtitle="Custom filterable report for Note for Approval records." 
      />

      <GenericReportFilter 
        onSearch={handleSearch} 
        onReset={handleReset} 
        functionOptions={functionOptions}
        projectOptions={projectOptions}
        contractorOptions={contractorOptions}
      />

      <div className="contentCard">
        <TableToolbar
          right={
            data.length > 0 ? (
              <ExportToExcel 
                data={data} 
                columns={EXPORT_COLUMNS} 
                fileName="NFA_Generic_Report" 
                sheetName="Generic Report" 
              />
            ) : null
          }
        />
        <ReportTable 
          columns={TABLE_COLUMNS} 
          data={displayData} 
          loading={loading}
          emptyMessage={hasSearched ? "No NFA details found. Please refine your search" : "Please click on Search to load the report"} 
        />
      </div>
    </div>
  );
};

export default GenericReport;
