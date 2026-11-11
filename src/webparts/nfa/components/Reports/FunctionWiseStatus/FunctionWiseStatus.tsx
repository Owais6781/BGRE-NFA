import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import TableToolbar from '../../../../../Share/Components/TableToolbar';
import ExportToExcel from '../../../../../Share/Components/ExportToExcel';
import { SearchInput } from '../../../../../Share/Components/SearchInput';
import { Skeleton } from 'primereact/skeleton';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { ReportService, IFunctionStatusRecord } from '../../../../../Share/Services/ReportService';
import { DropdownService, IDropdownOption } from '../../../../../Share/Services/DropdownService';
import FunctionWiseStatusFilter, { FilterValues } from './FunctionWiseStatusFilter';
import PageTitle from '../../../../../Share/Components/PageTitle';
import "../../../../../Share/styles/global.scss";

const COLUMNS = [
  { key: 'functionName', label: 'Function Name' },
  { key: 'approved',     label: 'Approved'      },
  { key: 'pending',      label: 'Pending'       },
  { key: 'query',        label: 'Query'         },
  { key: 'delegated',    label: 'Delegated'     },
  { key: 'rejected',     label: 'Rejected'      },
  { key: 'grandTotal',   label: 'Grand Total'   },
];

interface IFunctionWiseStatusProps {
  context: WebPartContext;
  sp: SPFI;
}

const FunctionWiseStatus: React.FC<IFunctionWiseStatusProps> = ({ context, sp }) => {
  const [data, setData] = useState<IFunctionStatusRecord[]>([]);
  const [functionOptions, setFunctionOptions] = useState<IDropdownOption[]>([]);
  const [filters, setFilters] = useState<FilterValues | null>(null);
  const [loading, setLoading] = useState(false);
  const [search, setSearch] = useState('');

  const loadData = async (fromDate: string, toDate: string, functionName: string) => {
    setLoading(true);
    try {
      const reportService = new ReportService(sp);
      const records = await reportService.getFunctionWiseStatus(
        fromDate,
        toDate,
        functionName !== '0' ? functionName : undefined
      );
      setData(records);
      setFilters({ fromDate, toDate, functionName });
    } catch (err) {
      console.error("FunctionWiseStatus: Error loading report data:", err);
      setData([]);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    const initPage = async () => {
      setLoading(true);
      try {
        const dropdownService = new DropdownService(context);
        const funcs = await dropdownService.getFunctionMaster();
        setFunctionOptions(funcs);

        // Default: today to tomorrow
        const todayStr = new Date().toISOString().split('T')[0];
        const tomorrowStr = new Date(Date.now() + 1 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];

        const reportService = new ReportService(sp);
        const records = await reportService.getFunctionWiseStatus(todayStr, tomorrowStr, '0');
        setData(records);
        setFilters({ fromDate: todayStr, toDate: tomorrowStr, functionName: '0' });
      } catch (err) {
        console.error("FunctionWiseStatus: Error initializing page data:", err);
      } finally {
        setLoading(false);
      }
    };
    initPage();
  }, [context, sp]);

  const handleSearch = (filterVals: FilterValues) => {
    loadData(filterVals.fromDate, filterVals.toDate, filterVals.functionName);
  };

  const handleReset = () => {
    setData([]);
    setFilters(null);
    setSearch('');
  };

  const displayData = useMemo(() => {
    if (!search) return data;
    return data.filter(row => 
      row.functionName.toLowerCase().includes(search.toLowerCase())
    );
  }, [data, search]);

  const totals = useMemo(() => {
    return displayData.reduce((acc, row) => ({
      approved: acc.approved + row.approved,
      pending: acc.pending + row.pending,
      query: acc.query + row.query,
      delegated: acc.delegated + row.delegated,
      rejected: acc.rejected + row.rejected,
      grandTotal: acc.grandTotal + row.grandTotal
    }), {
      approved: 0,
      pending: 0,
      query: 0,
      delegated: 0,
      rejected: 0,
      grandTotal: 0
    });
  }, [displayData]);

  return (
    <div className="pageContainer">
      <PageTitle 
        title="Note for Approval(NFA) | Function wise Status Report" 
        subtitle="Analyze approval requests and workflows grouped by department/function." 
      />

      {/* Filter Card */}
      <div className="mb-4">
        <FunctionWiseStatusFilter 
          onSearch={handleSearch} 
          onReset={handleReset} 
          functionOptions={functionOptions}
        />
      </div>

      <div className="contentCard">
        <TableToolbar
          right={
            <div className="flex items-center gap-3">
              <SearchInput value={search} onChange={setSearch} />
              {displayData.length > 0 && (
                <ExportToExcel 
                  data={displayData} 
                  columns={COLUMNS} 
                  fileName="Function_Wise_Status_Report" 
                  sheetName="Status Report" 
                />
              )}
            </div>
          }
        />

        <div className="tableWrapper mt-3">
          <table className="commonTable" id="tblDataMain">
            <thead>
              <tr>
                <th className="reportTableHeader" style={{ textAlign: 'left' }}>Function Name</th>
                <th className="reportTableHeader" style={{ textAlign: 'center' }}>Approved</th>
                <th className="reportTableHeader" style={{ textAlign: 'center' }}>Pending</th>
                <th className="reportTableHeader" style={{ textAlign: 'center' }}>Query</th>
                <th className="reportTableHeader" style={{ textAlign: 'center' }}>Delegated</th>
                <th className="reportTableHeader" style={{ textAlign: 'center' }}>Rejected</th>
                <th className="reportTableHeader" style={{ textAlign: 'center' }}>Grand Total</th>
              </tr>
            </thead>
            <tbody id="tbodyReport">
              {loading ? (
                <tr>
                  <td colSpan={7} className="reportTableCell text-center py-6" style={{ textAlign: 'center' }}>
                    <span className="inline-flex items-center gap-2 text-[#94a3b8]">
                      <span className="inline-block w-4 h-4 border-2 border-[#3b82f6] border-t-transparent rounded-full animate-spin" />
                      Loading…
                    </span>
                  </td>
                </tr>
              ) : displayData.length === 0 ? (
                <tr>
                  <td colSpan={7} className="reportTableCell text-center py-8" style={{ color: '#94a3b8', textAlign: 'center' }}>
                    {filters ? "No records found matching filters" : "Please click on search button to download report"}
                  </td>
                </tr>
              ) : (
                <>
                  {displayData.map((row, i) => (
                    <tr key={i} className={i % 2 === 0 ? 'dayWiseRowEven' : 'dayWiseRowOdd'}>
                      <td className="reportTableCell font-medium text-[#1e3a5f]">{row.functionName}</td>
                      <td className="reportTableCell text-center" style={{ textAlign: 'center' }}>{row.approved}</td>
                      <td className="reportTableCell text-center" style={{ textAlign: 'center' }}>{row.pending}</td>
                      <td className="reportTableCell text-center" style={{ textAlign: 'center' }}>{row.query}</td>
                      <td className="reportTableCell text-center" style={{ textAlign: 'center' }}>{row.delegated}</td>
                      <td className="reportTableCell text-center" style={{ textAlign: 'center' }}>{row.rejected}</td>
                      <td className="reportTableCell text-center font-bold text-[#1e3a5f]" style={{ textAlign: 'center' }}><b>{row.grandTotal}</b></td>
                    </tr>
                  ))}
                  <tr className="dayWiseGrandTotal font-bold" id="tfootReport">
                    <td className="reportTableCell font-bold">Grand Total</td>
                    <td className="reportTableCell text-center font-bold" style={{ textAlign: 'center' }}>{totals.approved}</td>
                    <td className="reportTableCell text-center font-bold" style={{ textAlign: 'center' }}>{totals.pending}</td>
                    <td className="reportTableCell text-center font-bold" style={{ textAlign: 'center' }}>{totals.query}</td>
                    <td className="reportTableCell text-center font-bold" style={{ textAlign: 'center' }}>{totals.delegated}</td>
                    <td className="reportTableCell text-center font-bold" style={{ textAlign: 'center' }}>{totals.rejected}</td>
                    <td className="reportTableCell text-center font-bold text-[#1e3a5f]" style={{ textAlign: 'center' }}><b>{totals.grandTotal}</b></td>
                  </tr>
                </>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

export default FunctionWiseStatus;
