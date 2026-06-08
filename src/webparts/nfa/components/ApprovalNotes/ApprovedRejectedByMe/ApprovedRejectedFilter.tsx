import * as React from 'react';
import { useState, useEffect } from 'react';
import { SlidersHorizontal, ChevronDown } from 'lucide-react';
import DatePickerField from '../../../../../Share/Components/Dateinput';
import DropdownField   from '../../../../../Share/Components/DropdownField';
import { IDropdownOption } from '../../../../../Share/Services/DropdownService';
import "../../../../../Share/styles/global.scss";
 
interface ApprovedRejectedFilterProps {
  filters: any;
  functionOptions: IDropdownOption[];
  projectOptions: IDropdownOption[];
  onSearch: (filters: {
    fromDate: string; toDate: string;
    function: string; status: string; project: string;
  }) => void;
  onClear: () => void;
}
 
const statusOptions = [
  { label: 'Approved', value: 'Approved' },
  { label: 'Rejected', value: 'Rejected' },
];
 
const ApprovedRejectedFilter: React.FC<ApprovedRejectedFilterProps> = ({ filters, onSearch, onClear, functionOptions, projectOptions }) => {
  const today   = new Date().toISOString().split('T')[0];
  const minus11 = new Date(Date.now() - 11 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
 
  const [fromDate,   setFromDate]   = useState(minus11);
  const [toDate,     setToDate]     = useState(today);
  const [funcVal,    setFuncVal]    = useState('');
  const [statusVal,  setStatusVal]  = useState('');
  const [projectVal, setProjectVal] = useState('');
  const [isOpen,     setIsOpen]     = useState(true);
 
  useEffect(() => {
    if (filters) {
      setFromDate(filters.fromDate || minus11);
      setToDate(filters.toDate || today);
      setFuncVal(filters.function || '');
      setStatusVal(filters.status || '');
      setProjectVal(filters.project || '');
    } else {
      setFromDate(minus11);
      setToDate(today);
      setFuncVal('');
      setStatusVal('');
      setProjectVal('');
    }
  }, [filters, today, minus11]);
 
  const handleSearch = () =>
    onSearch({ fromDate, toDate, function: funcVal, status: statusVal, project: projectVal });
 
  const handleClear = () => {
    setFromDate(minus11);
    setToDate(today);
    setFuncVal('');
    setStatusVal('');
    setProjectVal('');
    onClear();
  };
 
  return (
    <div className={`filterCard ${isOpen ? 'filterCard--open' : 'filterCard--closed'}`}>
      <div 
        className="filterCardHeader flex items-center justify-between cursor-pointer select-none"
        onClick={() => setIsOpen(!isOpen)}
      >
        <div className="flex items-center gap-2">
          <SlidersHorizontal size={14} className="text-[#AF8941]" />
          <span>Filters</span>
        </div>
        <ChevronDown 
          size={16} 
          className={`transform transition-transform duration-200 ${isOpen ? 'rotate-180' : ''}`} 
        />
      </div>
      <div className={`formSectionBody--collapsible ${isOpen ? 'formSectionBody--expanded' : 'formSectionBody--collapsed'}`}>
        <div className="filterCardBody">
          <div className="flex flex-wrap lg:flex-nowrap items-end gap-3.5 w-full">
            <div className="flex-1 min-w-[95px]">
              <DatePickerField label="From Date" value={fromDate} onChange={setFromDate} required />
            </div>
            <div className="flex-1 min-w-[95px]">
              <DatePickerField label="To Date"   value={toDate}   onChange={setToDate}   required />
            </div>
            <div className="flex-1 min-w-[95px]">
              <DropdownField label="Status"   value={statusVal}  onChange={setStatusVal}  options={statusOptions}   required />
            </div>
            <div className="flex-[1.2] min-w-[120px]">
              <DropdownField label="Project" value={projectVal} onChange={setProjectVal} options={projectOptions} required />
            </div>
            <div className="flex-[1.2] min-w-[120px]">
              <DropdownField label="Function" value={funcVal}    onChange={setFuncVal}    options={functionOptions} required />
            </div>
            <div className="flex-none flex gap-2 h-[38px] items-center">
              <button 
                className="btn btnPrimary h-full" 
                style={{ padding: '0 16px', minWidth: '75px', fontSize: '12px' }} 
                onClick={handleSearch}
              >
                Search
              </button>
              <button 
                className="btn btnDanger h-full" 
                style={{ padding: '0 16px', minWidth: '75px', fontSize: '12px' }} 
                onClick={handleClear}
              >
                Clear
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
 
export default ApprovedRejectedFilter;
