import * as React from 'react';
import { useState } from 'react';
import { SlidersHorizontal, ChevronDown } from 'lucide-react';
import DatePickerField from '../../../../../Share/Components/Dateinput';
import DropdownField   from '../../../../../Share/Components/DropdownField';
import CommonButton    from '../../../../../Share/Components/CommonButton';
import { IDropdownOption } from '../../../../../Share/Services/DropdownService';
import "../../../../../Share/styles/global.scss";

interface HomeFilterProps {
  functionOptions: IDropdownOption[];
  projectOptions: IDropdownOption[];
  onSearch: (filters: {
    fromDate: string; toDate: string;
    function: string; status: string; project: string;
  }) => void;
}

const statusOptions = [
  { label: 'Pending',  value: 'Pending'  },
  { label: 'Approved', value: 'Approved' },
  { label: 'Rejected', value: 'Rejected' },
];

const HomeFilter: React.FC<HomeFilterProps> = ({ onSearch, functionOptions, projectOptions }) => {
  const today   = new Date().toISOString().split('T')[0];
  const minus11 = new Date(Date.now() - 11 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];

  const [fromDate,   setFromDate]   = useState(minus11);
  const [toDate,     setToDate]     = useState(today);
  const [funcVal,    setFuncVal]    = useState('');
  const [statusVal,  setStatusVal]  = useState('');
  const [projectVal, setProjectVal] = useState('');
  const [isOpen,     setIsOpen]     = useState(true);

  const handleSearch = () =>
    onSearch({ fromDate, toDate, function: funcVal, status: statusVal, project: projectVal });

  return (
    <div className={`filterCard ${isOpen ? 'filterCard--open' : 'filterCard--closed'}`}>
      <div className="filterCardHeader cursor-pointer flex justify-between items-center" onClick={() => setIsOpen(!isOpen)}>
        <div className="flex items-center gap-2">
          <SlidersHorizontal size={14} />
          <span>Filters</span>
        </div>
        <div className={`transition-transform duration-300 ${isOpen ? 'rotate-180' : ''}`}>
          <ChevronDown size={16} />
        </div>
      </div>

      <div className={`formSectionBody--collapsible ${isOpen ? 'formSectionBody--expanded' : 'formSectionBody--collapsed'}`}>
        <div className="filterCardBody">
          <div className="flex flex-wrap items-end gap-x-4 gap-y-4">
            <div className="flex-1 min-w-[140px]">
              <DatePickerField label="From Date" value={fromDate} onChange={setFromDate} required />
            </div>
            <div className="flex-1 min-w-[140px]">
              <DatePickerField label="To Date"   value={toDate}   onChange={setToDate}   required />
            </div>
            <div className="flex-1 min-w-[150px]">
              <DropdownField label="Status"   value={statusVal}  onChange={setStatusVal}  options={statusOptions}   required />
            </div>
            <div className="flex-[1.5] min-w-[200px]">
              <DropdownField label="Project" value={projectVal} onChange={setProjectVal} options={projectOptions} required />
            </div>
            <div className="flex-1 min-w-[180px]">
              <DropdownField label="Function" value={funcVal}    onChange={setFuncVal}    options={functionOptions} required />
            </div>
            <div className="flex-none">
              <CommonButton text="Search" variant="primary" onClick={handleSearch} />
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default HomeFilter;
