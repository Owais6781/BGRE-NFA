import * as React from 'react';
import { useState } from 'react';
import { SlidersHorizontal } from 'lucide-react';
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

  const handleSearch = () =>
    onSearch({ fromDate, toDate, function: funcVal, status: statusVal, project: projectVal });

  return (
    <div className="filterCard">
      <div className="filterCardHeader">
        <SlidersHorizontal size={14} />
        Filters
      </div>
      <div className="filterCardBody">

        <div className="flex flex-wrap gap-x-6 gap-y-4">
          <div className="flex-1 min-w-[200px]">
            <DatePickerField label="From Date" value={fromDate} onChange={setFromDate} required />
          </div>
          <div className="flex-1 min-w-[200px]">
            <DatePickerField label="To Date"   value={toDate}   onChange={setToDate}   required />
          </div>
          <div className="flex-1 min-w-[200px]">
            <DropdownField label="Status"   value={statusVal}  onChange={setStatusVal}  options={statusOptions}   required />
          </div>
        </div>

        <div className="flex flex-wrap items-end gap-x-6 gap-y-4 mt-4">
          <div className="flex-1 min-w-[240px]">
            <DropdownField label="Project" value={projectVal} onChange={setProjectVal} options={projectOptions} required />
          </div>
          <div className="flex-1 min-w-[200px]">
            <DropdownField label="Function" value={funcVal}    onChange={setFuncVal}    options={functionOptions} required />
          </div>
          <div className="flex-none">
            <CommonButton text="Search" variant="primary" onClick={handleSearch} />
          </div>
        </div>

      </div>
    </div>
  );
};

export default HomeFilter;
