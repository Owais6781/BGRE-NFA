import * as React from 'react';
import { useState } from 'react';
import { SlidersHorizontal, Search } from 'lucide-react';
import DatePickerField from '../../../../../Share/Components/Dateinput';
import DropdownField   from '../../../../../Share/Components/DropdownField';
import CommonButton    from '../../../../../Share/Components/CommonButton';
import "../../../../../Share/styles/global.scss";

interface HomeFilterProps {
  onSearch: (filters: {
    fromDate: string; toDate: string;
    function: string; status: string; project: string;
  }) => void;
}

const functionOptions = [
  { label: 'IT',      value: 'IT'      },
  { label: 'HR',      value: 'HR'      },
  { label: 'Finance', value: 'Finance' },
  { label: 'Admin',   value: 'Admin'   },
];

const statusOptions = [
  { label: 'Pending',  value: 'Pending'  },
  { label: 'Approved', value: 'Approved' },
  { label: 'Rejected', value: 'Rejected' },
];

const projectOptions = [
  { label: 'Project A', value: 'A' },
  { label: 'Project B', value: 'B' },
];

const HomeFilter: React.FC<HomeFilterProps> = ({ onSearch }) => {
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
        </div>

        <div className="flex flex-wrap gap-x-6 gap-y-4 mt-4">
          <div className="flex-1 min-w-[200px]">
            <DropdownField label="Function" value={funcVal}    onChange={setFuncVal}    options={functionOptions} required />
          </div>
          <div className="flex-1 min-w-[200px]">
            <DropdownField label="Status"   value={statusVal}  onChange={setStatusVal}  options={statusOptions}   required />
          </div>
        </div>

        <div className="flex flex-wrap items-end gap-x-6 gap-y-4 mt-4">
          <div className="flex-1 min-w-[200px] max-w-[calc(50%-12px)]">
            <DropdownField label="Project" value={projectVal} onChange={setProjectVal} options={projectOptions} required />
          </div>
          <div className="flex flex-1 justify-end items-end">
            <CommonButton text="Search" variant="primary" onClick={handleSearch} />
          </div>
        </div>

      </div>
    </div>
  );
};

export default HomeFilter;
