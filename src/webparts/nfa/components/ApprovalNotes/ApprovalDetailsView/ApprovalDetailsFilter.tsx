import * as React from 'react';
import { useState } from 'react';
import { SlidersHorizontal } from 'lucide-react';
import DatePickerField from '../../../../../Share/Components/Dateinput';
import DropdownField   from '../../../../../Share/Components/DropdownField';
import CommonButton    from '../../../../../Share/Components/CommonButton';
import "../../../../../Share/styles/global.scss";

interface FilterValues {
  fromDate: string; toDate: string;
  function: string; status: string; project: string;
}

interface ApprovalDetailsFilterProps {
  onSearch: (filters: FilterValues) => void;
  onClear:  () => void;
}

const functionOptions = [
  { label: 'IT', value: 'IT' }, { label: 'HR', value: 'HR' },
  { label: 'GYM', value: 'GYM' }, { label: 'Finance', value: 'Finance' },
];
const statusOptions = [
  { label: 'Pending', value: 'Pending' }, { label: 'Approved', value: 'Approved' },
  { label: 'Rejected', value: 'Rejected' },
];
const projectOptions = [
  { label: 'Project A', value: 'A' }, { label: 'Project B', value: 'B' },
];

const ApprovalDetailsFilter: React.FC<ApprovalDetailsFilterProps> = ({ onSearch, onClear }) => {
  const today   = new Date().toISOString().split('T')[0];
  const minus18 = new Date(Date.now() - 18 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];

  const [fromDate,   setFromDate]   = useState(minus18);
  const [toDate,     setToDate]     = useState(today);
  const [funcVal,    setFuncVal]    = useState('');
  const [statusVal,  setStatusVal]  = useState('');
  const [projectVal, setProjectVal] = useState('');

  const handleClear = () => {
    setFromDate(minus18);
    setToDate(today);
    setFuncVal('');
    setStatusVal('');
    setProjectVal('');
    onClear();
  };

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
          <div className="flex flex-1 justify-end items-end gap-2">
            <CommonButton text="Search" variant="primary"
              onClick={() => onSearch({ fromDate, toDate, function: funcVal, status: statusVal, project: projectVal })} />
            <CommonButton text="Clear" variant="danger" onClick={handleClear} />
          </div>
        </div>

      </div>
    </div>
  );
};

export default ApprovalDetailsFilter;
