import * as React from 'react';
import { useState } from 'react';
import { SlidersHorizontal } from 'lucide-react';
import DatePickerField from '../../../../../Share/Components/Dateinput';
import DropdownField   from '../../../../../Share/Components/DropdownField';
import CommonButton    from '../../../../../Share/Components/CommonButton';
import "../../../../../Share/styles/global.scss";

import { IDropdownOption } from '../../../../../Share/Services/DropdownService';

interface FilterValues {
  fromDate: string;
  toDate: string;
  function: string;
  status: string;
  project: string;
}

interface ApprovedForFinanceFilterProps {
  onSearch: (filters: FilterValues) => void;
  onClear:  () => void;
  projectOptions: IDropdownOption[];
  functionOptions: IDropdownOption[];
}

const statusOptions = [
  { label: 'Pending', value: 'Pending' },
  { label: 'Approved', value: 'Approved' },
  { label: 'Completed', value: 'Completed' },
  { label: 'Rejected', value: 'Rejected' },
];

const ApprovedForFinanceFilter: React.FC<ApprovedForFinanceFilterProps> = ({ 
  onSearch, 
  onClear,
  projectOptions,
  functionOptions
}) => {
  const today   = new Date().toISOString().split('T')[0];
  const defaultFrom = new Date(Date.now() - 90 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];

  const [fromDate,   setFromDate]   = useState(defaultFrom);
  const [toDate,     setToDate]     = useState(today);
  const [funcVal,    setFuncVal]    = useState('');
  const [statusVal,  setStatusVal]  = useState('');
  const [projectVal, setProjectVal] = useState('');

  const handleClear = () => {
    setFromDate(defaultFrom);
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
        <div className="flex flex-wrap items-end gap-x-4 gap-y-3">
          <div className="flex-1 min-w-[150px] max-w-[200px]">
            <DatePickerField label="From Date" value={fromDate} onChange={setFromDate} />
          </div>
          <div className="flex-1 min-w-[150px] max-w-[200px]">
            <DatePickerField label="To Date"   value={toDate}   onChange={setToDate} />
          </div>
          <div className="flex-1 min-w-[160px] max-w-[220px]">
            <DropdownField label="Function" value={funcVal}    onChange={setFuncVal}    options={functionOptions || []} emptyLabel="All Functions" />
          </div>
          <div className="flex-1 min-w-[140px] max-w-[180px]">
            <DropdownField label="Status"   value={statusVal}  onChange={setStatusVal}  options={statusOptions} emptyLabel="All Statuses" />
          </div>
          <div className="flex-1 min-w-[160px] max-w-[220px]">
            <DropdownField label="Entity" value={projectVal} onChange={setProjectVal} options={projectOptions || []} emptyLabel="All Entities" />
          </div>
          <div className="flex items-center gap-2 pb-0.5">
            <CommonButton text="Search" variant="primary"
              onClick={() => onSearch({ fromDate, toDate, function: funcVal, status: statusVal, project: projectVal })} />
            <CommonButton text="Clear" variant="danger" onClick={handleClear} />
          </div>
        </div>
      </div>
    </div>
  );
};

export default ApprovedForFinanceFilter;
