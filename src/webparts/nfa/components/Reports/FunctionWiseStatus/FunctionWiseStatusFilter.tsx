import * as React from 'react';
import { useState } from 'react';
import { SlidersHorizontal } from 'lucide-react';
import DatePickerField from '../../../../../Share/Components/Dateinput';
import DropdownField from '../../../../../Share/Components/DropdownField';
import CommonButton from '../../../../../Share/Components/CommonButton';
import { IDropdownOption } from '../../../../../Share/Services/DropdownService';
import "../../../../../Share/styles/global.scss";

export interface FilterValues {
  fromDate: string;
  toDate: string;
  functionName: string;
}

interface FunctionWiseStatusFilterProps {
  onSearch: (filters: FilterValues) => void;
  onReset: () => void;
  functionOptions: IDropdownOption[];
}

const FunctionWiseStatusFilter: React.FC<FunctionWiseStatusFilterProps> = ({ 
  onSearch, 
  onReset, 
  functionOptions 
}) => {
  const today = new Date().toISOString().split('T')[0];
  const tomorrow = new Date(Date.now() + 1 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];

  const [fromDate, setFromDate] = useState(today);
  const [toDate, setToDate] = useState(tomorrow);
  const [functionName, setFunctionName] = useState('0');

  const handleReset = () => {
    setFromDate(today);
    setToDate(tomorrow);
    setFunctionName('0');
    onReset();
  };

  const optionsWithAll = [
    { label: '-All-', value: '0' },
    ...functionOptions.filter(o => o.value !== '0' && o.label !== '-All-')
  ];

  return (
    <div className="filterCard">
      <div className="filterCardHeader">
        <SlidersHorizontal size={14} />
        Function Wise Status — Filters
      </div>
      <div className="filterCardBody">
        <div className="flex flex-wrap items-end gap-x-4 gap-y-3">
          <div className="flex-1 min-w-[150px] max-w-[220px]">
            <DatePickerField label="From Date" value={fromDate} onChange={setFromDate} required />
          </div>
          <div className="flex-1 min-w-[150px] max-w-[220px]">
            <DatePickerField label="To Date" value={toDate} onChange={setToDate} required />
          </div>
          <div className="flex-1 min-w-[180px] max-w-[260px]">
            <DropdownField 
              label="Function Name" 
              value={functionName} 
              onChange={setFunctionName} 
              options={optionsWithAll} 
            />
          </div>
          <div className="flex items-center gap-2 pb-0.5">
            <CommonButton 
              text="Search" 
              variant="primary" 
              onClick={() => onSearch({ fromDate, toDate, functionName })} 
            />
            <CommonButton text="Reset" variant="danger" onClick={handleReset} />
          </div>
        </div>
      </div>
    </div>
  );
};

export default FunctionWiseStatusFilter;
