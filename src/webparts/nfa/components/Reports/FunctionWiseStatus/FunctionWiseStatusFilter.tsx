import * as React      from 'react';
import { useState }    from 'react';
import { SlidersHorizontal } from 'lucide-react';
import DatePickerField from '../../../../../Share/Components/Dateinput';
import DropdownField   from '../../../../../Share/Components/DropdownField';
import CommonButton    from '../../../../../Share/Components/CommonButton';
import "../../../../../Share/styles/global.scss";

import { IDropdownOption } from '../../../../../Share/Services/DropdownService';

interface FilterValues { fromDate: string; toDate: string; functionName: string; }
interface FunctionWiseStatusFilterProps {
  onSearch: (filters: FilterValues) => void;
  onReset:  () => void;
  functionOptions: IDropdownOption[];
}

const FunctionWiseStatusFilter: React.FC<FunctionWiseStatusFilterProps> = ({ onSearch, onReset, functionOptions }) => {
  const today     = new Date().toISOString().split('T')[0];
  const minus21   = new Date(Date.now() - 21 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];

  const [fromDate,     setFromDate]     = useState(minus21);
  const [toDate,       setToDate]       = useState(today);
  const [functionName, setFunctionName] = useState('');

  const handleReset = () => { setFromDate(minus21); setToDate(today); setFunctionName(''); onReset(); };

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
            <DatePickerField label="To Date"   value={toDate}   onChange={setToDate}   required />
          </div>
          <div className="flex-1 min-w-[180px] max-w-[260px]">
            <DropdownField label="Function Name" value={functionName} onChange={setFunctionName} options={functionOptions} required />
          </div>
          <div className="flex items-center gap-2 pb-0.5">
            <CommonButton text="Search" variant="primary" onClick={() => onSearch({ fromDate, toDate, functionName })} />
            <CommonButton text="Reset"  variant="danger"  onClick={handleReset} />
          </div>
        </div>
      </div>
    </div>
  );
};

export default FunctionWiseStatusFilter;
