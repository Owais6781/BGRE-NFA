import * as React      from 'react';
import { useState }    from 'react';
import { SlidersHorizontal } from 'lucide-react';
import DatePickerField from '../../../../../Share/Components/Dateinput';
import DropdownField   from '../../../../../Share/Components/DropdownField';
import CommonButton    from '../../../../../Share/Components/CommonButton';
import "../../../../../Share/styles/global.scss";

interface FilterValues { fromDate: string; toDate: string; functionName: string; }
interface FunctionWiseStatusFilterProps {
  onSearch: (filters: FilterValues) => void;
  onReset:  () => void;
}

const functionOptions = [
  { label: 'IT', value: 'IT' }, { label: 'HR', value: 'HR' },
  { label: 'GYM', value: 'GYM' }, { label: 'Finance', value: 'Finance' },
];

const FunctionWiseStatusFilter: React.FC<FunctionWiseStatusFilterProps> = ({ onSearch, onReset }) => {
  const today     = new Date().toISOString().split('T')[0];
  const yesterday = new Date(Date.now() - 1 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];

  const [fromDate,     setFromDate]     = useState(yesterday);
  const [toDate,       setToDate]       = useState(today);
  const [functionName, setFunctionName] = useState('');

  const handleReset = () => { setFromDate(yesterday); setToDate(today); setFunctionName(''); onReset(); };

  return (
    <div className="filterCard">
      <div className="filterCardHeader">
        <SlidersHorizontal size={14} />
        Function Wise Status — Filters
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
          <div className="flex-1 min-w-[200px] max-w-[calc(50%-12px)]">
            <DropdownField label="Function Name" value={functionName} onChange={setFunctionName} options={functionOptions} required />
          </div>
        </div>
        <div className="flex items-center gap-3 mt-4">
          <CommonButton text="Search" variant="primary" onClick={() => onSearch({ fromDate, toDate, functionName })} />
          <CommonButton text="Reset"  variant="danger"  onClick={handleReset} />
        </div>
      </div>
    </div>
  );
};

export default FunctionWiseStatusFilter;
