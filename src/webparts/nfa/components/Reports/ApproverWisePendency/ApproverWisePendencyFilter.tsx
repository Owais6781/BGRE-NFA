import * as React from 'react';
import { useState } from 'react';
import { SlidersHorizontal } from 'lucide-react';
import DatePickerField from '../../../../../Share/Components/Dateinput';
import DropdownField from '../../../../../Share/Components/DropdownField';
import CommonButton from '../../../../../Share/Components/CommonButton';
import "../../../../../Share/styles/global.scss";

export interface FilterValues {
  fromDate: string;
  toDate: string;
  status: string;
}

interface ApproverWisePendencyFilterProps {
  onSearch: (filters: FilterValues) => void;
  onReset: () => void;
}

const statusOptions = [
  { label: 'All', value: 'All' },
  { label: 'Pending', value: 'Pending' },
  { label: 'Query', value: 'Query' },
  { label: 'Delegated', value: 'Delegated' },
];

const ApproverWisePendencyFilter: React.FC<ApproverWisePendencyFilterProps> = ({ onSearch, onReset }) => {
  const tomorrow = new Date(Date.now() + 1 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
  const minus10 = new Date(Date.now() - 10 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];

  const [fromDate, setFromDate] = useState(minus10);
  const [toDate, setToDate] = useState(tomorrow);
  const [statusVal, setStatusVal] = useState('All');

  const handleReset = () => {
    setFromDate(minus10);
    setToDate(tomorrow);
    setStatusVal('All');
    onReset();
  };

  return (
    <div className="filterCard">
      <div className="filterCardHeader">
        <SlidersHorizontal size={14} />
        Approver Wise Pendency — Filters
      </div>
      <div className="filterCardBody">
        <div className="flex flex-wrap gap-x-6 gap-y-4">
          <div className="flex-1 min-w-[200px]">
            <DatePickerField label="From Date" value={fromDate} onChange={setFromDate} required />
          </div>
          <div className="flex-1 min-w-[200px]">
            <DatePickerField label="To Date" value={toDate} onChange={setToDate} required />
          </div>
        </div>
        <div className="flex flex-wrap gap-x-6 gap-y-4 mt-4">
          <div className="flex-1 min-w-[200px] max-w-[calc(50%-12px)]">
            <DropdownField label="Status" value={statusVal} onChange={setStatusVal} options={statusOptions} />
          </div>
        </div>
        <div className="flex items-center gap-3 mt-4">
          <CommonButton 
            text="Get Pendency Report" 
            variant="primary" 
            onClick={() => onSearch({ fromDate, toDate, status: statusVal })} 
          />
          <CommonButton text="Reset" variant="danger" onClick={handleReset} />
        </div>
      </div>
    </div>
  );
};

export default ApproverWisePendencyFilter;
