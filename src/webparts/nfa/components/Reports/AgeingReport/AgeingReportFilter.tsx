import * as React      from 'react';
import { useState }    from 'react';
import { SlidersHorizontal } from 'lucide-react';
import DatePickerField from '../../../../../Share/Components/Dateinput';
import CommonButton    from '../../../../../Share/Components/CommonButton';
import "../../../../../Share/styles/global.scss";

interface FilterValues { fromDate: string; toDate: string; }
interface AgeingReportFilterProps {
  onSearch: (filters: FilterValues) => void;
  onReset:  () => void;
}

const AgeingReportFilter: React.FC<AgeingReportFilterProps> = ({ onSearch, onReset }) => {
  const today   = new Date().toISOString().split('T')[0];
  const minus21 = new Date(Date.now() - 21 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];

  const [fromDate, setFromDate] = useState(minus21);
  const [toDate,   setToDate]   = useState(today);

  const handleReset = () => { setFromDate(minus21); setToDate(today); onReset(); };

  return (
    <div className="filterCard">
      <div className="filterCardHeader">
        <SlidersHorizontal size={14} />
        Ageing Report — Filters
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
        <div className="flex items-center gap-3 mt-4">
          <CommonButton text="Search" variant="primary" onClick={() => onSearch({ fromDate, toDate })} />
          <CommonButton text="Reset"  variant="danger"  onClick={handleReset} />
        </div>
      </div>
    </div>
  );
};

export default AgeingReportFilter;
