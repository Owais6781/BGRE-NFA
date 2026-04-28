import * as React      from 'react';
import { useState }    from 'react';
import { SlidersHorizontal } from 'lucide-react';
import DatePickerField from '../../../../../Share/Components/Dateinput';
import DropdownField   from '../../../../../Share/Components/DropdownField';
import CommonButton    from '../../../../../Share/Components/CommonButton';
import { InputField }  from '../../../../../Share/Components/InputField';
import "../../../../../Share/styles/global.scss";

interface FilterValues {
  fromDate: string; toDate: string; function: string;
  project: string; contractor: string;
  packageCostFrom: string; packageCostTo: string;
}
interface GenericReportFilterProps {
  onSearch: (filters: FilterValues) => void;
  onReset:  () => void;
}

const functionOptions  = [
  { label: 'IT', value: 'IT' }, { label: 'HR', value: 'HR' },
  { label: 'GYM', value: 'GYM' }, { label: 'Finance', value: 'Finance' },
  { label: 'Engineering', value: 'Engineering' }, { label: 'Project', value: 'Project' },
];
const projectOptions   = [
  { label: 'Project A', value: 'A' }, { label: 'Project B', value: 'B' }, { label: 'Project C', value: 'C' },
];
const contractorOptions = [
  { label: 'Contractor A', value: 'A' }, { label: 'Contractor B', value: 'B' }, { label: 'Contractor C', value: 'C' },
];

const GenericReportFilter: React.FC<GenericReportFilterProps> = ({ onSearch, onReset }) => {
  const today   = new Date().toISOString().split('T')[0];
  const minus16 = new Date(Date.now() - 16 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];

  const [fromDate,        setFromDate]        = useState(minus16);
  const [toDate,          setToDate]          = useState(today);
  const [functionVal,     setFunctionVal]     = useState('');
  const [projectVal,      setProjectVal]      = useState('');
  const [contractorVal,   setContractorVal]   = useState('');
  const [packageCostFrom, setPackageCostFrom] = useState('');
  const [packageCostTo,   setPackageCostTo]   = useState('');

  const handleReset = () => {
    setFromDate(minus16); setToDate(today);
    setFunctionVal(''); setProjectVal(''); setContractorVal('');
    setPackageCostFrom(''); setPackageCostTo('');
    onReset();
  };

  return (
    <div className="filterCard">
      <div className="filterCardHeader">
        <SlidersHorizontal size={14} />
        NFA Generic Report — Filters
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
            <DropdownField label="Function" value={functionVal} onChange={setFunctionVal} options={functionOptions} />
          </div>
        </div>
        <div className="flex flex-wrap gap-x-6 gap-y-4 mt-4">
          <div className="flex-1 min-w-[200px]">
            <DropdownField label="Project"            value={projectVal}    onChange={setProjectVal}    options={projectOptions} />
          </div>
          <div className="flex-1 min-w-[200px]">
            <DropdownField label="Contractor / Vendor" value={contractorVal} onChange={setContractorVal} options={contractorOptions} />
          </div>
        </div>
        <div className="flex flex-wrap gap-x-6 gap-y-4 mt-4">
          <div className="flex-1 min-w-[200px]">
            <InputField label="Package Cost From" value={packageCostFrom} onChange={setPackageCostFrom} type="number" placeholder="0.00" />
          </div>
          <div className="flex-1 min-w-[200px]">
            <InputField label="Package Cost To"   value={packageCostTo}   onChange={setPackageCostTo}   type="number" placeholder="0.00" />
          </div>
        </div>
        <div className="flex items-center gap-3 mt-4">
          <CommonButton text="Search" variant="primary"
            onClick={() => onSearch({ fromDate, toDate, function: functionVal, project: projectVal, contractor: contractorVal, packageCostFrom, packageCostTo })} />
          <CommonButton text="Reset" variant="danger" onClick={handleReset} />
        </div>
      </div>
    </div>
  );
};

export default GenericReportFilter;
