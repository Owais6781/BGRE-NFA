import * as React from 'react';
import { useState } from 'react';
import { SlidersHorizontal, AlertCircle } from 'lucide-react';
import DatePickerField from '../../../../../Share/Components/Dateinput';
import DropdownField from '../../../../../Share/Components/DropdownField';
import CommonButton from '../../../../../Share/Components/CommonButton';
import { InputField } from '../../../../../Share/Components/InputField';
import "../../../../../Share/styles/global.scss";

export interface FilterValues {
  fromDate: string;
  toDate: string;
  function: string;
  project: string;
  contractor: string;
  packageCostFrom: string;
  packageCostTo: string;
}

interface GenericReportFilterProps {
  onSearch: (filters: FilterValues) => void;
  onReset: () => void;
  functionOptions?: { label: string; value: string }[];
  projectOptions?: { label: string; value: string }[];
  contractorOptions?: { label: string; value: string }[];
}

const GenericReportFilter: React.FC<GenericReportFilterProps> = ({ 
  onSearch, 
  onReset,
  functionOptions = [],
  projectOptions = [],
  contractorOptions = []
}) => {
  const tomorrow = new Date(Date.now() + 1 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
  const minus15 = new Date(Date.now() - 15 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];

  const [fromDate, setFromDate] = useState(minus15);
  const [toDate, setToDate] = useState(tomorrow);
  const [functionVal, setFunctionVal] = useState('0');
  const [projectVal, setProjectVal] = useState('0');
  const [contractorVal, setContractorVal] = useState('0');
  const [packageCostFrom, setPackageCostFrom] = useState('');
  const [packageCostTo, setPackageCostTo] = useState('');
  const [validationError, setValidationError] = useState('');

  const validate = (): boolean => {
    if (packageCostFrom !== '' && packageCostTo === '') {
      setValidationError("Please enter package awarded cost to.");
      return false;
    }
    if (packageCostFrom === '' && packageCostTo !== '') {
      setValidationError("Please enter package awarded cost from.");
      return false;
    }
    if (packageCostFrom !== '' && packageCostTo !== '') {
      const from = parseFloat(packageCostFrom);
      const to = parseFloat(packageCostTo);
      if (from > to) {
        setValidationError("Package awarded cost from amount should be less than to amount.");
        return false;
      }
    }
    setValidationError('');
    return true;
  };

  const handleSearchClick = () => {
    if (validate()) {
      onSearch({ 
        fromDate, 
        toDate, 
        function: functionVal, 
        project: projectVal, 
        contractor: contractorVal, 
        packageCostFrom, 
        packageCostTo 
      });
    }
  };

  const handleReset = () => {
    setFromDate(minus15); 
    setToDate(tomorrow);
    setFunctionVal('0'); 
    setProjectVal('0'); 
    setContractorVal('0');
    setPackageCostFrom(''); 
    setPackageCostTo('');
    setValidationError('');
    onReset();
  };

  const funcsWithDefault = [
    { label: '-Select Function-', value: '0' },
    ...functionOptions.filter(o => o.value !== '0' && o.label !== '-Select Function-')
  ];

  const projsWithDefault = [
    { label: '-Select  Entity-', value: '0' },
    ...projectOptions.filter(o => o.value !== '0' && !o.label.startsWith('-Select'))
  ];

  const contractorsWithDefault = [
    { label: '-Select Contractor-', value: '0' },
    ...contractorOptions.filter(o => o.value !== '0' && !o.label.startsWith('-Select'))
  ];

  return (
    <div className="filterCard">
      <div className="filterCardHeader">
        <SlidersHorizontal size={14} />
        NFA Generic Report — Filters
      </div>
      <div className="filterCardBody">
        {validationError && (
          <div className="mb-4 p-3 rounded bg-red-50 border border-red-200 text-red-700 text-sm flex items-center gap-2">
            <AlertCircle size={16} />
            <span>{validationError}</span>
          </div>
        )}

        <div className="flex flex-wrap gap-x-6 gap-y-4">
          <div className="flex-1 min-w-[200px]">
            <DatePickerField label="From Date" value={fromDate} onChange={(v) => { setFromDate(v); setValidationError(''); }} required />
          </div>
          <div className="flex-1 min-w-[200px]">
            <DatePickerField label="To Date" value={toDate} onChange={(v) => { setToDate(v); setValidationError(''); }} required />
          </div>
        </div>

        <div className="flex flex-wrap gap-x-6 gap-y-4 mt-4">
          <div className="flex-1 min-w-[200px] max-w-[calc(50%-12px)]">
            <DropdownField label="Function" value={functionVal} onChange={setFunctionVal} options={funcsWithDefault} />
          </div>
        </div>

        <div className="flex flex-wrap gap-x-6 gap-y-4 mt-4">
          <div className="flex-1 min-w-[200px]">
            <DropdownField label="Entity" value={projectVal} onChange={setProjectVal} options={projsWithDefault} />
          </div>
          <div className="flex-1 min-w-[200px]">
            <DropdownField label="Contractor / Vendor" value={contractorVal} onChange={setContractorVal} options={contractorsWithDefault} />
          </div>
        </div>

        <div className="flex flex-wrap gap-x-6 gap-y-4 mt-4">
          <div className="flex-1 min-w-[200px]">
            <InputField 
              label="Package Cost From" 
              value={packageCostFrom} 
              onChange={(v) => { setPackageCostFrom(v); setValidationError(''); }} 
              type="number" 
              placeholder="0.00" 
            />
          </div>
          <div className="flex-1 min-w-[200px]">
            <InputField 
              label="Package Cost To" 
              value={packageCostTo} 
              onChange={(v) => { setPackageCostTo(v); setValidationError(''); }} 
              type="number" 
              placeholder="0.00" 
            />
          </div>
        </div>

        <div className="flex items-center gap-3 mt-4">
          <CommonButton text="Search" variant="primary" onClick={handleSearchClick} />
          <CommonButton text="Reset" variant="danger" onClick={handleReset} />
        </div>
      </div>
    </div>
  );
};

export default GenericReportFilter;
