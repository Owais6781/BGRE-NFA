import * as React from 'react';
import { useState } from 'react';
import {
  Building2, FileText, Wrench, DollarSign,
  Users, Paperclip, Settings, Star,
} from 'lucide-react';
import { InputField }       from '../../../../../Share/Components/InputField';
import { TextAreaField }    from '../../../../../Share/Components/TextAreaField';
import { NumberInput }      from '../../../../../Share/Components/NumberInput';
import { AttachmentUpload } from '../../../../../Share/Components/AttachmentUpload';
import { RichTextEditor }   from '../../../../../Share/Components/RichTextEditor';
import DropdownField        from '../../../../../Share/Components/DropdownField';
import CommonButton         from '../../../../../Share/Components/CommonButton';
import "../../../../../Share/styles/global.scss";

const approverOptions  = [{ label: 'User 1', value: 'user1' }, { label: 'User 2', value: 'user2' }, { label: 'User 3', value: 'user3' }];
const workflowOptions  = [{ label: 'MANUAL', value: 'MANUAL' }, { label: 'AUTO', value: 'AUTO' }];

const NewNFA: React.FC = () => {
  const today = new Date().toLocaleDateString('en-IN');

  const [form, setForm] = useState({
    originatorDept: '', date: today, budgeted: 'Non Budgeted',
    headSubHead: '', project: '', projectSubEntity: '',
    referenceSubject: '', introduction: '', chronology: '',
    commercialProposal: '', paymentTerms: '',
    liquidatedDamage: 'NA', contractPerf: 'NA',
    defectLiability: 'NA', workCompletion: 'Work must be completed and handed over to the satisfaction of Owner on or before -------',
    budgetRemarks: '', packageAwardedCost: '0.00', budget: '0.00',
    capexSerialNo: '', recommendation: '', contractorName: '',
    annexure: '', note1: '', note2: '', note3: '', note4: '',
    workflow: 'MANUAL', doaMatrix: '',
    approvers: Array(12).fill('') as string[],
    attachments: null as FileList | null,
  });

  const set = (key: string, val: any) => setForm(f => ({ ...f, [key]: val }));
  const setApprover = (i: number, val: string) => {
    const arr = [...form.approvers]; arr[i] = val;
    setForm(f => ({ ...f, approvers: arr }));
  };

  return (
    <div className="pageContainer">

      {/* Page Title */}
      <p className="formPageTitle">Fill Below Details to Submit Note for Approval</p>

      {/* ── Section 1: Basic Information ── */}
      <div className="formSection">
        <div className="formSectionHeader"><Building2 size={13} /> Basic Information</div>
        <div className="formSectionBody">
          <div className="flex flex-wrap gap-x-6 gap-y-4 mb-4">
            <div className="flex-1 min-w-[220px]">
              <DropdownField label="Originator Department" value={form.originatorDept}
                onChange={v => set('originatorDept', v)}
                options={[{ label: 'IT', value: 'IT' }, { label: 'HR', value: 'HR' }, { label: 'Finance', value: 'Finance' }]} />
            </div>
            <div className="flex-1 min-w-[220px]">
              <InputField label="Date" value={form.date} onChange={v => set('date', v)} disabled />
            </div>
          </div>
          <div className="flex flex-wrap gap-x-6 gap-y-4 mb-4">
            <div className="flex-1 min-w-[220px]">
              <DropdownField label="Budgeted / Non Budgeted" value={form.budgeted}
                onChange={v => set('budgeted', v)} required
                options={[{ label: 'Budgeted', value: 'Budgeted' }, { label: 'Non Budgeted', value: 'Non Budgeted' }]} />
            </div>
            <div className="flex-1 min-w-[220px]">
              <DropdownField label="Head / Sub-Head" value={form.headSubHead}
                onChange={v => set('headSubHead', v)}
                options={[{ label: 'Head A', value: 'A' }, { label: 'Head B', value: 'B' }]} />
            </div>
          </div>
          <div className="flex flex-wrap gap-x-6 gap-y-4">
            <div className="flex-1 min-w-[220px]">
              <DropdownField label="Project" value={form.project} onChange={v => set('project', v)} required
                options={[{ label: 'Project A', value: 'A' }, { label: 'Project B', value: 'B' }]} />
            </div>
            <div className="flex-1 min-w-[220px]">
              <DropdownField label="Project Sub-Entity" value={form.projectSubEntity}
                onChange={v => set('projectSubEntity', v)}
                options={[{ label: 'Entity 1', value: '1' }, { label: 'Entity 2', value: '2' }]} />
            </div>
          </div>
        </div>
      </div>

      {/* ── Section 2: NFA Details ── */}
      <div className="formSection">
        <div className="formSectionHeader"><FileText size={13} /> NFA Details</div>
        <div className="formSectionBody">
          <div className="mb-4">
            <TextAreaField label="Reference / Subject" value={form.referenceSubject}
              onChange={v => set('referenceSubject', v)} rows={2} placeholder="Reference / Subject" required />
          </div>
          <div className="mb-4">
            <RichTextEditor label="Introduction / Preamble / Background"
              value={form.introduction} onChange={v => set('introduction', v)} />
          </div>
          <div className="mb-4">
            <RichTextEditor label="Chronology" value={form.chronology} onChange={v => set('chronology', v)} />
          </div>
          <div className="mb-4">
            <RichTextEditor label="Commercial Proposal/s" value={form.commercialProposal} onChange={v => set('commercialProposal', v)} />
          </div>
          <div>
            <RichTextEditor label="Payment Terms" value={form.paymentTerms} onChange={v => set('paymentTerms', v)} />
          </div>
        </div>
      </div>

      {/* ── Section 3: Contract Terms ── */}
      <div className="formSection">
        <div className="formSectionHeader"><Wrench size={13} /> Contract Terms</div>
        <div className="formSectionBody">
          <div className="mb-4">
            <TextAreaField label="Liquidated Damage" value={form.liquidatedDamage}
              onChange={v => set('liquidatedDamage', v)} rows={3} required />
          </div>
          <div className="mb-4">
            <TextAreaField label="Contract Performance Guarantee" value={form.contractPerf}
              onChange={v => set('contractPerf', v)} rows={3} required />
          </div>
          <div className="mb-4">
            <TextAreaField label="Defect Liability Period" value={form.defectLiability}
              onChange={v => set('defectLiability', v)} rows={3} required />
          </div>
          <div>
            <TextAreaField label="Work Completion Period" value={form.workCompletion}
              onChange={v => set('workCompletion', v)} rows={3} required />
          </div>
        </div>
      </div>

      {/* ── Section 4: Financial Details ── */}
      <div className="formSection">
        <div className="formSectionHeader"><DollarSign size={13} /> Financial Details</div>
        <div className="formSectionBody">
          <div className="mb-4">
            <TextAreaField label="Budget Remarks" value={form.budgetRemarks}
              onChange={v => set('budgetRemarks', v)} rows={3} />
          </div>
          <div className="flex flex-wrap gap-x-6 gap-y-4 mb-4">
            <div className="flex-1 min-w-[220px]">
              <NumberInput label="Package Awarded Cost" value={form.packageAwardedCost}
                onChange={v => set('packageAwardedCost', v)} required />
            </div>
            <div className="flex-1 min-w-[220px]">
              <NumberInput label="Budget" value={form.budget} onChange={v => set('budget', v)} required />
            </div>
          </div>
          <div>
            <TextAreaField label="Capex Serial No" value={form.capexSerialNo}
              onChange={v => set('capexSerialNo', v)} rows={3} />
          </div>
        </div>
      </div>

      {/* ── Section 5: Recommendation ── */}
      <div className="formSection">
        <div className="formSectionHeader"><Star size={13} /> Recommendation</div>
        <div className="formSectionBody">
          <div className="mb-4">
            <RichTextEditor label="Recommendation with Justification" value={form.recommendation}
              onChange={v => set('recommendation', v)} />
          </div>
          <div>
            <InputField label="Contractor Name" value={form.contractorName}
              onChange={v => set('contractorName', v)} placeholder="Enter Contractor Name" required />
          </div>
        </div>
      </div>

      {/* ── Section 6: Supplementary ── */}
      <div className="formSection">
        <div className="formSectionHeader"><FileText size={13} /> Annexure & Notes</div>
        <div className="formSectionBody">
          <div className="mb-4">
            <TextAreaField label="Annexure / Enclosure (if any)" value={form.annexure}
              onChange={v => set('annexure', v)} rows={3} placeholder="Enter Annexure / Enclosure (if any)" />
          </div>
          {(['note1', 'note2', 'note3', 'note4'] as const).map((key, i) => (
            <div key={key} className="mb-4">
              <TextAreaField label={`Note-${i + 1}`} value={(form as any)[key]}
                onChange={v => set(key, v)} rows={3} placeholder={`Enter Note-${i + 1}`} />
            </div>
          ))}
        </div>
      </div>

      {/* ── Section 7: Workflow & Approvers ── */}
      <div className="formSection">
        <div className="formSectionHeader"><Settings size={13} /> Workflow & Approvers</div>
        <div className="formSectionBody">
          <div className="flex flex-wrap gap-x-6 gap-y-4 mb-5">
            <div className="flex-1 min-w-[220px]">
              <DropdownField label="Workflow Execution" value={form.workflow}
                onChange={v => set('workflow', v)} options={workflowOptions} required />
            </div>
            <div className="flex-1 min-w-[220px]">
              <DropdownField label="DOA Matrix" value={form.doaMatrix}
                onChange={v => set('doaMatrix', v)}
                options={[{ label: 'Matrix A', value: 'A' }, { label: 'Matrix B', value: 'B' }]} />
            </div>
          </div>

          <div className="formSectionHeader" style={{ margin: '0 -24px 16px', borderRadius: 0 }}>
            <Users size={13} /> Approver Levels
          </div>

          {[0, 4, 8].map((startIdx) => (
            <div key={startIdx} className="flex flex-wrap gap-x-4 gap-y-4 mb-4">
              {[0, 1, 2, 3].map((offset) => {
                const idx = startIdx + offset;
                return (
                  <div key={idx} className="flex-1 min-w-[180px]">
                    <DropdownField
                      label={`Approver Level-${idx + 1}`}
                      value={form.approvers[idx]}
                      onChange={v => setApprover(idx, v)}
                      options={approverOptions}
                      required={idx === 0}
                    />
                  </div>
                );
              })}
            </div>
          ))}
        </div>
      </div>

      {/* ── Section 8: Attachments ── */}
      <div className="formSection">
        <div className="formSectionHeader"><Paperclip size={13} /> Attachments</div>
        <div className="formSectionBody">
          <AttachmentUpload label="Attachment" onChange={(files) => set('attachments', files)} />
        </div>
      </div>

      {/* ── Action Buttons ── */}
      <div className="flex items-center justify-center gap-4 pt-2 pb-4">
        <CommonButton text="Submit"       variant="primary"   onClick={() => console.log('Submit', form)} />
        <CommonButton text="Save as Draft" variant="secondary" onClick={() => console.log('Draft', form)} />
        <CommonButton text="Cancel"       variant="danger"    onClick={() => console.log('Cancel')} />
      </div>

    </div>
  );
};

export default NewNFA;
