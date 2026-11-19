import { INFAFormData } from '../Services/NFAService';
import { stripHtml } from '../utils';
import * as reqMod from './html2pdf.bundle.min.js';
export interface IPdfApproverRow {
  levelLabel: string;
  designation: string;
  name: string;
  remarks: string;
  status: string;
  date: string;
}

export interface IPdfGeneratorData {
  form: INFAFormData;
  users: { label: string; value: string; email?: string; designation?: string; jobTitle?: string }[];
  history: any[];
  comments: any[];
}

/**
 * Formats a date into DD/MM/YYYY HH:MM AM/PM
 */
export const formatDateTime = (dateObj: Date): string => {

  const day = ("0" + dateObj.getDate()).slice(-2);
  const month = ("0" + (dateObj.getMonth() + 1)).slice(-2);
  const year = dateObj.getFullYear();

  let hours = dateObj.getHours();
  const minutes = ("0" + dateObj.getMinutes()).slice(-2);
  const ampm = hours >= 12 ? 'PM' : 'AM';
  hours = hours % 12;
  hours = hours ? hours : 12;

  return `${day}/${month}/${year} ${hours}:${minutes} ${ampm}`;
};

/**
 * Maps ordinal numbers for approver level labels (1st Approver, 2nd Approver, etc.)
 */
const getOrdinalLabel = (num: number): string => {
  const suffixes = ['th', 'st', 'nd', 'rd'];
  const v = num % 100;
  return `(${num}${suffixes[(v - 20) % 10] || suffixes[v] || suffixes[0]} Approver)`;
};

/**
 * Enhances narrative content for PDF:
 * 1. If HTML table exists, enhances with .excelDataTable styling, headers, and numeric cell alignments.
 * 2. If tab-separated text (TSV from Excel), converts to a neat .excelDataTable.
 * 3. If plain text with line breaks, wraps in pre-wrap styling to preserve all spacing and line breaks.
 */
export const formatPdfNarrativeContent = (content: string | undefined): string => {
  if (!content) return '';
  const trimmed = content.trim();
  if (!trimmed) return '';

  // 1. If HTML table exists
  if (trimmed.includes('<table') || trimmed.includes('<TABLE')) {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(trimmed, 'text/html');
      const tables = doc.querySelectorAll('table');

      tables.forEach((table) => {
        table.classList.add('excelDataTable');
        table.removeAttribute('width');
        table.style.width = '100%';
        table.style.maxWidth = '100%';
        table.style.tableLayout = 'auto';

        const trs = Array.from(table.querySelectorAll('tr'));
        trs.forEach((tr, rIdx) => {
          if (rIdx === 0 || tr.parentElement?.tagName.toLowerCase() === 'thead') {
            tr.classList.add('excelHeaderRow');
          }
          const cells = Array.from(tr.querySelectorAll('th, td'));
          cells.forEach((cell) => {
            cell.removeAttribute('width');
            const hCell = cell as HTMLElement;
            if (hCell.style) {
              hCell.style.width = '';
              hCell.style.maxWidth = 'none';
              hCell.style.position = 'static';
              hCell.style.wordBreak = 'break-word';
              hCell.style.overflowWrap = 'break-word';
            }
            const text = cell.textContent?.trim() || '';
            if (/^[-+]?₹?\s*[\d,]+(\.\d+)?%?$/.test(text) && !/^\d{2}[-/]\d{2}[-/]\d{4}$/.test(text)) {
              cell.classList.add('excelCellNumeric');
            }
          });
        });
      });

      return doc.body.innerHTML;
    } catch (e) {
      return trimmed;
    }
  }

  // 2. Tab-separated text (copied from Excel into a plain textarea)
  if (trimmed.includes('\t') && trimmed.includes('\n')) {
    const lines = trimmed.split(/\r?\n/).filter((l) => l.trim().length > 0);
    if (lines.length > 1) {
      const rows = lines.map((line) => line.split('\t').map((c) => c.trim()));
      return `
        <table class="excelDataTable">
          <thead>
            <tr class="excelHeaderRow">
              ${rows[0].map((h) => `<th>${h}</th>`).join('')}
            </tr>
          </thead>
          <tbody>
            ${rows.slice(1).map((r) => `
              <tr>
                ${r.map((c) => {
        const isNum = /^[-+]?₹?\s*[\d,]+(\.\d+)?%?$/.test(c) && !/^\d{2}[-/]\d{2}[-/]\d{4}$/.test(c);
        const classes = [isNum ? 'excelCellNumeric' : ''].filter(Boolean).join(' ');
        return `<td class="${classes}">${c}</td>`;
      }).join('')}
              </tr>
            `).join('')}
          </tbody>
        </table>
      `;
    }
  }

  // 3. Plain text: preserve newlines and spaces
  if (!/<[a-z][\s\S]*>/i.test(trimmed)) {
    return `<div style="white-space: pre-wrap; word-break: break-word;">${trimmed}</div>`;
  }

  return trimmed;
};

/**
 * Generates the clean HTML string representing the PDF template matching the exact BGRE layout.
 */
export const buildNfaPdfHtml = (data: IPdfGeneratorData): string => {
  const { form, users, history, comments } = data;

  const requestNo = form.requestNo || 'NFA-Draft';
  const statusRaw = form.status || 'CLOSED';
  const statusUpper = statusRaw.toUpperCase();

  // Status Badge Colors
  let statusBg = '#fef08a';
  let statusColor = '#854d0e';
  if (statusUpper.includes('APPROV')) {
    statusBg = '#dcfce7';
    statusColor = '#15803d';
  } else if (statusUpper.includes('REJECT')) {
    statusBg = '#fee2e2';
    statusColor = '#b91c1c';
  } else if (statusUpper.includes('PENDING')) {
    statusBg = '#fef3c7';
    statusColor = '#b45309';
  } else if (statusUpper.includes('CLOSE')) {
    statusBg = '#fef08a';
    statusColor = '#854d0e';
  }



  // ── Approver Workflow State ──
  const totalApprovers = form.approvers.filter((a) => Boolean(a)).length;
  const currentWfLevel = typeof form.level === 'number' ? form.level : parseInt(String(form.level || '1'), 10) || 1;
  const isOverallApproved = (statusUpper === 'APPROVED' || statusUpper === 'COMPLETED' || statusUpper === 'CLOSED') && !statusUpper.includes('PENDING');
  const isOverallRejected = statusUpper === 'REJECTED';
  const isDraftStatus = statusUpper.includes('DRAFT');

  // Level completed calculation (Requirement 2)
  let completedLevelNum = 0;
  if (isOverallApproved) {
    completedLevelNum = totalApprovers;
  } else if (currentWfLevel > 1) {
    completedLevelNum = currentWfLevel - 1;
  }
  const completedLevelText = completedLevelNum > 0
    ? `Level ${completedLevelNum} (${getOrdinalLabel(completedLevelNum).replace(/[()]/g, '')})`
    : 'None (Pending Level 1)';


  // Initiator and Date
  const initHistory = history.find((h) => h.action === 'Initiated');
  const initiatorName =
    initHistory?.user ||
    (form.initiatorId
      ? users.find((u) => u.value === String(form.initiatorId))?.label
      : '') ||
    'Mohd Ibad';
  const initiatedOn = initHistory?.date || form.date || '';

  // Formatted Approval Amount
  let formattedAmount = form.packageAwardedCost || '0.00';
  const numericAmount = parseFloat(String(form.packageAwardedCost || '0').replace(/,/g, ''));
  if (!isNaN(numericAmount) && numericAmount > 0) {
    formattedAmount = numericAmount.toLocaleString('en-IN', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
  }

  // Region / Tenant
  const region =
    form.location ||
    (form.projectSubEntity && form.projectSubEntity.length > 0
      ? form.projectSubEntity[0]
      : '') ||
    '';
  const tenantName = form.contractorName || '';
  const provisionedForBadDebts = form.budgetedNonBudgeted || '';


  // ── Construct Grid Items (Exclude empty pairs) ──
  const rawFields: { key: string; value: string }[] = [];

  const addField = (k: string, v: string | undefined): void => {
    if (v && v.trim() && v.trim() !== '-') {
      rawFields.push({ key: k, value: v.trim() });
    }
  };

  addField('REQUEST NO.', requestNo);
  addField('FUNCTION', form.originatorDept);
  addField('INITIATED BY', initiatorName);
  addField('INITIATED ON', initiatedOn);
  addField('ENTITY ', form.project);
  addField('SUB-ENTITY', region);
  addField('TYPE OF APPROVAL', form.natureOfApproval || form.workflow);
  addField('CONTRACTOR NAME', tenantName);
  addField('BUDGETED / NON BUDGETED', provisionedForBadDebts);
  addField('APPROVAL AMOUNT', `₹ ${formattedAmount}`);
  addField('APPROVAL COMPLETED UP TO', completedLevelText);

  // Also include other populated fields if present
  if (form.RFQRefNo) addField('RFQ REF NO', form.RFQRefNo);
  if (form.budget && parseFloat(form.budget) > 0) addField('BUDGET', `₹ ${form.budget}`);
  if (form.doaMatrix) addField('DOA MATRIX', form.doaMatrix);
  if (form.completeIndentPRReceivedFromDate) {
    const date = new Date(form.completeIndentPRReceivedFromDate);
    const formattedDate = date.toLocaleDateString('en-GB');
    addField('COMPLETE INDENT/PR RECEIVED FROM & DATE.', formattedDate);
  }
  if (form.boqSpecificationsRevisionsByDate) {
    const date = new Date(form.boqSpecificationsRevisionsByDate);
    const formattedDate = date.toLocaleDateString('en-GB');
    addField('BOQ/SPECIFICATIONS REVISIONS BY & DATE', formattedDate);
  }

  // Group fields into pairs of 2 for table rows
  const gridRows: { key1: string; val1: string; key2: string; val2: string }[] = [];
  for (let i = 0; i < rawFields.length; i += 2) {
    gridRows.push({
      key1: rawFields[i].key,
      val1: rawFields[i].value,
      key2: rawFields[i + 1] ? rawFields[i + 1].key : '',
      val2: rawFields[i + 1] ? rawFields[i + 1].value : ''
    });
  }

  // ── Narrative Sections (Only include populated fields) ──
  const narrativeSections: { title: string; content: string }[] = [];

  const addNarrative = (title: string, content: string | undefined): void => {
    if (content && stripHtml(content).trim().length > 0) {
      narrativeSections.push({ title, content: formatPdfNarrativeContent(content) });
    }
  };

  let subjectText = form.referenceSubject || '';
  if (isDraftStatus) {
    if (/^Draft\//i.test(subjectText.trim()) || subjectText.trim() === requestNo.trim()) {
      subjectText = '';
    }
  }
  addNarrative('SUBJECT', subjectText);
  addNarrative('INTRODUCTION / PREAMBLE / BACKGROUND', form.introduction);
  // Latest approved level remark (Requirements 3, 4, 5)
  let latestApprovedRemark = '';
  for (let lvl = totalApprovers; lvl >= 1; lvl--) {
    const isLvlApproved = isOverallApproved || currentWfLevel > lvl;
    if (isLvlApproved) {
      const c = comments.find((cm) => cm.level && parseInt(String(cm.level), 10) === lvl);
      if (c?.remarks && c.remarks.trim() && c.remarks.trim().toLowerCase() !== 'no remarks') {
        latestApprovedRemark = c.remarks.trim();
        break;
      }
      const h = history.find((hi) => hi.level && parseInt(String(hi.level), 10) === lvl);
      if (h?.remarks && h.remarks.trim() && h.remarks.trim().toLowerCase() !== 'no remarks') {
        latestApprovedRemark = h.remarks.trim();
        break;
      }
    }
  }
  const effectiveRemark = latestApprovedRemark || form.remark;
  addNarrative('Remarks', effectiveRemark);
  addNarrative('CHRONOLOGY', form.chronology);
  addNarrative('COMMERCIAL PROPOSAL', form.commercialProposal);
  addNarrative('PAYMENT TERMS', form.paymentTerms);
  // Contract Terms block
  const contractParts: string[] = [];
  if (form.liquidatedDamage && form.liquidatedDamage.trim() && !/^(na|n\/a)$/i.test(form.liquidatedDamage.trim())) {
    contractParts.push(`<strong>Liquidated Damage:</strong> ${form.liquidatedDamage}`);
  }
  if (form.contractPerf && form.contractPerf.trim() && !/^(na|n\/a)$/i.test(form.contractPerf.trim())) {
    contractParts.push(`<strong>Contract Performance:</strong> ${form.contractPerf}`);
  }
  if (form.defectLiability && form.defectLiability.trim() && !/^(na|n\/a)$/i.test(form.defectLiability.trim())) {
    contractParts.push(`<strong>Defect Liability Period:</strong> ${form.defectLiability}`);
  }
  if (form.workCompletion && form.workCompletion.trim() && !/^(na|n\/a)$/i.test(form.workCompletion.trim())) {
    contractParts.push(`<strong>Work Completion:</strong> ${form.workCompletion}`);
  }
  if (contractParts.length > 0) {
    narrativeSections.push({
      title: 'CONTRACT TERMS',
      content: formatPdfNarrativeContent(contractParts.join('<br/>'))
    });
  }


  addNarrative('BUDGET REMARKS', form.budgetRemarks);
  addNarrative('RECOMMENDATION WITH JUSTIFICATION', form.recommendation);
  addNarrative('ANNEXURE / ENCLOSURE', form.annexure);
  addNarrative('NOTE 1', form.note1);
  addNarrative('NOTE 2', form.note2);
  addNarrative('NOTE 3', form.note3);
  addNarrative('NOTE 4', form.note4);



  // ── Approver Rows parsed from WorkflowApprovalHistory ──
  const approverRows: IPdfApproverRow[] = [];
  const rawWfHistory = (
    form.workflowApprovalHistory ||
    form.workflowHistory ||
    (form as any).WorkflowApprovalHistory ||
    (form as any).WorkflowHistory ||
    ''
  ).trim();

  // 1. Try parsing HTML table if WorkflowApprovalHistory contains table markup
  if (rawWfHistory.includes('<table') || rawWfHistory.includes('<TABLE') || rawWfHistory.includes('<tr') || rawWfHistory.includes('<TR')) {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(rawWfHistory, 'text/html');
      const trs = Array.from(doc.querySelectorAll('tr'));
      trs.forEach((tr) => {
        const cells = Array.from(tr.querySelectorAll('td, th')).map((td) => td.textContent?.trim() || '');
        if (cells.length >= 5) {
          const levelLabel = cells[0];
          const designation = cells[1] || '';
          const name = cells[2] || '';
          const remarks = cells[3] || '';

          const rawStatus = cells[4] || 'Approved';
          const status = rawStatus.toLowerCase() === 'completed' ? 'Approved' : rawStatus;
          const date = cells[5] || '';
          if (levelLabel && (name || status)) {
            approverRows.push({
              levelLabel,
              designation,
              name,
              remarks,
              status,
              date
            });
          }
        }
      });
    } catch (e) {
      // ignore table parse error
    }
  }

  // 2. Parse text lines from WorkflowApprovalHistory
  if (approverRows.length === 0 && rawWfHistory) {
    const cleanHistoryText = rawWfHistory
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/br>/gi, '\n');

    const lines = cleanHistoryText.split('\n');
    const levelMap = new Map<number, IPdfApproverRow>();

    for (const rawLine of lines) {
      const cleanLine = rawLine
        .replace(/<[^>]*>/g, '')
        .replace(/&#58;/g, ':')
        .replace(/&nbsp;/g, ' ')
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .trim();

      if (!cleanLine) continue;

      // Match 1: NFA Request [Action] by [User] at Level [N]. on [Date]
      const actionMatch = cleanLine.match(
        /^NFA Request\s+(.*?)\s+by\s+(.*?)\s+at Level\s+(\d+)\.?(?:\s+on\s+(.*))?$/i
      );
      if (actionMatch) {
        const rawAction = actionMatch[1].trim();
        const userName = actionMatch[2].trim();
        const levelNum = parseInt(actionMatch[3].trim(), 10);
        const date = (actionMatch[4] || '').trim();

        const commentEntry = comments.find((c) => c.level && parseInt(String(c.level), 10) === levelNum);
        const histEntry = history.find((h) => h.level && parseInt(String(h.level), 10) === levelNum);
        const rawRemark = commentEntry?.remarks || histEntry?.remarks || '';
        const remarks = (rawRemark && rawRemark.trim().toLowerCase() !== 'no remarks') ? rawRemark.trim() : '';

        const userObj = users.find((u) => u.label?.trim().toLowerCase() === userName.toLowerCase());
        const approverDetail = form.approverDetails ? form.approverDetails[levelNum - 1] : null;
        const designation = (
          approverDetail?.designation ||
          userObj?.designation ||
          (userObj as any)?.jobTitle ||
          (userObj as any)?.department ||
          ''
        ).trim();

        // Display "Approved" in the PDF table even if stored as "Completed" in SharePoint
        const displayStatus = rawAction.toLowerCase() === 'completed' ? 'Approved' : rawAction;
        levelMap.set(levelNum, {
          levelLabel: getOrdinalLabel(levelNum),
          designation,
          name: userName,
          remarks,
          status: displayStatus,
          date
        });
        continue;
      }



      // Match 2: NFA Pending at [User] at level-[N] on [Date] OR Pending at [User] at level-[N] on [Date]
      const pendingMatch = cleanLine.match(
        /^(?:NFA\s+)?Pending at\s+(.*?)\s+at level-(\d+)(?:\s+on\s+(.*))?$/i
      );
      if (pendingMatch) {
        const userName = pendingMatch[1].trim();
        const levelNum = parseInt(pendingMatch[2].trim(), 10);
        const date = (pendingMatch[3] || '').trim();

        const existing = levelMap.get(levelNum);
        if (!existing || existing.status === 'Pending') {
          const userObj = users.find((u) => u.label?.trim().toLowerCase() === userName.toLowerCase());
          const approverDetail = form.approverDetails ? form.approverDetails[levelNum - 1] : null;
          const designation = (
            approverDetail?.designation ||
            userObj?.designation ||
            (userObj as any)?.jobTitle ||
            (userObj as any)?.department ||
            ''
          ).trim();

          levelMap.set(levelNum, {
            levelLabel: getOrdinalLabel(levelNum),
            designation,
            name: userName,
            remarks: '',
            status: 'Pending',
            date
          });
        }
        continue;
      }

      // Match 3: Query responded by [User] at Level [N]. on [Date]
      const queryRespMatch = cleanLine.match(
        /^Query responded by\s+(.*?)\s+at Level\s+(\d+)\.?(?:\s+on\s+(.*))?$/i
      );
      if (queryRespMatch) {
        const userName = queryRespMatch[1].trim();
        const levelNum = parseInt(queryRespMatch[2].trim(), 10);
        const date = (queryRespMatch[3] || '').trim();

        const existing = levelMap.get(levelNum);
        if (!existing || existing.status === 'Pending') {
          const userObj = users.find((u) => u.label?.trim().toLowerCase() === userName.toLowerCase());
          const approverDetail = form.approverDetails ? form.approverDetails[levelNum - 1] : null;
          const designation = (
            approverDetail?.designation ||
            userObj?.designation ||
            (userObj as any)?.jobTitle ||
            (userObj as any)?.department ||
            ''
          ).trim();

          levelMap.set(levelNum, {
            levelLabel: getOrdinalLabel(levelNum),
            designation,
            name: userName,
            remarks: '',
            status: 'Query Responded',
            date
          });
        }
        continue;
      }
    }

    const sortedLevels = Array.from(levelMap.keys()).sort((a, b) => a - b);
    for (const lvl of sortedLevels) {
      const row = levelMap.get(lvl);
      if (row) approverRows.push(row);
    }
  }

  // 3. Fallback to form.approvers if WorkflowApprovalHistory yielded no rows and not in draft
  if (approverRows.length === 0 && !isDraftStatus && form.approvers && form.approvers.length > 0) {
    form.approvers.forEach((appId, idx) => {
      if (!appId) return;

      const levelNum = idx + 1;
      const shouldInclude = isOverallApproved || levelNum <= currentWfLevel;
      if (!shouldInclude) return;

      const levelLabel = getOrdinalLabel(levelNum);
      const userObj = users.find((u) => u.value === String(appId));
      let userName = userObj ? userObj.label : `Approver ${levelNum}`;

      let status = 'Pending';
      if (isOverallApproved) {
        status = (levelNum === totalApprovers && statusUpper.includes('CLOSE')) ? 'Closed' : 'Approved';
      } else if (isOverallRejected) {
        if (levelNum < currentWfLevel) status = 'Approved';
        else if (levelNum === currentWfLevel) status = 'Rejected';
        else status = 'Pending';
      } else {
        if (levelNum < currentWfLevel) {
          status = 'Approved';
        } else {
          status = 'Pending';
        }
      }

      let remarks = '';
      let date = '';
      if (status === 'Approved' || status === 'Closed' || status === 'Rejected' || status === 'Completed') {
        const commentEntry = comments.find((c) => c.level && parseInt(String(c.level), 10) === levelNum);
        const histEntry = history.find((h) => h.level && parseInt(String(h.level), 10) === levelNum);

        const rawRemark = commentEntry?.remarks || histEntry?.remarks || '';
        remarks = (rawRemark && rawRemark.trim().toLowerCase() !== 'no remarks') ? rawRemark.trim() : '';
        date = histEntry?.date || commentEntry?.date || '';

        if (commentEntry?.user && commentEntry.user !== 'System') {
          userName = commentEntry.user;
        } else if (histEntry?.user && histEntry.user !== 'System') {
          userName = histEntry.user;
        }
      }

      const approverDetail = form.approverDetails ? form.approverDetails[idx] : null;
      const rawDesignation =
        approverDetail?.designation ||
        userObj?.designation ||
        (userObj as any)?.jobTitle ||
        (userObj as any)?.department;

      let designation = rawDesignation && rawDesignation.trim() ? rawDesignation.trim() : '';

      approverRows.push({
        levelLabel,
        designation,
        name: userName,
        remarks,
        status,
        date
      });
    });
  }

  const generatedDateStr = formatDateTime(new Date());

  // ── Render HTML ──
  return `
    <style>
      #nfa-pdf-root {
        width: 100%;
        max-width: 100%;
        margin: 0;
        padding: 0;
        box-sizing: border-box;
      }
      #nfa-pdf-root * {
        box-sizing: border-box;
      }
      .pdf-section {
        page-break-inside: avoid !important;
        break-inside: avoid !important;
        margin-bottom: 14px;
      }
      tr {
        page-break-inside: avoid !important;
        break-inside: avoid !important;
      }

      /* ── Reset and disable any position sticky from global styles ── */
      #nfa-pdf-root .excelCellPinned,
      .excelCellPinned {
        position: static !important;
        left: auto !important;
        top: auto !important;
        box-shadow: none !important;
        font-weight: 600 !important;
      }
      #nfa-pdf-root th, #nfa-pdf-root td {
        position: static !important;
      }

      /* ── Prevent any table or element from overflowing horizontally or cutting off ── */
      #nfa-pdf-root table,
      #nfa-pdf-root .excelDataTable,
      #nfa-pdf-root .pdf-narrative-content table {
        width: 100% !important;
        max-width: 100% !important;
        min-width: 0 !important;
        table-layout: auto !important;
        border-collapse: collapse !important;
        margin: 6px 0 !important;
        box-sizing: border-box !important;
      }

      #nfa-pdf-root th, #nfa-pdf-root td,
      #nfa-pdf-root .excelDataTable th, #nfa-pdf-root .excelDataTable td,
      #nfa-pdf-root .pdf-narrative-content table th, #nfa-pdf-root .pdf-narrative-content table td {
        padding: 5px 8px !important;
        border: 1px solid #cbd5e1 !important;
        line-height: 1.35 !important;
        vertical-align: middle !important;
        box-sizing: border-box !important;
        word-break: break-word !important;
        overflow-wrap: break-word !important;
        white-space: normal !important;
        font-size: 10.5px !important;
      }

      #nfa-pdf-root .excelDataTable tr:first-child th,
      #nfa-pdf-root .excelDataTable tr:first-child td,
      #nfa-pdf-root .excelDataTable thead th,
      #nfa-pdf-root .excelDataTable thead td,
      #nfa-pdf-root .excelDataTable tr.excelHeaderRow th,
      #nfa-pdf-root .excelDataTable tr.excelHeaderRow td,
      #nfa-pdf-root .pdf-narrative-content table tr:first-child th,
      #nfa-pdf-root .pdf-narrative-content table thead th {
        background-color: #f1f5f9 !important;
        color: #0f172a !important;
        font-weight: 700 !important;
        text-align: center !important;
        border-bottom: 2px solid #94a3b8 !important;
      }

      #nfa-pdf-root .excelDataTable tbody tr:nth-child(even),
      #nfa-pdf-root .excelDataTable tr:not(:first-child):nth-child(even),
      #nfa-pdf-root .pdf-narrative-content table tbody tr:nth-child(even) {
        background-color: #f8fafc !important;
      }

      #nfa-pdf-root .excelCellNumeric {
        text-align: right !important;
      }

      .pdf-narrative-content {
        padding: 8px 10px;
        background-color: #ffffff;
        font-size: 11px;
        color: #1e293b;
        line-height: 1.45;
        min-height: 24px;
        word-break: break-word;
        overflow-wrap: break-word;
        overflow: visible !important;
      }

      .pdf-narrative-content p {
        margin: 0 0 6px 0;
        line-height: 1.45;
        word-break: break-word;
        overflow-wrap: break-word;
      }
    </style>
    <div id="nfa-pdf-root" style="font-family: Arial, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; color: #1e293b; background: #ffffff; padding: 12px 14px; box-sizing: border-box; width: 100%; max-width: 100%; line-height: 1.4;">

      <!-- ── TOP HEADER ── -->
      <div class="pdf-section" style="display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #5d1020; padding-bottom: 12px; margin-bottom: 16px; page-break-inside: avoid; break-inside: avoid;">
        <div style="display: flex; align-items: baseline; gap: 14px;">
          <span style="font-size: 26px; font-weight: 900; color: #4A081E; letter-spacing: -0.5px; text-decoration: underline; text-underline-offset: 4px;">BGRE</span>
          <span style="font-size: 17px; font-weight: 700; color: #0f172a;">Note for Approval (NFA)</span>
        </div>
        <div style="display: flex; align-items: center; gap: 14px;">
          <span style="font-size: 15px; font-weight: 700; color: #0f172a;">${requestNo}</span>
          <span style="font-size: 11px; font-weight: 800; text-transform: uppercase; padding: 4px 12px; border-radius: 9999px; background-color: ${statusBg}; color: ${statusColor}; letter-spacing: 0.5px;">${statusUpper}</span>
        </div>
      </div>

      <!-- ── REQUEST DETAILS ── -->
      <div class="pdf-section" style="margin-bottom: 16px; border: 1px solid #cbd5e1; border-radius: 4px; overflow: visible; page-break-inside: avoid; break-inside: avoid;">
        <table style="width: 100%; border-collapse: collapse; font-size: 11px;">
          <tbody>
            ${gridRows
      .map(
        (row) => `
              <tr>
                <td style="background-color: #f8fafc; color: #475569; font-weight: 800; text-transform: uppercase; padding: 7px 10px; width: 20%; border: 1px solid #cbd5e1; word-break: break-word;">${row.key1
          }</td>
                <td style="background-color: #ffffff; color: #0f172a; font-weight: 600; padding: 7px 10px; width: 30%; border: 1px solid #cbd5e1; word-break: break-word;">${row.val1
          }</td>
                <td style="background-color: #f8fafc; color: #475569; font-weight: 800; text-transform: uppercase; padding: 7px 10px; width: 20%; border: 1px solid #cbd5e1; word-break: break-word;">${row.key2
          }</td>
                <td style="background-color: #ffffff; color: #0f172a; font-weight: 600; padding: 7px 10px; width: 30%; border: 1px solid #cbd5e1; word-break: break-word;">${row.val2
          }</td>
              </tr>
            `
      )
      .join('')}
          </tbody>
        </table>
      </div>

      <!-- ── NARRATIVE SECTIONS ── -->
      ${narrativeSections
      .map(
        (sec) => `
        <div class="pdf-section" style="margin-bottom: 14px; page-break-inside: avoid; break-inside: avoid;">
          <div style="color: #4A081E; font-size: 11px; font-weight: 800; text-transform: uppercase; padding: 4px 2px 4px 2px; letter-spacing: 0.5px; page-break-after: avoid; break-after: avoid;">
            ${sec.title}
          </div>
          <div style="border: 1px solid #cbd5e1; border-radius: 4px; overflow: visible;">
            <div class="pdf-narrative-content">
              ${sec.content}
            </div>
          </div>
        </div>
      `
      )
      .join('')}

      <!-- ── APPROVALS TABLE ── -->
      ${(approverRows.length > 0 || isDraftStatus)
      ? `
        <div class="pdf-section" style="margin-top: 18px; margin-bottom: 18px; page-break-inside: avoid; break-inside: avoid;">
          <div style="background-color: #f8fafc; text-align: center; color: #4A081E; font-size: 11px; font-weight: 800; text-transform: uppercase; padding: 6px 12px; letter-spacing: 0.5px; border: 1px solid #cbd5e1; border-bottom: none; ">
            APPROVALS
          </div>
          
          <table style="
  width: 100%;
  border-collapse: collapse;
  border-spacing: 0;
  font-size: 11px;
  border: 1px solid #cbd5e1;
  margin: 0;
  padding: 0;
">  
          <thead>
              <tr style="background-color: #4A081E; color: #ffffff;">
                <th style="padding: 8px 10px; font-weight: 800; text-align: left; text-transform: uppercase; border: 1px solid #4A081E; width: 14%;">LEVEL</th>
                <th style="padding: 8px 10px; font-weight: 800; text-align: left; text-transform: uppercase; border: 1px solid #4A081E; width: 18%;">DESIGNATION</th>
                <th style="padding: 8px 10px; font-weight: 800; text-align: left; text-transform: uppercase; border: 1px solid #4A081E; width: 20%;">NAME</th>
                <th style="padding: 8px 10px; font-weight: 800; text-align: left; text-transform: uppercase; border: 1px solid #4A081E; width: 24%;">REMARKS</th>
                <th style="padding: 8px 10px; font-weight: 800; text-align: center; text-transform: uppercase; border: 1px solid #4A081E; width: 12%;">STATUS</th>
                <th style="padding: 8px 10px; font-weight: 800; text-align: left; text-transform: uppercase; border: 1px solid #4A081E; width: 13%;">DATE</th>
              </tr>
            </thead>
            <tbody>
              ${isDraftStatus
        ? `
                  <tr style="background-color: #ffffff;">
                    <td colspan="6" style="padding: 14px 10px; text-align: center; color: #64748b; font-style: italic; border: 1px solid #cbd5e1; font-weight: 600; font-size: 11px;">
                      No Approval Action Recorded
                    </td>
                  </tr>
                `
        : approverRows
          .map((app, i) => {
            let badgeBg = '#f1f5f9';
            let badgeText = '#475569';
            if (app.status === 'Approved' || app.status === 'Completed') {
              badgeBg = '#dcfce7';
              badgeText = '#15803d';
            } else if (app.status === 'Closed') {
              badgeBg = '#e0f2fe';
              badgeText = '#0369a1';
            } else if (app.status === 'Rejected') {
              badgeBg = '#fee2e2';
              badgeText = '#b91c1c';
            } else if (app.status === 'Pending') {
              badgeBg = '#fef3c7';
              badgeText = '#b45309';
            }

            const rowBg = i % 2 === 1 ? '#f8fafc' : '#ffffff';

            return `
                  <tr style="background-color: ${rowBg};">
                    <td style="padding: 7px 10px; border: 1px solid #cbd5e1; font-weight: 700; color: #334155;">${app.levelLabel}</td>
                    <td style="padding: 7px 10px; border: 1px solid #cbd5e1; color: #475569;">${app.designation}</td>
                    <td style="padding: 7px 10px; border: 1px solid #cbd5e1; font-weight: 700; color: #0f172a;">${app.name}</td>
                    <td style="padding: 7px 10px; border: 1px solid #cbd5e1; color: #475569; word-break: break-word;">${app.remarks || ''}</td>
                    <td style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: center;">
                      <span style="display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 10px; font-weight: 800; background-color: ${badgeBg}; color: ${badgeText};">${app.status}</span>
                    </td>
                    <td style="padding: 7px 10px; border: 1px solid #cbd5e1; color: #475569; font-size: 10px;">${app.date || ''}</td>
                  </tr>
                `;
          })
          .join('')}
            </tbody>
          </table>
        </div>
      `
      : ''
    }

      <!-- ── FOOTER ── -->
      <div class="pdf-no-break" style="display: flex; justify-content: space-between; align-items: center; font-size: 10px; color: #64748b; border-top: 1px solid #e2e8f0; padding-top: 10px; margin-top: 16px; page-break-inside: avoid; break-inside: avoid;">
        <span>This is an official system-generated Note for Approval from BGRE and does not require physical signature.</span>
        <span>Generated on: ${generatedDateStr}</span>
      </div>

    </div>
  `;
};

// Cached local html2pdf module reference
let html2pdfModule: any = null;
let html2pdfPromise: Promise<any> | null = null;



const getHtml2Pdf = async (): Promise<any> => {
  const win = window as any;

  if (win.html2pdf) return win.html2pdf;
  if (html2pdfModule) return html2pdfModule;

  if (!html2pdfPromise) {
    html2pdfPromise = (async () => {
      try {
        const mod: any = await import('./html2pdf.bundle.min.js');

        const moduleValue = mod?.default || mod;

        if (moduleValue) return moduleValue;

        const fallback: any = await import('html2pdf.js');
        return fallback?.default || fallback;

      } catch (e) {
        console.warn('import html2pdf bundle failed:', e);

        try {
          const fallback: any = await import('html2pdf.js');
          return fallback?.default || fallback;
        } catch (e2) {
          console.warn('require html2pdf bundle failed:', e2);
          return null;
        }
      }
    })();
  }

  const result = await html2pdfPromise;

  if (result) {
    // eslint-disable-next-line require-atomic-updates
    html2pdfModule = result;
  } else {
    // eslint-disable-next-line require-atomic-updates
    html2pdfPromise = null;
  }

  return result;
};
/**
 * Downloads the generated PDF.
 * Directly saves the A4 PDF automatically to the Downloads folder.
 */
export const downloadNfaPdf = async (data: IPdfGeneratorData): Promise<void> => {
  const requestNo = (data.form.requestNo || 'NFA-Note').trim();
  // Sanitize filename for operating system (replace illegal characters with '-')
  const cleanFileName = requestNo.replace(/[/\\?%*:|"<>]/g, '-');

  const htmlContent = buildNfaPdfHtml(data);

  const html2pdf = await getHtml2Pdf();

  if (html2pdf) {
    const opt = {
      margin: [8, 8, 8, 8],
      filename: `${cleanFileName}.pdf`,
      image: { type: 'jpeg', quality: 0.98 },
      html2canvas: {
        scale: 2,
        useCORS: true,
        logging: false,
        backgroundColor: '#ffffff',
        scrollX: 0,
        scrollY: 0
      },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
      pagebreak: {
        mode: ['avoid-all', 'css', 'legacy'],
        avoid: ['.pdf-section', 'tr']
      }
    };

    await html2pdf().set(opt).from(htmlContent).save();
    return;
  }

  // Fallback: dedicated hidden iframe to trigger print dialog with pre-filled filename
  const iframe = document.createElement('iframe');
  iframe.style.position = 'fixed';
  iframe.style.right = '0';
  iframe.style.bottom = '0';
  iframe.style.width = '0';
  iframe.style.height = '0';
  iframe.style.border = 'none';
  document.body.appendChild(iframe);

  const doc = iframe.contentDocument || iframe.contentWindow?.document;
  if (doc) {
    doc.open();
    doc.write(`
      <!DOCTYPE html>
      <html>
        <head>
          <title>${cleanFileName}</title>
          <style>
            @page {
              size: A4 portrait;
              margin: 8mm;
            }
            body {
              margin: 0;
              padding: 0;
              background: #ffffff;
            }
            @media print {
              body {
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
              }
            }
          </style>
        </head>
        <body>
          ${htmlContent}
        </body>
      </html>
    `);
    doc.close();

    setTimeout(() => {
      const originalTitle = document.title;
      try {
        document.title = cleanFileName;
      } catch (e) {
        // ignore
      }

      iframe.contentWindow?.focus();
      iframe.contentWindow?.print();

      setTimeout(() => {
        try {
          document.title = originalTitle;
        } catch (e) {
          // ignore
        }
        if (document.body.contains(iframe)) {
          document.body.removeChild(iframe);
        }
      }, 4000);
    }, 400);
  }
};
