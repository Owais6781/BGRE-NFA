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
      narrativeSections.push({ title, content: content.trim() });
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
  if (form.liquidatedDamage && form.liquidatedDamage.trim() !== 'NA') {
    contractParts.push(`<strong>Liquidated Damage:</strong> ${form.liquidatedDamage}`);
  }
  if (form.contractPerf && form.contractPerf.trim() !== 'NA') {
    contractParts.push(`<strong>Contract Performance:</strong> ${form.contractPerf}`);
  }
  if (form.defectLiability && form.defectLiability.trim() !== 'NA') {
    contractParts.push(`<strong>Defect Liability Period:</strong> ${form.defectLiability}`);
  }
  if (form.workCompletion && form.workCompletion.trim() && !form.workCompletion.includes('-------')) {
    contractParts.push(`<strong>Work Completion:</strong> ${form.workCompletion}`);
  }
  if (contractParts.length > 0) {
    narrativeSections.push({
      title: 'CONTRACT TERMS',
      content: contractParts.join('<br/>')
    });
  }

 
  addNarrative('BUDGET REMARKS', form.budgetRemarks);
  addNarrative('RECOMMENDATION WITH JUSTIFICATION', form.recommendation);
  addNarrative('ANNEXURE / ENCLOSURE', form.annexure);
  addNarrative('NOTE 1', form.note1);
  addNarrative('NOTE 2', form.note2);
  addNarrative('NOTE 3', form.note3);
  addNarrative('NOTE 4', form.note4);

 

  // ── Approver Rows ──
  const approverRows: IPdfApproverRow[] = [];

  form.approvers.forEach((appId, idx) => {
    if (!appId) return;

    const levelNum = idx + 1;

    // Show only approved levels + the immediate next pending level (hide all subsequent pending levels)
    const shouldInclude = isOverallApproved || levelNum <= currentWfLevel;
    if (!shouldInclude) return;

    const levelLabel = getOrdinalLabel(levelNum);
    const userObj = users.find((u) => u.value === String(appId));
    let userName = userObj ? userObj.label : `Approver ${levelNum}`;

    // Determine status of this level
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

    // Only populate remarks and date for levels that have been approved/rejected
    // Each level must show strictly its own approver's submitted remark
    let remarks = '';
    let date = '';
    if (status === 'Approved' || status === 'Closed' || status === 'Rejected') {
      const commentEntry = comments.find((c) => c.level && parseInt(String(c.level), 10) === levelNum);
      const histEntry = history.find((h) => h.level && parseInt(String(h.level), 10) === levelNum);
      
      const rawRemark = commentEntry?.remarks || histEntry?.remarks || '';
      remarks = (rawRemark && rawRemark.trim().toLowerCase() !== 'no remarks') ? rawRemark.trim() : '';
      date = histEntry?.date || commentEntry?.date || '';

      // If comment or history has the actual person who acted on this level, use their name
      if (commentEntry?.user && commentEntry.user !== 'System') {
        userName = commentEntry.user;
      } else if (histEntry?.user && histEntry.user !== 'System') {
        userName = histEntry.user;
      }
    }

    // Actual user designation mentioned in the Approver ID field
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

  const generatedDateStr = formatDateTime(new Date());

  // ── Render HTML ──
  return `
    <style>
      * {
        box-sizing: border-box;
      }
      .pdf-section {
        page-break-inside: avoid !important;
        break-inside: avoid !important;
      }
      tr {
        page-break-inside: avoid !important;
        break-inside: avoid !important;
      }
      table {
        max-width: 100%;
        border-collapse: collapse;
      }
      img {
        max-width: 100%;
        height: auto;
      }
    </style>
    <div style="font-family: Arial, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; color: #1e293b; background: #ffffff; padding: 16px 20px; box-sizing: border-box; width: 100%; line-height: 1.4;">

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
      <div class="pdf-section" style="margin-bottom: 16px; border: 1px solid #cbd5e1; border-radius: 4px; overflow: hidden; page-break-inside: avoid; break-inside: avoid;">
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
          <div style="color: #4A081E; font-size: 11px; font-weight: 800; text-transform: uppercase; padding: 4px 2px 4px 2px; letter-spacing: 0.5px;">
            ${sec.title}
          </div>
          <div style="border: 1px solid #cbd5e1; border-radius: 4px; overflow: hidden;">
            <div style="padding: 10px 12px; background-color: #ffffff; font-size: 11.5px; color: #1e293b; line-height: 1.5; min-height: 28px; word-break: break-word; overflow-wrap: break-word;">
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
          <div style="background-color: #f8fafc; text-align: center; color: #4A081E; font-size: 11px; font-weight: 800; text-transform: uppercase; padding: 6px 12px; letter-spacing: 0.5px; border: 1px solid #cbd5e1; border-bottom: none;">
            APPROVALS
          </div>
          <table style="width: 100%; border-collapse: collapse; font-size: 11px; border: 1px solid #cbd5e1;">
            <thead>
              <tr style="background-color: #4A081E; color: #ffffff;">
                <th style="padding: 8px 10px; font-weight: 800; text-align: left; text-transform: uppercase; border: 1px solid #4A081E; width: 14%;">LEVEL</th>
                <th style="padding: 8px 10px; font-weight: 800; text-align: left; text-transform: uppercase; border: 1px solid #4A081E; width: 18%;">DESIGNATION</th>
                <th style="padding: 8px 10px; font-weight: 800; text-align: left; text-transform: uppercase; border: 1px solid #4A081E; width: 20%;">NAME</th>
                <th style="padding: 8px 10px; font-weight: 800; text-align: left; text-transform: uppercase; border: 1px solid #4A081E; width: 24%;">REMARKS</th>
                <th style="padding: 8px 10px; font-weight: 800; text-align: center; text-transform: uppercase; border: 1px solid #4A081E; width: 11%;">STATUS</th>
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
          if (app.status === 'Approved') {
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
      <div class="pdf-section" style="display: flex; justify-content: space-between; align-items: center; font-size: 10px; color: #64748b; border-top: 1px solid #e2e8f0; padding-top: 10px; margin-top: 16px; page-break-inside: avoid; break-inside: avoid;">
        <span>This is an official system-generated Note for Approval from BGRE and does not require physical signature.</span>
        <span>Generated on: ${generatedDateStr}</span>
      </div>

    </div>
  `;
};

// Cached local html2pdf module reference
let html2pdfModule: any = null;
let html2pdfPromise: Promise<any> | null = null;

// const getHtml2Pdf = async (): Promise<any> => {
//   const win = window as any;
//   if (win.html2pdf) return win.html2pdf;
//   if (html2pdfModule) return html2pdfModule;

//   try {
//     const mod: any = await import('./html2pdf.bundle.min.js');
//     html2pdfModule = mod && mod.default ? mod.default : mod;
//     if (!html2pdfModule) {
//   html2pdfModule = await import('html2pdf.js');
// }
//     if (html2pdfModule) return html2pdfModule;
//   } catch (e) {
//     console.warn('import html2pdf bundle failed, trying require:', e);
//     try {
//       html2pdfModule = reqMod && reqMod.default ? reqMod.default : reqMod;
//       if (html2pdfModule) return html2pdfModule;
//     } catch (e2) {
//       console.warn('require html2pdf bundle failed:', e2);
//     }
//   }

//   if (win.html2pdf) return win.html2pdf;
//   return null;
// };

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
        backgroundColor: '#ffffff'
      },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' },
      pagebreak: {
        mode: ['css', 'legacy'],
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
              size: A4 landscape;
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
