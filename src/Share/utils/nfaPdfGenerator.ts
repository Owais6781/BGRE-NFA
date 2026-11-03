import { INFAFormData } from '../Services/NFAService';
import { stripHtml } from '../utils';

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
  users: { label: string; value: string; email?: string }[];
  history: any[];
  comments: any[];
}

/**
 * Formats a date into DD/MM/YYYY HH:MM AM/PM
 */
export const formatDateTime = (dateObj: Date): string => {
  const day = String(dateObj.getDate()).padStart(2, '0');
  const month = String(dateObj.getMonth() + 1).padStart(2, '0');
  const year = dateObj.getFullYear();

  let hours = dateObj.getHours();
  const minutes = String(dateObj.getMinutes()).padStart(2, '0');
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
  addField('ORIGINATOR DEPARTMENT', form.originatorDept);
  addField('INITIATED BY', initiatorName);
  addField('INITIATED ON', initiatedOn);
  addField('LEGAL ENTITY NAME', form.project);
  addField('REGION', region);
  addField('TYPE OF APPROVAL', form.natureOfApproval || form.workflow);
  addField('TENANT NAME', tenantName);
  addField('PROVISIONED FOR BAD DEBTS', provisionedForBadDebts);
  addField('APPROVAL AMOUNT', `₹ ${formattedAmount}`);

  // Also include other populated fields if present
  if (form.RFQRefNo) addField('RFQ REF NO', form.RFQRefNo);
  if (form.budget && parseFloat(form.budget) > 0) addField('BUDGET', `₹ ${form.budget}`);
  if (form.doaMatrix) addField('DOA MATRIX', form.doaMatrix);

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

  addNarrative('SUBJECT', form.referenceSubject);
  addNarrative('INTRODUCTION', form.introduction);
  addNarrative('COMMERCIAL PROPOSAL', form.commercialProposal);
  addNarrative('ANNEXURE / ENCLOSURE', form.annexure);
  addNarrative('CHRONOLOGY', form.chronology);
  addNarrative('PAYMENT TERMS', form.paymentTerms);
  addNarrative('BUDGET REMARKS', form.budgetRemarks);
  addNarrative('RECOMMENDATION WITH JUSTIFICATION', form.recommendation);

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

  // ── Approver Rows ──
  const approverRows: IPdfApproverRow[] = [];
  const totalApprovers = form.approvers.filter((a) => Boolean(a)).length;

  form.approvers.forEach((appId, idx) => {
    if (!appId) return;

    const levelNum = idx + 1;
    const levelLabel = getOrdinalLabel(levelNum);
    const userObj = users.find((u) => u.value === String(appId));
    const userName = userObj ? userObj.label : `Approver ${levelNum}`;

    // Find in history / comments
    const histEntry = history.find(
      (h) =>
        (h.level && parseInt(h.level, 10) === levelNum) ||
        (h.user && userObj && h.user.toLowerCase() === userObj.label.toLowerCase())
    );

    const commentEntry = comments.find(
      (c) =>
        (c.level && parseInt(c.level, 10) === levelNum) ||
        (c.user && userObj && c.user.toLowerCase() === userObj.label.toLowerCase())
    );

    const remarks =
      commentEntry?.remarks ||
      histEntry?.remarks ||
      (histEntry?.action === 'Approved' ? '' : '');

    let status = 'Pending';
    if (form.level > levelNum || statusUpper.includes('CLOSE') || statusUpper.includes('APPROV')) {
      if (levelNum === totalApprovers && statusUpper.includes('CLOSE')) {
        status = 'Closed';
      } else {
        status = 'Approved';
      }
    } else if (form.level === levelNum) {
      status = 'Pending';
    }

    const date = histEntry?.date || commentEntry?.date || '';

    approverRows.push({
      levelLabel,
      designation: levelNum === 1 ? 'Software Engineer' : `${levelNum}th Level Approver`,
      name: userName,
      remarks,
      status,
      date
    });
  });

  const generatedDateStr = formatDateTime(new Date());

  // ── Render HTML ──
  return `
    <div style="font-family: Arial, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; color: #1e293b; background: #ffffff; padding: 24px 32px; box-sizing: border-box; max-width: 1000px; margin: 0 auto; line-height: 1.4;">

      <!-- ── TOP HEADER ── -->
      <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #5d1020; padding-bottom: 12px; margin-bottom: 16px;">
        <div style="display: flex; align-items: baseline; gap: 14px;">
          <span style="font-size: 26px; font-weight: 900; color: #4A081E; letter-spacing: -0.5px; text-decoration: underline; text-underline-offset: 4px;">BGRE</span>
          <span style="font-size: 17px; font-weight: 700; color: #0f172a;">Note for Approval (NFA)</span>
        </div>
        <div style="display: flex; align-items: center; gap: 14px;">
          <span style="font-size: 15px; font-weight: 700; color: #0f172a;">${requestNo}</span>
          <span style="font-size: 11px; font-weight: 800; text-transform: uppercase; padding: 4px 12px; border-radius: 9999px; background-color: ${statusBg}; color: ${statusColor}; letter-spacing: 0.5px;">${statusUpper}</span>
        </div>
      </div>

      <!-- ── METADATA GRID TABLE ── -->
      <table style="width: 100%; border-collapse: collapse; font-size: 11px; border: 1px solid #cbd5e1; margin-bottom: 16px;">
        <tbody>
          ${gridRows
            .map(
              (row) => `
            <tr>
              <td style="background-color: #f8fafc; color: #475569; font-weight: 800; text-transform: uppercase; padding: 7px 10px; width: 20%; border: 1px solid #cbd5e1;">${
                row.key1
              }</td>
              <td style="background-color: #ffffff; color: #0f172a; font-weight: 600; padding: 7px 10px; width: 30%; border: 1px solid #cbd5e1;">${
                row.val1
              }</td>
              <td style="background-color: #f8fafc; color: #475569; font-weight: 800; text-transform: uppercase; padding: 7px 10px; width: 20%; border: 1px solid #cbd5e1;">${
                row.key2
              }</td>
              <td style="background-color: #ffffff; color: #0f172a; font-weight: 600; padding: 7px 10px; width: 30%; border: 1px solid #cbd5e1;">${
                row.val2
              }</td>
            </tr>
          `
            )
            .join('')}
        </tbody>
      </table>

      <!-- ── NARRATIVE SECTIONS ── -->
      ${narrativeSections
        .map(
          (sec) => `
        <div style="margin-bottom: 12px; page-break-inside: avoid;">
          <div style="font-size: 11px; font-weight: 800; text-transform: uppercase; color: #334155; margin-bottom: 4px; letter-spacing: 0.5px;">
            ${sec.title}
          </div>
          <div style="border: 1px solid #cbd5e1; border-radius: 6px; padding: 8px 12px; background-color: #ffffff; font-size: 11.5px; color: #1e293b; line-height: 1.5; min-height: 28px;">
            ${sec.content}
          </div>
        </div>
      `
        )
        .join('')}

      <!-- ── APPROVALS TABLE ── -->
      ${
        approverRows.length > 0
          ? `
        <div style="margin-top: 18px; margin-bottom: 18px; page-break-inside: avoid;">
          <div style="text-align: center; font-size: 14px; font-weight: 800; color: #1e293b; margin-bottom: 8px;">
            Approvals
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
              ${approverRows
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
                    <td style="padding: 7px 10px; border: 1px solid #cbd5e1; color: #475569;">${app.remarks || ''}</td>
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
      <div style="display: flex; justify-content: space-between; align-items: center; font-size: 10px; color: #64748b; border-top: 1px solid #e2e8f0; padding-top: 10px; margin-top: 16px;">
        <span>This is an official system-generated Note for Approval from BGRE and does not require physical signature.</span>
        <span>Generated on: ${generatedDateStr}</span>
      </div>

    </div>
  `;
};

/**
 * Downloads the generated PDF.
 * Uses html2pdf.js dynamically if available, or opens a dedicated print/PDF rendering window.
 */
export const downloadNfaPdf = async (data: IPdfGeneratorData): Promise<void> => {
  const requestNo = (data.form.requestNo || 'NFA-Note').trim();
  // Sanitize filename for operating system (replace illegal characters with '-')
  const cleanFileName = requestNo.replace(/[/\\?%*:|"<>]/g, '-');

  const htmlContent = buildNfaPdfHtml(data);

  // Check if html2pdf is available on window or can be dynamically loaded
  const win = window as any;
  if (!win.html2pdf) {
    try {
      await new Promise<void>((resolve, reject) => {
        const script = document.createElement('script');
        script.src =
          'https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js';
        script.onload = () => resolve();
        script.onerror = () => reject(new Error('Failed to load html2pdf script'));
        document.head.appendChild(script);
      });
    } catch (e) {
      console.warn('html2pdf could not be loaded via CDN, falling back to print window:', e);
    }
  }

  if (win.html2pdf) {
    // Render in an offscreen container to ensure exact layout
    const container = document.createElement('div');
    container.innerHTML = htmlContent;
    container.style.position = 'fixed';
    container.style.left = '-9999px';
    container.style.top = '0';
    container.style.width = '1000px';
    container.style.backgroundColor = '#ffffff';
    document.body.appendChild(container);

    const opt = {
      margin: [8, 8, 8, 8],
      filename: `${cleanFileName}.pdf`,
      image: { type: 'jpeg', quality: 0.98 },
      html2canvas: { scale: 2, useCORS: true, logging: false },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' },
      pagebreak: { mode: ['avoid-all', 'css', 'legacy'] }
    };

    try {
      await win.html2pdf().set(opt).from(container).save();
    } finally {
      document.body.removeChild(container);
    }
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
