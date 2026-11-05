import { stripHtml } from "../utils";
export { stripHtml } from "../utils";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { NFADocService } from "./NFADocService";
import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users";

export const formatHistoryDate = (date: Date): string => {
  const pad = (num: number) => num < 10 ? '0' + num : String(num);
  const day = pad(date.getDate());
  const month = pad(date.getMonth() + 1);
  const year = date.getFullYear();

  let hours = date.getHours();
  const minutes = pad(date.getMinutes());
  const ampm = hours >= 12 ? 'PM' : 'AM';
  hours = hours % 12;
  hours = hours ? hours : 12;
  const strTime = `${hours}:${minutes} ${ampm}`;

  return `${day}/${month}/${year} ${strTime}`;
};

// ── Form Data Interface ──
export interface INFAFormData {
  project: string;
  projectSubEntity: string[];
  originatorDept: string;
  date: string;
  referenceSubject: string;
  requestNo?: string;
  introduction: string;
  location?: string;
  natureOfApproval?: string;

  completeIndentPRReceivedFromDate: string;
  boqSpecificationsRevisionsByDate: string;
  RFQRefNo: string;
  remark: string;
  Approvedrejectedby: string;
  Approvedrejectdatetime: string;

  chronology: string;
  commercialProposal: string;
  paymentTerms: string;
  liquidatedDamage: string;
  contractPerf: string;
  defectLiability: string;
  workCompletion: string;
  budgetRemarks: string;
  packageAwardedCost: string;
  budget: string;
  budgetedNonBudgeted: string;
  capexSerialNo?: string;
  recommendation: string;
  contractorName: string;
  annexure: string;
  note1: string;
  note2: string;
  note3: string;
  note4: string;
  workflow: string;
  doaMatrix: string;
  approvers: (string | number | null)[];
  approverDetails?: ({ id: number; name: string; designation: string; email?: string } | null)[];
  attachments?: FileList | null;
  uniqId: string;
  level: number;
  pendingAtId: number | null;
  PendingFrom : string;
  initiatorId?: number;
  commentHistory?: string;
  workflowHistory?: string;
  workflowHistoryHTML?: string;
  status?: string;
}

export interface INFASaveResponse {
  success: boolean;
  itemId: number | null;
  message: string;
  requestNo?: string;
}

export interface IValidationError {
  field: string;
  message: string;
}

export interface IValidationResult {
  isValid: boolean;
  errors: IValidationError[];
}

export type NFAStatus = "Draft" | "Pending";

export class NFAService {
  private sp: SPFI;
   private listName: string = "DigitalApprovalNotes";
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
    this.sp = spfi().using(SPFx(context));
  }

  public async resolveUserId(loginName: string): Promise<number | null> {
    try {
      const user = await this.sp.web.ensureUser(loginName);
      return user.Id;
    } catch (error) {
      console.error(error);
      return null;
    }
  }

  public validateForm(form: INFAFormData, status: NFAStatus): IValidationResult {
    const errors: IValidationError[] = [];
    if (!form.project.trim()) errors.push({ field: "project", message: "Entity  is required." });
    if (status === "Draft") return { isValid: errors.length === 0, errors };
    if (!form.projectSubEntity || form.projectSubEntity.length === 0) errors.push({ field: "projectSubEntity", message: "Sub-Entity is required." });

    if (!form.originatorDept.trim()) errors.push({ field: "originatorDept", message: "Function is required." });
    if (!stripHtml(form.referenceSubject)) errors.push({ field: "referenceSubject", message: "Subject is required." });

    // if (!stripHtml(form.introduction)) errors.push({ field: "introduction", message: "Introduction is required." });
    // if (!stripHtml(form.chronology)) errors.push({ field: "chronology", message: "Chronology is required." });
    // if (!stripHtml(form.commercialProposal)) errors.push({ field: "commercialProposal", message: "Commercial Proposals is required." });
 if (!form.liquidatedDamage || !form.liquidatedDamage.trim()) {
      errors.push({ field: "liquidatedDamage", message: "Liquidated Damage is required." });
    }
    if (!form.contractPerf || !form.contractPerf.trim()) {
      errors.push({ field: "contractPerf", message: "Contract Performance Guarantee is required." });
    }
    if (!form.defectLiability || !form.defectLiability.trim()) {
      errors.push({ field: "defectLiability", message: "Defect Liability Period is required." });
    }
    if (!form.workCompletion || !form.workCompletion.trim()) {
      errors.push({ field: "workCompletion", message: "Work Completion is required." });
    }

    
    if (!form.approvers[0]) errors.push({ field: "approvers", message: "Approver Level-1 is required." });

   
    if (form.workflow === 'AUTO' && (!form.doaMatrix || !form.doaMatrix.trim())) {
      errors.push({ field: "doaMatrix", message: "DOA Matrix is required for Auto workflow." });
    }

    return { isValid: errors.length === 0, errors };
  }



   private toSharePointDate(value: string): string | null {
    if (!value) return null;

    const match = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);

    if (match) {
      return `${match[1]}-${match[2]}-${match[3]}T12:00:00`;
    }

    const date = new Date(value);

    return Number.isNaN(date.getTime())
      ? null
      : date.toISOString();
  }


  
  private buildRequestBody(
    form: INFAFormData,
    status: NFAStatus,
    resolvedApprovers: (number | null)[],
    projectName: string,
    subEntityNames: string,
    pendingAtName: string = "",
    functionOptions: { label: string, value: string }[] = []
  ): Record<string, any> {
    const now = new Date();
    const cleanSubject = stripHtml(form.referenceSubject);

    const funcObj = functionOptions.find(
      f => f.value === String(form.originatorDept) || f.label === form.originatorDept
    );
    const functionName = funcObj ? funcObj.label : form.originatorDept;
    const functionId = funcObj ? funcObj.value : form.originatorDept;

    const body: Record<string, any> = {
      Title: cleanSubject,
      Project: projectName,
      ProjectSubEntity: subEntityNames,
      FunctionName: functionName,
      FunctionID: String(functionId || ""),
      Dept: functionName || "",
      Location: form.location || "",
      NatureofApproval: form.natureOfApproval || "",

      Status: status,
      SubjectMatter: cleanSubject,
      BriefDesc: form.introduction,

      RFQRefNo: form.RFQRefNo || "",
      Remark: form.remark || "",

      Chronology: form.chronology || "",
      CommercialProposal: form.commercialProposal || "",
      PaymentTerms: form.paymentTerms || "",
      LiquidatedDamage: form.liquidatedDamage || "",
      ContractPerformance: form.contractPerf || "",
      DefectLibilityPeriod: form.defectLiability || "",
      WorkCompletionPeriod: form.workCompletion || "",
      BudgetRemarks: form.budgetRemarks || "",
      ApprovalAmount: form.packageAwardedCost || "0.00",
      Budget: form.budget || "0.00",
      BudgetedNonBudgted: form.budgetedNonBudgeted || "",
      // CapexSerialNo: form.capexSerialNo || "",
      RecommendationJust: form.recommendation || "",
      ContractorName: form.contractorName || "",
      AnnexureEnclosure: form.annexure || "",
      AdditionalText1: form.note1 || "",
      AdditionalText2: form.note2 || "",
      AdditionalText3: form.note3 || "",
      AdditionalText4: form.note4 || "",
      WorkflowType: form.workflow || "",
      DOAClause: form.doaMatrix || "",
      RequestedOn: form.date || "",
      Month: now.toLocaleString('default', { month: 'long' }),
      Year: now.getFullYear().toString(),
      WFLevel: status === "Draft" ? "0" : "1",
      UniqID: form.uniqId || "",
      PendingAt: status === "Pending" ? (pendingAtName || "") : "",
      PendingFrom: status === "Pending" ? formatHistoryDate(now) : (form.PendingFrom || ""),
    };

    for (let i = 0; i < 12; i++) {
      const app = resolvedApprovers[i];
      const appVal = (typeof app === 'number') ? app : (app && !isNaN(parseInt(String(app), 10))) ? parseInt(String(app), 10) : null;
      body[`ApprovalLevel${i + 1}Id`] = appVal;
    }

    const prDate = this.toSharePointDate(form.completeIndentPRReceivedFromDate);
    if (prDate) {
      body.PRreceivedDate = prDate;
    }

    const boqDate = this.toSharePointDate(form.boqSpecificationsRevisionsByDate);
    if (boqDate) {
      body.BOQRevisionsDate = boqDate;
    }
    return body;
  }

  private async generateRequestNoLegacy(
    projectIdOrName: string,
    functionId: string,
    status: NFAStatus
  ): Promise<string> {
    const dtt = new Date();
    const yyyy = dtt.getFullYear();
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const mName = months[dtt.getMonth()];

    // Draft Request Number format: Draft/YYYY/Month/Counter
    if (status === "Draft") {
      try {
        const draftCounters = await this.sp.web.lists.getByTitle("CounterMaster").items
          .select("ID", "Title", "Counter", "FunctionAbbr", "FunctionID")
          .filter("FunctionID eq '99'")();

        if (draftCounters.length > 0) {
          const item = draftCounters[0];
          const counterVal = item.Counter || "1";
          const reqNo = `Draft/${yyyy}/${mName}/${counterVal}`;
          const nextCounter = String(parseInt(String(counterVal), 10) + 1);
          try {
            await this.sp.web.lists.getByTitle("CounterMaster").items.getById(item.ID).update({
              Counter: nextCounter
            });
          } catch (updateErr) {
            console.warn("NFAService: Error updating CounterMaster for draft:", updateErr);
          }
          return reqNo;
        }
      } catch (err) {
        console.warn("NFAService: Error reading CounterMaster for draft:", err);
      }
      return `Draft/${yyyy}/${mName}/1`;
    }

    // Pending (Submit) Request Number format: [NFASeries]/[FunctionAbbr]/[YYYY]/[Month]/[Counter]
    let projSeries = "";
    if (projectIdOrName) {
      try {
        let projItem: any = null;
        if (/^\d+$/.test(projectIdOrName)) {
          const pId = parseInt(projectIdOrName, 10);
          try {
            projItem = await this.sp.web.lists.getByTitle("Project Master").items.getById(pId).select("ID", "Title", "ProjectName", "NFASeries")();
          } catch {
            projItem = await this.sp.web.lists.getByTitle("ProjectMaster").items.getById(pId).select("ID", "Title", "ProjectName", "NFASeries")();
          }
        } else {
          try {
            const res = await this.sp.web.lists.getByTitle("Project Master").items
              .filter(`ProjectName eq '${projectIdOrName}' or Title eq '${projectIdOrName}'`)
              .select("ID", "Title", "ProjectName", "NFASeries")();
            if (res.length > 0) projItem = res[0];
          } catch {
            const res = await this.sp.web.lists.getByTitle("ProjectMaster").items
              .filter(`ProjectName eq '${projectIdOrName}' or Title eq '${projectIdOrName}'`)
              .select("ID", "Title", "ProjectName", "NFASeries")();
            if (res.length > 0) projItem = res[0];
          }
        }
        if (projItem && projItem.NFASeries) {
          projSeries = String(projItem.NFASeries).trim();
        }
      } catch (projErr) {
        console.warn("NFAService: Error resolving NFASeries from Project Master:", projErr);
      }
    }

    if (!projSeries) {
      const clean = (projectIdOrName || "NFA").trim();
      projSeries = clean.includes(" ") ? clean.split(/\s+/).map(w => w[0]).join("").toUpperCase() : clean.toUpperCase();
      projSeries = projSeries.replace(/[^A-Z0-9]/g, "");
    }

    let functionAbbr = "";
    let counterVal = "1";
    let counterItemId: number | null = null;

    try {
      const counterItems = await this.sp.web.lists.getByTitle("CounterMaster").items
        .select("ID", "Title", "Counter", "FunctionAbbr", "FunctionID")
        .top(500)();

      for (const c of counterItems) {
        if (String(c.FunctionID) === String(functionId)) {
          functionAbbr = c.FunctionAbbr || "";
        }
        if (String(c.FunctionID) === "0") {
          counterVal = c.Counter || "1";
          counterItemId = c.ID;
        }
      }
    } catch (cErr) {
      console.warn("NFAService: Error fetching CounterMaster items:", cErr);
    }

    let reqNo = "";
    if (functionAbbr) {
      reqNo = `${projSeries}/${functionAbbr}/${yyyy}/${mName}/${counterVal}`;
    } else {
      reqNo = `${projSeries}/${yyyy}/${mName}/${counterVal}`;
    }

    if (counterItemId) {
      try {
        const nextCounter = String(parseInt(String(counterVal), 10) + 1);
        await this.sp.web.lists.getByTitle("CounterMaster").items.getById(counterItemId).update({
          Counter: nextCounter
        });
      } catch (updErr) {
        console.warn("NFAService: Error updating CounterMaster Counter:", updErr);
      }
    }

    return reqNo;
  }

  private async writeWFLog(
    itemId: number,
    requestNo: string,
    form: INFAFormData,
    resolvedApprovers: (number | null)[],
    projectName: string,
    functionOptions: { label: string, value: string }[],
    subEntityNames: string = ""
  ): Promise<void> {
    try {
      const funcObj = functionOptions.find(f => f.value === String(form.originatorDept) || f.label === form.originatorDept);
      const functionName = funcObj ? funcObj.label : form.originatorDept;
      const functionId = funcObj ? funcObj.value : form.originatorDept;

      const pendingAtId = resolvedApprovers[0] || null;
      let pendingAtName = "";

      if (pendingAtId) {
        try {
          const user = await this.sp.web.siteUsers.getById(pendingAtId)();
          pendingAtName = user.Title || "";
        } catch (userErr) {
          console.error("NFAService: Error fetching site user for pendingAtId:", pendingAtId, userErr);
        }
      }


      // Only 1 row created on submission (see logBody2 below)

      // Log 2: Next Pending Task Log (Pending)
      const logBody2: Record<string, any> = {
        Title: requestNo,
        WFID: String(itemId),
        FunctionID: String(functionId),
        FunctionName: functionName,
        DOAClause: form.doaMatrix || "",
        Project: projectName,
        Amount: form.packageAwardedCost,
        Status: "Pending",
        PendingAt: pendingAtName || "",
        WFLevel: "1",
        Location: form.location || "",
        Department: functionName || "",
        Remarks: form.budgetRemarks || "",
        PartyName:form.contractorName||""
      };

      if (pendingAtId) {
        logBody2.PPLPendingAtId = pendingAtId;
      }

      await this.sp.web.lists.getByTitle("WFLogList").items.add(logBody2);

      // DashboardList Save Integration
      try {
        const getDueDate = (): Date => {
          const someDate = new Date();
          const dd = someDate.getDate();
          const mm = someDate.getMonth() + 1;
          const yyyy = someDate.getFullYear();
          const ddStr = dd < 10 ? '0' + dd : String(dd);
          const mmStr = mm < 10 ? '0' + mm : String(mm);
          const someFormattedDate = `${mmStr}/${ddStr}/${yyyy}`;
          const vDueDate = new Date(someFormattedDate);
          vDueDate.setDate(vDueDate.getDate() + 1);
          return vDueDate;
        };

        const now = new Date();
        const formattedDateTime = formatHistoryDate(now);
        const initiatorName = this.context.pageContext.user.displayName || "Initiator";

        // Dashboard Log: Current State (Pending)
        const dashboardBody: Record<string, any> = {
          Title: requestNo,
          Project: projectName,
          Entity:projectName,
          SubEntity:subEntityNames,
          Function: functionName,
          FunctionID: String(functionId),
          Amount: form.packageAwardedCost,
          WFID: String(itemId),
          WFLevel: "1",
          Status: "Pending",
          DOA: form.doaMatrix || "",
          WorkflowType: form.workflow || "",
          CurrentApproverId: pendingAtId,
          CurrentApprovalName: pendingAtName,
          CurrentApprovalID: pendingAtId ? String(pendingAtId) : "",
          PendingFrom: formattedDateTime,
          InitiatedBy: initiatorName,
          InitiatedOn: form.date || formattedDateTime,
          PendingFromDate: now,
          Location: form.location || "",
          Department: functionName || "",
          DueDate: getDueDate()
        };

        // Check if item already exists in DashboardList for this NFA
        const existingItems = await this.sp.web.lists.getByTitle("DashboardList").items
          .filter(`WFID eq '${itemId}'`)
          .select("Id")();

        if (existingItems.length > 0) {
          await this.sp.web.lists.getByTitle("DashboardList").items.getById(existingItems[0].Id).update(dashboardBody);
        } else {
          await this.sp.web.lists.getByTitle("DashboardList").items.add(dashboardBody);
        }

        try {
          await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId).update({
            PendingFrom: formattedDateTime
          });
        } catch (updateErr) {
          console.warn("NFAService: Error updating PendingFrom in DigitalApprovalNotes:", updateErr);
        }
      } catch (dashError) {
        console.error("NFAService: Error writing to DashboardList:", dashError);
      }

    } catch (logError) {
      console.error("NFAService: Error writing to WFLogList:", logError);
    }
  }

  public async saveNFA(
    form: INFAFormData,
    status: NFAStatus,
    projectOptions: { label: string, value: string }[],
    subEntityOptions: { label: string, value: string }[],
    functionOptions: { label: string, value: string }[]
  ): Promise<INFASaveResponse> {
    try {
      const resolvedApprovers = await Promise.all(
        form.approvers.map(async (a) => {
          if (!a) return null;
          if (typeof a === "number") return a;
          const strVal = String(a).trim();
          if (!strVal || strVal === "null") return null;
          if (/^\d+$/.test(strVal)) return parseInt(strVal);
          return await this.resolveUserId(strVal);
        })
      );

      const projectObj = projectOptions.find(p => p.value === form.project);
      const projectName = projectObj ? projectObj.label : form.project;
      const subEntityNames = form.projectSubEntity.map(id => subEntityOptions.find(o => o.value === id)?.label || id).join("; ");

      const funcObj = functionOptions.find(
        f => f.value === String(form.originatorDept) || f.label === form.originatorDept
      );
      const functionId = funcObj ? funcObj.value : form.originatorDept;

      const pendingAtId = resolvedApprovers[0] || null;
      let pendingAtName = "";
      if (pendingAtId) {
        try {
          const user = await this.sp.web.getUserById(pendingAtId)();
          pendingAtName = user.Title || "";
        } catch (e1) {
          try {
            const user = await this.sp.web.siteUsers.getById(pendingAtId)();
            pendingAtName = user.Title || "";
          } catch (userErr) {
            console.error("NFAService: Error fetching site user for pendingAtId:", pendingAtId, userErr);
          }
        }
      }

      const body = this.buildRequestBody(form, status, resolvedApprovers, projectName, subEntityNames, pendingAtName, functionOptions);

      const requestNo = await this.generateRequestNoLegacy(form.project, String(functionId), status);
      body.Title = requestNo;

      if (status === "Pending") {
        const initiatorName = this.context.pageContext.user.displayName || "Initiator";
        const dateStr = formatHistoryDate(new Date());

        const initWorkflowHistory = `NFA Request Initiated by ${initiatorName} on ${dateStr}</br>`;
        const pendingWorkflowHistory = pendingAtName
          ? `\nNFA Pending at ${pendingAtName} at level-1 on ${dateStr}</br>`
          : "";
        const fullWorkflowHistory = initWorkflowHistory + pendingWorkflowHistory;

        body.WorkflowHistory = fullWorkflowHistory;
        body.WorkflowApprovalHistory = fullWorkflowHistory;
      }

      const result = await this.sp.web.lists.getByTitle(this.listName).items.add(body);
      const itemId = result.Id;

      if (status === "Pending") {
        await this.writeWFLog(itemId, requestNo, form, resolvedApprovers, projectName, functionOptions, subEntityNames);
      }

      if (form.uniqId) {
        try {
          const docService = new NFADocService(this.context);
          await docService.updateDocumentsRequestNo(form.uniqId, requestNo);
        } catch (docErr) {
          console.error("NFAService: Error updating NFADocs Title with requestNo:", docErr);
        }
      }

      return { success: true, itemId: itemId, requestNo: requestNo, message: "Saved successfully!" };
    } catch (error) {
      console.error("NFAService: Error in saveNFA:", error);
      const errMsg = error instanceof Error ? error.message : "Error occurred during save.";
      return { success: false, itemId: null, message: errMsg };
    }
  }

  public async updateNFA(
    itemId: number,
    form: INFAFormData,
    status: NFAStatus,
    projectOptions: { label: string, value: string }[],
    subEntityOptions: { label: string, value: string }[],
    functionOptions: { label: string, value: string }[]
  ): Promise<INFASaveResponse> {
    try {
      const resolvedApprovers = await Promise.all(
        form.approvers.map(async (a) => {
          if (!a) return null;
          if (typeof a === "number") return a;
          const strVal = String(a).trim();
          if (!strVal || strVal === "null") return null;
          if (/^\d+$/.test(strVal)) return parseInt(strVal);
          return await this.resolveUserId(strVal);
        })
      );

      const projectObj = projectOptions.find(p => p.value === form.project);
      const projectName = projectObj ? projectObj.label : form.project;
      const subEntityNames = form.projectSubEntity.map(id => subEntityOptions.find(o => o.value === id)?.label || id).join("; ");

      const funcObj = functionOptions.find(
        f => f.value === String(form.originatorDept) || f.label === form.originatorDept
      );
      const functionId = funcObj ? funcObj.value : form.originatorDept;

      const pendingAtId = resolvedApprovers[0] || null;
      let pendingAtName = "";
      if (pendingAtId) {
        try {
          const user = await this.sp.web.getUserById(pendingAtId)();
          pendingAtName = user.Title || "";
        } catch (e1) {
          try {
            const user = await this.sp.web.siteUsers.getById(pendingAtId)();
            pendingAtName = user.Title || "";
          } catch (userErr) {
            console.error("NFAService: Error fetching site user for pendingAtId:", pendingAtId, userErr);
          }
        }
      }

      const body = this.buildRequestBody(form, status, resolvedApprovers, projectName, subEntityNames, pendingAtName, functionOptions);

      let requestNo = form.requestNo || "";
      if (status === "Pending" && (!requestNo || requestNo.startsWith("Draft"))) {
        requestNo = await this.generateRequestNoLegacy(form.project, String(functionId), "Pending");
      } else if (!requestNo) {
        requestNo = await this.generateRequestNoLegacy(form.project, String(functionId), status);
      }
      body.Title = requestNo;

      if (status === "Pending") {
        try {
          const item = await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId).select("WorkflowHistory", "WorkflowApprovalHistory")();
          const currentWorkflowHistory = item.WorkflowHistory || item.WorkflowApprovalHistory || "";
          const initiatorName = this.context.pageContext.user.displayName || "Initiator";
          const dateStr = formatHistoryDate(new Date());

          const initWorkflowHistory = `NFA Request Initiated by ${initiatorName} on ${dateStr}</br>`;
          const pendingWorkflowHistory = pendingAtName
            ? `\nNFA Pending at ${pendingAtName} at level-1 on ${dateStr}</br>`
            : "";
          const newWorkflowEntry = initWorkflowHistory + pendingWorkflowHistory;

          const updatedWorkflowHistory = currentWorkflowHistory
            ? `${currentWorkflowHistory}\n${newWorkflowEntry}`
            : newWorkflowEntry;

          body.WorkflowHistory = updatedWorkflowHistory;
          body.WorkflowApprovalHistory = updatedWorkflowHistory;
        } catch (hErr) {
          console.warn("NFAService: Error reading WorkflowHistory in updateNFA:", hErr);
        }
      }

      await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId).update(body);

      if (status === "Pending") {
        await this.writeWFLog(itemId, requestNo, form, resolvedApprovers, projectName, functionOptions, subEntityNames);
      }

      if (form.uniqId) {
        try {
          const docService = new NFADocService(this.context);
          await docService.updateDocumentsRequestNo(form.uniqId, requestNo);
        } catch (docErr) {
          console.error("NFAService: Error updating NFADocs Title with requestNo:", docErr);
        }
      }

      return { success: true, itemId, requestNo: requestNo, message: "Updated successfully!" };
    } catch (error) {
      console.error("NFAService: Error in updateNFA:", error);
      const errMsg = error instanceof Error ? error.message : "Error occurred during update.";
      return { success: false, itemId, message: errMsg };
    }
  }

  public async getNFA(itemId: number): Promise<INFAFormData | null> {
    try {
      const levels = Array.from({ length: 12 }, (_, i) => `ApprovalLevel${i + 1}`);
      // const item = await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId)
      //   .select("*", "PendingAt/Id", "PendingAt/Title", "Author/Id", "Author/Title", "CommentHistory", "WorkflowHistory", ...levels.map(l => `${l}/Id`), ...levels.map(l => `${l}/Title`))
      //   .expand("PendingAt", "Author", ...levels)();

      let item: any;
      try {
        item = await this.sp.web.lists
          .getByTitle(this.listName)
          .items.getById(itemId)
          .select(
            "*", "Author/Id", "Author/Title", "Author/EMail",
            ...levels.map(l => `${l}/Id`),
            ...levels.map(l => `${l}/Title`),
            ...levels.map(l => `${l}/EMail`),
            ...levels.map(l => `${l}/JobTitle`),
            ...levels.map(l => `${l}/Department`)
          )
          .expand("Author", ...levels)();
      } catch (expandErr) {
        console.warn("NFAService: Failed to expand ApprovalLevels, falling back:", expandErr);
        item = await this.sp.web.lists
          .getByTitle(this.listName)
          .items.getById(itemId)
          .select("*", "Author/Id", "Author/Title", "Author/EMail")
          .expand("Author")();
      }

      
      console.log("Raw SharePoint Item:", item);
      return {
        project: item.Project || "",
        projectSubEntity: item.ProjectSubEntity ? item.ProjectSubEntity.split("; ") : [],
        originatorDept:item.FunctionName || "",
        date: item.RequestedOn || "",
        referenceSubject: (() => {
          const subj = stripHtml(item.SubjectMatter || "");
          if (subj && !/^Draft\//i.test(subj.trim())) return subj;
          const title = stripHtml(item.Title || "");
          if (title && !/^Draft\//i.test(title.trim())) return title;
          return "";
        })(),
        requestNo: item.Title || "",
        introduction: item.BriefDesc || "",

        completeIndentPRReceivedFromDate: item.PRreceivedDate ? item.PRreceivedDate.substring(0, 10) : "",
        boqSpecificationsRevisionsByDate: item.BOQRevisionsDate ? item.BOQRevisionsDate.substring(0, 10) : "",
        RFQRefNo: item.RFQRefNo || "",
        remark: item.Remark || "",
          
        Approvedrejectedby: item.Approvedrejectedby || "",
        Approvedrejectdatetime: item.Approvedrejectdatetime || "",

        chronology: item.Chronology || "",
        commercialProposal: item.CommercialProposal || "",
        paymentTerms: item.PaymentTerms || "",
        liquidatedDamage: item.LiquidatedDamage || "",
        contractPerf: item.ContractPerformance || "",
        defectLiability: item.DefectLibilityPeriod || "",
        workCompletion: item.WorkCompletionPeriod || "",
        budgetRemarks: item.BudgetRemarks || "",
        packageAwardedCost: item.ApprovalAmount || "0.00",
        budget: item.Budget || "0.00",
        budgetedNonBudgeted: item. BudgetedNonBudgted || "",
        capexSerialNo: "",
        recommendation: item.RecommendationJust || "",
        contractorName: item.ContractorName || "",
        annexure: item.AnnexureEnclosure || "",
        note1: item.AdditionalText1 || "",
        note2: item.AdditionalText2 || "",
        note3: item.AdditionalText3 || "",
        note4: item.AdditionalText4 || "",
        workflow: item.WorkflowType || "",
        doaMatrix: item.DOAClause || "",
        location: item.Location || "",
        natureOfApproval: item.NatureofApproval || "",
        approvers: levels.map(l => item[l]?.Id || item[`${l}Id`] || ""),
        approverDetails: levels.map(l => {
          const app = item[l];
          if (app && app.Id) {
            return {
              id: app.Id,
              name: app.Title || "",
              designation: app.JobTitle || app.Department || "",
              email: app.EMail || ""
            };
          }
          return null;
        }),
        uniqId: item.UniqID || item.UniqId || "",
        level: parseInt(item.WFLevel || item.Level || "0", 10),
        pendingAtId: (() => { const lvl = parseInt(item.WFLevel || item.Level || "1", 10); return (lvl >= 1 && lvl <= 12) ? (item[levels[lvl - 1]]?.Id || item[`ApprovalLevel${lvl}Id`] || null) : null; })(),
        initiatorId: item.Author ? item.Author.Id : undefined,
        commentHistory: item.CommentHistory || "",
        workflowHistory: item.WorkflowHistory || "",
        workflowHistoryHTML: item.WorkflowHistoryHTML || "",
        status: item.Status || "",
        PendingFrom:item.PendingFrom || ""
      };
    } catch (error) {
      console.error("NFAService: Exception in getNFA:", error);
      return null;
    }
  }



  public async getNFAList(): Promise<any[]> {
    try {
      const currentUser = await this.sp.web.currentUser();
      const currentUserId = currentUser.Id;

      const items = await this.sp.web.lists.getByTitle(this.listName).items
        .select("*","Id", "Title", "Dept","ContractorName","ApprovalAmount","PendingFrom", "Approvedrejectedby","Approvedrejectdatetime","Project", "RequestedOn", "Status", "WFLevel", "ApprovalLevel1/Title")
        .expand("ApprovalLevel1")
        .filter(`AuthorId eq ${currentUserId}`)();

      return items.map((item: any) => ({
        id: item.Id,
        requestNo: item.Title,
        function: item.Dept,
        contractorName: item.ContractorName || "",
        project: item.Project,
        createdOn: item.RequestedOn || item.Created || "",
        status: item.Status,
        Approvedrejectedby: item.Approvedrejectedby,
        Approvedrejectdatetime : item.Approvedrejectdatetime,
        pendingAt: item.PendingAt || (item.ApprovalLevel1 ? item.ApprovalLevel1.Title : ""),
        level: parseInt(item.WFLevel || item.Level || "0", 10),
        PendingFrom:item.PendingFrom ,
        doaMatrix: item.DOAClause || "",
        packageAwardedCost: item.ApprovalAmount || "0.00",
      }));
    } catch (error) {
      console.error("Error fetching NFA list:", error);
      return [];
    }
  }

  public async getInboxData(): Promise<any[]> {
    try {
      const currentUser = await this.sp.web.currentUser();
      const currentUserId = currentUser.Id;
      const currentUserName = (currentUser.Title || "").trim().toLowerCase();

      const resultMap = new Map<number, any>();

      // 1. Check DashboardList for active tasks pending for current user
      try {
        const dashboardItems = await this.sp.web.lists.getByTitle("DashboardList").items
          .filter(`Status eq 'Pending' and (CurrentApproverId eq ${currentUserId} or CurrentApprovalID eq '${currentUserId}')`)
          .top(500)();

        for (const d of dashboardItems) {
          const wfid = parseInt(d.WFID || d.Id, 10);
          if (!wfid || isNaN(wfid)) continue;

          const createdDate = d.Created ? new Date(d.Created) : new Date();
          const today = new Date();
          const diffTime = Math.abs(today.getTime() - createdDate.getTime());
          const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

          resultMap.set(wfid, {
            id: wfid,
            requestNo: d.Title || "",
            function: d.Function || d.Department || "",
            project: d.Project || "",
            createdDate: d.InitiatedOn || d.PendingFrom || createdDate.toLocaleDateString('en-IN'),
            status: d.Status || "Pending",
            pendingAt: d.CurrentApprovalName || currentUser.Title || "",
            pendingAtId: currentUserId,
            level: parseInt(d.WFLevel || "1", 10),
            amount: d.Amount || "0.00",
            initiator: d.InitiatedBy || "",
            agingDays: diffDays
          });
        }
      } catch (dashErr) {
        console.warn("NFAService: Error fetching from DashboardList:", dashErr);
      }

      // 2. Fetch active queries pending for current user from QueryMgt
      try {
        let queryMgtItems: any[] = [];
        try {
          try {
            queryMgtItems = await this.sp.web.lists.getByTitle("QueryMgtList").items
              .filter(`Status eq 'Pending' and QueryToId eq ${currentUserId}`)
              .select("Id", "Title", "WFID", "Function", "Project", "Amount", "WFLevel", "QueriedOn", "Created", "QueryById", "QueryBy/Title")
              .expand("QueryBy")
              .top(500)();
          } catch { }
          if (queryMgtItems.length === 0) {
            queryMgtItems = await this.sp.web.lists.getByTitle("QueryMgt").items
            .filter(`Status eq 'Pending' and QueryToId eq ${currentUserId}`)
            .select("Id", "Title", "WFID", "Function", "Project", "Amount", "WFLevel", "QueriedOn", "Created", "QueryById", "QueryBy/Title")
            .expand("QueryBy")
            .top(500)();
        } }catch {
          try {
          
            queryMgtItems = await this.sp.web.lists.getByTitle("QueryMgtList").items
              .filter(`Status eq 'Pending' and QueryToId eq ${currentUserId}`)
              .select("Id", "Title", "WFID", "Function", "Project", "Amount", "WFLevel", "QueriedOn", "Created", "QueryById", "QueryBy/Title")
              .expand("QueryBy")
              .top(500)();
          } catch { }
          if (queryMgtItems.length === 0) {
            queryMgtItems = await this.sp.web.lists.getByTitle("QueryMgt").items
              .filter(`Status eq 'Pending' and QueryToId eq ${currentUserId}`)
              .select("Id", "Title", "WFID", "Function", "Project", "Amount", "WFLevel", "QueriedOn", "Created", "QueryById")
              .top(500)();
          } 
        }
      
        for (const q of queryMgtItems) {
          const wfid = parseInt(q.WFID || q.Id, 10);
          if (!wfid || isNaN(wfid)) continue;

          const createdDate = q.Created ? new Date(q.Created) : new Date();
          const today = new Date();
          const diffTime = Math.abs(today.getTime() - createdDate.getTime());
          const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

          const queryByName = q.QueryBy ? q.QueryBy.Title : "Approver";

          resultMap.set(wfid, {
            id: wfid,
            requestNo: q.Title || "",
            function: q.Function || "",
            project: q.Project || "",
            createdDate: q.QueriedOn || createdDate.toLocaleDateString('en-IN'),
            status: "Query",
            pendingAt: currentUser.Title,
            pendingAtId: currentUserId,
            level: parseInt(q.WFLevel || "1", 10),
            amount: q.Amount || "0.00",
            initiator: queryByName,
            agingDays: diffDays
          });
        
      } }catch (qErr) {
        console.warn("NFAService: Error fetching queries from QueryMgt:", qErr);
      }

      // 3. Fetch active queries from DashboardList (QueryPendingToID eq current user)
      try {
        const queryDashItems = await this.sp.web.lists.getByTitle("DashboardList").items
          .filter(`Status eq 'Query' and (QueryPendingToID eq '${currentUserId}' or QueryPendingToID eq ${currentUserId})`)
          .top(500)();

        for (const d of queryDashItems) {
          const wfid = parseInt(d.WFID || d.Id, 10);
          if (!wfid || isNaN(wfid)) continue;

          if (!resultMap.has(wfid)) {
            const createdDate = d.Created ? new Date(d.Created) : new Date();
            const today = new Date();
            const diffTime = Math.abs(today.getTime() - createdDate.getTime());
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

            resultMap.set(wfid, {
              id: wfid,
              requestNo: d.Title || "",
              function: d.Function || d.Department || "",
              project: d.Project || "",
              createdDate: d.QueryPendingOn || d.PendingFrom || createdDate.toLocaleDateString('en-IN'),
              status: "Query",
              pendingAt: currentUser.Title,
              pendingAtId: currentUserId,
              level: parseInt(d.WFLevel || "1", 10),
              amount: d.Amount || "0.00",
              initiator: d.CurrentApprovalName || "Approver",
              agingDays: diffDays
            });
          }
        }
      } catch (qDashErr) {
        console.warn("NFAService: Error fetching queries from DashboardList:", qDashErr);
      }

      // 4. Also check DigitalApprovalNotes directly
      try {
        const levels = Array.from({ length: 12 }, (_, i) => `ApprovalLevel${i + 1}Id`);
        const items = await this.sp.web.lists.getByTitle(this.listName).items
          .select("Id", "Title", "Dept", "Project", "RequestedOn", "Status", "PendingAt", "WFLevel", "ApprovalAmount", "Author/Title", "Created", ...levels)
          .expand("Author")
          .filter("(Status eq 'Pending' or Status eq 'Query')")
          .top(1000)();

        for (const item of items) {
          const wfid = item.Id;
          const status = item.Status || "";
          const levelNum = parseInt(item.WFLevel || "1", 10);
          const currentLevelKey = `ApprovalLevel${levelNum}Id`;
          const currentApproverAtLevelId = item[currentLevelKey];

          const pendingAtText = (item.PendingAt || "").trim().toLowerCase();

          // Check if user matches by ID at current level, or Level 1 ID, or PendingAt display name
          const isUserMatch =
            ((status || "").toLowerCase() === "pending" &&
             ((currentApproverAtLevelId && currentApproverAtLevelId === currentUserId) ||
              (levelNum === 1 && item.ApprovalLevel1Id === currentUserId) ||
              (pendingAtText && pendingAtText === currentUserName))) ||
            ((status || "").toLowerCase() === "query" &&
             (pendingAtText && pendingAtText === currentUserName));


          if (isUserMatch && !resultMap.has(wfid)) {
            const createdDate = new Date(item.Created);
            const today = new Date();
            const diffTime = Math.abs(today.getTime() - createdDate.getTime());
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

            resultMap.set(wfid, {
              id: wfid,
              requestNo: item.Title,
              function: item.Dept,
              project: item.Project,
              createdDate: item.RequestedOn || createdDate.toLocaleDateString('en-IN'),
              status: status,
              pendingAt: item.PendingAt || currentUser.Title,
              pendingAtId: currentUserId,
              level: levelNum,
              amount: item.ApprovalAmount || "0.00",
              initiator: item.Author ? item.Author.Title : "",
              agingDays: diffDays
            });
          }
        }
      } catch (nfaErr) {
        console.error("NFAService: Error fetching from DigitalApprovalNotes:", nfaErr);
      }

      return Array.from(resultMap.values());
    } catch (error) {
      console.error("Error fetching Inbox data:", error);
      return [];
    }
  }

  public async getHistory(nfaId: number): Promise<any[]> {
    // Audit logs are now stored directly in the WorkflowHistory column of the DigitalApprovalNotes list item.
    // Returning an empty array forces the UI to parse it from the loaded item properties.
    return [];
  }
}
