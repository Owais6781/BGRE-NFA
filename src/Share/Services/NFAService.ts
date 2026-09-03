import { WebPartContext } from "@microsoft/sp-webpart-base";
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
  capexSerialNo: string;
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
  attachments?: FileList | null;
  uniqId: string;
  level: number;
  pendingAtId: number | null;
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
    if (!form.project.trim()) errors.push({ field: "project", message: "Project is required." });
    if (status === "Draft") return { isValid: errors.length === 0, errors };

    if (!form.originatorDept.trim()) errors.push({ field: "originatorDept", message: "Dept is required." });
    if (!form.referenceSubject.trim()) errors.push({ field: "referenceSubject", message: "Subject is required." });
    if (!form.projectSubEntity || form.projectSubEntity.length === 0) errors.push({ field: "projectSubEntity", message: "Sub-Entity is required." });

    const stripHtml = (html: string): string => (html || "").replace(/<[^>]*>/g, "").replace(/&nbsp;/g, " ").trim();
    if (!stripHtml(form.introduction)) errors.push({ field: "introduction", message: "Introduction is required." });
    if (!stripHtml(form.chronology)) errors.push({ field: "chronology", message: "Chronology is required." });
    if (!stripHtml(form.commercialProposal)) errors.push({ field: "commercialProposal", message: "Commercial Proposals is required." });

    if (!form.approvers[0]) errors.push({ field: "approvers", message: "Approver Level-1 is required." });

    const cost = parseFloat(String(form.packageAwardedCost || "").replace(/,/g, ''));
    if (isNaN(cost) || cost <= 0) {
      errors.push({ field: "packageAwardedCost", message: "Awarded Cost is required and must be greater than 0." });
    }
    const budgetVal = parseFloat(String(form.budget || "").replace(/,/g, ''));
    if (isNaN(budgetVal) || budgetVal <= 0) {
      errors.push({ field: "budget", message: "Budget is required and must be greater than 0." });
    }

    return { isValid: errors.length === 0, errors };
  }

  private buildRequestBody(
    form: INFAFormData,
    status: NFAStatus,
    resolvedApprovers: (number | null)[],
    projectName: string,
    subEntityNames: string
  ): Record<string, any> {
    const now = new Date();
    const body: Record<string, any> = {
      Title: form.referenceSubject,
      Project: projectName,
      ProjectSubEntity: subEntityNames,
      Dept: form.originatorDept,
      Status: status,
      SubjectMatter: form.referenceSubject,
      BriefDesc: form.introduction,
      Chronology: form.chronology,
      CommercialProposal: form.commercialProposal,
      PaymentTerms: form.paymentTerms,
      LiquidatedDamage: form.liquidatedDamage,
      ContractPerformance: form.contractPerf,
      DefectLibilityPeriod: form.defectLiability,
      WorkCompletionPeriod: form.workCompletion,
      BudgetRemarks: form.budgetRemarks,
      ApprovalAmount: form.packageAwardedCost,
      Budget: form.budget,
      RecommendationJust: form.recommendation,
      ContractorName: form.contractorName,
      AnnexureEnclosure: form.annexure,
      AdditionalText1: form.note1,
      AdditionalText2: form.note2,
      AdditionalText3: form.note3,
      AdditionalText4: form.note4,
      WorkflowType: form.workflow,
      DOAClause: form.doaMatrix,
      RequestedOn: form.date,
      Month: now.toLocaleString('default', { month: 'long' }),
      Year: now.getFullYear().toString(),
      PendingAtId: resolvedApprovers[0] || null, 
      Level: 1, 
      UniqId: form.uniqId || "",
    };

    // Mapping 12 levels of approvers to ApprovalLevel1Id, ApprovalLevel2Id, etc.
    for (let i = 0; i < 12; i++) {
      const fieldName = `ApprovalLevel${i + 1}Id`;
      body[fieldName] = resolvedApprovers[i] || null;
    }

    return body;
  }

  private generateRequestNo(project: string, dateStr: string, status: NFAStatus, itemId: number): string {
    let projectInitials = "NFA";
    if (project) {
      const clean = project.trim();
      if (clean.includes(" ")) {
        projectInitials = clean.split(/\s+/).map(w => w[0]).join("").toUpperCase();
      } else {
        projectInitials = clean.toUpperCase();
      }
    }
    projectInitials = projectInitials.replace(/[^A-Z0-9]/g, "");

    let dateObj = new Date();
    if (dateStr && dateStr.includes("/")) {
      const parts = dateStr.split("/");
      if (parts.length === 3) {
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10) - 1;
        const year = parseInt(parts[2], 10);
        const parsedDate = new Date(year, month, day);
        if (!isNaN(parsedDate.getTime())) {
          dateObj = parsedDate;
        }
      }
    }

    const year = dateObj.getFullYear().toString();
    const month = dateObj.toLocaleString('en-US', { month: 'short' }).toUpperCase();

    let suffix = "Draf";
    if (status !== "Draft") {
      if (itemId < 10) {
        suffix = "00" + itemId;
      } else if (itemId < 100) {
        suffix = "0" + itemId;
      } else {
        suffix = String(itemId);
      }
    }
    return `${projectInitials}-${year}-${month}-${suffix}`;
  }

  private async writeWFLog(
    itemId: number,
    requestNo: string,
    form: INFAFormData,
    resolvedApprovers: (number | null)[],
    projectName: string,
    functionOptions: { label: string, value: string }[]
  ): Promise<void> {
    try {
      const funcObj = functionOptions.find(f => f.value === form.originatorDept);
      const functionName = funcObj ? funcObj.label : form.originatorDept;

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

      // Log 1: Completed Action Log (Submitted)
      const logBody1: Record<string, any> = {
        Title: requestNo,
        WFID: String(itemId),
        FunctionID: form.originatorDept,
        FunctionName: functionName,
        Project: projectName,
        Amount: form.packageAwardedCost,
        Status: "Submitted",
        PendingAt: "",
        WFLevel: "0",
        Remarks: form.budgetRemarks || ""
      };

      await this.sp.web.lists.getByTitle("WFLogList").items.add(logBody1);

      // Log 2: Next Pending Task Log (Pending)
      const logBody2: Record<string, any> = {
        Title: requestNo,
        WFID: String(itemId),
        FunctionID: form.originatorDept,
        FunctionName: functionName,
        Project: projectName,
        Amount: form.packageAwardedCost,
        Status: "Pending",
        PendingAt: pendingAtName,
        WFLevel: "1",
        Remarks: form.budgetRemarks || ""
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
        const initiatorName = this.context.pageContext.user.displayName || "Initiator";

        // Dashboard Log: Current State (Pending)
        const dashboardBody: Record<string, any> = {
          Title: requestNo,
          Function: functionName,
          Project: projectName,
          Amount: form.packageAwardedCost,
          FunctionID: form.originatorDept,
          WFID: String(itemId),
          WFLevel: "1",
          Status: "Pending",
          DOA: form.doaMatrix || "",
          CurrentApproverId: pendingAtId,
          CurrentApprovalName: pendingAtName,
          CurrentApprovalID: pendingAtId ? String(pendingAtId) : "",
          PendingFrom: initiatorName,
          InitiatedBy: initiatorName,
          InitiatedOn: form.date || formatHistoryDate(now),
          PendingFromDate: now,
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
      const body = this.buildRequestBody(form, status, resolvedApprovers, projectName, subEntityNames);

      if (status === "Draft") {
        body.Title = this.generateRequestNo(projectName, form.date, "Draft", 0);
      }

      const result = await this.sp.web.lists.getByTitle(this.listName).items.add(body);
      const itemId = result.Id;

      let requestNo = body.Title;

      if (status === "Pending") {
        requestNo = this.generateRequestNo(projectName, form.date, "Pending", itemId);
        const initiatorName = this.context.pageContext.user.displayName || "Initiator";
        const dateStr = formatHistoryDate(new Date());

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

        const initWorkflowHistory = `NFA Request Initiated by ${initiatorName} on ${dateStr}</br>`;
        const pendingWorkflowHistory = pendingAtName
          ? `\nNFA Pending at ${pendingAtName} at level-1 on ${dateStr}</br>`
          : "";
        const fullWorkflowHistory = initWorkflowHistory + pendingWorkflowHistory;

        await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId).update({
          Title: requestNo,
          WorkflowHistory: fullWorkflowHistory
        });
        await this.writeWFLog(itemId, requestNo, form, resolvedApprovers, projectName, functionOptions);
      }

      return { success: true, itemId: itemId, message: "Saved successfully!" };
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
      const body = this.buildRequestBody(form, status, resolvedApprovers, projectName, subEntityNames);

      const requestNo = this.generateRequestNo(projectName, form.date, status, itemId);
      body.Title = requestNo;

      if (status === "Pending") {
        try {
          const item = await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId).select("WorkflowHistory")();
          const currentWorkflowHistory = item.WorkflowHistory || "";
          const initiatorName = this.context.pageContext.user.displayName || "Initiator";
          const dateStr = formatHistoryDate(new Date());

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

          const initWorkflowHistory = `NFA Request Initiated by ${initiatorName} on ${dateStr}</br>`;
          const pendingWorkflowHistory = pendingAtName
            ? `\nNFA Pending at ${pendingAtName} at level-1 on ${dateStr}</br>`
            : "";
          const newWorkflowEntry = initWorkflowHistory + pendingWorkflowHistory;
          
          body.WorkflowHistory = currentWorkflowHistory
            ? `${currentWorkflowHistory}\n${newWorkflowEntry}`
            : newWorkflowEntry;
        } catch (hErr) {
          console.warn("NFAService: Error reading WorkflowHistory in updateNFA:", hErr);
        }
      }

      await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId).update(body);

      if (status === "Pending") {
        await this.writeWFLog(itemId, requestNo, form, resolvedApprovers, projectName, functionOptions);
      }

      return { success: true, itemId, message: "Updated successfully!" };
    } catch (error) {
      console.error("NFAService: Error in updateNFA:", error);
      const errMsg = error instanceof Error ? error.message : "Error occurred during update.";
      return { success: false, itemId, message: errMsg };
    }
  }

  public async getNFA(itemId: number): Promise<INFAFormData | null> {
    try {
      const levels = Array.from({ length: 12 }, (_, i) => `ApprovalLevel${i + 1}`);
      const item = await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId)
        .select("*", "PendingAt/Id", "PendingAt/Title", "Author/Id", "Author/Title", "CommentHistory", "WorkflowHistory", ...levels.map(l => `${l}/Id`), ...levels.map(l => `${l}/Title`))
        .expand("PendingAt", "Author", ...levels)();

      return {
        project: item.Project || "",
        projectSubEntity: item.ProjectSubEntity ? item.ProjectSubEntity.split("; ") : [],
        originatorDept: item.Dept || "",
        date: item.RequestedOn || "",
        referenceSubject: item.SubjectMatter || item.Title || "",
        requestNo: item.Title || "",
        introduction: item.BriefDesc || "",
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
        approvers: levels.map(l => item[l]?.Id || ""),
        uniqId: item.UniqId || "",
        level: item.Level || 0,
        pendingAtId: item.PendingAt ? item.PendingAt.Id : null,
        initiatorId: item.Author ? item.Author.Id : undefined,
        commentHistory: item.CommentHistory || "",
        workflowHistory: item.WorkflowHistory || "",
        workflowHistoryHTML: item.WorkflowHistoryHTML || "",
        status: item.Status || ""
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
        .select("Id", "Title", "Dept", "Project", "RequestedOn", "Status", "PendingAt/Title", "Level")
        .expand("PendingAt")
        .filter(`AuthorId eq ${currentUserId}`)();

      return items.map((item: any) => ({
        id: item.Id,
        requestNo: item.Title,
        function: item.Dept,
        project: item.Project,
        createdOn: item.RequestedOn,
        status: item.Status,
        pendingAt: item.PendingAt ? item.PendingAt.Title : "",
        level: item.Level || 0
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

      const items = await this.sp.web.lists.getByTitle(this.listName).items
        .select("Id", "Title", "Dept", "Project", "RequestedOn", "Status", "PendingAt/Id", "PendingAt/Title", "Level", "ApprovalAmount", "Author/Title", "Created")
        .expand("PendingAt", "Author")
        .filter(`(Status eq 'Pending' or Status eq 'Query') and PendingAt/Id eq ${currentUserId}`)();

      return items.map((item: any) => {
        const createdDate = new Date(item.Created);
        const today = new Date();
        const diffTime = Math.abs(today.getTime() - createdDate.getTime());
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

        return {
          id: item.Id,
          requestNo: item.Title,
          function: item.Dept,
          project: item.Project,
          createdDate: item.RequestedOn || createdDate.toLocaleDateString('en-IN'),
          status: item.Status,
          pendingAt: item.PendingAt ? item.PendingAt.Title : "",
          pendingAtId: item.PendingAt ? item.PendingAt.Id : null,
          level: item.Level || 0,
          amount: item.ApprovalAmount || "0.00",
          initiator: item.Author ? item.Author.Title : "",
          agingDays: diffDays
        };
      });
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
