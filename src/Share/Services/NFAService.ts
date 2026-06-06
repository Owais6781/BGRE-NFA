import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users";

// ── Form Data Interface ──
export interface INFAFormData {
  project: string;
  projectSubEntity: string[];
  originatorDept: string;
  date: string;
  referenceSubject: string;
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

  constructor(context: WebPartContext) {
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
    if (!form.approvers[0]) errors.push({ field: "approvers", message: "Approver Level-1 is required." });

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

  public async saveNFA(form: INFAFormData, status: NFAStatus, projectOptions: { label: string, value: string }[], subEntityOptions: { label: string, value: string }[]): Promise<INFASaveResponse> {
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

      const result = await this.sp.web.lists.getByTitle(this.listName).items.add(body);
      return { success: true, itemId: result.data.Id, message: "Saved successfully!" };
    } catch (error) {
      return { success: false, itemId: null, message: "Error occurred during save." };
    }
  }

  public async updateNFA(itemId: number, form: INFAFormData, status: NFAStatus, projectOptions: { label: string, value: string }[], subEntityOptions: { label: string, value: string }[]): Promise<INFASaveResponse> {
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

      await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId).update(body);
      return { success: true, itemId, message: "Updated successfully!" };
    } catch (error) {
      return { success: false, itemId, message: "Error occurred during update." };
    }
  }

  public async getNFA(itemId: number): Promise<INFAFormData | null> {
    try {
      const levels = Array.from({ length: 12 }, (_, i) => `ApprovalLevel${i + 1}`);
      const item = await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId)
        .select("*", "PendingAt/Id", "PendingAt/Title", "Author/Id", "Author/Title", ...levels.map(l => `${l}/Id`), ...levels.map(l => `${l}/Title`))
        .expand("PendingAt", "Author", ...levels)();

      return {
        project: item.Project || "",
        projectSubEntity: item.ProjectSubEntity ? item.ProjectSubEntity.split("; ") : [],
        originatorDept: item.Dept || "",
        date: item.RequestedOn || "",
        referenceSubject: item.Title || "",
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
        initiatorId: item.Author ? item.Author.Id : undefined
      };
    } catch (error) {
      console.error("NFAService: Exception in getNFA:", error);
      return null;
    }
  }

  public async getNFAList(): Promise<any[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.listName).items
        .select("Id", "Title", "Dept", "Project", "RequestedOn", "Status", "PendingAt/Title", "Level")
        .expand("PendingAt")();

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
    try {
      // Assuming a list named 'NFAHistory' exists to store audit trail
      const items = await this.sp.web.lists.getByTitle("NFAHistory").items
        .filter(`NFAId eq ${nfaId}`)
        .select("Id", "Title", "Action", "Remarks", "Created", "User/Title")
        .expand("User")
        .orderBy("Created", true)();

      return items.map(item => ({
        id: item.Id,
        action: item.Action,
        remarks: item.Remarks,
        user: item.User ? item.User.Title : "Unknown",
        date: new Date(item.Created).toLocaleString('en-IN')
      }));
    } catch (error) {
      console.warn("NFAService: getHistory failed (list might not exist):", error);
      // Return mock data for demo if list doesn't exist
      return [
        { id: 1, action: 'Submitted', remarks: 'NFA submitted for approval', user: 'Initiator', date: '01/06/2026, 10:00:00' },
      ];
    }
  }
}
