import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";

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
  private context: WebPartContext;
  private listName: string = "DigitalApprovalNotes";

  constructor(context: WebPartContext) {
    this.context = context;
  }

  public async resolveUserId(loginName: string): Promise<number | null> {
    try {
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const endpoint = `${siteUrl}/_api/web/ensureuser`;
      const response = await this.context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
          },
          body: JSON.stringify({ logonName: loginName }),
        }
      );
      if (response.ok) {
        const data = await response.json();
        return data.Id;
      }
      return null;
    } catch (error) {
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
      PendingAtId: resolvedApprovers[0] || null, // Updated to save as Person ID
      UniqId: form.uniqId || "",
    };

    for (let i = 0; i < 12; i++) {
      const fieldName = `ApprovalLevel${i + 1}Id`;
      body[fieldName] = resolvedApprovers[i] || null;
    }

    return body;
  }

  public async saveNFA(form: INFAFormData, status: NFAStatus, projectOptions: { label: string, value: string }[], subEntityOptions: { label: string, value: string }[]): Promise<INFASaveResponse> {
    try {
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items`;
      const resolvedApprovers = await Promise.all(
        form.approvers.map(async (a) => {
          if (typeof a === "string" && a) return await this.resolveUserId(a);
          return typeof a === "number" ? a : null;
        })
      );
      const projectObj = projectOptions.find(p => p.value === form.project);
      const projectName = projectObj ? projectObj.label : form.project;
      const subEntityNames = form.projectSubEntity.map(id => subEntityOptions.find(o => o.value === id)?.label || id).join("; ");
      const body = this.buildRequestBody(form, status, resolvedApprovers, projectName, subEntityNames);

      const options: ISPHttpClientOptions = {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "odata-version": "",
        },
        body: JSON.stringify(body),
      };

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        endpoint, SPHttpClient.configurations.v1, options
      );

      if (!response.ok) {
        const errorText = await response.text();
        return { success: false, itemId: null, message: `Failed: ${errorText}` };
      }

      const result = await response.json();
      return { success: true, itemId: result.Id, message: "Saved successfully!" };
    } catch (error) {
      return { success: false, itemId: null, message: "Error occurred during save." };
    }
  }

  public async updateNFA(itemId: number, form: INFAFormData, status: NFAStatus, projectOptions: { label: string, value: string }[], subEntityOptions: { label: string, value: string }[]): Promise<INFASaveResponse> {
    try {
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items(${itemId})`;

      const resolvedApprovers = await Promise.all(
        form.approvers.map(async (a) => {
          if (typeof a === "string" && a) return await this.resolveUserId(a);
          return typeof a === "number" ? a : null;
        })
      );

      const projectObj = projectOptions.find(p => p.value === form.project);
      const projectName = projectObj ? projectObj.label : form.project;
      const subEntityNames = form.projectSubEntity.map(id => subEntityOptions.find(o => o.value === id)?.label || id).join("; ");
      const body = this.buildRequestBody(form, status, resolvedApprovers, projectName, subEntityNames);

      const options: ISPHttpClientOptions = {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE",
        },
        body: JSON.stringify(body),
      };

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        endpoint, SPHttpClient.configurations.v1, options
      );

      if (!response.ok) return { success: false, itemId, message: "Update failed." };

      return { success: true, itemId, message: "Updated successfully!" };
    } catch (error) {
      return { success: false, itemId, message: "Error occurred during update." };
    }
  }

  public async getNFA(itemId: number): Promise<INFAFormData | null> {
    try {
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const levels = Array.from({ length: 12 }, (_, i) => `ApprovalLevel${i + 1}`);
      const select = `*,${levels.map(l => `${l}/Id`).join(',')}`;
      const expand = levels.join(',');
      const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items(${itemId})?$select=${select}&$expand=${expand}`;

      const response = await this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
      if (!response.ok) return null;

      const item = await response.json();

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
        uniqId: item.UniqId || ""
      };
    } catch (error) {
      console.error("Error fetching NFA:", error);
      return null;
    }
  }

  public async getNFAList(): Promise<any[]> {
    try {
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const select = "Id,Title,Dept,Project,RequestedOn,Status,PendingAt/Title,Level";
      const expand = "PendingAt";
      const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items?$select=${select}&$expand=${expand}`;

      const response = await this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
      if (!response.ok) return [];

      const data = await response.json();
      return data.value.map((item: any) => ({
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
}
