import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface IFunctionStatusRecord {
  functionName: string;
  approved: number;
  pending: number;
  query: number;
  delegated: number;
  rejected: number;
  grandTotal: number;
}

export interface IAgeingStatusRecord {
  id?: number;
  pendingAt: string;
  approvalNotes: string;
  status: string;
  function: string;
  amount: string;
  ageingFromCreated: number;
  ageingFromLastApprover: number;
}

export interface IDayWiseAgeingRecord {
  pendingDays: number;
  count: number;
}

export interface IGenericReportRecord {
  id: number;
  nfaNo: string;
  subjectMatter: string;
  function: string;
  project: string;
  contractor: string;
  packageAwarded: string;
  budgetedAmount: string;
  status: string;
  createdBy: string;
  createdOn: string;
}

export interface IGenericReportFilterParams {
  fromDate?: string;
  toDate?: string;
  functionId?: string;
  functionName?: string;
  project?: string;
  projectId?: string;
  contractor?: string;
  packageCostFrom?: string;
  packageCostTo?: string;
}

export interface IApproverPendencyRecord {
  approverName: string;
  noOfPendingTask: number;
  status: string;
}

export interface IApproverPendencyFilterParams {
  fromDate?: string;
  toDate?: string;
  status?: string;
}

export class ReportService {
  private sp: SPFI;
  private listName: string = "DigitalApprovalNotes";

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  public async getFunctionWiseStatus(
    fromDate?: string,
    toDate?: string,
    functionIdOrName?: string
  ): Promise<IFunctionStatusRecord[]> {
    try {
      // 1. Fetch functions from FunctionMaster list
      let funcList: { id: string; title: string }[] = [];
      const listNames = ["FunctionMaster", "Function Master"];
      for (const name of listNames) {
        try {
          const items = await this.sp.web.lists.getByTitle(name).items
            .select("Id", "Title")
            .orderBy("Title", true)
            .top(500)();
          if (items && items.length > 0) {
            funcList = items.map((f: any) => ({
              id: String(f.Id),
              title: String(f.Title || "").trim()
            })).filter(f => f.title);
            break;
          }
        } catch {
          // try next name
        }
      }

      // 2. Fetch items from DashboardList
      const dashboardItems = await this.sp.web.lists.getByTitle("DashboardList").items
        .select("Id", "Title", "Function", "FunctionID", "Status", "Created")
        .top(5000)();

      const parseDate = (val: any): Date | null => {
        if (!val) return null;
        if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
        const str = String(val).trim();
        const dmyMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (dmyMatch) {
          return new Date(parseInt(dmyMatch[3], 10), parseInt(dmyMatch[2], 10) - 1, parseInt(dmyMatch[1], 10));
        }
        const d = new Date(str);
        return isNaN(d.getTime()) ? null : d;
      };

      const from = fromDate ? new Date(fromDate) : null;
      if (from) from.setHours(0, 0, 0, 0);

      const to = toDate ? new Date(toDate) : null;
      if (to) to.setHours(23, 59, 59, 999);

      // Filter by Created date in memory
      const filteredItems = dashboardItems.filter(item => {
        if (from || to) {
          if (!item.Created) return false;
          const cDate = parseDate(item.Created);
          if (!cDate) return false;
          if (from && cDate < from) return false;
          if (to && cDate > to) return false;
        }
        return true;
      });

      // 3. Aggregate counts per function
      const recordMap: Record<string, IFunctionStatusRecord> = {};
      funcList.forEach(fn => {
        recordMap[fn.title] = {
          functionName: fn.title,
          approved: 0,
          pending: 0,
          query: 0,
          delegated: 0,
          rejected: 0,
          grandTotal: 0
        };
      });

      filteredItems.forEach(item => {
        const itemFuncId = String(item.FunctionID || "").trim();
        const itemFuncName = String(item.Function || "").trim().toLowerCase();

        // Find matching function from master
        const matched = funcList.find(fn => 
          (itemFuncId && fn.id === itemFuncId) || 
          (itemFuncName && fn.title.toLowerCase() === itemFuncName)
        );

        const targetFunctionName = matched ? matched.title : (item.Function || "");
        if (!targetFunctionName) return;

        if (!recordMap[targetFunctionName]) {
          recordMap[targetFunctionName] = {
            functionName: targetFunctionName,
            approved: 0,
            pending: 0,
            query: 0,
            delegated: 0,
            rejected: 0,
            grandTotal: 0
          };
        }

        const status = (item.Status || "").toLowerCase().trim();
        if (status === "approved" || status === "completed") {
          recordMap[targetFunctionName].approved++;
        } else if (status === "pending") {
          recordMap[targetFunctionName].pending++;
        } else if (status === "query" || status === "query delegated") {
          recordMap[targetFunctionName].query++;
        } else if (status === "delegated") {
          recordMap[targetFunctionName].delegated++;
        } else if (status === "rejected") {
          recordMap[targetFunctionName].rejected++;
        }
      });

      // 4. Calculate grand total for each row
      let results: IFunctionStatusRecord[] = Object.keys(recordMap).map(fName => {
        const rec = recordMap[fName];
        rec.grandTotal = rec.approved + rec.pending + rec.query + rec.delegated + rec.rejected;
        return rec;
      });

      // 5. If specific function selected
      if (functionIdOrName && functionIdOrName !== "0" && functionIdOrName !== "All" && functionIdOrName !== "-All-") {
        const filterStr = functionIdOrName.trim().toLowerCase();
        const matchedFunc = funcList.find(f => f.id === functionIdOrName || f.title.toLowerCase() === filterStr);
        const targetTitle = matchedFunc ? matchedFunc.title.toLowerCase() : filterStr;

        results = results.filter(r => 
          r.functionName.toLowerCase() === targetTitle
        );
      }

      return results.sort((a, b) => a.functionName.localeCompare(b.functionName));
    } catch (error) {
      console.error("ReportService: Error in getFunctionWiseStatus:", error);
      return [];
    }
  }

  public async getAgeingReport(
    fromDate?: string,
    toDate?: string,
    functionName?: string
  ): Promise<IAgeingStatusRecord[]> {
    try {
      // Helper function to parse dates formatted as DD/MM/YYYY or ISO strings
      const parseDate = (val: any): Date | null => {
        if (!val) return null;
        if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
        const str = String(val).trim();
        const dmyMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (dmyMatch) {
          const day = parseInt(dmyMatch[1], 10);
          const month = parseInt(dmyMatch[2], 10) - 1;
          const year = parseInt(dmyMatch[3], 10);
          return new Date(year, month, day);
        }
        const d = new Date(str);
        return isNaN(d.getTime()) ? null : d;
      };

      // Fetch pending items from DashboardList
      let items: any[] = [];
      try {
        items = await this.sp.web.lists.getByTitle("DashboardList").items
          .select(
            "WFID",
            "Title",
            "Created",
            "Status",
            "PendingFromDate",
            "PendingFrom",
            "InitiatedOn",
            "CurrentApprovalName",
            "CurrentApprover",
            "Function",
            "Amount"
          )
          .filter(
            "(Status eq 'Pending' or Status eq 'Query' or Status eq 'Delegated' or Status eq 'Query Delegated')"
          )
          .top(5000)();
      } catch {
        items = await this.sp.web.lists.getByTitle("DashboardList").items
          .filter(
            "(Status eq 'Pending' or Status eq 'Query' or Status eq 'Delegated' or Status eq 'Query Delegated')"
          )
          .top(5000)();
      }

      const today = new Date();

      // Aggregate and calculate ageing in memory
      let records = items.map(item => {
        const initiatedOnDate = parseDate(item.InitiatedOn) || parseDate(item.Created);
        const pendingFromDate = parseDate(item.PendingFromDate) || parseDate(item.PendingFrom) || parseDate(item.Created);

        const ageingFromCreated = initiatedOnDate 
          ? Math.floor((today.getTime() - initiatedOnDate.getTime()) / (1000 * 60 * 60 * 24)) 
          : 0;

        const ageingFromLastApprover = pendingFromDate 
          ? Math.floor((today.getTime() - pendingFromDate.getTime()) / (1000 * 60 * 60 * 24)) 
          : 0;

        const pendingAtVal = item.CurrentApprovalName || (typeof item.CurrentApprover === 'string' ? item.CurrentApprover : item.CurrentApprover?.Title) || "Unknown";

        const parsedWfid = item.WFID ? parseInt(item.WFID, 10) : NaN;
        const recordId = !isNaN(parsedWfid) ? parsedWfid : (typeof item.Id === 'number' ? item.Id : undefined);

        return {
          id: recordId,
          pendingAt: pendingAtVal,
          approvalNotes: item.Title || "",
          status: item.Status || "Pending",
          function: item.Function || "",
          amount: item.Amount || "0.00",
          ageingFromCreated: ageingFromCreated >= 0 ? ageingFromCreated : 0,
          ageingFromLastApprover: ageingFromLastApprover >= 0 ? ageingFromLastApprover : 0,
          initiatedDate: initiatedOnDate,
          functionRaw: item.Function
        };
      });

      // Filter by date range (InitiatedOn / Created) in memory
      if (fromDate || toDate) {
        const from = fromDate ? new Date(fromDate) : null;
        if (from) from.setHours(0, 0, 0, 0);

        const to = toDate ? new Date(toDate) : null;
        if (to) to.setHours(23, 59, 59, 999);

        records = records.filter(item => {
          if (!item.initiatedDate) return true;
          if (from && item.initiatedDate < from) return false;
          if (to && item.initiatedDate > to) return false;
          return true;
        });
      }

      // Filter by Function if selected
      if (functionName && functionName !== "0" && functionName !== "-All-") {
        records = records.filter(item => item.functionRaw === functionName);
      }

      return records.map(r => ({
        id: r.id,
        pendingAt: r.pendingAt,
        approvalNotes: r.approvalNotes,
        status: r.status,
        function: r.function,
        amount: r.amount,
        ageingFromCreated: r.ageingFromCreated,
        ageingFromLastApprover: r.ageingFromLastApprover
      }));
    } catch (error) {
      console.error("ReportService: Error in getAgeingReport:", error);
      return [];
    }
  }

  public async getDayWiseAgeingReport(
    fromDate?: string,
    toDate?: string
  ): Promise<IDayWiseAgeingRecord[]> {
    try {
      const items = await this.sp.web.lists.getByTitle("DashboardList").items
        .select(
          "WFID",
          "Title",
          "Created",
          "Status",
          "PendingFromDate",
          "QueryPendingOn",
          "DelegatedOn",
          "CurrentApprovalName",
          "Function",
          "Amount"
        )
        .filter(
          "(Status eq 'Pending' or Status eq 'Query' or Status eq 'Delegated' or Status eq 'Query Delegated')"
        )
        .top(5000)();

      const parseDate = (val: any): Date | null => {
        if (!val) return null;
        if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
        const str = String(val).trim();
        const dmyMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (dmyMatch) {
          const day = parseInt(dmyMatch[1], 10);
          const month = parseInt(dmyMatch[2], 10) - 1;
          const year = parseInt(dmyMatch[3], 10);
          return new Date(year, month, day);
        }
        const d = new Date(str);
        return isNaN(d.getTime()) ? null : d;
      };

      // Filter by Created date in memory
      let filtered = items;
      if (fromDate || toDate) {
        const from = fromDate ? new Date(fromDate) : null;
        if (from) from.setHours(0, 0, 0, 0);

        const to = toDate ? new Date(toDate) : null;
        if (to) to.setHours(23, 59, 59, 999);

        filtered = filtered.filter(item => {
          if (!item.Created) return false;
          const cDate = parseDate(item.Created);
          if (!cDate) return false;
          if (from && cDate < from) return false;
          if (to && cDate > to) return false;
          return true;
        });
      }

      const today = new Date();
      today.setHours(0, 0, 0, 0);

      const daysCountMap: Record<number, number> = {};

      filtered.forEach(item => {
        let pendingDate: Date | null = null;
        const st = (item.Status || "").toLowerCase();

        if (st === "pending") {
          pendingDate = parseDate(item.PendingFromDate);
        } else if (st === "query") {
          pendingDate = parseDate(item.QueryPendingOn);
        } else if (st === "delegated" || st === "query delegated") {
          pendingDate = parseDate(item.DelegatedOn);
        }

        if (!pendingDate) {
          pendingDate = parseDate(item.Created);
        }

        if (pendingDate) {
          const pDate = new Date(pendingDate);
          pDate.setHours(0, 0, 0, 0);
          const diffMs = today.getTime() - pDate.getTime();
          const days = Math.max(0, Math.round(diffMs / (1000 * 60 * 60 * 24)));
          daysCountMap[days] = (daysCountMap[days] || 0) + 1;
        }
      });

      const uniqueDays = Object.keys(daysCountMap)
        .map(Number)
        .sort((a, b) => a - b);

      return uniqueDays.map(d => ({
        pendingDays: d,
        count: daysCountMap[d]
      }));
    } catch (error) {
      console.error("ReportService: Error in getDayWiseAgeingReport:", error);
      return [];
    }
  }

  public async getContractorList(): Promise<string[]> {
    try {
      const items = await this.sp.web.lists.getByTitle("DigitalApprovalNotes").items
        .select("ContractorName")
        .top(5000)();
      const names: string[] = items
        .map((i: any) => String(i.ContractorName || "").trim())
        .filter((name: string) => Boolean(name));
      const distinct: string[] = Array.from(new Set(names)).sort((a: string, b: string) => a.localeCompare(b));
      return distinct;
    } catch (e) {
      console.error("ReportService: Error fetching contractors:", e);
      return [];
    }
  }

  public async getProjectList(): Promise<{ label: string; value: string }[]> {
    const listNames = ["Project Master", "ProjectMaster"];
    for (const name of listNames) {
      try {
        const items = await this.sp.web.lists
          .getByTitle(name)
          .items
          .select("Id", "Title", "ProjectName", "IsActive")
          .filter("IsActive eq 1 or IsActive eq '1'")
          .orderBy("ProjectName", true)
          .top(500)();

        if (items && items.length > 0) {
          return items.map((item: any) => {
            const displayName = item.Title && item.ProjectName 
              ? `${item.Title}-${item.ProjectName}` 
              : (item.ProjectName || item.Title || String(item.Id));
            return {
              label: displayName,
              value: String(item.Id)
            };
          });
        }
      } catch {
        // try next list name
      }
    }

    // Fallback: extract distinct Project values directly from DigitalApprovalNotes
    try {
      const notes = await this.sp.web.lists
        .getByTitle("DigitalApprovalNotes")
        .items
        .select("Project")
        .top(3000)();
      const projNames: string[] = notes
        .map((n: any) => String(n.Project || "").trim())
        .filter((p: string) => Boolean(p));
      const distinct: string[] = Array.from(new Set(projNames)).sort();
      return distinct.map((p: string) => ({ label: p, value: p }));
    } catch {
      return [];
    }
  }

  public async getGenericReport(
    filters: IGenericReportFilterParams
  ): Promise<IGenericReportRecord[]> {
    try {
      const items = await this.sp.web.lists.getByTitle("DigitalApprovalNotes").items
        .select(
          "ID",
          "Title",
          "SubjectMatter",
          "FunctionName",
          "FunctionID",
          "Project",
          "ContractorName",
          "ApprovalAmount",
          "Budget",
          "Status",
          "Created",
          "NFAType",
          "Author/Title"
        )
        .expand("Author")
        .filter("Status ne 'Draft'")
        .top(5000)();

      const parseDate = (val: any): Date | null => {
        if (!val) return null;
        if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
        const str = String(val).trim();
        const dmyMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (dmyMatch) {
          return new Date(parseInt(dmyMatch[3], 10), parseInt(dmyMatch[2], 10) - 1, parseInt(dmyMatch[1], 10));
        }
        const d = new Date(str);
        return isNaN(d.getTime()) ? null : d;
      };

      const from = filters.fromDate ? new Date(filters.fromDate) : null;
      if (from) from.setHours(0, 0, 0, 0);

      const to = filters.toDate ? new Date(filters.toDate) : null;
      if (to) to.setHours(23, 59, 59, 999);

      const costFrom = filters.packageCostFrom ? parseFloat(filters.packageCostFrom) : null;
      const costTo = filters.packageCostTo ? parseFloat(filters.packageCostTo) : null;

      const filtered = items.filter(item => {
        if (item.NFAType === "Invoice") return false;

        // Date filter
        if (from || to) {
          if (!item.Created) return false;
          const cDate = parseDate(item.Created);
          if (!cDate) return false;
          if (from && cDate < from) return false;
          if (to && cDate > to) return false;
        }

        // Project / Entity filter
        if (
          filters.project &&
          filters.project !== "0" &&
          filters.project !== "-Select Project-" &&
          filters.project !== "-Select Entity-" &&
          filters.project !== "-Select Project / Entity-"
        ) {
          const itemProj = (item.Project || "").toLowerCase().trim();
          const targetProj = filters.project.toLowerCase().trim();
          const targetId = filters.projectId ? String(filters.projectId).trim().toLowerCase() : "";

          // 1. Direct or substring match on project name
          const matchByName = Boolean(
            itemProj && (
              itemProj.includes(targetProj) ||
              targetProj.includes(itemProj)
            )
          );

          // 2. If targetProj has "PRJ-01-Downtown" (code-name pattern), check the name part
          const splitProj = targetProj.includes('-') 
            ? targetProj.split('-').slice(1).join('-').trim() 
            : "";
          const matchBySplit = Boolean(splitProj && itemProj.includes(splitProj));

          // 3. Match by ID if numeric ID was stored or matches
          const matchById = Boolean(
            targetId && (
              itemProj === targetId ||
              String((item as any).ProjectID || "") === targetId
            )
          );

          if (!matchByName && !matchBySplit && !matchById) {
            return false;
          }
        }

        // Contractor filter
        if (filters.contractor && filters.contractor !== "0" && filters.contractor !== "-Select Contractor-") {
          const itemContractor = (item.ContractorName || "").toLowerCase();
          const targetContractor = filters.contractor.toLowerCase();
          if (!itemContractor.includes(targetContractor)) return false;
        }

        // Function filter
        if (filters.functionId && filters.functionId !== "0" && filters.functionId !== "-Select Function-") {
          const funcIdMatch = String(item.FunctionID || "") === String(filters.functionId);
          const funcNameMatch = (item.FunctionName || "").toLowerCase() === (filters.functionName || "").toLowerCase();
          if (!funcIdMatch && !funcNameMatch) return false;
        }

        // Package cost filter
        const amount = parseFloat(String(item.ApprovalAmount || "0").replace(/,/g, ''));
        if (costFrom !== null && !isNaN(costFrom) && amount < costFrom) return false;
        if (costTo !== null && !isNaN(costTo) && amount > costTo) return false;

        return true;
      });

      const formatDate = (dateVal: any): string => {
        const d = parseDate(dateVal);
        if (!d) return "";
        const day = d.getDate() < 10 ? '0' + d.getDate() : String(d.getDate());
        const month = (d.getMonth() + 1) < 10 ? '0' + (d.getMonth() + 1) : String(d.getMonth() + 1);
        const year = d.getFullYear();
        return `${day}/${month}/${year}`;
      };

      return filtered.map(item => ({
        id: item.ID,
        nfaNo: item.Title || "",
        subjectMatter: item.SubjectMatter || "",
        function: item.FunctionName || "",
        project: item.Project || "",
        contractor: item.ContractorName || "",
        packageAwarded: item.ApprovalAmount || "0.00",
        budgetedAmount: item.Budget || "0.00",
        status: item.Status || "",
        createdBy: item.Author?.Title || "",
        createdOn: formatDate(item.Created)
      }));
    } catch (error) {
      console.error("ReportService: Error in getGenericReport:", error);
      return [];
    }
  }

  public async getApproverWisePendency(
    filters: IApproverPendencyFilterParams
  ): Promise<IApproverPendencyRecord[]> {
    try {
      const items = await this.sp.web.lists.getByTitle("DashboardList").items
        .select(
          "Title",
          "Status",
          "CurrentApprovalName",
          "CurrentApprovalID",
          "QueryPendingToName",
          "QueryPendingToID",
          "DelegatedToName",
          "DelegatedToID",
          "PendingFromDate",
          "Created"
        )
        .filter(
          "(Status eq 'Pending' or Status eq 'Query' or Status eq 'Delegated' or Status eq 'Query Delegated')"
        )
        .top(5000)();

      const parseDate = (val: any): Date | null => {
        if (!val) return null;
        if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
        const str = String(val).trim();
        const dmyMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (dmyMatch) {
          return new Date(parseInt(dmyMatch[3], 10), parseInt(dmyMatch[2], 10) - 1, parseInt(dmyMatch[1], 10));
        }
        const d = new Date(str);
        return isNaN(d.getTime()) ? null : d;
      };

      const from = filters.fromDate ? new Date(filters.fromDate) : null;
      if (from) from.setHours(0, 0, 0, 0);

      const to = filters.toDate ? new Date(filters.toDate) : null;
      if (to) to.setHours(23, 59, 59, 999);

      const countMap: Record<string, { approverName: string; status: string; count: number }> = {};

      items.forEach(item => {
        const itemDate = parseDate(item.PendingFromDate) || parseDate(item.Created);
        if (from || to) {
          if (!itemDate) return;
          if (from && itemDate < from) return;
          if (to && itemDate > to) return;
        }

        const rawStatus = (item.Status || "").toLowerCase();
        let approverName = "";
        let effectiveStatus = "";

        if (rawStatus === "pending") {
          approverName = item.CurrentApprovalName || "Unknown";
          effectiveStatus = "Pending";
        } else if (rawStatus === "query" || rawStatus === "query delegated") {
          approverName = item.QueryPendingToName || item.CurrentApprovalName || "Unknown";
          effectiveStatus = "Query";
        } else if (rawStatus === "delegated") {
          approverName = item.DelegatedToName || item.CurrentApprovalName || "Unknown";
          effectiveStatus = "Delegated";
        }

        if (!approverName || approverName.trim() === "") return;
        approverName = approverName.trim();

        const filterSt = (filters.status || "All").toLowerCase();
        if (filterSt !== "all" && filterSt !== "" && filterSt !== effectiveStatus.toLowerCase()) {
          return;
        }

        const key = `${approverName}###${effectiveStatus}`;
        if (!countMap[key]) {
          countMap[key] = {
            approverName,
            status: effectiveStatus,
            count: 0
          };
        }
        countMap[key].count += 1;
      });

      const records: IApproverPendencyRecord[] = Object.keys(countMap)
        .map(k => ({
          approverName: countMap[k].approverName,
          noOfPendingTask: countMap[k].count,
          status: countMap[k].status
        }))
        .filter(r => r.noOfPendingTask > 0)
        .sort((a, b) => a.approverName.localeCompare(b.approverName));

      return records;
    } catch (error) {
      console.error("ReportService: Error in getApproverWisePendency:", error);
      return [];
    }
  }
}
