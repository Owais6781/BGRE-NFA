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
  pendingAt: string;
  approvalNotes: string;
  status: string;
  function: string;
  amount: string;
  ageingFromCreated: number;
  ageingFromLastApprover: number;
}

export class ReportService {
  private sp: SPFI;
  private listName: string = "DigitalApprovalNotes";

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  public async getFunctionWiseStatus(
    fromDate: string,
    toDate: string,
    functionName?: string
  ): Promise<IFunctionStatusRecord[]> {
    try {
      // 1. Fetch functions from FunctionMaster list
      let functions: string[] = [];
      if (functionName) {
        functions = [functionName];
      } else {
        const funcItems = await this.sp.web.lists.getByTitle("FunctionMaster").items
          .select("Title")
          .orderBy("Title", true)
          .top(500)();
        functions = funcItems.map(f => f.Title).filter(t => t);
      }

      // 2. Fetch all lifecycle items from DigitalApprovalNotes (select Dept, Status, and Created date)
      let notesItems = await this.sp.web.lists.getByTitle("DigitalApprovalNotes").items
        .select("Id", "Dept", "Status", "Created")
        .top(5000)();

      // 3. Fetch active status tracking from DashboardList to determine active queries and delegations
      let dashboardItems = await this.sp.web.lists.getByTitle("DashboardList").items
        .select("WFID", "Status", "DelegatedToName")
        .top(5000)();

      const delegatedWfids = new Set<string>();
      const queryWfids = new Set<string>();
      dashboardItems.forEach((d: any) => {
        if (d.WFID) {
          if (d.DelegatedToName) {
            delegatedWfids.add(d.WFID);
          }
          if (d.Status === "Query") {
            queryWfids.add(d.WFID);
          }
        }
      });

      // Filter notes items by date range in memory
      if (fromDate || toDate) {
        const from = fromDate ? new Date(fromDate) : null;
        if (from) from.setHours(0, 0, 0, 0);

        const to = toDate ? new Date(toDate) : null;
        if (to) to.setHours(23, 59, 59, 999);

        notesItems = notesItems.filter(item => {
          if (!item.Created) return false;
          const itemDate = new Date(item.Created);
          if (from && itemDate < from) return false;
          if (to && itemDate > to) return false;
          return true;
        });
      }

      // If a specific function is filtered
      if (functionName) {
        notesItems = notesItems.filter(item => item.Dept === functionName);
      }

      // 4. Aggregate counts per function
      const recordMap: Record<string, IFunctionStatusRecord> = {};
      functions.forEach(fName => {
        recordMap[fName] = {
          functionName: fName,
          approved: 0,
          pending: 0,
          query: 0,
          delegated: 0,
          rejected: 0,
          grandTotal: 0
        };
      });

      notesItems.forEach(item => {
        const dept = item.Dept || "";
        if (recordMap[dept]) {
          const status = (item.Status || "").toLowerCase();
          const wfIdStr = item.Id ? item.Id.toString() : "";

          if (status === "approved") {
            recordMap[dept].approved++;
          } else if (status === "rejected") {
            recordMap[dept].rejected++;
          } else if (status === "query" || queryWfids.has(wfIdStr)) {
            recordMap[dept].query++;
          } else if (status === "pending") {
            if (delegatedWfids.has(wfIdStr)) {
              recordMap[dept].delegated++;
            } else {
              recordMap[dept].pending++;
            }
          }
        }
      });

      // Calculate Grand Total for each row
      const results: IFunctionStatusRecord[] = [];
      functions.forEach(fName => {
        const rec = recordMap[fName];
        rec.grandTotal = rec.approved + rec.pending + rec.query + rec.delegated + rec.rejected;
        results.push(rec);
      });

      return results;
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
      // Fetch pending items from DashboardList
      let items = await this.sp.web.lists.getByTitle("DashboardList").items
        .filter("(Status eq 'Pending' or Status eq 'Query')")
        .select("Title", "Function", "Amount", "Status", "CurrentApprovalName", "InitiatedOn", "PendingFromDate")
        .top(5000)();

      // Aggregate and calculate ageing in memory
      let records = items.map(item => {
        const initiatedOnDate = item.InitiatedOn ? new Date(item.InitiatedOn) : null;
        const pendingFromDate = item.PendingFromDate ? new Date(item.PendingFromDate) : null;
        const today = new Date();

        const ageingFromCreated = initiatedOnDate 
          ? Math.floor((today.getTime() - initiatedOnDate.getTime()) / (1000 * 60 * 60 * 24)) 
          : 0;

        const ageingFromLastApprover = pendingFromDate 
          ? Math.floor((today.getTime() - pendingFromDate.getTime()) / (1000 * 60 * 60 * 24)) 
          : 0;

        return {
          pendingAt: item.CurrentApprovalName || "Unknown",
          approvalNotes: item.Title || "",
          status: item.Status || "",
          function: item.Function || "",
          amount: item.Amount || "0.00",
          ageingFromCreated: ageingFromCreated >= 0 ? ageingFromCreated : 0,
          ageingFromLastApprover: ageingFromLastApprover >= 0 ? ageingFromLastApprover : 0,
          initiatedOnRaw: item.InitiatedOn,
          functionRaw: item.Function
        };
      });

      // Filter by date range (InitiatedOn) in memory
      if (fromDate || toDate) {
        const from = fromDate ? new Date(fromDate) : null;
        if (from) from.setHours(0, 0, 0, 0);

        const to = toDate ? new Date(toDate) : null;
        if (to) to.setHours(23, 59, 59, 999);

        records = records.filter(item => {
          if (!item.initiatedOnRaw) return false;
          const itemDate = new Date(item.initiatedOnRaw);
          if (from && itemDate < from) return false;
          if (to && itemDate > to) return false;
          return true;
        });
      }

      // Filter by Function if selected
      if (functionName) {
        records = records.filter(item => item.functionRaw === functionName);
      }

      return records.map(r => ({
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
}
