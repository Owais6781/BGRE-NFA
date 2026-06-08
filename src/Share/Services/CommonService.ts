import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users";

export interface ApprovedRejectedRecord {
  id: number;
  requestNo: string;
  function: string;
  project: string;
  createdBy: string;
  createdOn: string;
  createdOnRaw?: string;
  status: string;
  level: number;
  approvedRejectedBy: string;
  approvedRejectedDate: string;
  approvedRejectedDateRaw?: string;
}

export class CommonService {
  private context: WebPartContext;
  private sp: SPFI;
  private logListName: string = "WFLogList";
  private primaryListName: string = "DigitalApprovalNotes";

  constructor(context: WebPartContext, sp: SPFI) {
    this.context = context;
    this.sp = sp;
  }

  public async getApprovedRejectedByMe(): Promise<ApprovedRejectedRecord[]> {
    try {
      const currentUser = await this.sp.web.currentUser();
      const currentUserId = currentUser.Id;

      // Query WFLogList where AuthorId equals current user's ID and Status is 'Approved' or 'Rejected'
      const logItems = await this.sp.web.lists.getByTitle(this.logListName).items
        .filter(`AuthorId eq ${currentUserId} and (Status eq 'Approved' or Status eq 'Rejected')`)
        .select("Title", "WFID", "FunctionName", "Project", "Status", "WFLevel", "Created", "PPLPendingAt/Title")
        .expand("PPLPendingAt")();

      if (!logItems || logItems.length === 0) {
        return [];
      }

      // Collect unique WFIDs
      const wfIdsSet = new Set<number>();
      logItems.forEach((item: any) => {
        if (item.WFID) {
          const idNum = parseInt(item.WFID, 10);
          if (!isNaN(idNum)) {
            wfIdsSet.add(idNum);
          }
        }
      });

      const wfIds = Array.from(wfIdsSet);
      const notesMap = new Map<number, { createdBy: string; createdOn: string }>();

      if (wfIds.length > 0) {
        // Query DigitalApprovalNotes in chunks of 40 to prevent OData filter string from exceeding URL limits
        const chunkSize = 40;
        for (let i = 0; i < wfIds.length; i += chunkSize) {
          const chunk = wfIds.slice(i, i + chunkSize);
          const filterQuery = chunk.map(id => `Id eq ${id}`).join(" or ");

          const noteItems = await this.sp.web.lists.getByTitle(this.primaryListName).items
            .filter(filterQuery)
            .select("Id", "Author/Title", "RequestedOn", "Created")
            .expand("Author")();

          noteItems.forEach((note: any) => {
            notesMap.set(note.Id, {
              createdBy: note.Author ? note.Author.Title : "Unknown",
              createdOn: note.RequestedOn || note.Created || ""
            });
          });
        }
      }

      // Format and combine records
      const records: ApprovedRejectedRecord[] = logItems.map((log: any) => {
        const nfaId = parseInt(log.WFID, 10) || 0;
        const noteMeta = notesMap.get(nfaId);

        let createdOnFormatted = "";
        if (noteMeta && noteMeta.createdOn) {
          try {
            const d = new Date(noteMeta.createdOn);
            if (!isNaN(d.getTime())) {
              createdOnFormatted = d.toLocaleDateString('en-IN');
            } else {
              createdOnFormatted = noteMeta.createdOn;
            }
          } catch {
            createdOnFormatted = noteMeta.createdOn;
          }
        }

        let actionDateFormatted = "";
        if (log.Created) {
          try {
            const d = new Date(log.Created);
            if (!isNaN(d.getTime())) {
              actionDateFormatted = d.toLocaleDateString('en-IN') + " " + d.toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' });
            } else {
              actionDateFormatted = log.Created;
            }
          } catch {
            actionDateFormatted = log.Created;
          }
        }

        return {
          id: nfaId,
          requestNo: log.Title || "",
          function: log.FunctionName || "",
          project: log.Project || "",
          createdBy: noteMeta ? noteMeta.createdBy : "Unknown",
          createdOn: createdOnFormatted,
          createdOnRaw: noteMeta ? noteMeta.createdOn : "",
          status: log.Status || "",
          level: parseInt(log.WFLevel, 10) || 1,
          approvedRejectedBy: log.PPLPendingAt ? log.PPLPendingAt.Title : (currentUser.Title || ""),
          approvedRejectedDate: actionDateFormatted,
          approvedRejectedDateRaw: log.Created || ""
        };
      });

      return records;
    } catch (error) {
      console.error("CommonService: Error in getApprovedRejectedByMe:", error);
      throw error;
    }
  }
}
