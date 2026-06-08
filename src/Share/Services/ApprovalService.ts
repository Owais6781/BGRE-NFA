import { WebPartContext } from "@microsoft/sp-webpart-base";
import { formatHistoryDate } from "./NFAService";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users";

export class ApprovalService {
  private sp: SPFI;
  private listName: string = "DigitalApprovalNotes";

  constructor(context: WebPartContext, sp: SPFI) {
    this.sp = sp;
  }

  private async getAndUpdateCommentHistory(
    itemId: number,
    action: string,
    remarks: string,
    currentUserId: number,
    updateBody: Record<string, any>
  ): Promise<any> {
    const item = await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId)
      .select("CommentHistory", "WorkflowHistory", "Title", "Project", "Dept", "ApprovalAmount", "Level", "Author/Title", "RequestedOn", "DOAClause")
      .expand("Author")();

    let userName = "Unknown User";
    if (currentUserId) {
      try {
        const user = await this.sp.web.siteUsers.getById(currentUserId)();
        userName = user.Title || "Unknown User";
      } catch (uErr) {
        console.error("ApprovalService: Error getting user for CommentHistory:", uErr);
      }
    }

    const now = new Date();
    const dateString = now.toLocaleString("en-IN");
    const newCommentEntry = `[${dateString}] ${userName} (${action}): ${remarks || "No remarks"}`;

    const dateStr = formatHistoryDate(now);
    const newWorkflowEntry = `NFA Request ${action} by ${userName} at Level ${item.Level || 1}. on ${dateStr}</br>`;

    const currentCommentHistory = item.CommentHistory || "";
    const currentWorkflowHistory = item.WorkflowHistory || "";

    updateBody.CommentHistory = currentCommentHistory 
      ? `${currentCommentHistory}\n${newCommentEntry}`
      : newCommentEntry;

    updateBody.WorkflowHistory = currentWorkflowHistory
      ? `${currentWorkflowHistory}\n${newWorkflowEntry}`
      : newWorkflowEntry;

    await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId).update(updateBody);

    return item;
  }

  private async writeWFLog(
    item: any,
    itemId: number,
    remarks: string,
    action: string,
    newLevel: number | undefined,
    newPendingAtId: number | null,
    currentUserId: number
  ): Promise<void> {
    try {
      let pendingAtName = "";
      if (newPendingAtId) {
        try {
          const user = await this.sp.web.siteUsers.getById(newPendingAtId)();
          pendingAtName = user.Title || "";
        } catch (uErr) {
          console.error("ApprovalService: Error getting user for WFLogList:", uErr);
        }
      }

      // Log 1: Completed Action Log (PPLPendingAt = null)
      const logBody1: Record<string, any> = {
        Title: item.Title || "NFA",
        WFID: String(itemId),
        FunctionID: item.Dept || "",
        FunctionName: item.Dept || "",
        Project: item.Project || "",
        Amount: String(item.ApprovalAmount || "0.00"),
        Status: action,
        PendingAt: "",
        WFLevel: String(item.Level || 1),
        Remarks: remarks || ""
      };

      await this.sp.web.lists.getByTitle("WFLogList").items.add(logBody1);

      // Log 2: Next Pending Task Log (only if newPendingAtId is not null and action is not Rejected)
      if (newPendingAtId && action !== "Rejected") {
        const logBody2: Record<string, any> = {
          Title: item.Title || "NFA",
          WFID: String(itemId),
          FunctionID: item.Dept || "",
          FunctionName: item.Dept || "",
          Project: item.Project || "",
          Amount: String(item.ApprovalAmount || "0.00"),
          Status: "Pending",
          PendingAt: pendingAtName,
          WFLevel: String(newLevel !== undefined ? newLevel : (item.Level || 1)),
          Remarks: ""
        };

        logBody2.PPLPendingAtId = newPendingAtId;

        await this.sp.web.lists.getByTitle("WFLogList").items.add(logBody2);
      }

      // DashboardList Save Integration
      try {
        let actionTakerName = "";
        if (currentUserId) {
          try {
            const user = await this.sp.web.siteUsers.getById(currentUserId)();
            actionTakerName = user.Title || "";
          } catch (uErr) {
            console.error("ApprovalService: Error getting action taker user:", uErr);
          }
        }

        const initiatorName = item.Author ? item.Author.Title : "";
        const initiatedOn = item.RequestedOn || "";
        const doaClause = item.DOAClause || "";

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

        let newStatus = action;
        if (action === "Approved") {
          newStatus = newPendingAtId ? "Pending" : "Approved";
        } else if (action === "Delegate") {
          newStatus = "Pending";
        }

        const isTerminal = (newStatus === "Approved" || newStatus === "Rejected");

        const dashboardBody: Record<string, any> = {
          Title: item.Title || "NFA",
          Function: item.Dept || "",
          Project: item.Project || "",
          Amount: String(item.ApprovalAmount || "0.00"),
          FunctionID: item.Dept || "",
          WFID: String(itemId),
          WFLevel: String(newLevel !== undefined ? newLevel : (item.Level || 1)),
          Status: newStatus,
          DOA: doaClause,
          CurrentApproverId: isTerminal ? null : newPendingAtId,
          CurrentApprovalName: isTerminal ? "" : pendingAtName,
          CurrentApprovalID: (isTerminal || !newPendingAtId) ? "" : String(newPendingAtId),
          PendingFrom: actionTakerName,
          InitiatedBy: initiatorName,
          InitiatedOn: initiatedOn,
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
      } catch (dashErr) {
        console.warn("ApprovalService: Failed to write to DashboardList:", dashErr);
      }

    } catch (logErr) {
      console.warn("ApprovalService: Failed to write to WFLogList:", logErr);
    }
  }

  public async approveNFA(
    itemId: number,
    currentLevel: number,
    approvers: (string | number | null)[],
    remarks: string,
    currentUserId: number,
    subject: string
  ): Promise<{ success: boolean; message: string }> {
    try {
      let nextLevel = currentLevel + 1;
      let nextApproverId: number | null = null;
      
      while (nextLevel <= 12) {
        const app = approvers[nextLevel - 1];
        if (app && String(app).trim() !== "" && String(app).trim() !== "null") {
          nextApproverId = parseInt(String(app));
          break;
        }
        nextLevel++;
      }
      
      const updateBody: Record<string, any> = {};
      
      if (nextApproverId) {
        updateBody.Level = nextLevel;
        updateBody.PendingAtId = nextApproverId;
        updateBody.Status = "Pending";
      } else {
        updateBody.Level = currentLevel;
        updateBody.PendingAtId = null;
        updateBody.Status = "Approved";
      }
      
      const item = await this.getAndUpdateCommentHistory(itemId, "Approved", remarks, currentUserId, updateBody);
      
      await this.writeWFLog(item, itemId, remarks, "Approved", updateBody.Level, updateBody.PendingAtId, currentUserId);
      
      return { success: true, message: nextApproverId ? "NFA approved and forwarded to next level." : "NFA fully approved!" };
    } catch (error) {
      console.error("ApprovalService: approveNFA failed:", error);
      return { success: false, message: "Error approving NFA." };
    }
  }

  public async rejectNFA(
    itemId: number,
    remarks: string,
    currentUserId: number,
    subject: string
  ): Promise<{ success: boolean; message: string }> {
    try {
      const updateBody = {
        Status: "Rejected",
        PendingAtId: null
      };
      
      const item = await this.getAndUpdateCommentHistory(itemId, "Rejected", remarks, currentUserId, updateBody);

      await this.writeWFLog(item, itemId, remarks, "Rejected", undefined, null, currentUserId);
      
      return { success: true, message: "NFA has been rejected." };
    } catch (error) {
      console.error("ApprovalService: rejectNFA failed:", error);
      return { success: false, message: "Error rejecting NFA." };
    }
  }

  public async queryNFA(
    itemId: number,
    remarks: string,
    queryUserId: number,
    currentUserId: number,
    subject: string
  ): Promise<{ success: boolean; message: string }> {
    try {
      const updateBody = {
        Status: "Query",
        PendingAtId: queryUserId
      };
      
      const item = await this.getAndUpdateCommentHistory(itemId, "Query", remarks, currentUserId, updateBody);

      await this.writeWFLog(item, itemId, remarks, "Query", undefined, queryUserId, currentUserId);
      
      return { success: true, message: "Query sent successfully." };
    } catch (error) {
      console.error("ApprovalService: queryNFA failed:", error);
      return { success: false, message: "Error sending query." };
    }
  }

  public async delegateNFA(
    itemId: number,
    remarks: string,
    delegateUserId: number,
    currentLevel: number,
    currentUserId: number,
    subject: string
  ): Promise<{ success: boolean; message: string }> {
    try {
      const updateBody: Record<string, any> = {
        PendingAtId: delegateUserId
      };
      
      if (currentLevel >= 1 && currentLevel <= 12) {
        updateBody[`ApprovalLevel${currentLevel}Id`] = delegateUserId;
      }
      
      const item = await this.getAndUpdateCommentHistory(itemId, "Delegate", remarks, currentUserId, updateBody);

      await this.writeWFLog(item, itemId, remarks, "Delegate", currentLevel, delegateUserId, currentUserId);
      
      return { success: true, message: "NFA delegated successfully." };
    } catch (error) {
      console.error("ApprovalService: delegateNFA failed:", error);
      return { success: false, message: "Error delegating NFA." };
    }
  }

  public async returnNFA(
    itemId: number,
    remarks: string,
    initiatorId: number,
    currentUserId: number,
    subject: string
  ): Promise<{ success: boolean; message: string }> {
    try {
      const updateBody = {
        Status: "Returned",
        PendingAtId: initiatorId,
        Level: 0
      };
      
      const item = await this.getAndUpdateCommentHistory(itemId, "Returned", remarks, currentUserId, updateBody);

      await this.writeWFLog(item, itemId, remarks, "Returned", 0, initiatorId, currentUserId);
      
      return { success: true, message: "NFA returned to initiator." };
    } catch (error) {
      console.error("ApprovalService: returnNFA failed:", error);
      return { success: false, message: "Error returning NFA." };
    }
  }
}
