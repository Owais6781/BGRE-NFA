import { WebPartContext } from "@microsoft/sp-webpart-base";
import { formatHistoryDate } from "./NFAService";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users";
import { SPHttpClient } from "@microsoft/sp-http";

export class ApprovalService {
  private sp: SPFI;
  private context: WebPartContext;
  private listName: string = "DigitalApprovalNotes";
  private queryMgtListName: string = "QueryMgt";

  constructor(context: WebPartContext, sp: SPFI) {
    this.context = context;
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
      .select("Status", "CommentHistory", "WorkflowHistory", "Title", "Project", "Dept", "ApprovalAmount", "Level", "Author/Title", "RequestedOn", "DOAClause")
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

    let queryUserName = "Unknown User";
    if (action === "Query" && updateBody.PendingAtId) {
      try {
        const u = await this.sp.web.siteUsers.getById(updateBody.PendingAtId)();
        queryUserName = u.Title || "Unknown User";
      } catch (uErr) {
        console.error("ApprovalService: Error getting queried user for WorkflowHistory:", uErr);
      }
    }

    let nextApproverName = "Unknown Approver";
    if (action === "Approved" && updateBody.PendingAtId) {
      try {
        const u = await this.sp.web.siteUsers.getById(updateBody.PendingAtId)();
        nextApproverName = u.Title || "Unknown Approver";
      } catch (uErr) {
        console.error("ApprovalService: Error getting next approver for WorkflowHistory:", uErr);
      }
    }

    const now = new Date();
    const dateString = now.toLocaleString("en-IN");
    const newCommentEntry = `[${dateString}] ${userName} (${action}) [Level ${item.Level || 1}]: ${remarks || "No remarks"}`;

    const dateStr = formatHistoryDate(now);
    let newWorkflowEntry = "";
    if (action === "Query") {
      newWorkflowEntry = `Pending for Query response to ${queryUserName} at level-${item.Level || 1} on ${dateStr}</br>`;
    } else {
      newWorkflowEntry = `NFA Request ${action} by ${userName} at Level ${item.Level || 1}. on ${dateStr}</br>`;
      if (action === "Approved" && updateBody.PendingAtId) {
        newWorkflowEntry += `\nNFA Pending at ${nextApproverName} at level-${updateBody.Level || (item.Level + 1)} on ${dateStr}</br>`;
      }
    }

    const currentCommentHistory = item.CommentHistory || "";
    const currentWorkflowHistory = item.WorkflowHistory || "";

    const isNormalAction = (action === "Approved" || action === "Rejected" || action === "Delegate");
    const isUnderQuery = item.Status === "Query";

    const bodyToUpdate = { ...updateBody };

    if (isNormalAction && !isUnderQuery) {
      bodyToUpdate.CommentHistory = currentCommentHistory 
        ? `${currentCommentHistory}\n${newCommentEntry}`
        : newCommentEntry;
    }

    bodyToUpdate.WorkflowHistory = currentWorkflowHistory
      ? `${currentWorkflowHistory}\n${newWorkflowEntry}`
      : newWorkflowEntry;

    await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId).update(bodyToUpdate);

    if (isUnderQuery) {
      // It's a query response. Update the Response column in QueryMgt list.
      try {
        const queryItems = await this.sp.web.lists.getByTitle(this.queryMgtListName).items
          .filter(`WFID eq '${itemId}' and QueryToId eq ${currentUserId}`)
          .orderBy("ID", false)
          .select("Id")();
        
        if (queryItems.length > 0) {
          // Update the latest query record with the response
          await this.sp.web.lists.getByTitle(this.queryMgtListName).items.getById(queryItems[0].Id).update({
            Response: remarks || "",
            Status: "Responded"
          });
        }
      } catch (qErr) {
        console.error("ApprovalService: Error updating QueryMgt response:", qErr);
      }
    }

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

      // Log 2: Next Pending Task Log (only if newPendingAtId is not null and action is not Rejected/Query)
      if (newPendingAtId && action !== "Rejected" && action !== "Query") {
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

        const isQueryAction = (action === "Query");

        const dashboardBody: Record<string, any> = {
          Title: item.Title || "NFA",
          Function: item.Dept || "",
          Project: item.Project || "",
          Amount: String(item.ApprovalAmount || "0.00"),
          FunctionID: item.Dept || "",
          WFID: String(itemId),
          WFLevel: isQueryAction ? String(item.Level || 1) : String(newLevel !== undefined ? newLevel : (item.Level || 1)),
          Status: newStatus,
          DOA: doaClause,
          CurrentApproverId: isTerminal ? null : (isQueryAction ? currentUserId : newPendingAtId),
          CurrentApprovalName: isTerminal ? "" : (isQueryAction ? actionTakerName : pendingAtName),
          CurrentApprovalID: isTerminal ? "" : (isQueryAction ? String(currentUserId) : (newPendingAtId ? String(newPendingAtId) : "")),
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

        const updatePayload: Record<string, any> = { ...dashboardBody };
        if (action === "Query") {
          updatePayload.QueryPendingToID = String(newPendingAtId || "");
          updatePayload.QueryPendingToName = pendingAtName;
          updatePayload.QueryPendingOn = formatHistoryDate(now);
          updatePayload.QueryById = currentUserId;
        } else if (action === "Delegate") {
          updatePayload.DelegatedToID = String(newPendingAtId || "");
          updatePayload.DelegatedToName = pendingAtName;
        } else {
          updatePayload.QueryPendingToID = "";
          updatePayload.QueryPendingToName = "";
          updatePayload.QueryPendingOn = null;
          updatePayload.QueryById = null;
          updatePayload.DelegatedToID = "";
          updatePayload.DelegatedToName = "";
        }

        try {
          if (existingItems.length > 0) {
            await this.sp.web.lists.getByTitle("DashboardList").items.getById(existingItems[0].Id).update(updatePayload);
          } else {
            await this.sp.web.lists.getByTitle("DashboardList").items.add(updatePayload);
          }
        } catch (err) {
          console.warn("ApprovalService: Dashboard update failed with primary payload, trying fallback (QueryBy as text field)...", err);
          const fallbackPayload: Record<string, any> = { ...dashboardBody };
          if (action === "Query") {
            fallbackPayload.QueryPendingToID = String(newPendingAtId || "");
            fallbackPayload.QueryPendingToName = pendingAtName;
            fallbackPayload.QueryPendingOn = formatHistoryDate(now);
            fallbackPayload.QueryBy = actionTakerName;
          } else if (action === "Delegate") {
            fallbackPayload.DelegatedToID = String(newPendingAtId || "");
            fallbackPayload.DelegatedToName = pendingAtName;
          } else {
            fallbackPayload.QueryPendingToID = "";
            fallbackPayload.QueryPendingToName = "";
            fallbackPayload.QueryPendingOn = null;
            fallbackPayload.QueryBy = "";
            fallbackPayload.DelegatedToID = "";
            fallbackPayload.DelegatedToName = "";
          }

          if (existingItems.length > 0) {
            await this.sp.web.lists.getByTitle("DashboardList").items.getById(existingItems[0].Id).update(fallbackPayload);
          } else {
            await this.sp.web.lists.getByTitle("DashboardList").items.add(fallbackPayload);
          }
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
    queryUserId: number | number[] | string,
    currentUserId: number,
    subject: string
  ): Promise<{ success: boolean; message: string }> {
    try {
      // Parse queryUserId into an array of numbers
      let userIds: number[] = [];
      if (Array.isArray(queryUserId)) {
        userIds = queryUserId.map(id => typeof id === 'string' ? parseInt(id, 10) : id).filter(id => !isNaN(id));
      } else if (typeof queryUserId === 'string') {
        userIds = queryUserId.split(',').map(id => parseInt(id.trim(), 10)).filter(id => !isNaN(id));
      } else if (typeof queryUserId === 'number' && !isNaN(queryUserId)) {
        userIds = [queryUserId];
      }

      if (userIds.length === 0) {
        return { success: false, message: "No valid query users specified." };
      }

      // For standard workflow updates, we set PendingAtId to the first user in the list
      const primaryQueryUserId = userIds[0];

      const updateBody = {
        Status: "Query",
        PendingAtId: primaryQueryUserId
      };
      
      const item = await this.getAndUpdateCommentHistory(itemId, "Query", remarks, currentUserId, updateBody);

      await this.writeWFLog(item, itemId, remarks, "Query", undefined, primaryQueryUserId, currentUserId);

      // Add records to QueryMgt list (one record per user)
      for (const uid of userIds) {
        const queryMgtBody = {
          Title: item.Title || "NFA",
          WFID: String(itemId),
          Project: item.Project || "",
          DOA: item.DOAClause || "",
          Function: item.Dept || "",
          WorkflowType: "NFA",
          Amount: String(item.ApprovalAmount || "0.00"),
          QueryToId: uid,
          QueryById: currentUserId,
          Query: remarks || "",
          Status: "Pending",
          QueriedOn: new Date().toLocaleString("en-IN"),
          WFLevel: String(item.Level || 1),
          FunctionID: item.Dept || "",
          NFAType: "NFA"
        };

        try {
          await this.sp.web.lists.getByTitle(this.queryMgtListName).items.add(queryMgtBody);
        } catch (listErr) {
          console.error(`ApprovalService: Failed to write to ${this.queryMgtListName} for user ${uid}:`, listErr);
        }
      }
      
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

  // Attachment helper to upload files directly to QueryMgt list item
  private async attachFileToQueryMgtItem(
    itemId: number,
    file: File
  ): Promise<void> {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const fileName = encodeURIComponent(file.name);
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${this.queryMgtListName}')/items(${itemId})/AttachmentFiles/add(FileName='${fileName}')`;

    const arrayBuffer = await file.arrayBuffer();

    const response = await this.context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/octet-stream",
        },
        body: arrayBuffer,
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to attach file to QueryMgt item: ${errorText}`);
    }
  }

  public async respondToQuery(
    itemId: number,
    remarks: string,
    currentUserId: number,
    files: File[]
  ): Promise<{ success: boolean; message: string }> {
    try {
      // 1. Get the active query record from QueryMgt list (with robust fallback)
      let queryItems = await this.sp.web.lists.getByTitle(this.queryMgtListName).items
        .filter(`WFID eq '${itemId}' and QueryToId eq ${currentUserId} and Status eq 'Pending'`)
        .orderBy("ID", false)
        .select("Id", "QueryById", "QueryBy/Title")
        .expand("QueryBy")();

      // Fallback 1: Any pending query for this NFA
      if (queryItems.length === 0) {
        queryItems = await this.sp.web.lists.getByTitle(this.queryMgtListName).items
          .filter(`WFID eq '${itemId}' and Status eq 'Pending'`)
          .orderBy("ID", false)
          .select("Id", "QueryById", "QueryBy/Title")
          .expand("QueryBy")();
      }

      // Fallback 2: Any query for this NFA (latest first)
      if (queryItems.length === 0) {
        queryItems = await this.sp.web.lists.getByTitle(this.queryMgtListName).items
          .filter(`WFID eq '${itemId}'`)
          .orderBy("ID", false)
          .select("Id", "QueryById", "QueryBy/Title")
          .expand("QueryBy")();
      }

      if (queryItems.length === 0) {
        return { success: false, message: "No active query found for this item." };
      }

      const activeQuery = queryItems[0];
      const queryById = activeQuery.QueryById;
      const queryByName = activeQuery.QueryBy ? activeQuery.QueryBy.Title : "";

      if (!queryById) {
        return { success: false, message: "Could not find who queried this NFA." };
      }

      // 2. Fetch the NFA item from DigitalApprovalNotes
      const item = await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId)
        .select("Status", "WorkflowHistory", "Title", "Project", "Dept", "ApprovalAmount", "Level", "Author/Title", "RequestedOn", "DOAClause")
        .expand("Author")();

      // 3. Get current user details
      let userName = "Unknown User";
      if (currentUserId) {
        try {
          const user = await this.sp.web.siteUsers.getById(currentUserId)();
          userName = user.Title || "Unknown User";
        } catch (uErr) {
          console.error("ApprovalService: Error getting user details:", uErr);
        }
      }

      const now = new Date();
      const dateStr = formatHistoryDate(now);

      // 4. Update the active query in QueryMgt list
      await this.sp.web.lists.getByTitle(this.queryMgtListName).items.getById(activeQuery.Id).update({
        Response: remarks || "",
        Status: "Responded",
        RespondedOn: now.toLocaleString("en-IN")
      });

      // 5. Upload files directly to this QueryMgt list item (no popup)
      if (files && files.length > 0) {
        for (const file of files) {
          try {
            await this.attachFileToQueryMgtItem(activeQuery.Id, file);
          } catch (fileErr) {
            console.error(`ApprovalService: Error attaching file ${file.name} to QueryMgt:`, fileErr);
          }
        }
      }

      // 6. Update DigitalApprovalNotes item
      const newWorkflowEntry = `Query responded by ${userName} at Level ${item.Level || 1}. on ${dateStr}</br>`;
      const currentWorkflowHistory = item.WorkflowHistory || "";
      const updatedWorkflowHistory = currentWorkflowHistory
        ? `${currentWorkflowHistory}\n${newWorkflowEntry}`
        : newWorkflowEntry;

      await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId).update({
        Status: "Pending",
        PendingAtId: queryById,
        WorkflowHistory: updatedWorkflowHistory
      });

      // 7. Write Log (Responded Log) in WFLogList (only one log)
      const logBody: Record<string, any> = {
        Title: item.Title || "NFA",
        WFID: String(itemId),
        FunctionID: item.Dept || "",
        FunctionName: item.Dept || "",
        Project: item.Project || "",
        Amount: String(item.ApprovalAmount || "0.00"),
        Status: "Responded",
        PendingAt: "",
        WFLevel: String(item.Level || 1),
        Remarks: remarks || ""
      };
      await this.sp.web.lists.getByTitle("WFLogList").items.add(logBody);

      // 8. Update DashboardList
      try {
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

        const dashboardBody: Record<string, any> = {
          Title: item.Title || "NFA",
          Function: item.Dept || "",
          Project: item.Project || "",
          Amount: String(item.ApprovalAmount || "0.00"),
          FunctionID: item.Dept || "",
          WFID: String(itemId),
          WFLevel: String(item.Level || 1),
          Status: "Pending",
          DOA: doaClause,
          CurrentApproverId: queryById,
          CurrentApprovalName: queryByName,
          CurrentApprovalID: String(queryById),
          PendingFrom: userName,
          InitiatedBy: initiatorName,
          InitiatedOn: initiatedOn,
          PendingFromDate: now,
          DueDate: getDueDate(),
          QueryPendingToID: "",
          QueryPendingToName: "",
          QueryPendingOn: null,
          QueryById: null,
          QueryBy: "",
          DelegatedToID: "",
          DelegatedToName: ""
        };

        const existingItems = await this.sp.web.lists.getByTitle("DashboardList").items
          .filter(`WFID eq '${itemId}'`)
          .select("Id")();

        if (existingItems.length > 0) {
          await this.sp.web.lists.getByTitle("DashboardList").items.getById(existingItems[0].Id).update(dashboardBody);
        } else {
          await this.sp.web.lists.getByTitle("DashboardList").items.add(dashboardBody);
        }
      } catch (dashErr) {
        console.warn("ApprovalService: Failed to update DashboardList:", dashErr);
      }

      return { success: true, message: "Query response submitted successfully." };
    } catch (error) {
      console.error("ApprovalService: respondToQuery failed:", error);
      return { success: false, message: "Error submitting query response." };
    }
  }
}
