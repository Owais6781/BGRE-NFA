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
      .select("Status", "CommentHistory", "WorkflowHistory", "WorkflowApprovalHistory", "Title", "Project", "ProjectSubEntity", "Dept", "ApprovalAmount", "WFLevel", "Author/Title", "RequestedOn", "DOAClause", "WorkflowType")
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
    const newCommentEntry = `[${dateString}] ${userName} (${action}) [Level ${item.WFLevel || item.Level || 1}]: ${remarks || "No remarks"}`;

    const isFinalApproval = (action === "Approved" && !updateBody.PendingAtId);
    const dateStr = formatHistoryDate(now);
    let newWorkflowEntry = "";
    if (action === "Query") {
      newWorkflowEntry = `Pending for Query response to ${queryUserName} at level-${item.WFLevel || item.Level || 1} on ${dateStr}</br>`;
    } else {
      const historyAction = isFinalApproval ? "Completed" : action;
      newWorkflowEntry = `NFA Request ${historyAction} by ${userName} at Level ${item.WFLevel || item.Level || 1}. on ${dateStr}</br>`;
      if (action === "Approved" && updateBody.PendingAtId) {
        newWorkflowEntry += `\nNFA Pending at ${nextApproverName} at level-${updateBody.WFLevel || parseInt(updateBody.WFLevel, 10) || (parseInt(item.WFLevel || item.Level || "1", 10) + 1)} on ${dateStr}</br>`;
      }
    }

    const currentCommentHistory = item.CommentHistory || "";
    const currentWorkflowHistory = item.WorkflowHistory || item.WorkflowApprovalHistory || "";

    const isNormalAction = (action === "Approved" || action === "Rejected" || action === "Delegate");
    const isUnderQuery = item.Status === "Query";

    const isTerminal = isFinalApproval || action === "Rejected";

    const bodyToUpdate = { ...updateBody };
    bodyToUpdate.PendingFrom = isTerminal ? "" : dateStr;
    delete bodyToUpdate.PendingAtId;

    if (action === "Approved") {
      bodyToUpdate.PendingAt = (!isFinalApproval && nextApproverName && nextApproverName !== "Unknown Approver") ? nextApproverName : "";
    } else if (action === "Query") {
      bodyToUpdate.PendingAt = queryUserName && queryUserName !== "Unknown User" ? queryUserName : "";
    } else if (action === "Delegate") {
      bodyToUpdate.PendingAt = nextApproverName && nextApproverName !== "Unknown Approver" ? nextApproverName : "";
    } else if (action === "Returned") {
      bodyToUpdate.PendingAt = item.Author ? item.Author.Title : "";
    } else if (action === "Rejected") {
      bodyToUpdate.PendingAt = "";
      bodyToUpdate.PendingFrom = "";
    }

    if (isNormalAction && !isUnderQuery) {
      bodyToUpdate.CommentHistory = currentCommentHistory
        ? `${currentCommentHistory}\n${newCommentEntry}`
        : newCommentEntry;
    }

    const updatedWorkflowHistory = currentWorkflowHistory
      ? `${currentWorkflowHistory}\n${newWorkflowEntry}`
      : newWorkflowEntry;

    bodyToUpdate.WorkflowHistory = updatedWorkflowHistory;
    bodyToUpdate.WorkflowApprovalHistory = updatedWorkflowHistory;

    if ((!userName || userName === "Unknown User") && this.context?.pageContext?.user?.displayName) {
      userName = this.context.pageContext.user.displayName;
    }

    if (bodyToUpdate.Status === "Approved" || bodyToUpdate.Status === "Rejected") {
      bodyToUpdate.Approvedrejectedby = userName;
      bodyToUpdate.Approvedrejectdatetime = dateStr;
    }

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
            Status: "Responded",
            WFLevel: String(item.WFLevel || item.Level || 1),
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
    currentUserId: number,
    functionOptions: { label: string, value: string }[] = []
  ): Promise<void> {


    const funcObj = functionOptions.find(
      f =>
        f.value === String(item.FunctionID || item.originatorDept || "") ||
        f.label === String(item.FunctionName || item.Dept || item.originatorDept || "")
    );

    let functionName = funcObj
      ? funcObj.label
      : String(item.FunctionName || item.Dept || "");

    let functionId = funcObj
      ? funcObj.value
      : String(item.FunctionID || "");

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




      let actionTakerName = "";
      if (currentUserId) {
        try {
          const user = await this.sp.web.siteUsers.getById(currentUserId)();
          actionTakerName = user.Title || "";
        } catch (uErr) {
          console.error("ApprovalService: Error getting action taker user:", uErr);
        }
      }

      const currentLevelStr = String(item.WFLevel || item.Level || 1);
      if (!functionId || !functionName) {
        try {
          const existingLogs = await this.sp.web.lists.getByTitle("WFLogList").items
            .filter(`WFID eq '${itemId}' and FunctionID ne null`)
            .select("FunctionID", "FunctionName")
            .top(1)();
          if (existingLogs.length > 0) {
            if (!functionId && existingLogs[0].FunctionID) functionId = String(existingLogs[0].FunctionID);
            if (!functionName && existingLogs[0].FunctionName) functionName = String(existingLogs[0].FunctionName);
          }
        } catch (logErr) {
          console.warn("ApprovalService: could not fetch Function from WFLogList:", logErr);
        }
      }

      if (!functionId && functionName) {
        try {
          const fItems = await this.sp.web.lists.getByTitle("FunctionMaster").items
            .filter(`Title eq '${functionName}'`)
            .select("Id", "Title")
            .top(1)();
          if (fItems.length > 0) {
            functionId = String(fItems[0].Id);
          }
        } catch (fmErr) {
          console.warn("ApprovalService: could not fetch from FunctionMaster:", fmErr);
        }
      }

      const isApproved = (action === "Approved");
      const isRejected = (action === "Rejected");
      const isQuery = (action === "Query");
    const wfType = String(item.WorkflowType || item.workflow || "MANUAL").trim().toUpperCase();
      let pendingAtValue = "";
      if (isApproved || isRejected) {
        pendingAtValue = actionTakerName;
      } else if (isQuery) {
        pendingAtValue = pendingAtName;
      } else {
        pendingAtValue = pendingAtName || actionTakerName;
      }

      let pplPendingAtValue: number | null = null;
      if (isApproved || isRejected) {
        pplPendingAtValue = currentUserId;
      } else if (isQuery) {
        pplPendingAtValue = newPendingAtId;
      } else {
        pplPendingAtValue = newPendingAtId || currentUserId;
      }
    
      const logBody1: Record<string, any> = {
        Title: item.Title || "NFA",
        WFID: String(itemId),
        FunctionID: String(functionId || ""),
        FunctionName: functionName || "",
        Department: functionName || "",
        Project: item.Project || "",
        Amount: String(item.ApprovalAmount || "0.00"),
        Status: (action === "Approved" && !newPendingAtId) ? "Completed" : action,
        PendingAt: pendingAtValue,
        PPLPendingAtId: pplPendingAtValue,
        WFLevel: String(item.WFLevel || item.Level || 1),
        Remarks: remarks || "",

      };

      await this.sp.web.lists.getByTitle("WFLogList").items.add(logBody1);
   
      // Log 2: Next Pending Row (only if newPendingAtId is not null and action is Approved)
      if (newPendingAtId && action === "Approved") {
        const logBody2: Record<string, any> = {
          Title: item.Title || "NFA",
          WFID: String(itemId),
          FunctionID: String(functionId || ""),
          FunctionName: functionName || "",
          Department: functionName || "",
          DOAClause: item.DOAClause || "",
          Project: item.Project || "",
          Amount: String(item.ApprovalAmount || "0.00"),
          Status: "Pending",
          PendingAt: pendingAtName,
          WFLevel: String(newLevel !== undefined ? newLevel : (parseInt(currentLevelStr, 10) + 1)),
          Remarks: ""
        };
    
        logBody2.PPLPendingAtId = newPendingAtId;

        await this.sp.web.lists.getByTitle("WFLogList").items.add(logBody2);
      }

      // DashboardList Save Integration
      try {
        // actionTakerName already resolved above
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

        const isTerminal = (newStatus === "Approved" || newStatus === "Completed" || newStatus === "Rejected");

        const isQueryAction = (action === "Query");
const dashboardWFLevel = isQueryAction
  ? String(item.WFLevel || item.Level || 1)
  : String(newLevel ?? item.WFLevel ?? item.Level ?? 1);

        const dashboardBody: Record<string, any> = {
          Title: item.Title || "NFA",
          FunctionID: String(functionId || ""),
          Function: functionName || "",
          Department: functionName || "",
          DOAClause: item.DOAClause || "",
          Entity: item.Project || "",
          SubEntity: item.ProjectSubEntity || "",
          WorkflowType: wfType,
          Project: item.Project || "",
          Amount: String(item.ApprovalAmount || "0.00"),

          WFID: String(itemId),
          WFLevel: dashboardWFLevel,

          Status: newStatus,
          DOA: doaClause,
          CurrentApproverId: isTerminal ? null : (isQueryAction ? currentUserId : newPendingAtId),
          CurrentApproverID: isTerminal ? "" : (isQueryAction ? String(currentUserId) : (newPendingAtId ? String(newPendingAtId) : "")),
          CurrentApprovalName: isTerminal ? "" : (isQueryAction ? actionTakerName : pendingAtName),
          CurrentApprovalID: isTerminal ? "" : (isQueryAction ? String(currentUserId) : (newPendingAtId ? String(newPendingAtId) : "")),
          PendingFrom: isTerminal ? "" : formatHistoryDate(now),
          InitiatedBy: initiatorName,
          InitiatedOn: initiatedOn,
          PendingFromDate: isTerminal ? null : now,
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
          updatePayload.QueryPendingOn = "";
          updatePayload.QueryById = 0;
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
            fallbackPayload.QueryPendingOn = "";
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
        try {
          const isQueryAction = action === "Query";
          // actionTakerName already resolved above
          if (currentUserId) {
            try {
              const user = await this.sp.web.siteUsers.getById(currentUserId)();
              actionTakerName = user.Title || "";
            } catch (userErr) {
              console.error("Failed to get action taker name:", userErr);
            }
          }



          const targetWFLevel = isQueryAction ? String(item.WFLevel || item.Level || 1) : String(newLevel !== undefined ? newLevel : (item.WFLevel || item.Level || 1));
          let newStatus = action;
          if (action === "Approved") {
            newStatus = newPendingAtId ? "Pending" : "Approved";
          } else if (action === "Delegate") {
            newStatus = "Pending";
          }
          const isTerminal =
            newStatus === "Approved" || newStatus === "Completed" || newStatus === "Rejected";
          const existingDash = await this.sp.web.lists.getByTitle("DashboardList").items
            .filter(`WFID eq '${itemId}'`)
            .select("Id")();
          if (existingDash.length > 0) {
            await this.sp.web.lists.getByTitle("DashboardList").items.getById(existingDash[0].Id).update({
              WFLevel: targetWFLevel,
              Status: newStatus,
              CurrentApproverId: isTerminal ? null : (isQueryAction ? currentUserId : newPendingAtId),
              CurrentApproverID: isTerminal ? "" : (isQueryAction ? String(currentUserId) : (newPendingAtId ? String(newPendingAtId) : "")),
              CurrentApprovalName: isTerminal ? "" : (isQueryAction ? actionTakerName : pendingAtName),
              CurrentApprovalID: isTerminal ? "" : (isQueryAction ? String(currentUserId) : (newPendingAtId ? String(newPendingAtId) : "")),
              PendingFrom: isTerminal ? "" : formatHistoryDate(new Date())
            });
          }
        } catch (coreErr) {
          console.warn("ApprovalService: Core fallback failed to update WFLevel in DashboardList:", coreErr);
        }
      }

    } catch (logErr) {
      console.warn("ApprovalService: Failed to write to WFLogList:", logErr);
    }
  }

  public async closeAllPendingQueriesForNFA(
    itemId: number,
    resolutionStatus: string = "Resolved",
    resolutionRemarks?: string,
    excludeId?: number
  ): Promise<void> {
    const listNames = ["QueryMgt", "QueryMgtList"];
    for (const listTitle of listNames) {
      try {
        let filterStr = `WFID eq '${itemId}' and Status eq 'Pending'`;
        if (excludeId) {
          filterStr += ` and Id ne ${excludeId}`;
        }
        const pendingItems = await this.sp.web.lists.getByTitle(listTitle).items
          .filter(filterStr)
          .select("Id")();
        for (const qItem of pendingItems) {
          try {
            const updatePayload: Record<string, any> = {
              Status: resolutionStatus
            };
            if (resolutionRemarks) {
              updatePayload.Response = resolutionRemarks;
              updatePayload.RespondedOn = new Date().toLocaleString("en-IN");
            }
            await this.sp.web.lists.getByTitle(listTitle).items.getById(qItem.Id).update(updatePayload);
          } catch (itemErr) {
            console.warn(`ApprovalService: Error updating query item ${qItem.Id} in ${listTitle}:`, itemErr);
          }
        }
      } catch (listErr) {
        // List may not exist under this title
      }
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

      if (!nextApproverId) {
        try {
          const levels = Array.from({ length: 12 }, (_, i) => `ApprovalLevel${i + 1}Id`);
          const dbItem = await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId)
            .select(...levels)();
          let checkLvl = currentLevel + 1;
          while (checkLvl <= 12) {
            const val = dbItem[`ApprovalLevel${checkLvl}Id`];
            if (val && !isNaN(parseInt(String(val), 10))) {
              nextApproverId = parseInt(String(val), 10);
              nextLevel = checkLvl;
              break;
            }
            checkLvl++;
          }
        } catch (fetchErr) {
          console.warn("ApprovalService: Fallback approver lookup error:", fetchErr);
        }
      }

      const updateBody: Record<string, any> = {};

      if (nextApproverId) {
        updateBody.WFLevel = String(nextLevel);
        updateBody.PendingAtId = nextApproverId;
        updateBody.Status = "Pending";
      } else {
        updateBody.WFLevel = String(currentLevel);
        updateBody.PendingAtId = null;
        updateBody.Status = "Approved";
      }

      const item = await this.getAndUpdateCommentHistory(itemId, "Approved", remarks, currentUserId, updateBody);

      await this.writeWFLog(item, itemId, remarks, "Approved", parseInt(updateBody.WFLevel, 10), updateBody.PendingAtId, currentUserId);

      // Clean up / remove all pending queries for this NFA so they are removed from the Inbox of everyone to whom they were assigned
      try {
        await this.closeAllPendingQueriesForNFA(itemId, "Resolved", "NFA approved");
      } catch (qCloseErr) {
        console.warn("ApprovalService: Error closing pending queries on approve:", qCloseErr);
      }

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

      try {
        await this.closeAllPendingQueriesForNFA(itemId, "Resolved", "NFA rejected");
      } catch (qCloseErr) {
        console.warn("ApprovalService: Error closing pending queries on reject:", qCloseErr);
      }

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
    subject: string,
    functionOptions: { label: string, value: string }[] = []
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


      const funcObj = functionOptions.find(
        f =>
          f.value === String(item.FunctionID || item.originatorDept || "") ||
          f.label === String(item.FunctionName || item.Dept || item.originatorDept || "")
      );

      let functionName = funcObj
        ? funcObj.label
        : String(item.FunctionName || item.Dept || "");

      let functionId = funcObj
        ? funcObj.value
        : String(item.FunctionID || "");

      // Ensure FunctionID is the actual numeric ID and FunctionName is the text name
      if (!functionId || !functionName) {
        try {
          const existingLogs = await this.sp.web.lists.getByTitle("WFLogList").items
            .filter(`WFID eq '${itemId}' and FunctionID ne null`)
            .select("FunctionID", "FunctionName")
            .top(1)();
          if (existingLogs.length > 0) {
            if (!functionId && existingLogs[0].FunctionID) functionId = String(existingLogs[0].FunctionID);
            if (!functionName && existingLogs[0].FunctionName) functionName = String(existingLogs[0].FunctionName);
          }
        } catch (logErr) {
          console.warn("ApprovalService: could not fetch Function from WFLogList:", logErr);
        }
      }

      if (!functionId && functionName) {
        try {
          const fItems = await this.sp.web.lists.getByTitle("FunctionMaster").items
            .filter(`Title eq '${functionName}'`)
            .select("Id", "Title")
            .top(1)();
          if (fItems.length > 0) {
            functionId = String(fItems[0].Id);
          }
        } catch (fmErr) {
          console.warn("ApprovalService: could not fetch from FunctionMaster:", fmErr);
        }
      }

      let wfType = String(item.WorkflowType || item.workflow || "").trim();
      if (!wfType) {
        try {
          const nfaItem = await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId)
            .select("WorkflowType")();
          if (nfaItem && nfaItem.WorkflowType) {
            wfType = String(nfaItem.WorkflowType).trim();
          }
        } catch { }
      }
      if (!wfType) {
        wfType = "MANUAL";
      }
      const finalWfType = wfType.toUpperCase();

      // Add records to QueryMgt list (one record per user)
      for (const uid of userIds) {
        const queryMgtBody = {
          Title: item.Title || "NFA",
          WFID: String(itemId),
          Project: item.Project || "",
          SubEntity: item.ProjectSubEntity || "",
          DOA: item.DOAClause || "",
          FunctionID: String(functionId || ""),
          Function: functionName || "",
          Department: functionName || "",
          DOAClause: item.DOAClause || "",
          WorkflowType: finalWfType,
          Amount: String(item.ApprovalAmount || "0.00"),
          QueryToId: uid,
          QueryById: currentUserId,
          Query: remarks || "",
          Status: "Pending",
          QueriedOn: new Date().toLocaleString("en-IN"),
          WFLevel: String(item.WFLevel || item.Level || 1),
          Location: item.Location || "",
          NFAType:"",
        };

        try {
          await this.sp.web.lists.getByTitle(this.queryMgtListName).items.add(queryMgtBody);
        } catch (listErr) {
          console.warn(`ApprovalService: Initial write to ${this.queryMgtListName} failed, trying safe schema:`, listErr);
          try {
            await this.sp.web.lists.getByTitle(this.queryMgtListName).items.add({
              Title: item.Title || "NFA",
              WFID: String(itemId),
              Department: functionName || "",
              Function: functionName || "",
              FunctionID: String(functionId || ""),
              Project: item.Project || "",
              SubEntity: item.ProjectSubEntity || "",
              Amount: String(item.ApprovalAmount || "0.00"),
              QueryToId: uid,
              QueryById: currentUserId,
              Query: remarks || "",
              Status: "Pending",
              QueriedOn: new Date().toLocaleString("en-IN"),
              WFLevel: String(item.WFLevel || item.Level || 1),
            WorkflowType: finalWfType,
            });
          } catch (retryErr) {
            console.error(`ApprovalService: Failed safe write to ${this.queryMgtListName}:`, retryErr);
          }
        }
      }

      // Update DashboardList with Query status
      try {
        let queryToName = "";
        try {
          const targetUser = await this.sp.web.siteUsers.getById(primaryQueryUserId)();
          queryToName = targetUser.Title || "";
        } catch { }

        const now = new Date();
        const dashItems = await this.sp.web.lists.getByTitle("DashboardList").items
          .filter(`WFID eq '${itemId}'`)
          .select("Id")();
        if (dashItems.length > 0) {
          await this.sp.web.lists.getByTitle("DashboardList").items.getById(dashItems[0].Id).update({
            Status: "Query",
            QueryPendingOn: now.toLocaleString("en-IN"),
            QueryById: String(currentUserId),
            QueryPendingToID: String(primaryQueryUserId),
            QueryPendingToName: queryToName,
            PendingFrom: formatHistoryDate(now)
          });
        }
      } catch (dashErr) {
        console.warn("ApprovalService: Failed to update DashboardList on query:", dashErr);
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
        WFLevel: "0"
      };

      const item = await this.getAndUpdateCommentHistory(itemId, "Returned", remarks, currentUserId, updateBody);

      await this.writeWFLog(item, itemId, remarks, "Returned", 0, initiatorId, currentUserId);

      try {
        await this.closeAllPendingQueriesForNFA(itemId, "Resolved", "NFA returned");
      } catch (qCloseErr) {
        console.warn("ApprovalService: Error closing pending queries on return:", qCloseErr);
      }

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
      // 1. Fetch the NFA item from DigitalApprovalNotes first
      const item = await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId)
        .select("Status", "WorkflowHistory", "WorkflowApprovalHistory", "Title", "Project", "ProjectSubEntity", "Dept", "ApprovalAmount", "WFLevel", "Author/Title", "RequestedOn", "DOAClause", "WorkflowType")
        .expand("Author")();

      // 2. Get the active query record from QueryMgt list (with robust fallback)
      let queryItems: any[] = [];
      try {
        queryItems = await this.sp.web.lists.getByTitle(this.queryMgtListName).items
          .filter(`WFID eq '${itemId}' and QueryToId eq ${currentUserId} and Status eq 'Pending'`)
          .orderBy("ID", false)
          .select("Id", "QueryById", "QueryBy/Title")
          .expand("QueryBy")();
      } catch { }

      if (queryItems.length === 0) {
        try {
          queryItems = await this.sp.web.lists.getByTitle(this.queryMgtListName).items
            .filter(`(WFID eq '${itemId}' or Title eq '${item.Title}') and Status eq 'Pending'`)
            .orderBy("ID", false)
            .select("Id", "QueryById", "QueryBy/Title")
            .expand("QueryBy")();
        } catch { }
      }

      if (queryItems.length === 0) {
        try {
          queryItems = await this.sp.web.lists.getByTitle(this.queryMgtListName).items
            .filter(`WFID eq '${itemId}' or Title eq '${item.Title}'`)
            .orderBy("ID", false)
            .select("Id", "QueryById", "QueryBy/Title")
            .expand("QueryBy")();
        } catch { }
      }
      const activeQuery = queryItems.length > 0 ? queryItems[0] : null;
      if (false) { (this.sp.web.lists as any)
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

      }
      let queryById: number | undefined = activeQuery ? activeQuery.QueryById : undefined;
      let queryByName: string = activeQuery ? (activeQuery.QueryBy?.Title || "") : "";

      if (!queryById) {
        try {
          const queryLogs = await this.sp.web.lists.getByTitle("WFLogList").items
            .filter(`WFID eq '${itemId}' and Status eq 'Query'`)
            .orderBy("Created", false)
            .select("Author/Id", "Author/Title")
            .expand("Author")();
          if (queryLogs.length > 0 && queryLogs[0].Author) {
            queryById = queryLogs[0].Author.Id;
            queryByName = queryLogs[0].Author.Title || "";
          }
        } catch { }
        if (!queryById) {
          try {
            const dashItems = await this.sp.web.lists.getByTitle("DashboardList").items
              .filter(`WFID eq '${itemId}'`)
              .select("CurrentApproverId", "CurrentApprovalName")();
            if (dashItems.length > 0) {
              queryById = dashItems[0].CurrentApproverId;
              queryByName = dashItems[0].CurrentApprovalName || "";
            }
          } catch { }
        }

        if (!queryById) {
          return { success: false, message: "Could not find who queried this NFA." };
        }
        if (queryById && !queryByName) {
          try {
            const user = await this.sp.web.siteUsers.getById(queryById)();
            queryByName = user.Title || "";
          } catch { }
        }
      }

      if (false) (this.sp.web.lists as any)

        .select("Status", "WorkflowHistory", "Title", "Project", "ProjectSubEntity", "Dept", "ApprovalAmount", "WFLevel", "Author/Title", "RequestedOn", "DOAClause", "WorkflowType")
        .expand("Author")();

      // 3. Get current user details
      let isApproved = false;
      let isQuery = true;
      let actionTakerName = "";
      let pendingAtName = "";
      let newPendingAtId: number | null = null;
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
      pendingAtName = userName;
      const dateStr = formatHistoryDate(now);

      // 4. Update or create record in QueryMgt
      let queryMgtId = activeQuery ? activeQuery.Id : null;
      if (activeQuery) {
        try {
          await this.sp.web.lists.getByTitle(this.queryMgtListName).items.getById(activeQuery.Id).update({
            Response: remarks || "",
            Status: "Responded",
            RespondedOn: now.toLocaleString("en-IN")
          });
        } catch (uErr) {
          console.warn("ApprovalService: Error updating QueryMgt item:", uErr);
        }
      } else {
        try {
          const addRes = await this.sp.web.lists.getByTitle(this.queryMgtListName).items.add({
            Title: item.Title || "NFA",
            WFID: String(itemId),
            QueryToId: currentUserId,
            QueryById: queryById,
            Response: remarks || "",
            Status: "Responded",
            RespondedOn: now.toLocaleString("en-IN"),
            WFLevel: String(item.WFLevel || item.Level || 1)
          });
          queryMgtId = addRes.data.Id;
        } catch (addErr) {
          console.warn("ApprovalService: Error creating QueryMgt record on respond:", addErr);
        }
      }

      // Also mark any other pending query records for this NFA for other assigned users as Responded
      // so that the query is removed from the inbox of everyone to whom it was assigned
      try {
        const otherRemarks = `Query responded by ${userName}`;
        await this.closeAllPendingQueriesForNFA(itemId, "Responded", otherRemarks, activeQuery ? activeQuery.Id : undefined);
      } catch (qCloseErr) {
        console.warn("ApprovalService: Error resolving other pending queries on respond:", qCloseErr);
      }
      if (false) {
      await this.sp.web.lists.getByTitle(this.queryMgtListName).items.getById(activeQuery.Id).update({
        Response: remarks || "",
        Status: "Responded",
        RespondedOn: now.toLocaleString("en-IN")
      });

      } // 5. Upload files directly to this QueryMgt list item (no popup)
      if (queryMgtId && files && files.length > 0) {
        for (const file of files) {
          try {
            await this.attachFileToQueryMgtItem(queryMgtId, file);
          } catch (fileErr) {
            console.error(`ApprovalService: Error attaching file ${file.name} to QueryMgt:`, fileErr);
          }
        }
      }

      // 6. Update DigitalApprovalNotes item
      const newWorkflowEntry = `Query responded by ${userName} at Level ${item.WFLevel || item.Level || 1}. on ${dateStr}</br>`;
      const currentWorkflowHistory = item.WorkflowHistory || item.WorkflowApprovalHistory || "";
      const updatedWorkflowHistory = currentWorkflowHistory
        ? `${currentWorkflowHistory}\n${newWorkflowEntry}`
        : newWorkflowEntry;

      await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId).update({
        Status: "Pending",
        PendingAt: queryByName,
        PendingFrom: dateStr,
        WorkflowHistory: updatedWorkflowHistory,
        WorkflowApprovalHistory: updatedWorkflowHistory
      });

      // 7. Write Log (Responded Log) in WFLogList (only one log)
      let respFunctionId = "";
      let respFunctionName = item.Dept || "";
      try {
        const existingLogs = await this.sp.web.lists.getByTitle("WFLogList").items
          .filter(`WFID eq '${itemId}' and FunctionID ne null`)
          .select("FunctionID", "FunctionName")
          .top(1)();
        if (existingLogs.length > 0) {
          respFunctionId = String(existingLogs[0].FunctionID || "");
          if (existingLogs[0].FunctionName) respFunctionName = String(existingLogs[0].FunctionName);
        }
      } catch { }

      if (!respFunctionId && respFunctionName) {
        try {
          const fItems = await this.sp.web.lists.getByTitle("FunctionMaster").items
            .filter(`Title eq '${respFunctionName}'`)
            .select("Id", "Title")
            .top(1)();
          if (fItems.length > 0) {
            respFunctionId = String(fItems[0].Id);
          }
        } catch { }
      }
      const pendingAtValue = queryByName || userName;
      const pplPendingAtValue = queryById || currentUserId;

      const logBody: Record<string, any> = {
        Title: item.Title || "NFA",
        WFID: String(itemId),
        FunctionID: respFunctionId || item.Dept || "",
        FunctionName: respFunctionName || item.Dept || "",
        Project: item.Project || "",
        Amount: String(item.ApprovalAmount || "0.00"),
        Status: "Responded",
        PendingAt: pendingAtValue,
        PPLPendingAtId: pplPendingAtValue,
        WFLevel: String(item.WFLevel || item.Level || 1),
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
          Entity: item.Project || "",
          SubEntity: item.ProjectSubEntity || "",

          Project: item.Project || "",
          Amount: String(item.ApprovalAmount || "0.00"),
          FunctionID: respFunctionId || item.Dept || "",
          WFID: String(itemId),
          WFLevel: String(item.WFLevel || item.Level || 1),
          Status: "Pending",
          DOA: doaClause,
          CurrentApproverId: queryById,
          CurrentApproverID: String(queryById),
          CurrentApprovalName: queryByName,
          CurrentApprovalID: String(queryById),
          PendingFrom: dateStr,
          InitiatedBy: initiatorName,
          InitiatedOn: initiatedOn,
          PendingFromDate: now,
          DueDate: getDueDate(),
          QueryPendingToID: "",
          QueryPendingToName: "",
          QueryPendingOn: "",
          QueryById: "0",
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
