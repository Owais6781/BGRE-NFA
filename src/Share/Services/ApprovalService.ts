import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export class ApprovalService {
  private sp: SPFI;
  private listName: string = "DigitalApprovalNotes";

  constructor(context: WebPartContext, sp: SPFI) {
    this.sp = sp;
  }

  private async addHistoryLog(
    itemId: number,
    action: string,
    remarks: string,
    currentUserId: number,
    subject: string
  ): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle("NFAHistory").items.add({
        Title: subject || "Action Note",
        Action: action,
        Remarks: remarks,
        NFAId: itemId,
        UserId: currentUserId
      });
    } catch (historyError) {
      console.warn("ApprovalService: Failed to write to NFAHistory list (list may not exist):", historyError);
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
      
      await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId).update(updateBody);
      
      await this.addHistoryLog(itemId, "Approved", remarks, currentUserId, subject);
      
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
      await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId).update({
        Status: "Rejected",
        PendingAtId: null
      });
      
      await this.addHistoryLog(itemId, "Rejected", remarks, currentUserId, subject);
      
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
      await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId).update({
        Status: "Query",
        PendingAtId: queryUserId
      });
      
      await this.addHistoryLog(itemId, "Query", remarks, currentUserId, subject);
      
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
      
      await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId).update(updateBody);
      
      await this.addHistoryLog(itemId, "Delegate", remarks, currentUserId, subject);
      
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
      await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId).update({
        Status: "Returned",
        PendingAtId: initiatorId,
        Level: 0
      });
      
      await this.addHistoryLog(itemId, "Returned", remarks, currentUserId, subject);
      
      return { success: true, message: "NFA returned to initiator." };
    } catch (error) {
      console.error("ApprovalService: returnNFA failed:", error);
      return { success: false, message: "Error returning NFA." };
    }
  }
}
