import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import { IAttachmentEntry } from "../Components/AttachmentPopup";

// ── Saved Document Interface ──
export interface ISavedDocument {
  id: number;
  documentType: string;
  fileName: string;
  fileSize: number;
  uniqId: string;
  serverRelativeUrl?: string;
}

export class NFADocService {
  private sp: SPFI;
  private listName: string = "NFADoc";

  constructor(context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
  }

  // ── Generate Unique ID ──
  public generateUniqId(): string {
    const timestamp = Date.now().toString(36);
    const random = Math.random().toString(36).substring(2, 8);
    return `NFA-${timestamp}-${random}`.toUpperCase();
  }

  // ── Save All Entries ──
  public async saveDocuments(
    entries: IAttachmentEntry[],
    uniqId: string
  ): Promise<ISavedDocument[]> {
    const savedDocs: ISavedDocument[] = [];

    for (const entry of entries) {
      // 1. Create list item
      const itemResult = await this.sp.web.lists.getByTitle(this.listName).items.add({
        Title: entry.documentType,
        UniqId: uniqId,
      });

      // 2. Attach file to item
      await itemResult.item.attachmentFiles.add(entry.file.name, entry.file);

      // 3. Track saved document
      savedDocs.push({
        id: itemResult.data.Id,
        documentType: entry.documentType,
        fileName: entry.file.name,
        fileSize: entry.file.size,
        uniqId,
      });
    }

    return savedDocs;
  }

  // ── Delete Document ──
  public async deleteDocument(itemId: number): Promise<void> {
    await this.sp.web.lists.getByTitle(this.listName).items.getById(itemId).delete();
  }

  public async getDocumentsByUniqId(uniqId: string): Promise<ISavedDocument[]> {
    try {
      // Filter by UniqId and Expand AttachmentFiles to get the actual file details
      const items = await this.sp.web.lists.getByTitle(this.listName).items
        .filter(`UniqId eq '${uniqId}'`)
        .select("Id", "Title", "UniqId", "AttachmentFiles")
        .expand("AttachmentFiles")();

      // Map the list items to the ISavedDocument interface
      return items.map((item: any) => ({
        id: item.Id,
        documentType: item.Title, // We stored the type in the Title field
        fileName: item.AttachmentFiles && item.AttachmentFiles.length > 0 ? item.AttachmentFiles[0].FileName : "Unknown",
        fileSize: item.AttachmentFiles && item.AttachmentFiles.length > 0 ? item.AttachmentFiles[0].Length : 0,
        uniqId: item.UniqId,
        serverRelativeUrl: item.AttachmentFiles && item.AttachmentFiles.length > 0 ? item.AttachmentFiles[0].ServerRelativeUrl : ""
      }));
    } catch (error) {
      console.error("Error fetching documents by UniqId:", error);
      return [];
    }
    }


}