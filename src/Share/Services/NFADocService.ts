import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IAttachmentEntry } from "../Components/AttachmentPopup";

// ── Saved Document Interface ──
export interface ISavedDocument {
  id: number;
  documentType: string;
  fileName: string;
  fileSize: number;
  uniqId: string;
}

export class NFADocService {
  private context: WebPartContext;
  private listName: string = "NFADoc";

  constructor(context: WebPartContext) {
    this.context = context;
  }

  // ── Generate Unique ID ──
  public generateUniqId(): string {
    const timestamp = Date.now().toString(36);
    const random = Math.random().toString(36).substring(2, 8);
    return `NFA-${timestamp}-${random}`.toUpperCase();
  }

  // ── Save Single Document ──
  private async createListItem(
    documentType: string,
    uniqId: string
  ): Promise<number> {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items`;

    const body = JSON.stringify({
      Title: documentType,
      UniqId: uniqId,
    });

    const response: SPHttpClientResponse = await this.context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "odata-version": "",
        },
        body,
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to create item: ${errorText}`);
    }

    const result = await response.json();
    return result.Id;
  }

  // ── Attach File to List Item ──
  private async attachFileToItem(
    itemId: number,
    file: File
  ): Promise<void> {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const fileName = encodeURIComponent(file.name);
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items(${itemId})/AttachmentFiles/add(FileName='${fileName}')`;

    const arrayBuffer = await file.arrayBuffer();

    const response: SPHttpClientResponse = await this.context.spHttpClient.post(
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
      throw new Error(`Failed to attach file: ${errorText}`);
    }
  }

  // ── Save All Entries ──
  public async saveDocuments(
    entries: IAttachmentEntry[],
    uniqId: string
  ): Promise<ISavedDocument[]> {
    const savedDocs: ISavedDocument[] = [];

    for (const entry of entries) {
      // 1. Create list item
      const itemId = await this.createListItem(entry.documentType, uniqId);

      // 2. Attach file to item
      await this.attachFileToItem(itemId, entry.file);

      // 3. Track saved document
      savedDocs.push({
        id: itemId,
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
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items(${itemId})`;

    const response: SPHttpClientResponse =
      await this.context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE",
          },
        }
      );

    if (!response.ok) {
      throw new Error("Failed to delete document");
    }
  }

  public async getDocumentsByUniqId(uniqId: string): Promise<ISavedDocument[]> {
      try {
         const siteUrl = this.context.pageContext.web.absoluteUrl;
       // Filter by UniqId and Expand AttachmentFiles to get the actual file details
        const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items?$filter=UniqId eq'${uniqId}'&$expand=AttachmentFiles`;
   
        const response: SPHttpClientResponse = await this.context.spHttpClient.get(
          endpoint,
          SPHttpClient.configurations.v1
        );
   
        if (!response.ok) return [];
   
        const data = await response.json();
   
        // Map the list items to the ISavedDocument interface
        return data.value.map((item: any) => ({
          id: item.Id,
          documentType: item.Title, // We stored the type in the Title field
          fileName: item.AttachmentFiles && item.AttachmentFiles.length > 0 ?item.AttachmentFiles[0].FileName : "Unknown",
          fileSize: item.AttachmentFiles && item.AttachmentFiles.length > 0 ? item.AttachmentFiles[0].Length: 0,
         uniqId: item.UniqId
        }));
      } catch (error) {
        console.error("Error fetching documents by UniqId:", error);
        return [];
      }
    }

}