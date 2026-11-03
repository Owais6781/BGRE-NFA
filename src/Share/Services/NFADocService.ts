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
  serverRelativeUrl?: string;
}

export class NFADocService {
  private context: WebPartContext;
  private listName: string = "NFADocs";

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
    title: string,
    uniqId: string
  ): Promise<number> {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items`;

    const body = JSON.stringify({
      Title: title,
      UniqNo: uniqId,
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
    uniqId: string,
    requestNo?: string,
    onProgress?: (current: number, total: number, fileName: string) => void
  ): Promise<ISavedDocument[]> {
    const savedDocs: ISavedDocument[] = [];
    const total = entries.length;
    let current = 0;

    for (const entry of entries) {
      if (onProgress) {
        onProgress(current, total, entry.file.name);
      }
      
      // 1. Create list item with documentType in Title
      const docTitle = entry.documentType || "Attachment";
      const itemId = await this.createListItem(docTitle, uniqId);

      // 2. Attach file to item
      await this.attachFileToItem(itemId, entry.file);

      // 3. Track saved document
      savedDocs.push({
        id: itemId,
        documentType: docTitle,
        fileName: entry.file.name,
        fileSize: entry.file.size,
        uniqId,
      });

      current++;
      if (onProgress) {
        onProgress(current, total, entry.file.name);
      }
    }

    return savedDocs;
  }

  // ── Update Title with RequestNo for All Documents of UniqId ──
  public async updateDocumentsRequestNo(uniqId: string, requestNo: string): Promise<void> {
    // Retain document Type in Title column as per specification (do not overwrite with requestNo)
    return;
  }

  // ── Update Document Type (Title column in NFADocs list) ──
  public async updateDocumentType(itemId: number, newType: string): Promise<void> {
    if (!itemId || !newType) return;
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items(${itemId})`;

    const response: SPHttpClientResponse = await this.context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE",
        },
        body: JSON.stringify({ Title: newType }),
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to update document type: ${errorText}`);
    }
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

  public async getDocumentsByUniqId(
    uniqId?: string,
    nfaId?: number,
    requestNo?: string
  ): Promise<ISavedDocument[]> {
    const docMap = new Map<string, ISavedDocument>();
    const siteUrl = this.context.pageContext.web.absoluteUrl;

    // 1. Fetch from NFADocs list by UniqNo
    if (uniqId && uniqId.trim()) {
      try {
        const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items?$select=Id,Title,UniqNo&$filter=UniqNo eq '${encodeURIComponent(uniqId.trim())}'&$expand=AttachmentFiles`;
        const res = await this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
        if (res.ok) {
          const data = await res.json();
          if (data && data.value) {
            for (const item of data.value) {
              if (item.AttachmentFiles && item.AttachmentFiles.length > 0) {
                for (const att of item.AttachmentFiles) {
                  const key = (att.ServerRelativeUrl || att.FileName).toLowerCase();
                  if (!docMap.has(key)) {
                    docMap.set(key, {
                      id: item.Id,
                      documentType: item.Title || "Attachment",
                      fileName: att.FileName || "Unknown",
                      fileSize: att.Length || 0,
                      uniqId: item.UniqNo || uniqId,
                      serverRelativeUrl: att.ServerRelativeUrl || ""
                    });
                  }
                }
              }
            }
          }
        }
      } catch (err) {
        console.warn("NFADocService: Error querying NFADocs by UniqNo:", err);
      }
    }

    // 2. Fetch from NFADocs list by Title = requestNo (fallback for when UniqNo differs)
    if (requestNo && requestNo.trim()) {
      try {
        const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items?$select=Id,Title,UniqNo&$filter=Title eq '${encodeURIComponent(requestNo.trim())}'&$expand=AttachmentFiles`;
        const res = await this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
        if (res.ok) {
          const data = await res.json();
          if (data && data.value) {
            for (const item of data.value) {
              if (item.AttachmentFiles && item.AttachmentFiles.length > 0) {
                for (const att of item.AttachmentFiles) {
                  const key = (att.ServerRelativeUrl || att.FileName).toLowerCase();
                  if (!docMap.has(key)) {
                    docMap.set(key, {
                      id: item.Id,
                      documentType: item.Title || "Attachment",
                      fileName: att.FileName || "Unknown",
                      fileSize: att.Length || 0,
                      uniqId: item.UniqNo || uniqId || "",
                      serverRelativeUrl: att.ServerRelativeUrl || ""
                    });
                  }
                }
              }
            }
          }
        }
      } catch (err) {
        console.warn("NFADocService: Error querying NFADocs by requestNo Title:", err);
      }
    }

    // 3. Fetch from DigitalApprovalNotes item direct attachments (legacy system support)
    if (nfaId && !isNaN(nfaId) && nfaId > 0) {
      try {
        const endpoint = `${siteUrl}/_api/web/lists/getbytitle('DigitalApprovalNotes')/items(${nfaId})/AttachmentFiles`;
        const res = await this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
        if (res.ok) {
          const data = await res.json();
          if (data && data.value) {
            for (const att of data.value) {
              const key = (att.ServerRelativeUrl || att.FileName).toLowerCase();
              if (!docMap.has(key)) {
                docMap.set(key, {
                  id: nfaId,
                  documentType: "Attachment",
                  fileName: att.FileName || "Unknown",
                  fileSize: att.Length || 0,
                  uniqId: uniqId || "",
                  serverRelativeUrl: att.ServerRelativeUrl || ""
                });
              }
            }
          }
        }
      } catch (err) {
        console.warn("NFADocService: Error querying DigitalApprovalNotes direct attachments:", err);
      }
    }

    return Array.from(docMap.values());
  }
}
