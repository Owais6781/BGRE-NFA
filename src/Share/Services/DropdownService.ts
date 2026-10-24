import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { WebPartContext } from "@microsoft/sp-webpart-base";

// ── Interface for dropdown option ──
export interface IDropdownOption {
  label: string;
  value: string;
}

// ── Interface for Project Master list item ──
export interface IProjectMasterItem {
  Id: number;
  ProjectName: string;
}


export interface ISubEntityItem {
  Id: number;
  EntityName: string;
  ProjectId: number;
}

// ── Main Service Class ──
export class DropdownService {
  private sp: ReturnType<typeof spfi>;

  constructor(context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
  }

  // ─────────────────────────────────────────────
  // Project — from "Project Master"
  // ─────────────────────────────────────────────
  public async getProjects(): Promise<IDropdownOption[]> {
    try {
      const items: IProjectMasterItem[] = await this.sp.web.lists
        .getByTitle("ProjectMaster")
        .items
        .select("Id", "ProjectName")
        .orderBy("ProjectName", true)
        .top(500)();

      return items.map((item) => ({
        label: item.ProjectName,
        value: item.Id.toString(),    // or item.Id.toString() if you need ID
      }));
    } catch (error) {
      console.error("Error fetching Projects:", error);
      return [];
    }
  }


  // SubEntity — Filtered by Project Selection
  public async getSubEntity(projectId: string): Promise<IDropdownOption[]> {
    try {
      if (!projectId) {
        return [];
      }
      const items: ISubEntityItem[] = await this.sp.web.lists
        .getByTitle("SubEntity")
        .items
        .select("Id", "EntityName", "ProjectID")
        .filter(`ProjectID eq ${projectId}`)
        .orderBy("EntityName", true)
        .top(500)();

      return items.map((item) => ({
        label: item.EntityName,
        value: item.Id.toString(),
      }));
    } catch (error) {
      console.error("Error fetching SubEntity:", error);
      return [];
    }
  }


  //  FunctionMaster — from "FunctionMaster"
  public async getFunctionMaster(): Promise<IDropdownOption[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle("FunctionMaster")
        .items
        .select("Id", "Title")
        .orderBy("Title", true)
        .top(500)();
  

    
      return items.map((item) => ({
        label: item.Title,
        value: String(item.Id),

      }));
    } catch (error) {
      console.error("Error fetching Function Master data:", error);
      return [];
    }
  }

  // ─────────────────────────────────────────────
  // GENERIC: Fetch from any list (reusable)
  // ─────────────────────────────────────────────
  public async getDropdownOptions(
    listName: string,
    labelField: string = "Title",
    valueField: string = "Title",
    orderByField: string = "Title",
    ascending: boolean = true,
    topCount: number = 500
  ): Promise<IDropdownOption[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(listName)
        .items
        .select(labelField, valueField)
        .orderBy(orderByField, ascending)
        .top(topCount)();

      return items.map((item: any) => ({
        label: item[labelField],
        value: item[valueField]?.toString(),
      }));
    } catch (error) {
      console.error(`Error fetching dropdown from "${listName}":`, error);
      return [];
    }
  }
}