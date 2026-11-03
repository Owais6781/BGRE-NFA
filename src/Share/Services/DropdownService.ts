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

// ── Standard FunctionMaster items previously implemented ──
export interface IFunctionMasterItem {
  id: number;
  title: string;
  abbr?: string;
}

export const DEFAULT_FUNCTION_MASTER_LIST: IFunctionMasterItem[] = [
  { id: 1, title: 'Admin & Operation', abbr: 'ADM' },
  { id: 2, title: 'Sales & Revenues', abbr: 'SAL' },
  { id: 3, title: 'Engineering', abbr: 'ENG' },
  { id: 4, title: 'Finance', abbr: 'FIN' },
  { id: 5, title: 'HR', abbr: 'HR' },
  { id: 6, title: 'IT', abbr: 'IT' },
  { id: 7, title: 'GYM', abbr: 'GYM' },
  { id: 8, title: 'Project', abbr: 'PRJ' },
  { id: 9, title: 'Procurement', abbr: 'PROC' },
  { id: 10, title: 'Commercial', abbr: 'COMM' },
  { id: 11, title: 'Legal', abbr: 'LEG' },
  { id: 12, title: 'Accounts', abbr: 'ACC' },
  { id: 13, title: 'Contracts', abbr: 'CON' },
  { id: 14, title: 'Billing', abbr: 'BIL' }
];

export const DEFAULT_FUNCTION_MASTER_OPTIONS: IDropdownOption[] = DEFAULT_FUNCTION_MASTER_LIST.map((f) => ({
  label: f.title,
  value: String(f.id)
}));

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
        .select("Id", "ProjectName", "IsActive")
        .filter("IsActive eq 1")
        .orderBy("ProjectName", true)
        .top(500)();

      return items.map((item) => ({
        label: item.ProjectName,
        value: item.Id.toString(),
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

  /**
   * Restores previously implemented FunctionMaster items to SharePoint list if missing.
   */
  public async restoreFunctionMasterItems(): Promise<{ added: string[]; existing: string[]; failed: string[] }> {
    const result = { added: [] as string[], existing: [] as string[], failed: [] as string[] };
    const listNames = ["FunctionMaster", "Function Master"];
    let targetList: any = null;

    for (const name of listNames) {
      try {
        const list = this.sp.web.lists.getByTitle(name);
        await list.select("Id")();
        targetList = list;
        break;
      } catch {
        // try next
      }
    }

    if (!targetList) {
      console.warn("DropdownService: Neither FunctionMaster nor 'Function Master' list was found on the SharePoint site.");
      return result;
    }

    const existingTitles = new Set<string>();
    try {
      const items = await targetList.items.select("Id", "Title").top(500)();
      (items || []).forEach((it: any) => {
        if (it && it.Title) {
          existingTitles.add(String(it.Title).trim().toLowerCase());
        }
      });
    } catch (e) {
      console.warn("DropdownService: Error reading existing FunctionMaster items before restore:", e);
    }

    for (const fn of DEFAULT_FUNCTION_MASTER_LIST) {
      if (existingTitles.has(fn.title.trim().toLowerCase())) {
        result.existing.push(fn.title);
        continue;
      }

      try {
        try {
          await targetList.items.add({
            Title: fn.title,
            IsActive: 1
          });
        } catch {
          await targetList.items.add({
            Title: fn.title,
            IsActive: "1"
          });
        }
        result.added.push(fn.title);
      } catch (addErr) {
        console.error(`DropdownService: Failed to add function "${fn.title}" to FunctionMaster:`, addErr);
        result.failed.push(fn.title);
      }
    }

    return result;
  }

  // ─────────────────────────────────────────────
  // FunctionMaster — from "FunctionMaster" with resilient fallback & auto-restore
  // ─────────────────────────────────────────────
  public async getFunctionMaster(): Promise<IDropdownOption[]> {
    try {
      let items: any[] = [];
      const listNames = ["FunctionMaster", "Function Master"];

      for (const name of listNames) {
        try {
          // 1. Try with IsActive eq 1
          items = await this.sp.web.lists
            .getByTitle(name)
            .items
            .select("Id", "Title", "IsActive")
            .filter("IsActive eq 1")
            .orderBy("Title", true)
            .top(500)();
          if (items && items.length > 0) break;
        } catch {
          try {
            // 2. Try with text filter (IsActive eq '1')
            items = await this.sp.web.lists
              .getByTitle(name)
              .items
              .select("Id", "Title", "IsActive")
              .filter("IsActive eq '1'")
              .orderBy("Title", true)
              .top(500)();
            if (items && items.length > 0) break;
          } catch {
            try {
              // 3. Fallback without filter, filtering in JavaScript
              const all = await this.sp.web.lists
                .getByTitle(name)
                .items
                .select("Id", "Title", "IsActive")
                .orderBy("Title", true)
                .top(500)();
              items = (all || []).filter(
                (it: any) => it.IsActive == 1 || it.IsActive === "1" || it.IsActive === true || it.IsActive == null
              );
              if (items && items.length > 0) break;
            } catch {
              // Try next list name
            }
          }
        }
      }

      // If SharePoint list has no items or was emptied, auto-restore in background and return default options
      if (!items || items.length === 0) {
        console.warn("DropdownService: FunctionMaster list empty or returned 0 items. Restoring default items...");
        this.restoreFunctionMasterItems().catch((err) =>
          console.warn("DropdownService: Auto-restore background notice:", err)
        );
        return DEFAULT_FUNCTION_MASTER_OPTIONS;
      }

      return items.map((item) => ({
        label: item.Title,
        value: String(item.Id),
      }));
    } catch (error) {
      console.error("Error fetching Function Master data, falling back to default options:", error);
      return DEFAULT_FUNCTION_MASTER_OPTIONS;
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
