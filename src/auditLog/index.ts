import { Helper, Types } from "gd-sprest-bs";
import { List } from "../common/list";
import { Configuration } from "./cfg";

// Audit Log Item Creation
export interface IAuditLogItemCreation {
    LogComment?: string;
    LogData?: string;
    ParentItemId: number;
    ParentListName: string;
    LogUserId: number;
    Title: string;
}

// Audit Log Item
export interface IAuditLogItem extends IAuditLogItemCreation, Types.SP.IListItem {
    LogUser?: {
        EMail: string;
        Id: number;
        LoginName: string;
        Title: string;
    };
}

// Audit Log Properties
export interface IAuditLogProps {
    listName: string;
    onInitError?: (...args) => void;
    onInitialized?: () => void;
    webUrl?: string;
}

/**
 * Audit Log
 */
export class AuditLog {
    // Configuration
    private _cfg: Helper.ISPConfig = null;
    get Configuration(): Helper.ISPConfig { return this._cfg; }
    get AuditListName(): string {
        // Return the list name
        return this.Configuration._configuration.ListCfg[0].ListInformation.Title;
    }

    // List
    private _list: List<IAuditLogItem> = null;
    get List(): List<IAuditLogItem> { return this._list; }

    // Constructor
    constructor(props: IAuditLogProps) {
        // Set the configuration
        this._cfg = Configuration;

        // Create the list component
        this._list = new List<IAuditLogItem>({
            listName: props.listName,
            webUrl: props.webUrl,
            onInitError: props.onInitError,
            onInitialized: props.onInitialized,
            itemQuery: {
                Filter: "Id eq 0"
            }
        });
    }

    // Gets data for an associated item
    getItems(itemId: number, onQuery?: (query: Types.IODataQuery) => Types.IODataQuery): PromiseLike<IAuditLogItem[]> {
        // Set the odata query
        let odata: Types.IODataQuery = {
            Expand: ["LogUser"],
            Filter: `ParentListName eq '${this.List.ListName}' and ParentItemId eq ${itemId}`,
            OrderBy: ["Created desc"],
            Select: ["*", "LogUser/EMail", "LogUser/Id", "LogUser/LoginName", "LogUser/Title"]
        };

        // Call the event
        odata = onQuery ? onQuery(odata) : odata;

        // Get the list data
        return this.List.refresh(odata);
    }

    // Installs the list
    install(): PromiseLike<void> {
        // Installs the list
        return this.Configuration.install();
    }

    // Logs an item to the audit log
    logItem(values: IAuditLogItemCreation): PromiseLike<IAuditLogItem> {
        // Create the item
        return this.List.createItem(values);
    }

    // Removes the audit log list
    uninstall(): PromiseLike<void> {
        // Remove the audit log list
        return this.Configuration.uninstall();
    }
}