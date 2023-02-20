import { ContextInfo, Helper, Types, Web } from "gd-sprest-bs";
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

/**
 * Audit Log
 */
export class AuditLog<T = IAuditLogItem> {
    // Configuration
    private _cfg: Helper.ISPConfig = null;
    get Configuration(): Helper.ISPConfig { return this._cfg; }
    get AuditListName(): string {
        // Return the list name
        return this.Configuration._configuration.ListCfg[0].ListInformation.Title;
    }

    // List name
    private _listName: string = null;
    get ListName(): string { return this._listName; }

    // Web url
    private _webUrl: string = null;
    get WebUrl(): string { return this._webUrl; }
    set WebUrl(value: string) { this._webUrl = value; }

    // Constructor
    constructor(listName: string) {
        // Save the associated list
        this._listName = listName;

        // Set the configuration
        this._cfg = Configuration;

        // Default the web url
        this.WebUrl = ContextInfo.webServerRelativeUrl;
    }

    // Gets data for an associated item
    getItems(itemId: number, onQuery?: (query: Types.IODataQuery) => Types.IODataQuery) {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Set the odata query
            let odata: Types.IODataQuery = {
                Expand: ["LogUser"],
                Filter: `ParentListName eq '${this.ListName}' and ParentItemId eq ${itemId}`,
                OrderBy: ["Created desc"],
                Select: ["*", "LogUser/EMail", "LogUser/Id", "LogUser/LoginName", "LogUser/Title"]
            };

            // Call the event
            odata = onQuery ? onQuery(odata) : odata;

            // Get the list data
            Web(this.WebUrl).Lists(this.AuditListName).Items().query(odata).execute(items => {
                // Resolve the request
                resolve(items.results as any);
            }, reject);
        });
    }

    // Logs an item to the audit log
    logItem(values: IAuditLogItemCreation): PromiseLike<T> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Create the item
            Web(this.WebUrl).Lists(this.AuditListName).Items().add(values).execute(resolve as any, reject);
        });
    }

    // Installs the list
    install(): PromiseLike<void> {
        // Installs the list
        return this.Configuration.install();
    }

    // Removes the audit log list
    uninstall(): PromiseLike<void> {
        // Remove the audit log list
        return this.Configuration.uninstall();
    }
}