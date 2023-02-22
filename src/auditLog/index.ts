import { Components, Helper, Types } from "gd-sprest-bs";
import * as jQuery from "jquery";
import { Modal } from "../common/modal";
import { List } from "../common/list";
import { LoadingDialog } from "../common/loadingDialog";
import { DataTable, IDataTableProps } from "../dashboard/table";
import { Configuration } from "./cfg";

// Audit Log Item Creation
export interface IAuditLogItemCreation {
    LogComment?: string;
    LogData?: string;
    ParentId: string;
    ParentListName: string;
    LogUserId?: number;
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

// Audit Log View Props
export interface IAuditLogViewProps {
    id: string;
    listName: string;
    onQuery?: (query: Types.IODataQuery) => Types.IODataQuery;
    onTableRendering?: (props: IDataTableProps) => IDataTableProps;
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

        // Update the list name
        let listName = props.listName ? props.listName : this._cfg._configuration.ListCfg[0].ListInformation.Title;
        props.listName ? this._cfg._configuration.ListCfg[0].ListInformation.Title = listName : null;

        // Create the list component
        this._list = new List<IAuditLogItem>({
            listName: listName,
            webUrl: props.webUrl,
            onInitError: props.onInitError,
            onInitialized: props.onInitialized,
            itemQuery: {
                Filter: "Id eq 0"
            }
        });
    }

    // Method to view the audit log information for an item
    private displayModal(items: IAuditLogItem[], viewProps: IAuditLogViewProps) {
        // Clear the modal
        Modal.clear();

        // Set the size
        Modal.setType(Components.ModalTypes.Large);

        // Set the header
        Modal.setHeader("Audit History");

        // Set the properties
        let dtProps: IDataTableProps = {
            el: Modal.BodyElement,
            rows: items,
            dtProps: {
                dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
                columnDefs: [
                    {
                        "targets": [3],
                        "orderable": false,
                        "searchable": false
                    }
                ],
                createdRow: function (row, data, index) {
                    jQuery('td', row).addClass('align-middle');
                },
                drawCallback: function (settings) {
                    let api = new jQuery.fn.dataTable.Api(settings) as any;
                    jQuery(api.context[0].nTable).removeClass('no-footer');
                    jQuery(api.context[0].nTable).addClass('tbl-footer');
                    jQuery(api.context[0].nTable).addClass('table-striped');
                    jQuery(api.context[0].nTableWrapper).find('.dataTables_info').addClass('text-center');
                    jQuery(api.context[0].nTableWrapper).find('.dataTables_length').addClass('pt-2');
                    jQuery(api.context[0].nTableWrapper).find('.dataTables_paginate').addClass('pt-03');
                },
                headerCallback: function (thead, data, start, end, display) {
                    jQuery('th', thead).addClass('align-middle');
                },
                // Set the empty text
                language: {
                    emptyTable: "No logs exist for this item."
                },
                // Order by the 1st column by default; ascending
                order: [[0, "desc"]]
            },
            columns: [
                {
                    name: "Created",
                    title: "Created"
                },
                {
                    name: "Title",
                    title: "Title"
                },
                {
                    name: "",
                    title: "User",
                    onRenderCell: (el, col, item: IAuditLogItem) => {
                        // Render the user information
                        item.LogUser ? el.innerHTML = item.LogUser.Title : null;
                    }
                },
                {
                    name: "LogComment",
                    title: "Comment"
                }
            ]
        };

        // Call the event
        dtProps = viewProps.onTableRendering ? viewProps.onTableRendering(dtProps) : dtProps;

        // Render the table
        new DataTable(dtProps);

        // Show the modal
        Modal.show();
    }

    // Gets data for an associated item, matching the ParentId value
    getItems(id: string, listName: string, onQuery?: (query: Types.IODataQuery) => Types.IODataQuery): PromiseLike<IAuditLogItem[]> {
        // Set the odata query
        let odata: Types.IODataQuery = {
            Expand: ["LogUser"],
            Filter: `ParentListName eq '${listName}' and ParentId eq '${id}'`,
            OrderBy: ["Created desc"],
            Select: ["*", "LogUser/EMail", "LogUser/Id", "LogUser/Title"]
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

    // Method to view the audit log by specified id
    viewLog(props: IAuditLogViewProps) {
        // Display a loading dialog
        LoadingDialog.setHeader("Loading Audit History");
        LoadingDialog.setBody("This will close after the information is loaded...");
        LoadingDialog.show();

        // Load the information
        this.getItems(props.id, props.listName, props.onQuery).then(items => {
            // Display the modal
            this.displayModal(items, props);

            // Hide the loading dialog
            LoadingDialog.hide();
        });
    }
}