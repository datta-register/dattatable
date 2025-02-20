import { Components } from "gd-sprest-bs";

// DataTables.net
import * as $ from "jquery";
import "datatables.net";
import "datatables.net-bs5";

/**
 * Data Table
 */
export interface IDataTable {
    addRow: (row: any) => void;
    datatable: any;
    filter: (idx: number, value?: string) => void;
    filterExact: (idx: number, value?: string) => void;
    filterMulti: (idx: number, values?: string[]) => void;
    onRendering?: (dtProps: any) => any;
    onRendered?: (el?: HTMLElement, dt?: any) => void;
    refresh: (rows: any[]) => void;
    search: (value?: string) => void;
}

/**
 * Properties
 */
export interface IDataTableProps {
    className?: string;
    columns: Components.ITableColumn[];
    dtProps?: any;
    el: HTMLElement;
    onRendering?: (dtProps: any) => any;
    onRendered?: (el?: HTMLElement, dt?: any) => void;
    rows?: any[];
}

/**
 * Data Table
 */
export class DataTable implements IDataTable {
    private _datatable = null;
    private _props: IDataTableProps = null;

    // Constructor
    constructor(props: IDataTableProps) {
        // Save the properties
        this._props = props;

        // Set the default properties
        this._props.dtProps = this._props.dtProps || {
            drawCallback: function (settings) {
                let api = new $.fn.dataTable.Api(settings) as any;

                // Styling option for striped rows
                $(api.context[0].nTable).addClass('table-striped');

                // Align the text to be centered
                $(api.context[0].nTableWrapper).find('.dt-info').parent().removeClass('col-md d-md-flex justify-content-between align-items-center');
                $(api.context[0].nTableWrapper).find('.dt-info').addClass('text-center');

                // Remove the label spacing to align with paging
                $(api.context[0].nTableWrapper).find('.dt-length label').addClass('d-none');

                // Push paging to the end
                $(api.context[0].nTableWrapper).find('.dt-paging').addClass('d-flex justify-content-end mx-0 px-0');

                // Add spacing for the footer
                $(api.context[0].nTableWrapper).find('.row:last-child').addClass('mb-1');
            },
            language: {
                lengthMenu: "_MENU_",
                paginate: {
                    first: "First",
                    last: "Last",
                    next: "Next",
                    previous: "Previous"
                }
            },
            layout: {
                top: null,
                topStart: null,
                topEnd: null,
                bottom: "info",
                bottomStart: "pageLength",
                bottomEnd: {
                    paging: {
                        type: "full_numbers"
                    }
                }
            }
        };

        // Render the table
        this.refresh(props.rows);
    }

    // Adds a row to the table
    addRow(row: any) {
        // See if this is an array
        if (row && typeof (row.length) === "number") {
            // Add the row
            this._datatable.row.add(row).draw(false);
        } else {
            let newRow = [];

            // Parse the columns
            for (let i = 0; i < this._props.columns.length; i++) {
                let column = this._props.columns[i];

                // Append the value
                newRow.push(row[column.name] || "");
            }

            // Add the row
            this._datatable.row.add(newRow).draw(false);
        }
    }

    // Applies the datatables.net plugin
    private applyPlugin(table: Components.ITable) {
        // Call the rendering event
        this._props.dtProps = this._props.onRendering ? this._props.onRendering(this._props.dtProps) : this._props.dtProps;

        // Render the datatable
        this._datatable = $(table.el).DataTable(this._props.dtProps);

        // Call the rendered event in a separate thread to ensure the dashboard object is created
        setTimeout(() => {
            this._props.onRendered ? this._props.onRendered(this._props.el, this._datatable) : null;
        }, 50);
    }

    /** Public Interface */

    // Datatables.net object
    get datatable(): any { return this._datatable; }

    // The data table element
    get el(): HTMLElement { return this._props.el; }

    // Filters the status
    filter(idx: number, value: string = "") {
        // Set the filter
        this._datatable.column(idx).search(value.replace(/[-[/\]{}()*+?.,\\^$#\s]/g, '\\$&'), true, false).draw();
    }

    // Filters the status
    filterExact(idx: number, value: string = "") {
        // Set the filter
        this._datatable.column(idx).search("^" + value.replace(/[-[/\]{}()*+?.,\\^$#\s]/g, '\\$&') + "$", true, false).draw();
    }

    // Filters multiple values against the status
    filterMulti(idx: number, values: string[] = []) {
        // Parse the values
        for (let i = 0; i < values.length; i++) {
            // Update the value
            values[i] = values[i].replace(/\|/g, '\\$&');
        }

        // Filter the values
        this.filter(idx, values.join('|'));
    }

    // Method to reload the data
    refresh(rows: any[] = []) {
        // See if the datatable exists
        if (this._datatable != null) {
            // Clear the datatable
            this._datatable.clear();
            this._datatable.destroy();
            this._datatable = null;
        }

        // Clear the datatable element
        while (this._props.el.firstChild) { this._props.el.removeChild(this._props.el.firstChild); }

        // Render the data table
        let table = Components.Table({
            className: this._props.className,
            el: this._props.el,
            rows,
            columns: this._props.columns
        });

        // Apply the plugin
        this.applyPlugin(table);
    }

    // Searches the datatable
    search(value: string = "") {
        // Search the table
        this._datatable.search(value).draw();
    }
}