import * as dtnet from "datatables.net-dt";
import { Components } from "gd-sprest-bs";
import { IColumn, IColumnControl, getColumnControl, getColumns, getDefaultOrder } from "./columnControl";

// DataTables.net
import "datatables.net-bs5";
import "datatables.net-columncontrol-bs5";

/**
 * Data Table
 */
export interface IDataTable {
    addRow: (row: any) => void;
    datatable: any;
    filter: (idx: number, value?: string) => void;
    filterExact: (idx: number, value?: string) => void;
    filterMulti: (idx: number, values?: string[]) => void;
    filterMultiExact: (idx: number, values?: string[]) => void;
    onRendering?: (dtProps: any) => any;
    onRendered?: (el?: HTMLElement, dt?: any) => void;
    refresh: (rows: any[]) => void;
    removeRow: (rowIdx: number) => void;
    search: (value?: string) => void;
    updateCell: (row: number, column: number, value) => void;
    updateRow: (rowIdx: number, data) => void;
}

/**
 * Properties
 */
export interface IDataTableProps {
    className?: string;
    columnControls?: IColumnControl[];
    columns: IColumn[];
    dtProps?: any;
    el?: HTMLElement;
    onRendering?: (dtProps: any) => any;
    onRendered?: (el?: HTMLElement, dt?: any) => void;
    rows?: any[];
}

/**
 * Data Table
 */
export class DataTable implements IDataTable {
    private _addRowItems = [];
    private _datatable = null;
    private _props: IDataTableProps = null;
    private _table: Components.ITable = null;

    // Constructor
    constructor(props: IDataTableProps) {
        // Save the properties
        this._props = props;

        // Render the table
        this.refresh(props.rows);
    }

    // Adds row(s) to the table
    addRow(row: any) {
        // Set the rows array
        (typeof (row.length) === "number" ? row : [row]).forEach(row => {
            let newRow = [];

            // Parse the columns
            for (let i = 0; i < this._props.columns.length; i++) {
                let column = this._props.columns[i];
                let value = typeof (row[column.name]) != "undefined" ? row[column.name] : "";

                // Append the value
                newRow.push(value);
            }

            // Save a reference to the original data for the event
            let uniqueId = Date.now();
            newRow.push(uniqueId);
            this._addRowItems.push({ rowId: uniqueId, row });

            // Add the row
            this._datatable.row.add(newRow);
        });

        // Refresh the table
        this._datatable.draw(false);
    }

    // Returns the default properties
    private getDefaultProperties(items: any[] = []) {
        // Return the default properties
        return {
            data: items,
            columnControl: getColumnControl(this._props.columnControls),
            columns: getColumns(this._props.columns),
            order: getDefaultOrder(this._props.columns),
            initComplete: (dt) => {
                // Add the bootstrap classes
                dt.table.classList.add("table");
                dt.table.classList.add("table-striped");

                // Remove the margin class
                dt.tableWrapper.querySelector(".dt-layout-table").classList.remove("mt-2");
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
    }

    // Method to render the table
    private render(items: any[]) {
        // Create the datatable element
        let el = document.createElement("div");
        el.classList.add("bs");
        this._props?.el.appendChild(el);

        // Set the table element
        let elTable = document.createElement("table");
        elTable.innerHTML = `<thead></thead><tbody></tbody>`;
        el.appendChild(elTable);

        // Add the custom class names
        (this._props.className || "").split(' ').forEach(className => {
            // Add the class name
            className ? elTable.classList.add(className) : null;
        });

        // Get the default properties
        let dtProps = this.getDefaultProperties(items);

        // Call the rendering event
        dtProps = this._props.onRendering ? this._props.onRendering(dtProps) : dtProps;

        // Render the table
        let dtLib: any = dtnet["DataTable"] || dtnet;
        this._datatable = new dtLib(elTable, dtProps);

        // Call the rendered event in a separate thread to ensure the dashboard object is created
        setTimeout(() => {
            this._props.onRendered ? this._props.onRendered(this._props.el) : null;
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

    // Filters multiple values against the status
    filterMultiExact(idx: number, values: string[] = []) {
        // Parse the values
        for (let i = 0; i < values.length; i++) {
            // Update the value
            values[i] = values[i].replace(/\|/g, '\\$&');
        }

        // Filter the values
        this._datatable.column(idx).search("^(" + values.join('|') + ")$", true, false).draw();
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

        // Render the table
        this.render(rows);
    }

    // Removes a row by index
    removeRow(rowIdx: number) {
        // Remove the row by index
        this._datatable.row(rowIdx).draw(false);
    }

    // Searches the datatable
    search(value: string = "") {
        // Search the table
        this._datatable.search(value).draw();
    }

    // Updates a cell in the datatable
    updateCell(row: number, column: number, value) {
        // Update the cell
        let elCell = this._datatable.cell({ row, column }).node();
        if (elCell) {
            // Update the cell
            this._table.updateColumn(elCell, column, value);
        } else {
            // Update the cell
            this._datatable.cell({ row, column }).data(value).draw(false);
        }
    }

    // Updates a row in the datatable
    updateRow(rowIdx: number, data) {
        // Update the row
        let elRow = this._datatable.row(rowIdx).node();
        if (elRow) {
            // Set the hidden columns
            let hiddenColIndexes = [];
            this._datatable.columns().every(idx => {
                // See if the column is hidden
                if (this._datatable.column(idx).visible() === false) {
                    // Add the index
                    hiddenColIndexes.push(idx);
                }
            });

            // Update the row
            this._table.updateRow(elRow, data, hiddenColIndexes);

            // Update the datatable
            this._datatable.draw(false);
        }
    }
}