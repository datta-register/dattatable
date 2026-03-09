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
    filterMultiExact: (idx: number, values?: string[]) => void;
    onRendering?: (dtProps: any) => any;
    onRendered?: (el?: HTMLElement, dt?: any) => void;
    refresh: (rows: any[]) => void;
    search: (value?: string) => void;
    updateCell: (row: number, column: number, value) => void;
    updateRow: (rowIdx: number, data) => void;
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
    private _addRowItems = [];
    private _datatable = null;
    private _props: IDataTableProps = null;
    private _table: Components.ITable = null;

    // Constructor
    constructor(props: IDataTableProps) {
        // Save the properties
        this._props = props;

        // Set the default properties
        this._props.dtProps = this._props.dtProps || {
            // This will call the render event if the addRow method is used
            createdRow: (row, data, dataIndex) => {
                // Get the row
                let rowId = data[data.length - 1];
                let rowItem = null;
                for (let i = 0; i < this._addRowItems.length; i++) {
                    // See if this is the target row
                    if (this._addRowItems[i].rowId == rowId) {
                        // Remove it from the array
                        rowItem = this._addRowItems.splice(i, 1)[0].row;
                        break;
                    }
                }

                // See if the row item was found, otherwise it wasn't rendered from the addRows function
                if (rowItem) {
                    // Parse the columns
                    for (let i = 0; i < this._props.columns.length; i++) {
                        let column = this._props.columns[i];

                        // See if an event exists
                        if (column.onRenderCell) {
                            // Call it
                            column.onRenderCell(row.children[i], column, rowItem || data, dataIndex);
                        }
                    }
                }
            },
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

        // Render the data table
        this._table = Components.Table({
            className: this._props.className,
            el: this._props.el,
            rows,
            columns: this._props.columns
        });

        // Apply the plugin
        this.applyPlugin(this._table);
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
            // Update the row
            this._table.updateRow(elRow, data);

            // Update the datatable
            this._datatable.draw(false);
        }
    }
}