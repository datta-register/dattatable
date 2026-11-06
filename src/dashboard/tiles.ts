import * as dtnet from "datatables.net-dt";
import { Components } from "gd-sprest-bs";

// DataTables.net
import "datatables.net-bs5";

export interface ITiles {
    add: (items: any[]) => void;
    filter: (idx: number, value?: string) => void;
    filterExact: (idx: number, value?: string) => void;
    filterMulti: (idx: number, values?: string[]) => void;
    filterMultiExact: (idx: number, values?: string[]) => void;
    refresh?: (items: any[]) => void;
    search?: (value: string) => void;
}

interface IColumn {
    description?: string;
    filter?: boolean;
    name: string;
    title?: string;
    visible?: boolean;
}

interface IColumnControl {
    order?: boolean;
    orderAsc?: boolean;
    orderDesc?: boolean;
    orderRemove?: boolean;
    orderClear?: boolean;
    orderAddAsc?: boolean;
    orderAddDesc?: boolean;
    reorder?: boolean;
    rowGroup?: boolean;
    rowGroupAdd?: boolean;
    rowGroupClear?: boolean;
    rowGroupRemove?: boolean;
    search?: boolean;
    searchList?: boolean;
    searchDropdown?: boolean;
    statusOnly?: boolean;
    spacer?: boolean;
    title?: boolean;
    visibleDropdown?: boolean;
}

export interface ITilesProps {
    bodyFields?: string[];
    bodyTemplate?: string;
    className?: string;
    classNameBody?: string;
    classNameCard?: string;
    classNameFooter?: string;
    classNameHeader?: string;
    columnControls?: IColumnControl[];
    columns?: IColumn[];
    colSize?: number;
    el?: HTMLElement;
    items?: any[];
    onBodyRendered?: (el?: HTMLElement, item?: any) => void;
    onCardBodyRendering?: (cardBody?: Components.ICardBody, item?: any) => void;
    onCardRendered?: (el?: HTMLElement, item?: any) => void;
    onCardRendering?: (cardProps?: Components.ICardProps, item?: any) => void;
    onClick?: (el?: HTMLElement, item?: Components.ICardBody) => void;
    onFooterRendered?: (el?: HTMLElement, item?: any) => void;
    onHeaderRendered?: (el?: HTMLElement, item?: any) => void;
    onRendering?: (dtProps: any) => any;
    onRendered?: (el?: HTMLElement) => void;
    onSubTitleRendered?: (el?: HTMLElement, item?: any) => void;
    onTitleRendered?: (el?: HTMLElement, item?: any) => void;
    showFooter?: boolean;
    showHeader?: boolean;
    subTitleFields?: string[];
    subTitleTemplate?: string;
    titleFields?: string[];
    titleTemplate?: string;
}

/**
 * Tiles
 */
export class Tiles implements ITiles {
    private _datatable: any = null;
    private _props: ITilesProps;

    // Constructor
    constructor(props: ITilesProps) {
        // Save the properties
        this._props = props;

        // Render the tile
        this.render(this._props.items);
    }

    // Adds a tile to the table
    add(items: any[]) {
        // Add the item
        this._datatable.rows.add(items);

        // Redraw the table to show the new tile
        this._datatable.draw(false);
    }

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

    // Generate the card properties
    private generateCard(item: any): Components.ICardProps {
        // See if the body fields exist
        let bodyContent = this._props.bodyTemplate || "";
        if (this._props.bodyFields) {
            // Parse the fields
            for (let i = 0; i < this._props.bodyFields.length; i++) {
                let field = this._props.bodyFields[i];
                let value = typeof (item[field]) != "undefined" ? item[field] : "";

                // See if there is a template
                if (this._props.bodyTemplate) {
                    // Replace the values
                    let pattern = new RegExp("({" + field + "})", "g")
                    bodyContent = bodyContent.replace(pattern, value);
                } else {
                    // Append the value
                    bodyContent += value;
                }
            }
        }

        // See if the sub-title fields exist
        let subTitleContent = this._props.subTitleTemplate || "";
        if (this._props.subTitleFields) {
            // Parse the fields
            for (let i = 0; i < this._props.subTitleFields.length; i++) {
                let field = this._props.subTitleFields[i];
                let value = typeof (item[field]) != "undefined" ? item[field] : "";

                // See if there is a template
                if (this._props.subTitleTemplate) {
                    // Replace the values
                    let pattern = new RegExp("({" + field + "})", "g")
                    subTitleContent = subTitleContent.replace(pattern, value);
                } else {
                    // Append the value
                    subTitleContent += value;
                }
            }
        }

        // See if the sub-title fields exist
        let titleContent = this._props.titleTemplate || "";
        if (this._props.titleFields) {
            // Parse the fields
            for (let i = 0; i < this._props.titleFields.length; i++) {
                let field = this._props.titleFields[i];
                let value = typeof (item[field]) != "undefined" ? item[field] : "";

                // See if there is a template
                if (this._props.titleTemplate) {
                    // Replace the values
                    let pattern = new RegExp("({" + field + "})", "g")
                    titleContent = titleContent.replace(pattern, value);
                } else {
                    // Append the value
                    titleContent += value;
                }
            }
        }

        // Set the card props
        let cardProps: Components.ICardProps = {
            onRender: (el, card) => {
                // Add the default class
                el.classList.add("h-100");

                // Set the class name
                if (this._props.classNameCard) {
                    this._props.classNameCard.trim().split(' ').forEach(className => {
                        // Add the class name
                        el.classList.add(className);
                    });
                }

                // Call the event
                this._props.onCardRendered ? this._props.onCardRendered(el, card) : null;
            },
            body: [{
                className: this._props.classNameBody,
                content: bodyContent,
                data: item,
                subTitle: subTitleContent,
                title: titleContent,
                onRender: (el, card) => {
                    let item = card.data;

                    // Call the events
                    this._props.onTitleRendered ? this._props.onTitleRendered(el.querySelector(".card-title"), item) : null;
                    this._props.onSubTitleRendered ? this._props.onSubTitleRendered(el.querySelector(".card-subtitle"), item) : null;
                    this._props.onBodyRendered ? this._props.onBodyRendered(el, item) : null;
                }
            }],
            onClick: (card, ev) => {
                // Call the event
                this._props.onClick ? this._props.onClick(ev.currentTarget as any, card) : null;
            }
        };

        // See if we are showing the header
        let showHeader = typeof (this._props.showHeader) === "boolean" ? this._props.showHeader : true;
        if (showHeader) {
            cardProps.header = {
                className: this._props.classNameHeader,
                onRender: (el) => {
                    this._props.onHeaderRendered ? this._props.onHeaderRendered(el, item) : null;
                }
            };
        }

        // See if we are showing the footer
        let showFooter = typeof (this._props.showFooter) === "boolean" ? this._props.showFooter : true;
        if (showFooter) {
            cardProps.footer = {
                className: this._props.classNameFooter,
                onRender: (el) => {
                    this._props.onFooterRendered ? this._props.onFooterRendered(el, item) : null;
                }
            };
        }

        // Call the events
        this._props.onCardBodyRendering ? this._props.onCardBodyRendering(cardProps.body[0], item) : null;
        this._props.onCardRendering ? this._props.onCardRendering(cardProps, item) : null;

        // Return the card properties
        return cardProps;
    }

    // Returns the column control
    private getColumnControl() {
        let columnControls = [];

        // Parse the column controls
        (this._props.columnControls || []).forEach(controlProps => {
            let colControl = [];

            // Add the properties
            controlProps?.rowGroup ? colControl.push("rowGroup") : null;
            controlProps?.rowGroupAdd ? colControl.push("rowGroupAdd") : null;
            controlProps?.rowGroupClear ? colControl.push("rowGroupClear") : null;
            controlProps?.rowGroupRemove ? colControl.push("rowGroupRemove") : null;
            controlProps?.order ? colControl.push("order") : null;
            controlProps?.reorder ? colControl.push("reorder") : null;
            controlProps?.search ? colControl.push("search") : null;
            controlProps?.searchDropdown ? colControl.push("searchDropdown") : null;
            controlProps?.searchList ? colControl.push(["searchList"]) : null;
            controlProps?.visibleDropdown ? colControl.push("colVisDropdown") : null;

            // Append the control
            colControl.length > 0 ? columnControls.push(colControl.length > 1 ? colControl : colControl[0]) : null;
        });

        // Return the column controls or the default options
        return columnControls.length > 0 ? columnControls : ["order"];
    }

    // Returns the columns
    private getColumns() {
        let colProps = this._props.columns || [];
        let columns: object[] = [];

        // Set the default column
        columns.push({
            data: null,
            render: (data, type, row) => {
                // See if this is the card
                if (type === 'display') {
                    // Return the card element
                    return Components.Card(this.generateCard(data)).el;
                }
                // Else, see if this is filter/sort
                else if (type === 'filter') {
                    let fields = {};

                    // Parse the columns
                    colProps.forEach(col => {
                        // See if this is a filter field
                        if (col.filter) { fields[col.name] = true; }
                    });

                    // Parse the fields
                    (this._props.bodyFields || []).forEach(field => { fields[field] = true; });
                    (this._props.titleFields || []).forEach(field => { fields[field] = true; });
                    (this._props.subTitleFields || []).forEach(field => { fields[field] = true; });

                    // Parse the fields
                    let values = [];
                    Object.keys(fields).forEach(field => {
                        if (row[field]) {
                            // Add the value
                            values.push(row[field]);
                        }
                    });

                    // Return the values
                    return values.join(' ');
                }

                return data;
            }
        });

        // Parse the filter columns
        colProps.forEach(colProp => {
            // Set the column
            let column: object = {
                data: colProp.name || "",
                title: colProp.title || colProp.name || "",
                visible: colProp.visible == false ? false : true
            };

            // Add the column
            columns.push(column);
        });

        // Return the columns
        return columns;
    }

    // Returns the default properties for the datatable
    private getDefaultProperties(items: any[]) {
        // Get the column size
        let colSize = this._props.colSize > 0 ? this._props.colSize : 4;
        if (colSize > 4) { colSize = 4; }

        // Determine the paging based on the column size
        let lengthMenu = [];
        switch (colSize) {
            case 1:
                lengthMenu = [4, 8, 16, 32];
                break;
            case 2:
                lengthMenu = [8, 16, 32, 64];
                break;
            case 3:
                lengthMenu = [12, 24, 48, 96];
                break;
            case 4:
                lengthMenu = [16, 32, 64, 128];
                break;
        }

        // Return the default properties
        return {
            columnControl: this.getColumnControl(),
            columns: this.getColumns(),
            data: items,
            lengthMenu,
            pageLength: lengthMenu[0],
            initComplete: (dt) => {
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
            },
            createdRow: (row, data, rowIdx) => {
                // Add the default classes
                row.classList.add("d-flex")
                row.classList.add("col")
                row.firstChild?.classList.add("h-100");

                // Add the class based on the column size
                switch (colSize) {
                    case 1:
                        row.classList.add("col-12")
                        break;
                    case 2:
                        row.classList.add("col-6")
                        break;
                    case 3:
                        row.classList.add("col-4")
                        break;
                    case 4:
                        row.classList.add("col-3")
                        break;
                }
            }
        };
    }

    // Method to reload the data
    refresh(items: any[] = []) {
        // See if the datatable exists
        if (this._datatable != null) {
            // Clear the datatable
            this._datatable.clear();
            this._datatable.destroy();
            this._datatable = null;
        }

        // Clear the datatable element
        while (this._props.el.firstChild) { this._props.el.removeChild(this._props.el.firstChild); }

        // Render the accordion
        this.render(items);
    }

    // Renders the dashboard
    private render(items: any[]) {
        // Create the datatable element
        let el = document.createElement("div");
        el.classList.add("bs");
        this._props?.el.appendChild(el);

        // Set the table element
        let elTable = document.createElement("table");
        elTable.innerHTML = `<thead><th></th></thead><tbody class="row"></tbody>`;
        el.appendChild(elTable);

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

    // Searches the tile
    search(value: string = "") {
        // Search the table
        this._datatable.search(value).draw();
    }
}