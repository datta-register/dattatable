export interface IColumn {
    columnControl?: object | object[];
    description?: string;
    filter?: boolean;
    name: string;
    order?: "asc" | "desc";
    orderIdx?: number;
    render?: (data, type: string, row, meta) => any;
    title?: string;
    visible?: boolean;
}

export interface IColumnControl {
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

// Returns the columns
export function getColumns(props: IColumn[] = []): object[] {
    let columns: object[] = [];

    // Parse the filter columns
    props.forEach(colProp => {
        // Set the column
        let column: object = {
            data: colProp.name || "",
            name: colProp.name || "",
            title: colProp.title || colProp.name || "",
            visible: colProp.visible == false ? false : true
        };

        // Set the properties
        colProp.columnControl ? column["columnControl"] = colProp.columnControl : null;
        colProp.render ? column["render"] = colProp.render : null;

        // Add the column
        columns.push(column);
    });

    // Return the columns
    return columns;
}

// Returns the column control
export function getColumnControl(props: IColumnControl[] = []) {
    let columnControls = [];

    // Parse the column controls
    props.forEach(controlProps => {
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

// Returns the default column order
export function getDefaultOrder(props: IColumn[] = [], incrementIdx: boolean = false): [number, string][] {
    let order: [number, string][] = [];

    // Parse the columns
    let columns: { [key: number]: IColumn } = {};
    for (let i = 0; i < props.length; i++) {
        let colProp = props[i];

        // See if the order is defined
        if (colProp.order) {
            // Add the column
            columns[i] = colProp;
        }
    }

    // Parse the columns
    Object.keys(columns).sort().forEach(idx => {
        // Add the order
        order.push([parseInt(idx) + (incrementIdx ? 1 : 0), columns[idx].order]);
    })

    // Return the order
    return order;
}