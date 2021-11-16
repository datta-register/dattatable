import { Components } from "gd-sprest-bs";
import { ItemForm } from "../itemForm";
import { FilterSlideout, IFilterItem } from "./filter";
import { Footer } from "./footer";
import { Header } from "./header";
import { Navigation } from "./navigation";
import { DataTable, IDataTable } from "./table";

// Export the components
export * from "./filter";
export * from "./footer";
export * from "./header";
export * from "./navigation";
export * from "./table";

// Dashboard
export interface IDashboardProps {
    el: HTMLElement;
    footer?: {
        items?: Components.INavbarItem[];
        itemsEnd?: Components.INavbarItem[];
        onRendering?: (props: Components.INavbarProps) => void;
        onRendered?: (el?: HTMLElement) => void;
    };
    filters?: {
        items: IFilterItem[];
        onClear?: () => void;
        onRendered?: (el?: HTMLElement) => void;
    }
    header?: {
        onRendering?: (props: Components.IJumbotronProps) => void;
        onRendered?: (el?: HTMLElement) => void;
        title?: string;
    },
    hideFilter?: boolean;
    hideFooter?: boolean;
    hideHeader?: boolean;
    hideNavigation?: boolean;
    navigation?: {
        title?: string | HTMLElement;
        items?: Components.INavbarItem[];
        itemsEnd?: Components.INavbarItem[];
        onFilterRendered?: (el: HTMLElement) => void;
        onRendering?: (props: Components.INavbarProps) => void;
        onRendered?: (el?: HTMLElement) => void;
    };
    table?: {
        columns: Components.ITableColumn[];
        dtProps?: any;
        onRendered?: (el?: HTMLElement, dt?: any) => void;
        rows?: any[];
    }
    onRendered?: (el?: HTMLElement) => void;
    useModal?: boolean;
}

/**
 * Dashboard
 */
export class Dashboard {
    private _dt: IDataTable = null;
    private _filters: FilterSlideout = null;
    private _props: IDashboardProps = null;

    // Constructor
    constructor(props: IDashboardProps) {
        // Set the properties
        this._props = props;

        // Set the flag
        typeof (props.useModal) === "boolean" ? ItemForm.UseModal = props.useModal : null;

        // Render the dashboard
        this.render();

        // Call the render event
        props.onRendered ? props.onRendered(this._props.el) : null;
    }

    // Renders the component
    private render() {
        // Create the filters
        this._filters = new FilterSlideout({
            filters: this._props.filters ? this._props.filters.items : [],
            onClear: this._props.filters ? this._props.filters.onClear : null,
            onRendered: this._props.filters ? this._props.filters.onRendered : null
        });

        // Render the template
        let elTemplate = document.createElement("div");
        elTemplate.classList.add("dashboard");
        elTemplate.innerHTML = `
        <div class="row">
            <div id="navigation" class="col"></div>
        </div>
        <div class="row">
            <div id="header" class="col mx-75 rounded-bottom"></div>
        </div>
        <div class="row">
            <div id="datatable" class="col"></div>
        </div>
        <div class="row">
            <div id="footer" class="col"></div>
        </div>`.trim();
        this._props.el.appendChild(elTemplate);

        // See if we are hiding the header
        if (this._props.hideHeader) {
            // Hide the element
            this._props.el.querySelector("#header").classList.add("d-none");
        } else {
            // Render the header
            let header = this._props.header || {};
            new Header({
                el: this._props.el.querySelector("#header"),
                onRendering: this._props.header ? this._props.header.onRendering : null,
                onRendered: this._props.header ? this._props.header.onRendered : null,
                title: header.title
            });
            // Update the navigation rounded corners
            this._props.el.querySelector("#navigation nav").classList.remove("rounded");
            this._props.el.querySelector("#navigation nav").classList.add("rounded-top");
        }

        // See if we are hiding the navigation
        if (this._props.hideNavigation) {
            // Hide the element
            this._props.el.querySelector("#navigation").classList.add("d-none");
        } else {
            // Render the navigation
            let navigation = this._props.navigation || {};
            new Navigation({
                el: this._props.el.querySelector("#navigation"),
                hideFilter: this._props.hideFilter,
                items: navigation.items,
                itemsEnd: navigation.itemsEnd,
                title: navigation.title,
                onFilterRendered: navigation.onFilterRendered,
                onRendering: navigation.onRendering,
                onRendered: navigation.onRendered,
                onSearch: value => {
                    // Search the data table
                    this._dt.search(value);
                },
                onShowFilter: () => {
                    // Show the filter
                    this._filters.show();
                },
            });
        }

        // Render the data table
        this._dt = new DataTable({
            columns: this._props.table ? this._props.table.columns : null,
            dtProps: this._props.table ? this._props.table.dtProps : null,
            el: this._props.el.querySelector("#datatable"),
            onRendered: this._props.table ? this._props.table.onRendered : null,
            rows: this._props.table ? this._props.table.rows : null
        });

        // See if we are hiding the footer
        if (this._props.hideFooter) {
            // Hide the element
            this._props.el.querySelector("#footer").classList.add("d-none");
        } else {
            // Render the footer
            let footer = this._props.footer || {};
            new Footer({
                el: this._props.el.querySelector("#footer"),
                items: footer.items,
                itemsEnd: footer.itemsEnd,
                onRendering: footer.onRendering,
                onRendered: footer.onRendered
            });
        }
    }

    /**
     * Public Interface
     */

    // Filter the table
    filter(idx: number, value?: string) {
        // Filter the table
        this._dt.filter(idx, value);
    }

    // Returns a filter checkbox group by its key
    getFilter(key: string) { return this._filters.getFilter(key); }

    // Hides the filter
    hideFilter() { this._filters.hide(); }

    // Refresh the table
    refresh(rows: any[]) {
        // Refresh the table
        this._dt.refresh(rows);
    }

    // Search the table
    search(value?: string) {
        // Search the table
        this._dt.search(value);
    }

    // Sets a filter checkbox group value
    setFilterValue(key: string, value?: string | string[]) { return this._filters.setFilterValue(key, value); }

    // Shows the filter
    showFilter() { this._filters.show(); }
}
