import { Components } from "gd-sprest-bs";
import { ItemForm } from "../common";
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
    hideFooter?: boolean;
    hideHeader?: boolean;
    hideNavigation?: boolean;
    hideSubNavigation?: boolean;
    navigation?: {
        showFilter?: boolean;
        showSearch?: boolean;
        title?: string | HTMLElement;
        items?: Components.INavbarItem[];
        itemsEnd?: Components.INavbarItem[];
        onFilterRendered?: (el: HTMLElement) => void;
        onRendering?: (props: Components.INavbarProps) => void;
        onRendered?: (el?: HTMLElement) => void;
        onSearchRendered?: (el: HTMLElement) => void;
        onShowFilter?: () => void;
    };
    subNavigation?: {
        showFilter?: boolean;
        showSearch?: boolean;
        title?: string | HTMLElement;
        items?: Components.INavbarItem[];
        itemsEnd?: Components.INavbarItem[];
        onFilterRendered?: (el: HTMLElement) => void;
        onRendering?: (props: Components.INavbarProps) => void;
        onRendered?: (el?: HTMLElement) => void;
        onSearchRendered?: (el: HTMLElement) => void;
        onShowFilter?: () => void;
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
    private _navigation: Navigation = null;
    private _props: IDashboardProps = null;

    // Constructor
    constructor(props: IDashboardProps) {
        // Set the properties
        this._props = props;

        // Set the flag
        typeof (props.useModal) === "boolean" ? ItemForm.UseModal = props.useModal : null;

        // Render the dashboard
        this.render();

        // Let the object get instaniated before calling the event
        setTimeout(() => {
            // Call the render event
            props.onRendered ? props.onRendered(this._props.el) : null;
        }, 10);
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
            <div id="header" class="col mx-75"></div>
        </div>
        <div class="row">
            <div id="sub-navigation" class="col"></div>
        </div>
        <div class="row">
            <div id="datatable" class="col"></div>
        </div>
        <div class="row">
            <div id="footer" class="col"></div>
        </div>`.trim();
        this._props.el.appendChild(elTemplate);

        // Set/Default the visibility flags
        let headerIsVisible = (this._props.hideHeader == null || this._props.hideHeader == true) && this._props.header == null ? false : true;
        let navIsVisible = this._props.hideNavigation ? false : true;
        let subNavIsVisible = (this._props.hideSubNavigation == null || this._props.hideSubNavigation == true) && this._props.subNavigation == null ? false : true;

        // See if we are hiding the navigation
        if (!navIsVisible) {
            // Hide the element
            this._props.el.querySelector("#navigation").classList.add("d-none");
        } else {
            // Render the navigation
            let navProps = this._props.navigation || {};
            this._navigation = new Navigation({
                el: this._props.el.querySelector("#navigation"),
                hideFilter: navProps.showFilter != null ? !navProps.showFilter : false,
                hideSearch: navProps.showSearch != null ? !navProps.showSearch : false,
                items: navProps.items,
                itemsEnd: navProps.itemsEnd,
                title: navProps.title,
                onFilterRendered: navProps.onFilterRendered,
                onRendering: props => {
                    // Set the default classname
                    props.className = props.className || ("bg-sharepoint rounded" + (headerIsVisible || subNavIsVisible ? "-top" : ""));

                    // Set the default type
                    props.type = typeof (props.type) === "number" ? props.type : Components.NavbarTypes.Dark;

                    // Call the rendering event if it exists
                    navProps.onRendering ? navProps.onRendering(props) : null;
                },
                onRendered: navProps.onRendered,
                onSearch: value => {
                    // Search the data table
                    this._dt.search(value);
                },
                onSearchRendered: navProps.onSearchRendered,
                onShowFilter: () => {
                    // Show the filter
                    this._filters.show();

                    // Call the event
                    navProps.onShowFilter ? navProps.onShowFilter() : null;
                },
            });
        }

        // See if we are hiding the header
        if (!headerIsVisible) {
            // Hide the element
            this._props.el.querySelector("#header").classList.add("d-none");
        } else {
            let elHeader = this._props.el.querySelector("#header") as HTMLElement;

            // Render the header
            let header = this._props.header || {};
            new Header({
                el: elHeader,
                onRendering: this._props.header ? this._props.header.onRendering : null,
                onRendered: this._props.header ? this._props.header.onRendered : null,
                title: header.title
            });

            // See if the sub-nav is not visible
            if (!navIsVisible && !subNavIsVisible) {
                // Set the class name
                elHeader.classList.add("rounded");
            } else if (navIsVisible && !subNavIsVisible) {
                // Set the class name
                elHeader.classList.add("rounded-bottom");
            } else if (!navIsVisible && subNavIsVisible) {
                // Set the class name
                elHeader.classList.add("rounded-top");
            }
        }

        // See if we are hiding the sub-navigation
        if (!subNavIsVisible) {
            // Hide the element
            this._props.el.querySelector("#sub-navigation").classList.add("d-none");
        } else {
            // Render the navigation
            let navProps = this._props.subNavigation || {};
            new Navigation({
                el: this._props.el.querySelector("#sub-navigation"),
                hideFilter: navProps.showFilter != null ? !navProps.showFilter : true,
                hideSearch: navProps.showSearch != null ? !navProps.showSearch : true,
                items: navProps.items,
                itemsEnd: navProps.itemsEnd,
                title: navProps.title,
                onFilterRendered: navProps.onFilterRendered,
                onRendering: props => {
                    // Set the default classname
                    props.className = props.className || ("rounded" + (!headerIsVisible && !navIsVisible ? "" : "-bottom"));
                    props.className = "sub-nav " + props.className;

                    // Set the default type
                    props.type = typeof (props.type) === "number" ? props.type : Components.NavbarTypes.Light;

                    // Call the rendering event if it exists
                    navProps.onRendering ? navProps.onRendering(props) : null;
                },
                onRendered: navProps.onRendered,
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

    // Clears the filters
    clearFilter() {
        // Clear the filters
        this._filters.clear();
    }

    // Filter the table
    filter(idx: number, value?: string) {
        // Filter the table
        this._dt.filter(idx, value);
    }

    // Filter the table by multiple values
    filterMulti(idx: number, values?: string[]) {
        // Filter the table
        this._dt.filterMulti(idx, values);
    }

    // Returns a filter checkbox group by its key
    getFilter(key: string) { return this._filters.getFilter(key); }

    // Returns the current search value
    getSearchValue() { return this._navigation ? this._navigation.getSearchValue() : ""; }

    // Hides the filter
    hideFilter() { this._filters.hide(); }

    // Refresh the table
    refresh(rows: any[]) {
        // Refresh the table
        this._dt.refresh(rows);

        // See if a search value exists
        let searchValue = this._navigation ? this._navigation.getSearchValue() : "";
        if (searchValue) {
            // Apply the search value
            this.search(searchValue);
        }
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
