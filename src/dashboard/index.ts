import { Components } from "gd-sprest-bs";
import { ItemForm } from "../common";
import { Accordion, IAccordion } from "./accordion";
import { FilterSlideout, IFilterItem } from "./filter";
import { Footer } from "./footer";
import { Header } from "./header";
import { Navigation } from "./navigation";
import { DataTable, IDataTable } from "./table";
import { Tiles, ITiles } from "./tiles";

// Export the components
export * from "./accordion";
export * from "./filter";
export * from "./footer";
export * from "./header";
export * from "./navigation";
export * from "./table";
export * from "./tiles";

// Dashboard
export interface IDashboardProps {
    accordion?: {
        bodyFields?: string[];
        bodyTemplate?: string;
        filterFields?: string[];
        items: any[];
        onItemBodyRender?: (el?: HTMLElement, item?: any) => void;
        onItemClick?: (el?: HTMLElement, item?: any) => void;
        onItemHeaderRender?: (el?: HTMLElement, item?: any) => void;
        onItemRender?: (el?: HTMLElement, item?: any) => void;
        onPaginationClick?: (pageNumber?: number) => void;
        onPaginationRender?: (el?: HTMLElement) => void;
        paginationLimit?: number;
        showPagination?: boolean;
        titleFields?: string[];
        titleTemplate?: string;
    }
    el: HTMLElement;
    footer?: {
        items?: Components.INavbarItem[];
        itemsEnd?: Components.INavbarItem[];
        onRendering?: (props: Components.INavbarProps) => void;
        onRendered?: (el?: HTMLElement) => void;
    }
    filters?: {
        items: IFilterItem[];
        onClear?: () => void;
        onRendered?: (el?: HTMLElement) => void;
        onShowFilter?: () => void;
    }
    header?: {
        onRendering?: (props: Components.IJumbotronProps) => void;
        onRendered?: (el?: HTMLElement) => void;
        title?: string;
    }
    hideFooter?: boolean;
    hideHeader?: boolean;
    hideNavigation?: boolean;
    hideSubNavigation?: boolean;
    navigation?: {
        searchPlaceholder?: string;
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
    }
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
    }
    table?: {
        columns: Components.ITableColumn[];
        dtProps?: any;
        onRendering?: (dtProps: any) => any;
        onRendered?: (el?: HTMLElement, dt?: any) => void;
        rows?: any[];
    }
    tiles?: {
        bodyFields?: string[];
        bodyTemplate?: string;
        colSize?: number;
        filterFields?: string[];
        items: any[];
        onBodyRendered?: (el?: HTMLElement, item?: any) => void;
        onCardRendered?: (el?: HTMLElement, item?: any) => void;
        onCardRendering?: (item?: Components.ICardProps) => void;
        onColumnRendered?: (el?: HTMLElement, item?: any) => void;
        onFooterRendered?: (el?: HTMLElement, item?: any) => void;
        onHeaderRendered?: (el?: HTMLElement, item?: any) => void;
        onPaginationClick?: (pageNumber?: number) => void;
        onPaginationRendered?: (el?: HTMLElement) => void;
        onSubTitleRendered?: (el?: HTMLElement, item?: any) => void;
        onTitleRendered?: (el?: HTMLElement, item?: any) => void;
        paginationLimit?: number;
        showFooter?: boolean;
        showHeader?: boolean;
        showPagination?: boolean;
        subTitleFields?: string[];
        subTitleTemplate?: string;
        titleFields?: string[];
        titleTemplate?: string;
    }
    onRendered?: (el?: HTMLElement) => void;
    useModal?: boolean;
}

/**
 * Dashboard
 */
export class Dashboard {
    private _accordion: IAccordion = null;
    private _dt: IDataTable = null;
    private _filters: FilterSlideout = null;
    private _navigation: Navigation = null;
    private _props: IDashboardProps = null;
    private _tiles: ITiles = null;

    // Determines if we are rendering an accordion
    private get IsAccordion(): boolean { return this._props.accordion ? true : false; }

    // Determines if we are rendering tiles
    private get IsTiles(): boolean { return this._props.tiles ? true : false; }

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
            // Apply the filters
            this.applyFilters();

            // Call the render event
            props.onRendered ? props.onRendered(this._props.el) : null;
        }, 10);
    }

    // Applies the default active filters
    private applyFilters() {
        // See if filters exist
        if (this._props.filters?.items) {
            // Parse the filters
            for (let i = 0; i < this._props.filters.items.length; i++) {
                let applyFilter = false;
                let filter = this._props.filters.items[i];
                let value: any = filter.multi ? [] : "";

                // Parse the items
                for (let j = 0; j < filter?.items.length; j++) {
                    let item = filter.items[j];

                    // See if this one is active
                    if (item.isSelected) {
                        // Set the flag
                        applyFilter = true;

                        // Set the value
                        filter.multi ? value.push(item.label) : value = item.label;
                    }
                }

                // See if we are applying the filter
                if (applyFilter) {
                    // Apply the filter
                    filter.onFilter(value);
                }
            }
        }
    }

    // Renders the component
    private render() {
        // Create the filters
        this._filters = new FilterSlideout({
            filters: this._props.filters ? this._props.filters.items : [],
            onClear: this._props.filters ? this._props.filters.onClear : null,
            onRendered: this._props.filters ? this._props.filters.onRendered : null,
            onShowFilter: this._props.filters ? this._props.filters.onShowFilter : null,
            onFilter: this.IsAccordion || this.IsTiles ? (value, item) => {
                let values = typeof (value) === "string" ? [value] : value;

                // See if this is an accordion
                if (this.IsAccordion) {
                    // Filter the accordion
                    this._accordion.filter(values, item);
                } else {
                    // Filter the tiles
                    this._tiles.filter(values, item)
                }
            } : null
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
            <div id="${this.IsAccordion ? "accordion" : (this.IsTiles ? "tiles" : "datatable")}" class="col"></div>
        </div>
        <div class="row">
            <div id="footer" class="col"></div>
        </div>`.trim();
        this._props.el.appendChild(elTemplate);

        // Set/Default the visibility flags
        let headerIsVisible = typeof (this._props.hideHeader) === "boolean" ? !this._props.hideHeader : (this._props.header == null ? false : true);
        let navIsVisible = typeof (this._props.hideNavigation) === "boolean" ? !this._props.hideNavigation : true;
        let subNavIsVisible = typeof (this._props.hideSubNavigation) === "boolean" ? !this._props.hideSubNavigation : (this._props.subNavigation == null ? false : true);

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
                    // See if we are rendering an accordion
                    if (this.IsAccordion) {
                        // Search the accordion
                        this._accordion.search(value);
                    }
                    // Else, see if we are rendering tiles
                    else if (this.IsTiles) {
                        // Search the tiles
                        this._tiles.search(value);
                    } else {
                        // Search the data table
                        this._dt.search(value);
                    }
                },
                onSearchRendered: navProps.onSearchRendered,
                onShowFilter: () => {
                    // Show the filter
                    this._filters.show();

                    // Call the event
                    navProps.onShowFilter ? navProps.onShowFilter() : null;
                },
                searchPlaceholder: navProps.searchPlaceholder
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
                    // See if we are rendering an accordion
                    if (this.IsAccordion) {
                        // Search the accordion
                        this._accordion.search(value);
                    }
                    // Else, see if we are rendering tiles
                    else if (this.IsTiles) {
                        // Search the tiles
                        this._tiles.search(value);
                    } else {
                        // Search the data table
                        this._dt.search(value);
                    }
                },
                onShowFilter: () => {
                    // Show the filter
                    this._filters.show();
                },
            });
        }

        // See if we are rendering an accordion
        if (this.IsAccordion) {
            // Render the accordion
            this._accordion = new Accordion({
                ...{ el: this._props.el.querySelector("#accordion") },
                ...this._props.accordion
            });
        }
        // Else, see if we are rendering tiles
        else if (this.IsTiles) {
            // Render the tiles
            this._tiles = new Tiles({
                ...{ el: this._props.el.querySelector("#tiles") },
                ...this._props.tiles
            })
        } else {
            // Render the data table
            this._dt = new DataTable({
                ...{ el: this._props.el.querySelector("#datatable") },
                ...this._props.table
            });
        }

        // See if we are hiding the footer
        if (this._props.hideFooter) {
            // Hide the element
            this._props.el.querySelector("#footer").classList.add("d-none");
        } else {
            // Render the footer
            let footer = this._props.footer || {};
            new Footer({
                className: "bg-sharepoint rounded-bottom",
                el: this._props.el.querySelector("#footer"),
                items: footer.items,
                itemsEnd: footer.itemsEnd,
                onRendering: footer.onRendering,
                onRendered: footer.onRendered
            });
        }
    }

    /**
     * Component References
     */

    get Datatable(): IDataTable { return this._dt; }
    get Filters(): FilterSlideout { return this._filters; }
    get Navigation(): Navigation { return this._navigation; }

    /**
     * Public Interface
     */

    // Clears the filters
    clearFilter() {
        // Clear the filters
        this._filters.clear();
    }

    // Filter the table
    filter(idx: number, value?: string, exactMatchFl?: boolean) {
        // See if we have an accordion
        if (this.IsAccordion) {
            // Filter the accordion
            this._accordion.filter(value);
        }
        // Else, see if we are rendering tiles
        else if (this.IsTiles) {
            // Filter the tiles
            this._tiles.filter(value);
        } else {
            // If no value is specified, then we don't want to filter by exact value
            exactMatchFl && value ? this._dt.filterExact(idx, value) : this._dt.filter(idx, value);
        }
    }

    // Filter the accordion
    filterAccordion(values?: string | string[]) {
        // See if we have an accordion
        if (this.IsAccordion) {
            // Filter the accordion
            this._accordion.filter(typeof (values) === "string" ? [values] : values);
        }
    }

    // Filter the table by multiple values
    filterMulti(idx: number, values?: string[]) {
        // See if we have an accordion
        if (this.IsAccordion) {
            // Filter the accordion
            this._accordion.filter(values);
        }
        // Else, see if we have tiles
        else if (this.IsTiles) {
            // Filter the tiles
            this._tiles.filter(values);
        }
        else {
            // Filter the table
            this._dt.filterMulti(idx, values);
        }
    }

    // Filter the tiles
    filterTiles(values?: string | string[]) {
        // See if we have tiles
        if (this.IsTiles) {
            // Filter the tiles
            this._tiles.filter(typeof (values) === "string" ? [values] : values);
        }
    }

    // Returns a filter checkbox group by its key
    getFilter(key: string) { return this._filters.getFilter(key); }

    // Returns the current search value
    getSearchValue() { return this._navigation ? this._navigation.getSearchValue() : ""; }

    // Hides the filter
    hideFilter() { this._filters.hide(); }

    // Refresh the table
    refresh(rows: any[]) {
        // See if we have an accordion
        if (this.IsAccordion) {
            // Refresh the table
            this._accordion.refresh(rows);
        }
        // Else, see if we have tiles
        else if (this.IsTiles) {
            // Refresh the table
            this._tiles.refresh(rows);
        } else {
            // Refresh the table
            this._dt.refresh(rows);
        }

        // See if a search value exists
        let searchValue = this._navigation ? this._navigation.getSearchValue() : "";
        if (searchValue) {
            // Apply the search value
            this.search(searchValue);
        }
    }

    // Search the table
    search(value?: string) {
        // See if we have an accordion
        if (this.IsAccordion) {
            // Search the accordion
            this._accordion.search(value);
        }
        // Else, see if we have tiles
        else if (this.IsTiles) {
            // Search the tiles
            this._tiles.search(value);
        } else {
            // Search the table
            this._dt.search(value);
        }
    }

    // Sets a filter checkbox group value
    setFilterValue(key: string, value?: string | string[]) { return this._filters.setFilterValue(key, value); }

    // Shows the filter
    showFilter() { this._filters.show(); }
}
