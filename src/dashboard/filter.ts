import { Components, ThemeManager } from "gd-sprest-bs";
import { filter } from "gd-sprest-bs/build/icons/svgs/filter";
import { CanvasForm } from "../common/canvas";

/**
 * Filter Item
 */
export interface IFilterItem {
    header: string;
    items: Components.ICheckboxGroupItem[];
    multi?: boolean;
    onFilter?: (value: string | string[], item?: Components.ICheckboxGroupItem) => void;
    onSetFilterValue?: (value?: string | string[], item?: Components.ICheckboxGroupItem | Components.ICheckboxGroupItem[]) => string | string[];
}

/**
 * Properties
 */
export interface IFilterProps {
    filters: IFilterItem[];
    onClear?: () => void;
    onFilter?: (value: string | string[], item?: Components.ICheckboxGroupItem) => void;
    onRendered?: (el: HTMLElement) => void;
}

/**
 * Filter Slideout
 */
export class FilterSlideout {
    private _cbs: Array<Components.ICheckboxGroup> = null;
    private _el: HTMLElement = null;
    private _filters: IFilterItem[] = null;
    private _items: Array<Components.IAccordionItem> = null;
    private _onClear: () => void;
    private _onFilter: (value: string | string[], item?: Components.ICheckboxGroupItem) => void;

    constructor(props: IFilterProps) {
        // Save the properties
        this._filters = props.filters || [];
        this._onClear = props.onClear;
        this._onFilter = props.onFilter;

        // Initialize the variables
        this._cbs = [];
        this._items = [];

        // Generate the items
        this.generateFilters();

        // Call the render event
        props.onRendered ? props.onRendered(this._el) : null;
    }

    // Generates the filters
    private generateFilters() {
        // Create the filters element
        this._el = document.createElement("div");

        // Parse the filters
        for (let i = 0; i < this._filters.length; i++) {
            let fltr = this._filters[i];

            // Add the filter
            this._items.push(this.generateItem(fltr));
        }

        // Default the first filter to be displayed
        this._items.length > 0 ? this._items[0].showFl = true : null;

        // Render an accordion
        Components.Accordion({
            el: this._el,
            items: this._items
        });

        // Render a clear button
        Components.Tooltip({
            el: this._el,
            content: "Reset Filters",
            placement: Components.TooltipPlacements.Left,
            btnProps: {
                className: "float-end mt-3 p-1 pe-2",
                iconClassName: "me-1",
                iconType: filter,
                iconSize: 24,
                isSmall: true,
                text: "Reset",
                type: Components.ButtonTypes.OutlinePrimary,
                onClick: () => {
                    // Clear the filters
                    this.clear();
                }
            }
        });

        // Add the dark class if theme is inverted
        if (ThemeManager.IsInverted) {
            this._el.querySelectorAll("div.form-check.form-switch input[type=checkbox].form-check-input").forEach((el: HTMLElement) => {
                el.classList.add("dark");
            });
        }
    }

    // Generates the navigation dropdown items
    private generateItem(filter: IFilterItem) {
        // Create the item
        let item: Components.IAccordionItem = {
            header: filter.header,
            onRenderBody: el => {
                // Render the checkbox group
                this._cbs.push(Components.CheckboxGroup({
                    el,
                    items: filter.items,
                    multi: filter.multi,
                    type: Components.CheckboxGroupTypes.Switch,
                    onChange: (selectedCheckboxes: Components.ICheckboxGroupItem | Components.ICheckboxGroupItem[]) => {
                        // See if this is a single item
                        if (filter.multi) {
                            let values: string[] = [];

                            // Parse the items
                            let items = (selectedCheckboxes || []) as Components.ICheckboxGroupItem[];
                            for (let i = 0; i < items.length; i++) {
                                // Append the value
                                values.push(items[i].label);
                            }

                            // Execute the events
                            values = filter.onSetFilterValue ? filter.onSetFilterValue(values, selectedCheckboxes) as string[] : values;
                            filter.onFilter ? filter.onFilter(values, item) : null;
                            this._onFilter ? this._onFilter(values, item) : null;
                        } else {
                            let item = selectedCheckboxes as Components.ICheckboxGroupItem;
                            let filterValue: string = item ? item.label : "";

                            // Execute the events
                            filterValue = filter.onSetFilterValue ? filter.onSetFilterValue(filterValue, selectedCheckboxes) as string : filterValue;
                            filter.onFilter ? filter.onFilter(filterValue, item) : null;
                            this._onFilter ? this._onFilter(filterValue, item) : null;
                        }
                    }
                }));
            }
        };

        // Return the item
        return item;
    }

    // Clears the filters
    clear() {
        // Parse the filters
        for (let i = 0; i < this._cbs.length; i++) {
            // Clear the filter
            this._cbs[i].setValue("");
        }

        // Execute the event
        this._onClear ? this._onClear() : null;
    }

    // Gets a checkbox group by its name
    getFilter(key: string): Components.ICheckboxGroup {
        // Parse the items
        for (let i = 0; i < this._items.length; i++) {
            let item = this._items[i];

            // See if this is the target
            if (item.header == key) {
                // Return the checkbox
                return this._cbs[i];
            }
        }

        // Not found
        return null;
    }

    // Sets a checkbox group filter
    setFilterValue(key: string, value: string | string[]) {
        // Get the filter
        let filter = this.getFilter(key);
        if (filter) {
            // Set the value
            filter.setValue(value);
        }
    }

    // Hides the filter
    hide() { CanvasForm.hide(); }

    // Shows the filters
    show() {
        // Set the header and body
        CanvasForm.setHeader('<h5 class="m-0">Filters</h5>');
        CanvasForm.setBody(this._el || "<p>Loading the Filters...</p>");

        // Show the filters
        CanvasForm.show();
    }
}
