import { Components } from "gd-sprest-bs";
import { CanvasForm } from "../common";

/**
 * Filter Item
 */
export interface IFilterItem {
    header: string;
    items: Components.ICheckboxGroupItem[];
    multi?: boolean;
    onFilter?: (value: string | string[]) => void;
}

/**
 * Properties
 */
export interface IFilterProps {
    filters: IFilterItem[];
    onClear?: () => void;
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

    constructor(props: IFilterProps) {
        // Save the properties
        this._filters = props.filters || [];
        this._onClear = props.onClear;

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

        // Render a clear button
        Components.Button({
            el: this._el,
            className: "mb-3",
            text: "Clear Filters",
            type: Components.ButtonTypes.OutlineDanger,
            onClick: () => {
                // Parse the filters
                for (let i = 0; i < this._cbs.length; i++) {
                    // Clear the filter
                    this._cbs[i].setValue("");
                }

                // Execute the event
                this._onClear ? this._onClear() : null;
            }
        });

        // Parse the filters
        for (let i = 0; i < this._filters.length; i++) {
            let filter = this._filters[i];

            // Add the filter
            this._items.push(this.generateItem(filter));
        }

        // Default the first filter to be displayed
        this._items.length > 0 ? this._items[0].showFl = true : null;

        // Render an accordion
        Components.Accordion({
            el: this._el,
            items: this._items
        });
    }

    // Generates the navigation dropdown items
    private generateItem(filter: IFilterItem) {
        // Create the item
        let item: Components.IAccordionItem = {
            header: filter.header,
            onRender: el => {
                // Render the checkbox group
                this._cbs.push(Components.CheckboxGroup({
                    el,
                    items: filter.items,
                    multi: filter.multi,
                    type: Components.CheckboxGroupTypes.Switch,
                    onChange: (value: Components.ICheckboxGroupItem | Components.ICheckboxGroupItem[]) => {
                        // See if this is a single item
                        if (filter.multi) {
                            let values = [];

                            // Parse the items
                            let items = (value || []) as Components.ICheckboxGroupItem[];
                            for (let i = 0; i < items.length; i++) {
                                // Append the value
                                values.push(items[i].label);
                            }

                            // Execute the event
                            filter.onFilter ? filter.onFilter(values) : null;
                        } else {
                            let item = value as Components.ICheckboxGroupItem;

                            // Execute the event
                            filter.onFilter ? filter.onFilter(item ? item.label : "") : null;
                        }
                    }
                }));
            }
        };

        // Return the item
        return item;
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
        CanvasForm.setHeader("Filters");
        CanvasForm.setBody(this._el || "<p>Loading the Filters...</p>");

        // Show the filters
        CanvasForm.show();
    }
}