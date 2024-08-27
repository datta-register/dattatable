import { Components } from "gd-sprest-bs";

export interface IAccordion {
    filter?: (value: string | string[], item?: Components.ICheckboxGroupItem) => void;
    refresh?: (items: any[]) => void;
    search?: (value: string) => void;
}

export interface IAccordionProps {
    bodyFields?: string[];
    bodyTemplate?: string;
    className?: string;
    el: HTMLElement;
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

/**
 * Accordion
 */
export class Accordion implements IAccordion {
    private _accordion: Components.IAccordion = null;
    private _activeFilterValue: string = null;
    private _activeSearchValue: string = null;
    private _pagination: Components.IPagination = null;
    private _props: IAccordionProps;

    // Constructor
    constructor(props: IAccordionProps) {
        // Save the properties
        this._props = props;

        // Render the accordion
        this.refresh(this._props.items);
    }

    // Clears an item
    private clearItem(elItem: HTMLElement) {
        // Hide the item
        elItem.classList.add("d-none");

        // Remove the first/last item classes
        elItem.classList.remove("first-item")
        elItem.classList.remove("last-item")
    }

    // Filters the accordion
    filter(values: string[]) {
        this._activeFilterValue = values ? values.join('|') : "";

        // Parse all accordion items
        let items = this._props.el.querySelectorAll(".accordion-item");
        for (let i = 0; i < items.length; i++) {
            let elItem = items[i] as HTMLElement;

            // Clear the item
            this.clearItem(elItem);
        }

        // Render the items
        this.renderItems();
    }

    // Generates the accordion item
    private generateItem(item: any): Components.IAccordionItem {
        let filterValues = [];

        // See if the filter field is specified
        if (this._props.filterFields) {
            // Parse the fields
            for (let i = 0; i < this._props.filterFields.length; i++) {
                let filterField = this._props.filterFields[i];

                // Get the filter values
                let filters = item[filterField] || [];
                filters = filters["results"] || [filters];

                // Append the values
                filterValues = filterValues.concat(filters);
            }
        }

        // See if the body fields exist
        let bodyContent = this._props.bodyTemplate || "";
        if (this._props.bodyFields) {
            // Parse the fields
            for (let i = 0; i < this._props.bodyFields.length; i++) {
                let field = this._props.bodyFields[i];
                let value = item[field] || "";

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
        let titleContent = this._props.titleTemplate || "";
        if (this._props.titleFields) {
            // Parse the fields
            for (let i = 0; i < this._props.titleFields.length; i++) {
                let field = this._props.titleFields[i];
                let value = item[field] || "";

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

        // Return the item
        return {
            content: bodyContent,
            header: titleContent,
            onClick: this._props.onItemClick,
            onRenderBody: this._props.onItemBodyRender,
            onRenderHeader: this._props.onItemHeaderRender,
            onRender: (el, item) => {
                // See if filters exist
                if (filterValues && filterValues.length > 0) {
                    // Set the data filter value
                    el.setAttribute("data-filter", filterValues.join('|'));
                }

                // Call the event
                this._props.onItemRender ? this._props.onItemRender(el, item) : null
            }
        };
    }

    // Method to reload the data
    refresh(items: any[] = []) {
        // Clear the datatable element
        while (this._props.el.firstChild) { this._props.el.removeChild(this._props.el.firstChild); }

        // Render the accordion
        this.render(items);
    }

    // Renders the dashboard
    private render(items: any[]) {
        // Render the accordion
        this.renderAccordion(items);

        // Render the items
        this.renderItems();
    }

    // Renders the accordion
    private renderAccordion(items: any[]) {
        // Parse the items
        let accordionItems: Array<Components.IAccordionItem> = [];
        for (let i = 0; i < items.length; i++) {
            // Add an accordion item
            accordionItems.push(this.generateItem(items[i]));
        }

        // Render the accordion
        this._accordion = Components.Accordion({
            el: this._props.el,
            className: this._props.className,
            id: "accordion-list",
            items: accordionItems
        });
    }

    // Renders the items
    private renderItems() {
        let paginationLimit = this._props.paginationLimit || 10;

        // Get the pagination element
        let elPagination = this._props.el.querySelector(".accordion-pagination") as HTMLElement;
        if (elPagination) {
            // Clear the element
            while (elPagination.firstChild) { elPagination.removeChild(elPagination.firstChild); }
        } else {
            // Create the element
            elPagination = document.createElement("div");
            elPagination.classList.add("accordion-pagination");
            this._props.el.appendChild(elPagination);
        }

        // Get the elements as an array
        let elItems = Array.from(this._accordion.el.querySelectorAll(".accordion-item"));

        // Parse the items
        for (let i = elItems.length - 1; i >= 0; i--) {
            let elItem = elItems[i] as HTMLElement;
            let elContent = elItem.querySelector(".accordion-body") as HTMLElement;

            // See if this item is expanded
            if (elItem.querySelector(".accordion-collapse.show")) {
                // Collapse the item
                let btn = elItem.querySelector(".accordion-button") as HTMLButtonElement;
                btn?.click();
            }

            // See if there is an active filter
            if (this._activeFilterValue) {
                let filterValues = (elItem.dataset.filter || "").split('|');

                // Parse the active filters
                let activeFilters = this._activeFilterValue.split('|');
                let showItem = false;
                for (let j = 0; j < activeFilters.length; j++) {
                    // See if the item contains the filter value
                    if (filterValues.indexOf(activeFilters[j]) >= 0) {
                        // Set the flag and break from the loop
                        showItem = true;
                        break;
                    }
                }

                // See if we are hiding the item
                if (!showItem) {
                    // Clear the item
                    this.clearItem(elItem);

                    // Exclude the item from the array
                    elItems.splice(i, 1);

                    // Continue the loop
                    continue;
                }
            }

            // See if a search value exists
            if (this._activeSearchValue) {
                // See if the item doesn't contains the search value
                if (elItem.innerText.toLowerCase().indexOf(this._activeSearchValue) < 0 &&
                    elContent.innerText.toLowerCase().indexOf(this._activeSearchValue) < 0) {
                    // Clear the item
                    this.clearItem(elItem);

                    // Exclude the item from the array
                    elItems.splice(i, 1);

                    // Continue the loop
                    continue;
                }
            }
        }

        // Parse the active items to show
        for (let i = 0; i < paginationLimit; i++) {
            let elItem = elItems[i] as HTMLElement;

            // Ensure the item exists
            if (elItem) {
                // Clear the item
                this.clearItem(elItem);

                // Show the item
                elItem.classList.remove("d-none");

                // See if this is the first item
                if (i == 0) { elItem.classList.add("first-item"); }
            } else {
                // Set the class for the last item
                if (elItems[i - 1]) { elItems[i - 1].classList.add("last-item"); }

                // Break from the loop
                break;
            }
        }

        // Parse the active items to hide
        for (let i = paginationLimit; i < elItems.length; i++) {
            // Clear the item
            this.clearItem(elItems[i] as HTMLElement);
        }

        // Ensure items exist
        if (elItems.length > 0) {
            // Ensure the first item is expanded
            if (elItems[0].querySelector(".accordion-collapse.show") == null) {
                // Hide the button
                let btn = elItems[0].querySelector(".accordion-button") as HTMLButtonElement;
                btn?.click();
            }

            // Set the first item class
            elItems[0].classList.add("first-item");

            // Set the last item class
            let lastIdx = paginationLimit - 1;
            elItems[lastIdx < elItems.length ? lastIdx : elItems.length - 1].classList.add("last-item");
        }

        // See if we are showing pagination
        let showPagination = typeof (this._props.showPagination) === "boolean" ? this._props.showPagination : true;
        if (showPagination) {
            // Render the pagination
            this._pagination = Components.Pagination({
                el: elPagination,
                className: "d-flex justify-content-end pt-2",
                numberOfPages: Math.ceil(elItems.length / paginationLimit),
                onClick: (pageNumber) => {
                    // Parse the items
                    for (let i = 0; i < elItems.length; i++) {
                        let elItem = elItems[i] as HTMLElement;

                        // See if this item is expanded
                        if (elItem.querySelector(".accordion-collapse.show")) {
                            // Hide the button
                            let btn = elItem.querySelector(".accordion-button") as HTMLButtonElement;
                            btn?.click();
                        }

                        // Clear the item
                        this.clearItem(elItem);
                    }

                    // Parse the items to show
                    let startIdx = (pageNumber - 1) * paginationLimit;
                    for (let i = startIdx; i < startIdx + paginationLimit && i < elItems.length; i++) {
                        let elItem = elItems[i];

                        // Show the item
                        elItem.classList.remove("d-none");

                        // See if this is the first item
                        if (i == startIdx) {
                            // Set the first item class
                            elItem.classList.add("first-item");

                            // Expand the item
                            let btn = elItem.querySelector(".accordion-button") as HTMLButtonElement;
                            btn?.click();
                        }
                    }

                    // Set the last item class
                    let lastIdx = startIdx + paginationLimit - 1;
                    elItems[lastIdx < elItems.length ? lastIdx : elItems.length - 1].classList.add("last-item");

                    // Call the event
                    this._props.onPaginationClick ? this._props.onPaginationClick(pageNumber) : null;
                }
            });

            // Call the event
            this._props.onPaginationRender ? this._props.onPaginationRender(elPagination) : null;
        }
    }

    // Searches the accordion
    search(value: string) {
        // Set the search value
        this._activeSearchValue = (value || "").toLowerCase();

        // Render the items
        this.renderItems();
    }
}