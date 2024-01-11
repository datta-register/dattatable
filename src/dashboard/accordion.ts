import { Components } from "gd-sprest-bs";

export interface IAccordion {
    filter?: (string) => void;
    search?: (value: string) => void;
}

export interface IAccordionProps {
    bodyField?: string;
    el: HTMLElement;
    filterField?: string;
    items: any[];
    onItemClick?: (el?: HTMLElement, item?: any) => void;
    onItemRender?: (el?: HTMLElement, item?: any) => void;
    paginationLimit?: number;
    titleField?: string;
}

/**
 * Accordion
 */
export class Accordion implements IAccordion {
    private _accordion: Components.IAccordion = null;
    private _activeFilterClass: string = null;
    private _activeSearchFilter: string = null;
    private _pagination: Components.IPagination = null;
    private _props: IAccordionProps;

    // Constructor
    constructor(props: IAccordionProps) {
        // Save the properties
        this._props = props;

        // Render the accordion
        this.render();
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
    filter(value: string) {
        let className = value ? value.toLowerCase().replace(/ /g, "-") : null;
        this._activeFilterClass = className;

        // Parse all accordion items
        let items = this._props.el.querySelectorAll(".accordion-item");
        for (let i = 0; i < items.length; i++) {
            let elItem = items[i];

            // Show the item
            elItem.classList.remove("d-none");

            // Remove the first/last item classes
            elItem.classList.remove("first-item");
            elItem.classList.remove("last-item");

            // See if a class name exists
            if (className) {
                // See if this item doesn't matches
                if (!elItem.classList.contains(className)) {
                    // Hide the item
                    elItem.classList.add("d-none");
                }
            }
        }

        // Render the items
        this.renderItems();
    }

    // Renders the dashboard
    private render() {
        // Render the accordion
        this.renderAccordion();

        // Render the items
        this.renderItems();
    }

    // Renders the accordion
    private renderAccordion() {
        // Parse the items
        let accordionItems: Array<Components.IAccordionItem> = [];
        for (let i = 0; i < this._props.items.length; i++) {
            let item = this._props.items[i];
            let itemClassNames = [];

            // See if the filter field is specified
            if (this._props.filterField) {
                // Parse the selected filters
                let filters = item[this._props.filterField] || [];
                filters = filters["results"] || filters;
                for (let j = 0; j < filters.length; j++) {
                    let filter = filters[j].toLowerCase().replace(/ /g, '-');

                    // Add the filter as a class name
                    if (filter) { itemClassNames.push(filter); }
                }
            }

            // Add an accordion item
            accordionItems.push({
                className: itemClassNames.join(" "),
                content: (this._props.bodyField ? item[this._props.bodyField || "Description"] : null) || "",
                header: item[this._props.titleField || "Title"] || "",
                onClick: this._props.onItemClick,
                onRender: this._props.onItemRender,
                showFl: i == 0,
            });
        }

        // Render the accordion
        this._accordion = Components.Accordion({
            el: this._props.el,
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
        let elItems = Array.from(this._accordion.el.querySelectorAll(this._activeFilterClass ? "." + this._activeFilterClass : ".accordion-item"));

        // See if a search value exists
        if (this._activeSearchFilter) {
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


                // See if the item doesn't contains the search value
                if (elItem.innerText.toLowerCase().indexOf(this._activeSearchFilter) < 0 &&
                    elContent.innerText.toLowerCase().indexOf(this._activeSearchFilter) < 0) {
                    // Clear the item
                    this.clearItem(elItem);

                    // Exclude the item from the array
                    elItems.splice(i, 1);
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
            }
        });
    }

    // Searches the accordion
    search(value: string) {
        // Set the search value
        this._activeSearchFilter = (value || "").toLowerCase();

        // Render the items
        this.renderItems();
    }
}