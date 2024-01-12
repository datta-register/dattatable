import { Components } from "gd-sprest-bs";

export interface ITiles {
    filter?: (string) => void;
    search?: (value: string) => void;
}

export interface ITilesProps {
    bodyField?: string;
    el: HTMLElement;
    filterField?: string;
    items: any[];
    onCardRender?: (el?: HTMLElement, item?: any) => void;
    paginationLimit?: number;
    subTitleField?: string;
    titleField?: string;
}

/**
 * Tiles
 */
export class Tiles implements ITiles {
    private _tiles: Components.ICardGroup = null;
    private _activeFilterClass: string = null;
    private _activeSearchFilter: string = null;
    private _pagination: Components.IPagination = null;
    private _props: ITilesProps;

    // Constructor
    constructor(props: ITilesProps) {
        // Save the properties
        this._props = props;

        // Render the tile
        this.render();
    }

    // Filters the tile
    filter(value: string) {
        let className = value ? value.toLowerCase().replace(/ /g, "-") : null;
        this._activeFilterClass = className;

        // Parse all tile items
        let items = this._props.el.querySelectorAll(".card");
        for (let i = 0; i < items.length; i++) {
            let elItem = items[i];

            // Show the item
            elItem.classList.remove("d-none");
        }

        // Update the tiles
        this.updateTiles();
    }

    // Renders the dashboard
    private render() {
        // Render the tile
        this.renderTiles();

        // Update the tiles
        this.updateTiles();
    }

    // Renders the tiles
    private renderTiles() {
        // Parse the items
        let cards: Array<Components.ICardProps> = [];
        for (let i = 0; i < this._props.items.length; i++) {
            let item = this._props.items[i];
            let itemClassNames = [];

            // Get the filters
            let filters = item[this._props.filterField] || [];
            filters = filters["results"] || filters;

            // See if the filter field is specified
            if (this._props.filterField) {
                // Parse the selected filters
                for (let j = 0; j < filters.length; j++) {
                    let filter = filters[j].toLowerCase().replace(/ /g, '-');

                    // Add the filter as a class name
                    if (filter) { itemClassNames.push(filter); }
                }
            }

            // Add an tile
            cards.push({
                className: itemClassNames.join(" "),
                body: [{
                    content: (this._props.bodyField ? item[this._props.bodyField || "Description"] : null) || "",
                    data: filters,
                    subTitle: item[this._props.subTitleField] || "",
                    title: item[this._props.titleField || "Title"] || "",
                    onRender: this._props.onCardRender
                }]
            });
        }

        // Render the tiles
        this._tiles = Components.CardGroup({
            el: this._props.el,
            cards
        });
    }

    // Updates the visibility of the tiles
    private updateTiles() {
        let paginationLimit = this._props.paginationLimit || 10;

        // Get the pagination element
        let elPagination = this._props.el.querySelector(".tiles-pagination") as HTMLElement;
        if (elPagination) {
            // Clear the element
            while (elPagination.firstChild) { elPagination.removeChild(elPagination.firstChild); }
        } else {
            // Create the element
            elPagination = document.createElement("div");
            elPagination.classList.add("tiles-pagination");
            this._props.el.appendChild(elPagination);
        }

        // Get the elements as an array
        let elItems = Array.from(this._tiles.el.querySelectorAll(this._activeFilterClass ? "." + this._activeFilterClass : ".card"));

        // See if a search value exists
        if (this._activeSearchFilter) {
            // Parse the items
            for (let i = elItems.length - 1; i >= 0; i--) {
                let elItem = elItems[i] as HTMLElement;

                // See if the item doesn't contains the search value
                if (elItem.innerText.toLowerCase().indexOf(this._activeSearchFilter) < 0) {
                    // Hide the item
                    elItem.classList.add("d-none");

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
                // Show the item
                elItem.classList.remove("d-none");
            } else {
                // Set the class for the last item
                if (elItems[i - 1]) { elItems[i - 1].classList.add("last-item"); }

                // Break from the loop
                break;
            }
        }

        // Parse the active items to hide
        for (let i = paginationLimit; i < elItems.length; i++) {
            let elItem = elItems[i] as HTMLElement;

            // Hide the item
            elItem.classList.add("d-none");
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

                    // Hide the item
                    elItem.classList.add("d-none");
                }

                // Parse the items to show
                let startIdx = (pageNumber - 1) * paginationLimit;
                for (let i = startIdx; i < startIdx + paginationLimit && i < elItems.length; i++) {
                    let elItem = elItems[i];

                    // Show the item
                    elItem.classList.remove("d-none");
                }
            }
        });
    }

    // Searches the tile
    search(value: string) {
        // Set the search value
        this._activeSearchFilter = (value || "").toLowerCase();

        // Update the tiles
        this.updateTiles();
    }
}