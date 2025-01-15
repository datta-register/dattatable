import { Components } from "gd-sprest-bs";
import { filter } from "gd-sprest-bs/build/icons/svgs/filter";

/**
 * Navigation
 */
interface INavProps {
    el: HTMLElement;
    className?: string;
    hideFilter?: boolean;
    hideSearch?: boolean;
    iconSize?: number;
    iconType?: SVGImageElement | Function;
    items?: Components.INavbarItem[];
    itemsEnd?: Components.INavbarItem[];
    onFilterRendered?: (el: HTMLElement) => void;
    onRendering?: (props: Components.INavbarProps) => void;
    onRendered?: (el: HTMLElement) => void;
    onSearchRendered?: (el: HTMLElement) => void;
    onShowFilter?: () => void;
    onSearch?: (value: string) => void;
    searchPlaceholder?: string;
    title: string | HTMLElement;
}

/**
 * Navigation
 */
export class Navigation {
    private _nav: Components.INavbar;
    private _props: INavProps = null;

    // Constructor
    constructor(props: INavProps) {
        // Save the properties
        this._props = props;

        // Render the navigation
        this.render();

        // Call the render event
        props.onRendered ? props.onRendered(this._props.el) : null;
    }

    // Renders the component
    private render() {
        // See if an icon is defined
        let brand;
        if (this._props.iconType) {
            // Set the icon
            let size = this._props.iconSize || 32;
            let icon: HTMLElement = typeof (this._props.iconType) === "function" ? this._props.iconType(size) : this._props.iconType;

            // Set the text
            let text = document.createElement("div");
            text.classList.add("ms-2");
            text.append(this._props.title || "");

            // Set the brand
            brand = document.createElement("div");
            brand.classList.add("d-flex");
            brand.appendChild(icon);
            brand.appendChild(text);
        } else {
            // Set the brand
            brand = this._props.title;
        }

        // Define the default props
        let props: Components.INavbarProps = {
            brand,
            className: this._props.className,
            el: this._props.el,
            enableSearch: this._props.hideSearch != null ? !this._props.hideSearch : null,
            items: this._props.items,
            itemsEnd: this._props.itemsEnd,
            searchBox: {
                hideButton: true,
                onChange: this._props.onSearch,
                onSearch: this._props.onSearch,
                placeholder: this._props.searchPlaceholder || "Search this app",
            },
            onRendered: el => {
                // Update the collapse visibility
                el.classList.remove("navbar-expand-lg");
                el.classList.add("navbar-expand-sm");
                el.classList.add("rounded-top");
            }
        };

        // Call the rendering event
        this._props.onRendering ? this._props.onRendering(props) : null;

        // Render a navbar
        this._nav = Components.Navbar(props);

        // See if we are showing the filter
        if (this._props.hideFilter != true) {
            // Render the filter icon
            // Create a span to wrap the icon in
            let span = document.createElement("span");
            span.className = "bg-white d-inline-flex filter-icon ms-2 rounded";
            this._nav.el.firstElementChild.appendChild(span);

            // Render a tooltip
            let ttp = Components.Tooltip({
                el: span,
                content: "Filters",
                type: Components.TooltipTypes.Secondary,
                btnProps: {
                    // Render the filter button
                    iconType: filter,
                    iconSize: 28,
                    type: Components.ButtonTypes.OutlineSecondary
                },
            });

            // Call the event
            this._props.onShowFilter && ttp ? ttp.el.addEventListener("click", this._props.onShowFilter as any) : null;
            this._props.onFilterRendered ? this._props.onFilterRendered(span) : null;
        }

        // Call the event
        this._props.onSearchRendered ? this._props.onSearchRendered(this._props.el.querySelector("input[type='search']")) : null;
    }

    /**
     * Public Interface
     */

    // Returns the current search value
    getSearchValue() { return this._nav.getSearchValue(); }
}