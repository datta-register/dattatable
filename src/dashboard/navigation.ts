import { Components } from "gd-sprest-bs";
import { filter } from "gd-sprest-bs/build/icons/svgs/filter";

/**
 * Navigation
 */
interface INavProps {
    el: HTMLElement;
    hideFilter?: boolean;
    hideSearch?: boolean;
    items?: Components.INavbarItem[];
    itemsEnd?: Components.INavbarItem[];
    onFilterRendered?: (el: HTMLElement) => void;
    onRendering?: (props: Components.INavbarProps) => void;
    onRendered?: (el: HTMLElement) => void;
    onSearchRendered?: (el: HTMLElement) => void;
    onShowFilter?: () => void;
    onSearch?: (value: string) => void;
    title: string | HTMLElement;
}

/**
 * Navigation
 */
export class Navigation {
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
        // Define the default props
        let props: Components.INavbarProps = {
            el: this._props.el,
            brand: this._props.title,
            enableSearch: this._props.hideSearch != null ? !this._props.hideSearch : null,
            items: this._props.items,
            itemsEnd: this._props.itemsEnd,
            searchBox: {
                hideButton: true,
                onChange: this._props.onSearch,
                onSearch: this._props.onSearch,
                placeholder: "Search this app",
            }
        };

        // Call the rendering event
        this._props.onRendering ? this._props.onRendering(props) : null;

        // Render a navbar
        let nav = Components.Navbar(props);

        // See if we are showing the filter
        if (this._props.hideFilter != true) {
            // Render the filter icon
            // Create a span to wrap the icon in
            let span = document.createElement("span");
            span.className = "bg-white d-inline-flex filter-icon ms-2 rounded";
            nav.el.firstElementChild.appendChild(span);

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

            // Call the events
            this._props.onSearchRendered ? this._props.onSearchRendered(this._props.el.querySelector("input[type='search']")) : null;
            this._props.onShowFilter ? ttp.el.addEventListener("click", this._props.onShowFilter as any) : null;
            this._props.onFilterRendered ? this._props.onFilterRendered(this._props.el) : null;
        }
    }
}
