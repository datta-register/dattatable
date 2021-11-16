import { Components } from "gd-sprest-bs";
import { filter } from "gd-sprest-bs/build/icons/svgs/filter";

/**
 * Navigation
 */
interface INavProps {
    el: HTMLElement;
    hideFilter?: boolean;
    items?: Components.INavbarItem[];
    itemsEnd?: Components.INavbarItem[];
    onFilterRendered?: (el: HTMLElement) => void;
    onRendering?: (props: Components.INavbarProps) => void;
    onRendered?: (el: HTMLElement) => void;
    onShowFilter?: Function;
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
            className: "bg-sharepoint header rounded",
            items: this._props.items,
            itemsEnd: this._props.itemsEnd,
            searchBox: {
                hideButton: true,
                onChange: this._props.onSearch,
                onSearch: this._props.onSearch
            },
            type: Components.NavbarTypes.Dark
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
            span.className = "bg-white d-inline-flex filter-icon ms-2 nav-link rounded";
            nav.el.firstElementChild.appendChild(span);

            // Render a tooltip
            let ttp = Components.Tooltip({
                el: span,
                content: "Filters",
                type: Components.TooltipTypes.LightBorder,
                btnProps: {
                    // Render the filter button
                    iconType: filter,
                    iconSize: 28,
                    type: Components.ButtonTypes.OutlineSecondary
                },
            });
            this._props.onShowFilter ? ttp.el.addEventListener("click", this._props.onShowFilter as any) : null;
            
            // Call the render event
            this._props.onFilterRendered ? this._props.onFilterRendered(icon) : null;
        }
    }
}
