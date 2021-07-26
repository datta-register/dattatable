import { Components } from "gd-sprest-bs";
import { filterSquare } from "gd-sprest-bs/build/icons/svgs/filterSquare";

/**
 * Navigation
 */
interface INavProps {
    el: HTMLElement;
    hideFilter?: boolean;
    items: Components.INavbarItem[];
    itemsEnd: Components.INavbarItem[];
    onFilterRendered?: (el: HTMLElement) => void;
    onRendering?: (props: Components.INavbarProps) => void;
    onRendered?: (el: HTMLElement) => void;
    onShowFilter: Function;
    onSearch: (value: string) => void;
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
            }
        };

        // Call the rendering event
        this._props.onRendering ? this._props.onRendering(props) : null;

        // Render a navbar
        let nav = Components.Navbar(props);

        // Update the navbar color palate
        nav.el.classList.remove("navbar-light");
        nav.el.classList.add("navbar-dark");
        
        // See if we are showing the filter
        if (this._props.hideFilter != true) {
            // Render the filter icon
            let icon = document.createElement("div");
            icon.classList.add("filter-icon");
            icon.classList.add("nav-link");
            icon.classList.add("text-light");
            icon.style.cursor = "pointer";
            icon.appendChild(filterSquare());
            icon.addEventListener("click", this._props.onShowFilter as any);
            nav.el.firstElementChild.appendChild(icon);

            // Call the render event
            this._props.onFilterRendered ? this._props.onFilterRendered(icon) : null;
        }
    }
}
