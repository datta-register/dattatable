import { Components } from "gd-sprest-bs";

/**
 * Footer
 */
export interface IFooterProps {
    el: HTMLElement;
    items?: Components.INavbarItem[];
    itemsEnd?: Components.INavbarItem[];
    onRendering?: (props: Components.INavbarProps) => void;
    onRendered?: (el: HTMLElement) => void;
}

/**
 * Footer
 */
export class Footer {
    private _props: IFooterProps = null;

    // Constructor
    constructor(props: IFooterProps) {
        // Save the properties
        this._props = props;

        // Render the footer
        this.render();

        // Call the render event
        props.onRendered ? props.onRendered(this._props.el) : null;
    }

    // Renders the component
    private render() {
        // Define the default props
        let props: Components.INavbarProps = {
            el: this._props.el,
            items: this._props.items,
            itemsEnd: this._props.itemsEnd
        };

        // Call the rendering event
        this._props.onRendering ? this._props.onRendering(props) : null;

        // Render a navbar
        Components.Navbar(props);
    }
}