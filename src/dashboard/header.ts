import { Components } from "gd-sprest-bs";
import { IHeaderProps } from "./header.d";
export * from "./header";

/**
 * Header
 */
export class Header {
    private _props: IHeaderProps = null;

    // Constructor
    constructor(props: IHeaderProps) {
        // Save the properties
        this._props = props;

        // Render the header
        this.render();
    }

    // Renders the component
    private render() {
        // Render a jumbotron
        Components.Jumbotron({
            el: this._props.el,
            className: "header",
            lead: this._props.title
        });
    }
}