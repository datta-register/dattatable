import { Components } from "gd-sprest-bs";

/**
 * Header
 */
 export interface IHeaderProps {
    el: HTMLElement;
    onRender?: (el:HTMLElement) => void;
    title?: string;
}

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

        // Call the render event
        props.onRender ? props.onRender(this._props.el) : null;
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