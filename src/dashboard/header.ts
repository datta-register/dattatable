import { Components } from "gd-sprest-bs";

/**
 * Header
 */
export interface IHeaderProps {
    el: HTMLElement;
    onRendering?: (props: Components.IJumbotronProps) => void;
    onRendered?: (el: HTMLElement) => void;
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
        props.onRendered ? props.onRendered(this._props.el) : null;
    }

    // Renders the component
    private render() {
        // Define the default props
        let props: Components.IJumbotronProps = {
            el: this._props.el,
            className: "header",
            lead: this._props.title
        };

        // Call the rendering event
        this._props.onRendering ? this._props.onRendering(props) : null;

        // Render a jumbotron
        Components.Jumbotron(props);
    }
}