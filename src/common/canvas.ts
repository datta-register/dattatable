import { Components } from "gd-sprest-bs";

/**
 * Canvas Form
 */
export class CanvasForm {
    private static _canvas: Components.IOffcanvas = null;

    // Modal Body
    private static _elBody: HTMLElement = null;
    static get BodyElement(): HTMLElement { return this._elBody; }

    // Modal Header
    private static _elHeader: HTMLElement = null;
    static get HeaderElement(): HTMLElement { return this._elHeader; }

    // Constructor
    constructor() {
        // Render the canvas
        CanvasForm.render();
    }

    // Clears the canvas form
    static clear() {
        // Clear the header and body
        this.setHeader("");
        this.setBody("");

        // Set the default properties
        this.setAutoClose(true);
        this.setType(Components.OffcanvasTypes.End);
    }

    // Element
    static get el(): HTMLElement { return this._canvas.el as HTMLElement; }

    // Hides the canvas
    static hide() { this._canvas.hide(); }

    // Renders the canvas
    private static render() {
        // Create the element
        let el = document.createElement("div");
        el.id = "core-canvas";

        // Ensure the body exists
        if (document.body) {
            // Append the element
            document.body.appendChild(el);
        } else {
            // Create an event
            window.addEventListener("load", () => {
                // Append the element
                document.body.appendChild(el);
            });
        }

        // Render the canvas
        this._canvas = Components.Offcanvas({
            el,
            options: {
                autoClose: true,
                backdrop: true,
                focus: true,
                keyboard: true,
                scroll: true
            },
            onRenderBody: el => { this._elBody = el; },
            onRenderHeader: el => { this._elHeader = el; }
        });
    }

    // Sets the auto close flag
    static setAutoClose(value: boolean) { this._canvas.setAutoClose(value); }

    // Sets the body
    static setBody(content) {
        // Clear the body
        while (this._elBody.firstChild) { this._elBody.removeChild(this._elBody.firstChild); }

        // See if content exists
        if (content) {
            // See if this is text
            if (typeof (content) == "string") {
                // Set the html
                this._elBody.innerHTML = content;
            } else {
                // Append the element
                this._elBody.appendChild(content);
            }
        }
    }

    // Sets the header
    static setHeader(content) {
        // Clear the body
        while (this._elHeader.firstChild) { this._elHeader.removeChild(this._elHeader.firstChild); }

        // See if content exists
        if (content) {
            // See if this is text
            if (typeof (content) == "string") {
                // Set the html
                this._elHeader.innerHTML = content;
            } else {
                // Append the element
                this._elHeader.appendChild(content);
            }
        }
    }

    // Sets the modal type
    static setType(type) { this._canvas.setType(type); }

    // Shows the canvas
    static show() { this._canvas.show(); }
}

// Create an instance of the canvas form
new CanvasForm();