import { Components } from "gd-sprest-bs";

/**
 * Canvas Form
 */
export class _CanvasForm {
    private _canvas: Components.IOffcanvas = null;

    // Modal Body
    private _elBody: HTMLElement = null;
    get BodyElement(): HTMLElement { return this._elBody; }

    // Modal Header
    private _elHeader: HTMLElement = null;
    get HeaderElement(): HTMLElement { return this._elHeader; }

    // Constructor
    constructor() {
        // Render the canvas
        this.render();
    }

    // Clears the canvas form
    clear() {
        // Clear the header and body
        this.setHeader("");
        this.setBody("");
    }

    // Element
    get el(): HTMLElement { return this._canvas.el as HTMLElement; }

    // Hides the canvas
    hide() { this._canvas.hide(); }

    // Renders the canvas
    private render() {
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
    setAutoClose(value: boolean) { this._canvas.setAutoClose(value); }

    // Sets the body
    setBody(content) {
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
    setHeader(content) {
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
    setType(type) { this._canvas.setType(type); }

    // Shows the canvas
    show() { this._canvas.show(); }
}
export const CanvasForm = new _CanvasForm();