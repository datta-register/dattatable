import { Components } from "gd-sprest-bs";

/**
 * Canvas Form
 */
export class CanvasForm {
    private static _canvas: Components.IOffcanvas = null;
    private static _onCloseEvent: Function = null;

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
        this.setCloseEvent(null);
        this.setSize(0);
        this.setType(Components.OffcanvasTypes.End);

        // Ensure the header and body are visible
        this._elBody.classList.remove("d-none");
        this._elHeader.classList.remove("d-none");
    }

    // Element
    static get el(): HTMLElement { return this._canvas.el as HTMLElement; }

    // Hides the canvas
    static hide() {
        this._canvas.hide();

        // Call the close event
        this._onCloseEvent ? this._onCloseEvent() : null;
    }

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
            onRenderHeader: el => { this._elHeader = el; },
            onClose: () => {
                // Call the close event
                this._onCloseEvent ? this._onCloseEvent() : null;
            }
        });
    }

    // Sets the auto close flag
    static setAutoClose(value: boolean) { this._canvas.setAutoClose(value); }

    // Sets the close event
    static setCloseEvent(event) { this._onCloseEvent = event; }

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

    // Sets the modal size
    static setSize(size: number) { this._canvas.setSize(size); }

    // Sets the modal type
    static setType(type: number) { this._canvas.setType(type); }

    // Shows the canvas
    static show() { this._canvas.show(); }
}

// Create an instance of the canvas form
new CanvasForm();