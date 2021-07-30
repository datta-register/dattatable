import { Components } from "gd-sprest-bs";

/**
 * Modal
 */
class _Modal {
    private _modal: Components.IModal = null;
    private _onCloseEvent: Function = null;

    // Modal Body
    private _elBody: HTMLElement = null;
    get BodyElement(): HTMLElement { return this._elBody; }

    // Modal Footer
    private _elFooter: HTMLElement = null;
    get FooterElement(): HTMLElement { return this._elFooter; }

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
        // Clear the header, body and footer
        this.setHeader("");
        this.setBody("");
        this.setFooter("");
    }

    // Hides the modal
    hide() { this._modal.hide(); }

    // Renders the canvas
    private render() {
        // Create the element
        let el = document.createElement("div");
        el.id = "core-modal";
        document.body.appendChild(el);

        // Render the canvas
        this._modal = Components.Modal({
            el,
            isCentered: true,
            type: Components.ModalTypes.Large,
            options: {
                autoClose: false,
                backdrop: true,
                keyboard: true
            },
            onRenderBody: el => { this._elBody = el; },
            onRenderFooter: el => { this._elFooter = el; },
            onRenderHeader: el => { this._elHeader = el.querySelector(".modal-title");; },
            onClose: () => {
                // Call the close event
                this._onCloseEvent ? this._onCloseEvent() : null;
            }
        });
    }

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

    // Sets the close event
    setCloseEvent(event) { this._onCloseEvent = event; }

    // Sets the footer
    setFooter(content) {
        // Clear the body
        while (this._elFooter.firstChild) { this._elFooter.removeChild(this._elFooter.firstChild); }

        // See if content exists
        if (content) {
            // See if this is text
            if (typeof (content) == "string") {
                // Set the html
                this._elFooter.innerHTML = content;
            } else {
                // Append the element
                this._elFooter.appendChild(content);
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
    setType(type) { this._modal.setType(type); }

    // Shows the modal
    show() { this._modal.show(); }
}
export const Modal = new _Modal();
