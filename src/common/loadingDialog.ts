import { Components } from "gd-sprest-bs";

/**
 * Loading Dialog
 */
export class LoadingDialog {
    private static _el: HTMLElement = null;
    private static _elBody: HTMLElement = null;
    private static _elBackdrop: HTMLElement = null;
    private static _elHeader: HTMLElement = null;

    // Constructor
    constructor() {
        // Render the canvas
        LoadingDialog.render();
    }

    // Element
    static get el(): HTMLElement { return this._el; }

    // Hides the canvas
    static hide() {
        // Update the display
        this._el.style.display = "none";
        this._elBackdrop.style.display = "none";
    }

    // Renders the canvas
    private static render() {
        // Create the backdrop element
        this._elBackdrop = document.createElement("div");
        this._elBackdrop.id = "loading-dialog-backdrop";
        this._elBackdrop.style.display = "none";

        // Create the loading dialog element
        this._el = document.createElement("div");
        this._el.id = "loading-dialog";
        this._el.classList.add("bs");
        this._el.style.display = "none";
        this._el.innerHTML = "<div></div>";

        // Ensure the body exists
        if (document.body) {
            // Append the elements
            document.body.appendChild(this._elBackdrop);
            document.body.appendChild(this._el);
        } else {
            // Create an event
            window.addEventListener("load", () => {
                // Append the elements
                document.body.appendChild(this._elBackdrop);
                document.body.appendChild(this._el);
            });
        }

        // Update the main element
        let elMain = this._el.firstChild as HTMLElement;
        elMain.classList.add("d-flex");
        elMain.classList.add("flex-column");
        
        // Append the header
        this._elHeader = document.createElement("div");
        this._elHeader.classList.add("fs-4");
        this._elHeader.classList.add("p-2");
        elMain.appendChild(this._elHeader);

        // Render a spinner
        Components.Spinner({
            el: elMain,
            type: Components.SpinnerTypes.Primary
        });

        // Append the body
        this._elBody = document.createElement("div");
        this._elBody.classList.add("fs-6");
        this._elBody.classList.add("p-2");
        elMain.appendChild(this._elBody);
    }

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

    // Shows the canvas
    static show() {
        // Update the display
        this._el.style.display = "";
        this._elBackdrop.style.display = "";
    }
}

// Create an instance of the loading dialog
new LoadingDialog();
