import { Web } from "gd-sprest-bs";
import { Modal } from "./modal";

/**
 * Timeout
 */
export class Timeout {
    private static _loopId: number = null;
    private static _timeout: number = null;

    // Shows the canvas
    private static checkWeb() {
        // Query the current web
        Web().query({ Select: ["Title"] }).execute(
            // Success
            () => {
                // Do nothing
            },

            // Error getting the web information
            // Display a modal
            () => {
                // Clear the modal
                Modal.clear();

                // Set the header
                Modal.setHeader("Error Loading Web");

                // Set the body
                Modal.setBody("There was an error getting the current web. Please refresh the page.");

                // Show the modal
                Modal.show();

                // Stop the loop
                this.stop();
            }
        )
    }

    // Set the timeout value
    // Default value is 300000ms (5 min)
    static setTimer(timeout: number = 300000) {
        // Update the value
        this._timeout = timeout;
    }

    // Stops the timeout
    static stop() {
        // Ensure the loop id exists
        if (this._loopId) {
            // Stop the loop
            clearInterval(this._loopId);

            // Clear the loop id
            this._loopId = null;
        }
    }

    // Starts the timeout
    static start() {
        // Ensure it's not running
        if (this._loopId > 0) { return; }

        // Ensure the timer is set
        if (this._timeout == null) {
            // Set the default timer value
            this.setTimer();
        }

        // Start the interval
        this._loopId = setInterval(this.checkWeb, this._timeout) as any;
    }
}