import { ThemeManager } from "gd-sprest-bs";

/** Apply the Current Theme */
let _themeLoaded = false;
ThemeManager.load(true).then(() => { _themeLoaded = true; });
export function waitForTheme(): PromiseLike<void> {
    // Return a promise
    return new Promise((resolve) => {
        // See if the flag is set
        if (_themeLoaded) { resolve(); return; }

        // Wait for the theme to be loaded
        let counter = 0;
        let maxAttempts = 100;
        let loopId = setInterval(() => {
            // See if the theme was loaded
            if (_themeLoaded) {
                // Stop the loop
                clearInterval(loopId);

                // Resolve the request
                resolve();
            }
            // Else, see if we have hit the max attempts
            else if (++counter >= maxAttempts) {
                // Stop the loop
                clearInterval(loopId);

                // Resolve the request
                resolve();
            }
        }, 50);
    });
}