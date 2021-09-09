import { Helper, Web } from "gd-sprest-bs";

/**
 * Installation Required
 * Checks the SharePoint configuration file to see if an install is required.
 */
class _InstallationRequired {
    private _cfg: Helper.ISPConfig = null;
    private _report: string[] = null;

    // Checks the lists
    private checkLists(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
            // Parse the lists
            Helper.Executor(this._cfg._configuration.ListCfg, listCfg => {
                // Return a promise
                return new Promise((resolve) => {
                    // Get the list
                    Web().Lists(listCfg.ListInformation.Title).query({
                        Expand: ["Fields"]
                    }).execute(
                        // Success
                        list => {
                            // Parse the custom fields
                            Helper.Executor(listCfg.CustomFields, cfg => {
                                let exists = false;

                                // Parse the list fields
                                for (let i = 0; i < list.Fields.results.length; i++) {
                                    // See if the field exists
                                    if (list.Fields.results[i].InternalName == cfg.name) {
                                        // Set the flag
                                        exists = true;
                                        break;
                                    }
                                }

                                // See if the field doesn't exist
                                if (!exists) {
                                    // Field doesn't exist
                                    this._report.push("List '" + listCfg.ListInformation.Title + "' is missing a field: " + cfg.name);
                                }
                            }).then(() => {
                                // Resolve the promise
                                resolve(null);
                            });
                        },

                        // Error
                        () => {
                            // List doesn't exist
                            this._report.push("List '" + listCfg.ListInformation.Title + "' does not exist");

                            // Resolve the promise
                            resolve(null);
                        }
                    );
                });
            }).then(() => {
                // Resolve the request
                resolve();
            });
        });
    }

    // Checks the configuration to see if an installation is required
    requiresInstall(cfg: Helper.ISPConfig): PromiseLike<boolean> {
        // Save the configuration
        this._cfg = cfg;

        // Clear the report
        this._report = [];

        // Return a promise
        return new Promise((resolve) => {
            // Check the lists
            this.checkLists().then(() => {
                // Resolve the request
                resolve(this._report.length > 0);
            });
        });
    }
}
export const InstallationRequired = new _InstallationRequired();