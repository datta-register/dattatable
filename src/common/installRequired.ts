import { Components, Helper, Web } from "gd-sprest-bs";
import { LoadingDialog } from "./loadingDialog";
import { Modal } from "./modal";

/**
 * Installation Required
 * Checks the SharePoint configuration file to see if an install is required.
 */
export class InstallationRequired {
    private static _cfg: Helper.ISPConfig = null;
    private static _report: string[] = null;

    // Checks the lists
    private static checkLists(): PromiseLike<void> {
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
    static requiresInstall(cfg: Helper.ISPConfig): PromiseLike<boolean> {
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

    static showDialog() {
        // Set the header
        Modal.setHeader("Installation Required");

        // Set the body
        Modal.setBody(Components.Card({
            body: [
                {
                    text: "An installation is required. The following were missing in your environment:"
                },
                {
                    onRender: el => {
                        let items: Components.IListGroupItem[] = [];

                        // Parse the report
                        for (let i = 0; i < this._report.length; i++) {
                            // Add the item
                            items.push({ content: this._report[i] });
                        }

                        // Render a list
                        Components.ListGroup({
                            el,
                            items
                        });
                    }
                }
            ],
            footer: {
                onRender: el => {
                    let btnInstall: Components.IButton = null;
                    let btnRefresh: Components.IButton = null;

                    // Render the install button
                    Components.Tooltip({
                        el,
                        content: "Installs the SharePoint Assets",
                        btnProps: {
                            assignTo: btn => { btnInstall = btn; },
                            text: "Install",
                            type: Components.ButtonTypes.OutlineSuccess,
                            onClick: () => {
                                // Show a loading dialog
                                LoadingDialog.setHeader("Installing the Solution");
                                LoadingDialog.setBody("This will close after the assets are installed.");
                                LoadingDialog.show();

                                this._cfg.install().then(() => {
                                    // Hide the button
                                    btnInstall.hide();

                                    // Show the refresh button
                                    btnRefresh.show();

                                    // Close the dialog
                                    LoadingDialog.hide();
                                });
                            }
                        }
                    });

                    // Render the refresh button
                    Components.Tooltip({
                        el,
                        content: "Refresh the Page",
                        btnProps: {
                            assignTo: btn => { btnRefresh = btn; },
                            className: "d-none",
                            text: "Refresh",
                            type: Components.ButtonTypes.OutlinePrimary,
                            onClick: () => {
                                // Refresh the page
                                window.location.reload();
                            }
                        }
                    });
                }
            }
        }).el);

        // Show the modal
        Modal.show();
    }
}