import { Components, Helper, Site, Web } from "gd-sprest-bs";
import { LoadingDialog } from "./loadingDialog";
import { Modal } from "./modal";

// Show Dialog Properties
export interface IShowDialogProps {
    errors?: Components.IListGroupItem[];
    onHeaderRendered?(el: HTMLElement);
    onBodyRendered?(el: HTMLElement);
    onFooterRendered?(el: HTMLElement);
}

/**
 * Installation Required
 * Checks the SharePoint configuration file to see if an install is required.
 */
export class InstallationRequired {
    private static _cfg: Helper.ISPConfig = null;
    private static _report: string[] = null;

    // Custom Actions exists
    private static _customActionsExist: boolean = null;
    static get CustomActionsExist(): boolean { return this._customActionsExist; }

    // Checks the custom actions
    private static checkCustomActions(): PromiseLike<void> {
        // Clear the flag
        this._customActionsExist = true;

        // Method to check the web custom actions
        let checkWeb = () => {
            // Return a promise
            return new Promise(resolve => {
                // Ensure custom actions exist
                if (this._cfg._configuration.CustomActionCfg == null || this._cfg._configuration.CustomActionCfg.Web == null || this._cfg._configuration.CustomActionCfg.Web.length == 0) {
                    // Resolve the request
                    resolve(null);
                    return;
                }

                // Load the web custom actions
                let web = Web();
                web.UserCustomActions().execute(webCustomActions => {
                    // Parse the web custom actions
                    for (let i = 0; i < this._cfg._configuration.CustomActionCfg.Web.length; i++) {
                        let customAction = this._cfg._configuration.CustomActionCfg.Web[i];
                        let found = false;

                        // Parse the web custom actions
                        for (let j = 0; j < webCustomActions.results.length; j++) {
                            // See if they match
                            if (customAction.Title == webCustomActions.results[j].Title) {
                                // Set the flag
                                found = true;
                            }
                        }

                        // See if it was not found
                        if (!found) {
                            // Set the error
                            this._report.push("Web Custom Action '" + customAction.Title + "' is missing.");

                            // Set the flag
                            this._customActionsExist = false;
                        }
                    }

                    // Resolve the request
                    resolve(null);
                });
            });
        }

        // Method to check the site custom actions
        let checkSite = () => {
            return new Promise((resolve) => {
                // Ensure custom actions exist
                if (this._cfg._configuration.CustomActionCfg == null || this._cfg._configuration.CustomActionCfg.Site == null || this._cfg._configuration.CustomActionCfg.Site.length == 0) {
                    // Resolve the request
                    resolve(null);
                    return;
                }

                // Load the site custom actions
                Site().UserCustomActions().execute(siteCustomActions => {
                    // Parse the site custom actions
                    for (let i = 0; i < this._cfg._configuration.CustomActionCfg.Site.length; i++) {
                        let customAction = this._cfg._configuration.CustomActionCfg.Site[i];
                        let found = false;

                        // Parse the web custom actions
                        for (let j = 0; j < siteCustomActions.results.length; j++) {
                            // See if they match
                            if (customAction.Title == siteCustomActions.results[j].Title) {
                                // Set the flag
                                found = true;
                            }
                        }

                        // See if it was not found
                        if (!found) {
                            // Set the error
                            this._report.push("Site Custom Action '" + customAction.Title + "' is missing.");

                            // Set the flag
                            this._customActionsExist = false;
                        }
                    }

                    // Resolve the request
                    resolve(null);
                });
            });
        }

        // Return a promise
        return new Promise(resolve => {
            // Check the site & web
            Promise.all([
                checkSite(),
                checkWeb()
            ]).then(() => {
                // Resolve the request
                resolve();
            });
        });
    }

    // Lists exists
    private static _listsExist: boolean = null;
    static get ListsExist(): boolean { return this._listsExist; }

    // Checks the lists
    private static checkLists(): PromiseLike<void> {
        // Clear the flag
        this._listsExist = true;

        // Return a promise
        return new Promise((resolve) => {
            // Ensure lists exist
            if (this._cfg._configuration.ListCfg == null || this._cfg._configuration.ListCfg.length == 0) {
                // Resolve the request
                resolve(null);
                return;
            }

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

                            // Set the flag
                            this._listsExist = false;

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
            // Check the configuration
            Promise.all([
                // Check the custom actions
                this.checkCustomActions(),
                // Check the lists
                this.checkLists()
            ]).then(() => {
                // Resolve the request
                resolve(this._report.length > 0);
            });
        });
    }

    static showDialog(props: IShowDialogProps = {}) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Installation Required");

        // Set the body
        Modal.setBody(Components.Card({
            body: [
                {
                    text: props.errors && props.errors.length > 0 ?
                        "An installation is required. The following were missing in your environment:" :
                        "No errors were detected."
                },
                {
                    onRender: el => {
                        let items: Components.IListGroupItem[] = props.errors || [];

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
            ]
        }).el);

        // Render the install button
        let btnInstall: Components.IButton = null;
        Components.Tooltip({
            el: Modal.FooterElement,
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
        let btnRefresh: Components.IButton = null;
        Components.Tooltip({
            el: Modal.FooterElement,
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

        // Call the events
        props.onHeaderRendered ? props.onHeaderRendered(Modal.HeaderElement) : null;
        props.onBodyRendered ? props.onBodyRendered(Modal.BodyElement) : null;
        props.onFooterRendered ? props.onFooterRendered(Modal.FooterElement) : null;

        // Show the modal
        Modal.show();
    }
}