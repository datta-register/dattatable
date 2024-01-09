import { Components, Helper, SPTypes, Site, Web } from "gd-sprest-bs";
import { LoadingDialog } from "./loadingDialog";
import { Modal } from "./modal";

// Installation Properties
export interface IInstallProps {
    cfg: Helper.ISPConfig | Helper.ISPConfig[];
    onCompleted?: () => any;
    onError?: (cfg: Helper.ISPConfig) => void;
}

// Show Dialog Properties
export interface IShowDialogProps {
    errors?: Components.IListGroupItem[];
    onBodyRendered?(el: HTMLElement);
    onCompleted?: () => void;
    onFooterRendered?(el: HTMLElement);
    onHeaderRendered?(el: HTMLElement);
    onInstalled?: (cfg?: Helper.ISPConfig) => void;
}

/**
 * Installation Required
 * Checks the SharePoint configuration file to see if an install is required.
 */
export class InstallationRequired {
    private static _cfg: Helper.ISPConfig | Helper.ISPConfig[] = null;
    private static _report: string[] = null;
    private static _webUrl: string = null;

    // Custom Actions exists
    private static _customActionsExist: boolean = null;
    static get CustomActionsExist(): boolean { return this._customActionsExist; }

    // Checks the custom actions
    private static checkCustomActions(cfg: Helper.ISPConfig): PromiseLike<void> {
        // Clear the flag
        this._customActionsExist = true;

        // Method to check the web custom actions
        let checkWeb = () => {
            // Return a promise
            return new Promise(resolve => {
                // Ensure custom actions exist
                if (cfg._configuration.CustomActionCfg == null || cfg._configuration.CustomActionCfg.Web == null || cfg._configuration.CustomActionCfg.Web.length == 0) {
                    // Resolve the request
                    resolve(null);
                    return;
                }

                // Load the web custom actions
                let web = Web(this._webUrl);
                web.UserCustomActions().execute(webCustomActions => {
                    // Parse the web custom actions
                    for (let i = 0; i < cfg._configuration.CustomActionCfg.Web.length; i++) {
                        let customAction = cfg._configuration.CustomActionCfg.Web[i];
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
                if (cfg._configuration.CustomActionCfg == null || cfg._configuration.CustomActionCfg.Site == null || cfg._configuration.CustomActionCfg.Site.length == 0) {
                    // Resolve the request
                    resolve(null);
                    return;
                }

                // Load the site custom actions
                Site(this._webUrl).UserCustomActions().execute(siteCustomActions => {
                    // Parse the site custom actions
                    for (let i = 0; i < cfg._configuration.CustomActionCfg.Site.length; i++) {
                        let customAction = cfg._configuration.CustomActionCfg.Site[i];
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
    private static checkLists(cfg: Helper.ISPConfig): PromiseLike<void> {
        // Clear the flag
        this._listsExist = true;

        // Return a promise
        return new Promise((resolve) => {
            // Ensure lists exist
            if (cfg._configuration.ListCfg == null || cfg._configuration.ListCfg.length == 0) {
                // Resolve the request
                resolve(null);
                return;
            }

            // Parse the lists
            Helper.Executor(cfg._configuration.ListCfg, listCfg => {
                // Return a promise
                return new Promise((resolve) => {
                    // Get the list
                    Web(this._webUrl).Lists(listCfg.ListInformation.Title).query({
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

    // Determines if the user is an owner or admin
    private static checkUserPermissions(): PromiseLike<boolean> {
        // Return a promise
        return new Promise(resolve => {
            // Get the current user information
            Web(this._webUrl).CurrentUser().execute(
                // Success
                user => {
                    // Get the current user permissions
                    Web(this._webUrl).getUserEffectivePermissions(user.LoginName).execute(
                        // Success
                        permissions => {
                            // See if the user has manage web rights
                            resolve(Helper.hasPermissions(permissions.GetUserEffectivePermissions, SPTypes.BasePermissionTypes.ManageLists));
                        },

                        // Error
                        () => {
                            // Resolve the request
                            resolve(false);
                        }
                    );
                },

                // Error
                () => {
                    // Resolve the request
                    resolve(false);
                }
            );
        });
    }

    // Checks the configuration to see if an installation is required
    static requiresInstall(props: IInstallProps): PromiseLike<boolean> {
        // Save the configuration
        this._cfg = props.cfg;

        // Clear the report
        this._report = [];

        // Return a promise
        return new Promise((resolve, reject) => {
            // Check the user permissions
            this.checkUserPermissions().then(hasPermissions => {
                // See if they are not an owner or admin
                if (!hasPermissions) {
                    // Reject the request
                    reject();
                    return;
                }

                // Ensure this is an array
                let cfgs: Helper.ISPConfig[] = typeof ((this._cfg as []).length) === "number" ? this._cfg as any : [this._cfg];

                // Parse the configurations
                Helper.Executor(cfgs, cfg => {
                    // Return a promise
                    return new Promise(resolve => {
                        let numbOfErrors = this._report.length;

                        // Check the configuration
                        Promise.all([
                            // Check the custom actions
                            this.checkCustomActions(cfg),
                            // Check the lists
                            this.checkLists(cfg)
                        ]).then(() => {
                            // See if there are errors
                            if (this._report.length > numbOfErrors) {
                                // Execute the event
                                props.onError ? props.onError(cfg) : null;
                            }

                            // Execute the event and see if a promise was returned
                            let returnVal = props.onCompleted ? props.onCompleted() : null;
                            if (returnVal && typeof (returnVal["then"]) === "function") {
                                // Wait for the request to complete
                                returnVal.then(() => {
                                    // Resolve the request
                                    resolve(this._report.length > 0);
                                });
                            } else {
                                // Resolve the request
                                resolve(this._report.length > 0);
                            }
                        });
                    });
                }).then(() => {
                    // Resolve the request
                    resolve(this._report.length > 0);
                });
            });
        });
    }

    static showDialog(props: IShowDialogProps = {}) {
        // Clear the modal
        Modal.clear();

        // Ensure the configuration exists
        if (this._cfg == null || this._report == null) {
            // Show the error dialog
            this.showErrorDialog();
            return;
        }

        // Set the header
        Modal.setHeader("Installation Required");

        // Set the body
        Modal.setBody(Components.Card({
            body: [
                {
                    text: this._report.length > 0 || (props.errors && props.errors.length > 0) ?
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

                    // Ensure this is an array
                    let cfgs: Helper.ISPConfig[] = typeof ((this._cfg as []).length) === "number" ? this._cfg as any : [this._cfg];

                    // Parse the configurations
                    Helper.Executor(cfgs, cfg => {
                        // Return a promise
                        return new Promise(resolve => {
                            // Set the web url
                            this._webUrl = cfg.getWebUrl();

                            // Add the components
                            cfg.install().then(() => {
                                // Hide the button
                                btnInstall.hide();

                                // Show the refresh button
                                btnRefresh.show();

                                // Call the event
                                props.onInstalled ? props.onInstalled(cfg) : null;

                                // Check the next configuration
                                resolve(null);
                            });
                        });
                    }).then(() => {
                        // Close the dialog
                        LoadingDialog.hide();

                        // Call the event
                        props.onCompleted ? props.onCompleted() : null;
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

    private static showErrorDialog() {
        // Set the header
        Modal.setHeader("Not Configured");

        // Set the body
        Modal.setBody("The component has not been configured. The SPConfiguration definition hasn't been initialized.");

        // Show the dialog
        Modal.show();
    }
}