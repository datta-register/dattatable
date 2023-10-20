import { Components, Helper, Types, Web } from "gd-sprest-bs";
import { LoadingDialog } from "./loadingDialog";
import { Modal } from "./modal";

// List Security
export interface IListSecurity {
    createFl?: boolean;
    groups?: Types.SP.GroupCreationInformation[];
    listItems: IListSecurityItem[];
    onGroupCreating?: (props: Types.SP.GroupCreationInformation) => Types.SP.GroupCreationInformation;
    onGroupCreated?: (group: Types.SP.Group) => void;
    webUrl?: string;
}

// List Security Information
export interface IListSecurityItem {
    groupName: string;
    listName: string;
    permission: number | string;
}

/**
 * List Security
 * This component is used for setting the security group/permissions for a list.
 */
export class ListSecurity {
    private _ddlItems: Components.IDropdownItem[] = null;
    private _initFl: boolean = false;
    private _permissionTypes: { [key: number | string]: number } = null;
    private _props: IListSecurity = null;

    // Constructor
    constructor(props: IListSecurity) {
        // Save the properties
        this._props = props;

        // Get the permission types
        this.getPermissionTypes().then(() => {
            // See if we are creating the groups
            if (typeof (props.createFl) === "undefined" || props.createFl) {
                // Create the groups
                this.createGroups().then(() => {
                    // Set the flag
                    this._initFl = true;
                });
            } else {
                // Set the flag
                this._initFl = true;
            }
        });
    }

    // Configures a list
    private configureList(listInfo: IListSecurityItem): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Ensure a permission exists
            let permissionId = this._permissionTypes[listInfo.permission];
            if (permissionId > 0) {
                // Get the group id
                this.getGroupId(listInfo.groupName).then(
                    // Exists
                    groupId => {
                        // Add the group to the list
                        Web(this._props.webUrl).Lists(listInfo.listName).RoleAssignments().addRoleAssignment(groupId, permissionId).execute(resolve, () => {
                            // Log to the console
                            console.error(`[${listInfo.listName}] Error adding the group '${listInfo.groupName}' with permission ${listInfo.permission} to the list.`);

                            // Resolve the request
                            resolve();
                        });
                    },

                    // Error
                    () => {
                        // Log to the console
                        console.error(`[${listInfo.listName}] Site Group '${listInfo.groupName}' doesn't exist`);

                        // Resolve the request
                        resolve();
                    }
                );
            } else {
                // Resolve the request
                resolve();
            }
        });
    }

    // Configures the lists
    private configureLists(lists: IListSecurityItem[]) {
        let counter = 0;

        // Show a loading dialog
        LoadingDialog.setHeader("Configuring Lists");
        LoadingDialog.setBody("This will close after the lists are configured...");
        LoadingDialog.show();

        // Reset the lists
        this.resetPermissions(lists).then(() => {
            // Parse the lists
            Helper.Executor(lists, list => {
                // Return a promise
                return new Promise(resolve => {
                    // Update the loading dialog
                    LoadingDialog.setBody(`Configuring ${++counter} of ${lists.length} lists...`);

                    // Configure the list
                    this.configureList(list).then(resolve);
                });
            });
        });
    }

    // Creates the security groups
    private createGroups(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Parse the group names
            let groups = this._props.groups || [];
            Helper.Executor(groups, groupInfo => {
                // Return a promise
                return new Promise(resolve => {
                    // Get the group
                    this.getGroupId(groupInfo.Title).then(
                        // Group exists
                        group => {
                            // Check the next group
                            resolve(group);
                        },

                        // Group doesn't exist
                        () => {
                            // Create the properties
                            let props: Types.SP.GroupCreationInformation = {
                                Title: groupInfo.Title
                            };

                            // Call the event if it exists
                            props = this._props.onGroupCreating ? this._props.onGroupCreating(props) : props;

                            // Create the group
                            Web(this._props.webUrl).SiteGroups().add(props).execute(
                                // Successfully created the group
                                group => {
                                    // Set the group id
                                    this.setGroupId(groupInfo.Title.toLowerCase(), group);

                                    // Call the event
                                    this._props.onGroupCreated ? this._props.onGroupCreated(group) : null;

                                    // Check the next group
                                    resolve(group);
                                },

                                // Error creating the group
                                () => {
                                    // Log to the console
                                    console.error(`Site Group '${groupInfo.Title}' was unable to be created.`);

                                    // Check the next group
                                    resolve(null);
                                }
                            );
                        }
                    );
                });
            });
        });
    }

    // Gets the group id
    private _groups: { [key: string]: Types.SP.Group } = {};
    private getGroupId(groupName: string): PromiseLike<number> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let key = groupName.toLowerCase();

            // See if we have already have it
            if (this._groups[key]) {
                // Resolve the request
                resolve(this._groups[key].Id);
                return;
            }

            // Get the group information
            switch (key) {
                // Default owner's group
                case "owner":
                    Web(this._props.webUrl).AssociatedOwnerGroup().execute(group => { resolve(this.setGroupId(key, group)); }, reject);
                    break;

                // Default member's group
                case "member":
                    Web(this._props.webUrl).AssociatedMemberGroup().execute(group => { resolve(this.setGroupId(key, group)); }, reject);
                    break;

                // Default visitor's group
                case "visitor":
                    Web(this._props.webUrl).AssociatedVisitorGroup().execute(group => { resolve(this.setGroupId(key, group)); }, reject);
                    break;

                // Default
                default:
                    // Get the group
                    Web(this._props.webUrl).SiteGroups().getByName(groupName).execute(group => { resolve(this.setGroupId(key, group)); }, reject);
                    break;
            }
        });
    }
    // Saves the group information
    private setGroupId(key: string, group: Types.SP.Group) {
        // Save the info
        this._groups[key] = group;

        // Return the group id
        return group.Id;
    }

    // Render the modal
    private renderModal() {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("List Security");

        // Parse the lists
        let rows: Components.IFormRow[] = [];
        for (let i = 0; i < this._props.listItems.length; i++) {
            let listItem = this._props.listItems[i];

            // Add the row
            rows.push({
                isAutoSized: true,
                columns: [
                    {
                        control: {
                            name: "listName_" + i,
                            label: "List Name",
                            type: Components.FormControlTypes.Readonly,
                            value: listItem.listName
                        }
                    },
                    {
                        control: {
                            name: "groupName_" + i,
                            label: "Group Name",
                            type: Components.FormControlTypes.Readonly,
                            value: listItem.groupName
                        }
                    },
                    {
                        control: {
                            name: "permission_" + i,
                            label: "Permission",
                            type: Components.FormControlTypes.Dropdown,
                            items: this._ddlItems,
                            required: true,
                            value: listItem.permission
                        } as Components.IFormControlPropsDropdown
                    }
                ]
            });
        }

        // Render the form
        let form = Components.Form({
            el: Modal.BodyElement,
            rows
        });

        // Render the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            tooltips: [
                {
                    content: "Closes the modal",
                    btnProps: {
                        text: "Cancel",
                        type: Components.ButtonTypes.OutlineDanger,
                        onClick: () => {
                            // Close the modal
                            Modal.hide();
                        }
                    }
                },
                {
                    content: "Configures the security for the lists.",
                    btnProps: {
                        text: "Configure",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            let lists: IListSecurityItem[] = [];

                            // Get the form values
                            let tableData = form.getValues();

                            // Parse the # of rows
                            for (let i = 0; i < rows.length; i++) {
                                // Add the list information
                                lists.push({
                                    groupName: tableData["groupName_" + i],
                                    listName: tableData["listName_" + i],
                                    permission: tableData["permission_" + i]
                                });
                            }

                            // Configure the lists
                            this.configureLists(lists);
                        }
                    }
                }
            ]
        });

        // Show the modal
        Modal.show();
    }

    // Get the permission types
    private getPermissionTypes(): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Get the definitions
            Web().RoleDefinitions().query({
                OrderBy: ["Name"]
            }).execute(roleDefs => {
                // Clear the permissions and items
                this._ddlItems = [{ text: "No Permission" }];
                this._permissionTypes = {};

                // Parse the role definitions
                for (let i = 0; i < roleDefs.results.length; i++) {
                    let roleDef = roleDefs.results[i];

                    // Add the role and item
                    let value = roleDef.RoleTypeKind > 0 ? roleDef.RoleTypeKind : roleDef.Name;
                    this._permissionTypes[value] = roleDef.Id;
                    this._ddlItems.push({
                        data: roleDef,
                        text: roleDef.Name,
                        value: value.toString()
                    });
                }
            });
        });
    }

    // Reset list permissions
    private resetPermissions(listItems: IListSecurityItem[]) {
        // Parse the lists
        let lists = {};
        for (let i = 0; i < listItems.length; i++) {
            // Add the list
            lists[listItems[i].listName] = true;
        }

        // Return a promise
        return new Promise((resolve) => {
            // Parse the list keys
            for (let listName in lists) {
                let list = Web(this._props.webUrl).Lists(listName);

                // Reset the permissions
                list.resetRoleInheritance().execute();

                // Revert to the default
                list.breakRoleInheritance(false, true).execute(true);

                // Wait for the requests to complete
                list.done(resolve);
            }
        });
    }

    // Shows the modal
    show() {
        // Show a loading dialog
        LoadingDialog.setHeader("Loading the Security Information")
        LoadingDialog.setBody("This will close after the security information is read...");
        LoadingDialog.show();

        // Ensure the permissions are loaded
        let loopId = setInterval(() => {
            // See if the permissions exist
            if (this._ddlItems && this._permissionTypes && this._initFl) {
                // Hide the loading dialog
                LoadingDialog.hide();

                // Show the modal
                this.renderModal();

                // Stop the loop
                clearInterval(loopId);
            }
        }, 10);
    }
}