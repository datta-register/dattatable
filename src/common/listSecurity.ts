import { Components, Helper, Types, Web } from "gd-sprest-bs";
import { LoadingDialog } from "./loadingDialog";
import { Modal } from "./modal";

// List Security
export interface IListSecurity {
    groups?: Types.SP.GroupCreationInformation[];
    listItems: IListSecurityItem[];
    onGroupCreating?: (props: Types.SP.GroupCreationInformation) => Types.SP.GroupCreationInformation;
    onGroupCreated?: (group: Types.SP.Group) => void;
    onGroupsLoaded?: (groups?: { [key: string]: Types.SP.Group }) => void;
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
    private _permissionTypes: { [key: number | string]: number } = null;
    private _props: IListSecurity = null;

    // Security group information
    private _groups: { [key: string]: Types.SP.Group } = {};
    getGroup(groupName: string): Types.SP.Group { return this._groups[groupName.toLowerCase()]; }
    private setGroupId(key: string, group: Types.SP.Group) {
        // Save the info
        this._groups[key] = group;

        // Return the group id
        return group.Id;
    }

    // Security group information for users
    private _users: { [key: number]: Types.SP.Group[] }

    // Constructor
    constructor(props: IListSecurity) {
        // Save the properties
        this._props = props;

        // Get the permission types
        this.getPermissionTypes();

        // Load the groups
        this.loadGroups();
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
    private createGroups(createFl: boolean): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // See if we aren't creating the groups
            if (!createFl) {
                // Resolve the request
                resolve();
                return;
            }

            // Update the loading dialog
            LoadingDialog.setBody("Ensuring the security groups exist...");

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
                            // Update the loading dialog
                            LoadingDialog.setBody(`Creating the security group: ${groupInfo.Title}`);

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

    // Get the permission types
    private getPermissionTypes(): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Get the definitions
            Web(this._props.webUrl).RoleDefinitions().query({
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

                // Resolve the request
                resolve();
            });
        });
    }

    // Method to see if a user is w/in a group
    private isInGroup(userId: number, groupName: string): PromiseLike<boolean> {
        // Return a promise
        return new Promise(resolve => {
            // See if we have the user information
            if (this._users[userId]) {
                let isInGroup = false;

                // Parse the groups
                for (let i = 0; i < this._users[userId].length; i++) {
                    let group = this._users[userId][i];

                    // See if this is the group name
                    if (groupName == group.Title) {
                        // Set the flag
                        isInGroup = true;
                        break;
                    }
                }

                // Resolve the request
                resolve(isInGroup);
                return;
            }

            // Get the group information for this user
            Web(this._props.webUrl).SiteUsers(userId).Groups().execute(
                // Success
                groups => {
                    // Save the user information
                    this._users[userId] = groups.results;

                    // See if they are in the group
                    this.isInGroup(userId, groupName).then(resolve);
                },

                // Error
                () => {
                    // Log
                    console.error(`Unable to get the user information for id: ${userId}`);

                    // Error get the user information
                    resolve(false);
                }
            );
        });
    }

    // Loads the security groups
    private loadGroups() {
        // Parse the security groups
        Helper.Executor(this._props.groups, group => {
            // Load the group
            return this.getGroupId(group.Title);
        }).then(() => {
            // Call the event
            this._props.onGroupsLoaded ? this._props.onGroupsLoaded() : null;
        });
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
    show(createFl: boolean = true) {
        // Show a loading dialog
        LoadingDialog.setHeader("Loading the Security Information")
        LoadingDialog.setBody("Loading the security group information...");
        LoadingDialog.show();

        // Create the groups
        this.createGroups(createFl).then(() => {
            // Update the loading dialog
            LoadingDialog.setBody("Waiting for the permission types to be loaded...");

            // Ensure the permissions are loaded
            let loopId = setInterval(() => {
                // See if the permissions exist
                if (this._ddlItems && this._permissionTypes) {
                    // Hide the loading dialog
                    LoadingDialog.hide();

                    // Show the modal
                    this.renderModal();

                    // Stop the loop
                    clearInterval(loopId);
                }
            }, 10);
        });
    }
}