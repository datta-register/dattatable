import { Components, ContextInfo, Helper, Types, Web } from "gd-sprest-bs";
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

// List Security Default Groups
export const ListSecurityDefaultGroups = {
    Members: "member",
    Owners: "owner",
    Visitors: "visitor"
}

/**
 * List Security
 * This component is used for setting the security group/permissions for a list.
 */
export class ListSecurity {
    private _ddlItems: Components.IDropdownItem[] = null;
    private _permissionTypes: { [key: number | string]: number } = null;
    private _props: IListSecurity = null;

    // Current User
    private _currentUser: Types.SP.User = null;
    get CurrentUser(): Types.SP.User { return this._currentUser; }

    // Security group information
    private _groups: { [key: string]: Types.SP.Group } = {};
    getGroup(groupName: string): Types.SP.Group { return this._groups[groupName.toLowerCase()]; }
    private setGroup(key: string, group: Types.SP.GroupOData | Types.SP.Group) {
        // Save the info
        this._groups[key] = group as Types.SP.Group;
        this._groupUsers[key] = [];

        // See if the users exist
        if ((group as Types.SP.GroupOData).Users) {
            // Set the users for this group
            this._groupUsers[key] = (group as Types.SP.GroupOData).Users.results || [];
        }
    }

    // Security group users
    private _groupUsers: { [key: string]: Types.SP.User[] } = {};
    getGroupUsers(groupName: string): Types.SP.User[] { return this._groupUsers[groupName.toLowerCase()] || []; }
    private addUser(user: Types.SP.User, groupName: string) {
        // Ensure the group exists
        let key = groupName.toLowerCase();
        if (this._groupUsers[key]) {
            // Add the user to the group
            this._groupUsers[key].push(user);
        }
    }
    private removeUser(userId: number, groupName: string) {
        // Ensure the group exists
        let key = groupName.toLowerCase();
        if (this._groupUsers[key]) {
            // Find the idx of the user
            for (let i = 0; i < this._groupUsers[key].length; i++) {
                // See if this is the user
                let user = this._groupUsers[key][i];
                if (user.Id == userId) {
                    // Remove the user
                    this._groupUsers[key].splice(i, 1);
                    break;
                }
            }
        }
    }

    // Constructor
    constructor(props: IListSecurity) {
        // Save the properties
        this._props = props;

        // Get the current user information
        this.loadCurrentUser();

        // Get the permission types
        this.getPermissionTypes();

        // Load the groups
        this.loadGroups();
    }

    // Method to check if the user is w/in a group and add them otherwise
    addToGroup(userId: number, groupName: string): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // See if the user is in the group
            if (this.isInGroup(userId, groupName)) {
                // Resolve the request
                resolve();
                return;
            }

            // Add the user to the group
            this._groups[groupName].Users().addUserById(userId).execute(user => {
                // Add the user to the group
                this.addUser(user, groupName);

                // Resolve the request
                resolve();
            }, reject);
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
        // Return a promise
        return new Promise((resolve, reject) => {
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
                }).then(resolve, reject);
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
                        () => {
                            // Check the next group
                            resolve(null);
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
                                    this.setGroup(groupInfo.Title.toLowerCase(), group);

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
            }).then(resolve, reject);
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
                case ListSecurityDefaultGroups.Owners:
                    Web(this._props.webUrl).AssociatedOwnerGroup().query({
                        Expand: ["Users"]
                    }).execute(group => { this.setGroup(key, group); resolve(group.Id); }, reject);
                    break;

                // Default member's group
                case ListSecurityDefaultGroups.Members:
                    Web(this._props.webUrl).AssociatedMemberGroup().query({
                        Expand: ["Users"]
                    }).execute(group => { this.setGroup(key, group); resolve(group.Id); }, reject);
                    break;

                // Default visitor's group
                case ListSecurityDefaultGroups.Visitors:
                    Web(this._props.webUrl).AssociatedVisitorGroup().query({
                        Expand: ["Users"]
                    }).execute(group => { this.setGroup(key, group); resolve(group.Id); }, reject);
                    break;

                // Default
                default:
                    // Get the group
                    Web(this._props.webUrl).SiteGroups().getByName(groupName).query({
                        Expand: ["Users"]
                    }).execute(group => { this.setGroup(key, group); resolve(group.Id); }, reject);
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
    isInGroup(userId: number, groupName: string): boolean {
        // See if we have the user information
        let users = this.getGroupUsers(groupName);
        for (let i = 0; i < users.length; i++) {
            // See if the user is in this group
            if (users[i].Id == userId) {
                return true;
            }
        }

        // Not in the group
        return false;
    }

    // Loads the current user
    private loadCurrentUser() {
        // Load the current user
        Web(this._props.webUrl).CurrentUser().execute(
            // Success
            user => {
                // Set the user
                this._currentUser = user;
            },
            // Error
            () => {
                // Log
                console.error("Unable to load the current user");

                // Set the information from the context info
                this._currentUser = {
                    Email: ContextInfo.userEmail,
                    Id: ContextInfo.userId,
                    LoginName: ContextInfo.userLoginName,
                    Title: ContextInfo.userDisplayName,
                    UserPrincipalName: ContextInfo.userPrincipalName
                } as any;
            }
        );
    }

    // Loads the security groups
    private loadGroups() {
        let groupNames: string[] = [];

        // Parse the security groups
        let groups = this._props.groups || [];
        for (let i = 0; i < groups.length; i++) {
            // Add the group name
            groupNames.push(groups[i].Title);
        }

        // Parse the list items
        let loadDefaultOwner = false;
        let loadDefaultMember = false;
        let loadDefaultVisitor = false;
        for (let i = 0; i < this._props.listItems.length; i++) {
            // See if this item uses a default group name
            switch (this._props.listItems[i].groupName) {
                // Members Group
                case ListSecurityDefaultGroups.Members:
                    loadDefaultMember = true;
                    break;

                // Owners Group
                case ListSecurityDefaultGroups.Owners:
                    loadDefaultOwner = true;
                    break;

                // Visitors Group
                case ListSecurityDefaultGroups.Visitors:
                    loadDefaultVisitor = true;
                    break;
            }
        }

        // Append the group names
        loadDefaultMember ? groupNames.push(ListSecurityDefaultGroups.Members) : null;
        loadDefaultOwner ? groupNames.push(ListSecurityDefaultGroups.Owners) : null;
        loadDefaultVisitor ? groupNames.push(ListSecurityDefaultGroups.Visitors) : null;

        // Parse the security groups
        Helper.Executor(groupNames, groupName => {
            // Return a promise
            return new Promise(resolve => {
                // Load the group
                this.getGroupId(groupName).then(resolve, resolve);
            });
        }).then(() => {
            // Call the event
            this._props.onGroupsLoaded ? this._props.onGroupsLoaded(this._groups) : null;
        });
    }

    // Method remove a user from a specified group
    removeFromGroup(userId: number, groupName: string): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // See if the user is not in the group
            if (!this.isInGroup(userId, groupName)) {
                // Resolve the request
                resolve();
                return;
            }

            // Remove the user from the group
            this._groups[groupName].Users().removeById(userId).execute(() => {
                // Add the user to the group
                this.removeUser(userId, groupName);

                // Resolve the request
                resolve();
            }, reject);
        });
    }

    // Render the modal
    private renderModal(onComplete: () => void) {
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

        // Gets the list names
        let getListNames = () => {
            let lists: IListSecurityItem[] = [];

            // Get the form values
            let tableData = form.getValues();

            // Parse the # of rows
            for (let i = 0; i < rows.length; i++) {
                // Add the list information
                lists.push({
                    groupName: tableData["groupName_" + i],
                    listName: tableData["listName_" + i],
                    permission: tableData["permission_" + i].value
                });
            }

            // Return the list names
            return lists;
        }

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
                    content: "Configures the list to inherit permissions from the web.",
                    btnProps: {
                        text: "Inherit",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Reset the list permissions
                            this.resetPermissions(getListNames(), false).then(onComplete);
                        }
                    }
                },
                {
                    content: "Configures the security for the lists.",
                    btnProps: {
                        text: "Configure",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Configure the lists
                            this.configureLists(getListNames()).then(onComplete);
                        }
                    }
                }
            ]
        });

        // Show the modal
        Modal.show();
    }

    // Reset list permissions
    private resetPermissions(listItems: IListSecurityItem[], breakInheritance: boolean = true) {
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
                if (breakInheritance) { list.breakRoleInheritance(false, true).execute(true); }

                // Wait for the requests to complete
                list.done(resolve);
            }
        });
    }

    // Shows the modal
    show(createFl: boolean = true, onComplete?: () => void) {
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
                    this.renderModal(() => {
                        // Hide the loading dialog
                        LoadingDialog.hide();

                        // Call the event
                        onComplete ? onComplete() : null;
                    });

                    // Stop the loop
                    clearInterval(loopId);
                }
            }, 10);
        }, () => {
            // Hide the loading dialog
            LoadingDialog.hide();

            // Show an error
            Modal.setHeader("Security Groups");
            Modal.setBody("Error loading/creating the security groups. Please contact your admin.");
            Modal.show();
        });
    }
}