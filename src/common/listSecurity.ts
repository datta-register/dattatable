import { Components, Helper, Types, Web } from "gd-sprest-bs";
import { LoadingDialog } from "./loadingDialog";
import { Modal } from "./modal";

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
    private _listItems: IListSecurityItem[] = null;
    private _permissionTypes: { [key: number | string]: number } = null;
    private _webUrl: string = null;

    // Constructor
    constructor(listItems: IListSecurityItem[], webUrl?: string) {
        // Save the properties
        this._listItems = listItems;
        this._webUrl = webUrl;

        // Get the permission types
        this.getPermissionTypes();
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
                        Web(this._webUrl).Lists(listInfo.listName).RoleAssignments().addRoleAssignment(groupId, permissionId).execute(resolve, () => {
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

    // Gets the group id
    private _groupIds: { [key: string]: number } = {};
    private getGroupId(groupName: string): PromiseLike<number> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let key = groupName.toLowerCase();

            // See if we have already have it
            if (this._groupIds[key]) {
                // Resolve the request
                resolve(this._groupIds[key]);
                return;
            }

            // Saves the group information
            let saveGroupInfo = (group: Types.SP.Group) => {
                // Save the info
                this._groupIds[key] = group.Id;

                // Resolve the request
                resolve(group.Id);
            }

            // Get the group information
            switch (key) {
                // Default owner's group
                case "owner":
                    Web(this._webUrl).AssociatedOwnerGroup().execute(saveGroupInfo, reject);
                    break;

                // Default member's group
                case "member":
                    Web(this._webUrl).AssociatedMemberGroup().execute(saveGroupInfo, reject);
                    break;

                // Default visitor's group
                case "visitor":
                    Web(this._webUrl).AssociatedVisitorGroup().execute(saveGroupInfo, reject);
                    break;

                // Default
                default:
                    // Get the group
                    Web(this._webUrl).SiteGroups().getByName(groupName).execute(saveGroupInfo, reject);
                    break;
            }
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
        for (let i = 0; i < this._listItems.length; i++) {
            let listItem = this._listItems[i];

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
                let list = Web(this._webUrl).Lists(listName);

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
        // Ensure the permissions are loaded
        let loopId = setInterval(() => {
            // See if the permissions exist
            if (this._ddlItems && this._permissionTypes) {
                // Show the modal
                this.renderModal();

                // Stop the loop
                clearInterval(loopId);
            }
        }, 10);
    }
}