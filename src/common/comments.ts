import { Components, Helper, SPTypes, Web } from "gd-sprest-bs";
import * as moment from "moment";
import { CanvasForm } from "./canvas";
import { LoadingDialog } from "./loadingDialog";

// List Name
const LIST_NAME = "Comments";

// Comments Item
export interface ICommentsItem {
    AuthorId: number;
    Author: { Id: number; EMail: string; Title: string; }
    Comment: string;
    Created: string;
    Title: string;
}

/**
 * Comments
 */
export class Comments {
    private static _elModalBody: HTMLElement = null;
    private static _elModalFooter: HTMLElement = null;
    private static _elModalHeader: HTMLElement = null;
    private static _listId: string = null;
    private static _modal: Components.IModal = null;
    private static _webUrl: string = null;

    // Comment's List Id
    private static _commentsListId: string = null;
    static get CommentsListId(): string { return this._commentsListId; }

    // Adds a comment to an item
    static add(item: any, comment: string): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Ensure a comment exists
            if (comment) {
                // Show a loading dialog
                LoadingDialog.setHeader("Adding Comment");
                LoadingDialog.setBody("This dialog will close after the comment is added.");
                LoadingDialog.show();

                // Add the item
                Web(this._webUrl).Lists(LIST_NAME).Items().add({
                    Comment: comment,
                    Title: this._listId + "|" + item.Id
                }).execute(() => {
                    // Hide the loading dialog
                    LoadingDialog.hide();

                    // Resolve the request
                    resolve();
                }, reject);
            } else {
                // Resolve the request
                resolve();
            }
        });
    }

    // Configuration
    static configuration(cfg?: Helper.ISPConfig) {
        let commentsCfg = Helper.SPConfig({
            ListCfg: [
                {
                    ListInformation: {
                        Title: LIST_NAME,
                        BaseTemplate: SPTypes.ListTemplateType.GenericList,
                        Hidden: true,
                        NoCrawl: true
                    },
                    TitleFieldDisplayName: "Comment ID",
                    TitleFieldIndexed: true,
                    CustomFields: [
                        {
                            name: "Comment",
                            title: "Comment",
                            type: Helper.SPCfgFieldType.Note,
                            noteType: SPTypes.FieldNoteType.TextOnly
                        } as Helper.IFieldInfoNote
                    ],
                    ViewInformation: [
                        {
                            ViewName: "All Items",
                            ViewQuery: "Author eq [Me]",
                            ViewFields: ["LinkTitle", "Comment", "Author", "Created"]
                        }
                    ]
                }
            ]
        });

        // Append the comments to the configuration if it exists
        if (cfg && cfg._configuration.ListCfg && cfg._configuration.ListCfg.length > 0) {
            // Append the list configuration
            cfg._configuration.ListCfg = cfg._configuration.ListCfg.concat(commentsCfg._configuration.ListCfg);
        }

        // Return the configuration
        return cfg || commentsCfg;
    }

    // Initialization of the component
    static init(listName: string, webUrl?: string): PromiseLike<void> {
        // Create the modal element
        let el = document.createElement("div");
        el.id = "comments-modal";

        // Ensure the body exists
        if (document.body) {
            // Append the element
            document.body.appendChild(el);
        } else {
            // Create an event
            window.addEventListener("load", () => {
                // Append the element
                document.body.appendChild(el);
            });
        }

        // Render the canvas
        this._modal = Components.Modal({
            el,
            type: Components.ModalTypes.Medium,
            options: {
                autoClose: false,
                backdrop: true,
                centered: true,
                keyboard: true
            },
            onRenderBody: el => { this._elModalBody = el; },
            onRenderFooter: el => { this._elModalFooter = el; },
            onRenderHeader: el => { this._elModalHeader = el.querySelector(".modal-title");; }
        });

        // Save the web url
        this._webUrl = webUrl;

        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the list id
            Web(this._webUrl).Lists(LIST_NAME).query({
                Select: ["Id"]
            }).execute(list => {
                // Set the list id
                this._commentsListId = list.Id;
            }, reject);

            // Load the List Information
            Web(this._webUrl).Lists(listName).execute(list => {
                // Set the list id
                this._listId = list.Id;

                // Resolve the request
                resolve();
            }, reject);
        });
    }

    // Loads the comments for the item
    private static load(item: any): PromiseLike<ICommentsItem[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Ensure a list id exists
            if (this._listId) {
                // Load the comments for this item
                Web(this._webUrl).Lists(LIST_NAME).Items().query({
                    Expand: ["Author"],
                    Filter: "Title eq '" + this._listId + "|" + item.Id + "'",
                    OrderBy: ["Created desc"],
                    Select: ["Id", "Comment", "Created", "Title", "Author/Id", "Author/EMail", "Author/Title"]
                }).execute(items => {
                    // Resolve the request
                    resolve(items.results as any);
                }, reject);
            } else {
                // Reject the request
                reject("The target list doesn't exist. Please initialize the component before using it.");
            }
        });
    }

    // Displays the new comment form
    static new(item: any, el: HTMLElement) {
        // Set the header
        this._elModalHeader.innerHTML = "Add Comment";

        // Clear the body and footer
        while (this._elModalBody.firstChild) { this._elModalBody.removeChild(this._elModalBody.firstChild); }
        while (this._elModalFooter.firstChild) { this._elModalFooter.removeChild(this._elModalFooter.firstChild); }

        // Create the form
        let form = Components.Form({
            el: this._elModalBody,
            controls: [{
                name: "comment",
                label: "Comment",
                type: Components.FormControlTypes.TextArea,
                rows: 10,
                required: true,
                errorMessage: "A value is required"
            } as Components.IFormControlPropsTextField]
        });

        // Set the footer
        Components.TooltipGroup({
            el: this._elModalFooter,
            tooltips: [
                {
                    content: "Click to add a comment to the request.",
                    btnProps: {
                        text: "Add",
                        type: Components.ButtonTypes.OutlineSuccess,
                        onClick: () => {
                            // Ensure the form is valid
                            if (form.isValid()) {
                                // Hide the modal
                                this._modal.hide();

                                // Add the comment
                                this.add(item, form.getValues()["comment"]).then(() => {
                                    // View the comments
                                    this.view(item, el);
                                });
                            }
                        }
                    }
                },
                {
                    content: "Closes this dialog and views the comments",
                    btnProps: {
                        text: "Close",
                        type: Components.ButtonTypes.OutlineDanger,
                        onClick: () => {
                            // Hide the modal
                            this._modal.hide();

                            // Show the canvas form
                            el ? null : CanvasForm.show();
                        }
                    }
                }
            ]
        });

        // Show the form
        this._modal.show();
    }

    // Security configuration
    static security(cfg: { principalId: number, roleDefId: number }[], webUrl?: string): PromiseLike<void> {
        // Resets the list permissions
        let resetPermissions = () => {
            return new Promise(resolve => {
                let list = Web(webUrl).Lists(LIST_NAME);

                // Reset the inheritance
                list.resetRoleInheritance().execute();

                // Clear the permissions
                list.breakRoleInheritance(false, true).execute(resolve, true);
            });
        }

        // Return a promise
        return new Promise((resolve) => {
            let list = Web(webUrl).Lists(LIST_NAME);

            // Ensure a configuration exists
            if (cfg && cfg.length > 0) {
                // Reset the list permissions
                resetPermissions().then(() => {
                    // Parse the configuration
                    for (let i = 0; i < cfg.length; i++) {
                        let principalId = cfg[i].principalId;
                        let roleDefId = cfg[i].roleDefId;

                        // Update the list settings
                        list.RoleAssignments().addRoleAssignment(principalId, roleDefId).execute(() => {
                            // Log
                            console.log("[Comments List] The user/group (" + principalId + ") was added successfully.");
                        });
                    }

                    // Wait for the requests to complete
                    list.done(resolve);
                });
            } else {
                // Resolve the request
                resolve();
            }
        });
    }

    // Displays the comments in a canvas form
    static view(item: any, el?: HTMLElement) {
        // Display a loading dialog
        LoadingDialog.setHeader("Loading Comments");
        LoadingDialog.setBody("This dialog will close after the comments are loaded.");
        LoadingDialog.show();

        // Clear the element
        while (el.firstChild) { el.removeChild(el.firstChild); }

        // Load the items
        this.load(item).then(comments => {
            // See if we are not rendering to an element
            if (el == null) {
                // Clear the form
                CanvasForm.clear();

                // Set the header
                CanvasForm.setHeader("<h5>Comments</h5>");
            }

            // Render a nav
            Components.Navbar({
                el: el || CanvasForm.BodyElement,
                onRendered: el => {
                    // Update the collapse visibility
                    el.classList.remove("navbar-expand-lg");
                    el.classList.add("navbar-expand-sm");
                },
                itemsEnd: [{
                    className: "btn-outline-light",
                    isButton: true,
                    text: "+ Add Comment",
                    onClick: () => {
                        // Hide the form
                        el ? null : CanvasForm.hide();

                        // Display a modal to add the comment
                        this.new(item, el);
                    }
                }]
            });

            // Parse the comments
            for (let i = 0; i < comments.length; i++) {
                let comment = comments[i];

                // Render a toast
                Components.Toast({
                    el: el || CanvasForm.BodyElement,
                    headerText: comment.Author.Title,
                    mutedText: moment(new Date(comment.Created)).toString(),
                    body: (comment.Comment || "").replace(/\r?\n/g, '<br />'),
                    options: {
                        animation: true,
                        autohide: false
                    }
                });
            }

            // Hide the dialog
            LoadingDialog.hide();

            // Show the canvas form
            el ? null : CanvasForm.show();
        });
    }
}