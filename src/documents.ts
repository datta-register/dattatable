import { Components, ContextInfo, Helper, List, Types, Web } from "gd-sprest-bs";
import * as jQuery from "jquery";
import * as moment from "moment";
import { LoadingDialog } from "./common/loadingDialog";
import { formatBytes, formatTimeValue, getFileExt } from "./common/methods";
import { DataTable, IDataTableProps } from "./dashboard/table";
import { ItemForm, IItemFormEditProps } from "./itemForm";

/** Icons */

import { bookmarkPlus } from "gd-sprest-bs/build/icons/svgs/bookmarkPlus";
import { fileEarmark } from "gd-sprest-bs/build/icons/svgs/fileEarmark";
import { fileEarmarkArrowDown } from "gd-sprest-bs/build/icons/svgs/fileEarmarkArrowDown";
import { fileEarmarkArrowUp } from "gd-sprest-bs/build/icons/svgs/fileEarmarkArrowUp";
import { fileEarmarkBarGraph } from "gd-sprest-bs/build/icons/svgs/fileEarmarkBarGraph";
import { fileEarmarkBinary } from "gd-sprest-bs/build/icons/svgs/fileEarmarkBinary";
import { fileEarmarkCode } from "gd-sprest-bs/build/icons/svgs/fileEarmarkCode";
import { fileEarmarkExcel } from "gd-sprest-bs/build/icons/svgs/fileEarmarkExcel";
import { fileEarmarkImage } from "gd-sprest-bs/build/icons/svgs/fileEarmarkImage";
import { fileEarmarkMusic } from "gd-sprest-bs/build/icons/svgs/fileEarmarkMusic";
import { fileEarmarkPdf } from "gd-sprest-bs/build/icons/svgs/fileEarmarkPdf";
import { fileEarmarkPpt } from "gd-sprest-bs/build/icons/svgs/fileEarmarkPpt";
import { fileEarmarkPlay } from "gd-sprest-bs/build/icons/svgs/fileEarmarkPlay";
import { fileEarmarkRichtext } from "gd-sprest-bs/build/icons/svgs/fileEarmarkRichtext";
import { fileEarmarkSpreadsheet } from "gd-sprest-bs/build/icons/svgs/fileEarmarkSpreadsheet";
import { fileEarmarkText } from "gd-sprest-bs/build/icons/svgs/fileEarmarkText";
import { fileEarmarkWord } from "gd-sprest-bs/build/icons/svgs/fileEarmarkWord";
import { fileEarmarkZip } from "gd-sprest-bs/build/icons/svgs/fileEarmarkZip";
import { filterSquare } from "gd-sprest-bs/build/icons/svgs/filterSquare";
import { front } from "gd-sprest-bs/build/icons/svgs/front";
import { inputCursorText } from "gd-sprest-bs/build/icons/svgs/inputCursorText";
import { layoutTextSidebar } from "gd-sprest-bs/build/icons/svgs/layoutTextSidebar";
import { x } from "gd-sprest-bs/build/icons/svgs/x";

/** Action Button Types */

export enum ActionButtonTypes {
    Delete = 1,
    Download = 2,
    Edit = 3,
    Properties = 4,
    View = 5
}

/**
 * Properties
 */
export interface IDocumentsProps {
    canDelete?: boolean;
    canEdit?: boolean;
    canView?: boolean;
    docSetId?: number;
    el: HTMLElement;
    enableSearch?: boolean;
    listName: string;
    query?: Types.IODataQuery;
    onFilterRendered?: (el: HTMLElement) => void;
    onItemFormEditing?: {
        onCreateEditForm?: (props: Components.IListFormEditProps) => Components.IListFormEditProps;
        onFormButtonsRendering?: (buttons: Components.IButtonProps[]) => Components.IButtonProps[];
        onGetListInfo?: (props: Helper.IListFormProps) => Helper.IListFormProps;
        onSave?: (values: any) => any | PromiseLike<any>;
        onSetFooter?: (el: HTMLElement) => void;
        onSetHeader?: (el: HTMLElement) => void;
        onUpdate?: (item?: any) => void;
        onValidation?: (values?: any) => boolean | PromiseLike<boolean>;
        useModal?: boolean;
    };
    onItemFormViewing?: {
        onCreateViewForm?: (props: Components.IListFormDisplayProps) => Components.IListFormDisplayProps;
        onFormButtonsRendering?: (buttons: Components.IButtonProps[]) => Components.IButtonProps[];
        onGetListInfo?: (props: Helper.IListFormProps) => Helper.IListFormProps;
        onSetFooter?: (el: HTMLElement) => void;
        onSetHeader?: (el: HTMLElement) => void;
        useModal?: boolean;
    };
    onNavigationRendering?: (props: Components.INavbarProps) => void;
    onNavigationRendered?: (nav: Components.INavbar) => void;
    onRefresh?: () => void;
    onRendered?: () => void;
    onSearchRendered?: (el: HTMLElement) => void;
    onShowFilter?: Function;
    table?: {
        columns?: Components.ITableColumn[];
        dtProps?: any;
        onRendering?: (props?: IDataTableProps) => void;
        onRendered?: (el?: HTMLElement, dt?: any) => void;
    }
    templatesUrl?: string;
}

/**
 * Documents
 * Renders a data table containing the contents of a document library.
 */
export class Documents {
    private _el: HTMLElement = null;
    private _props: IDocumentsProps = null;
    private _tblProps: IDataTableProps = null;

    // Constructor
    constructor(props: IDocumentsProps) {
        // Save the properties
        this._props = props;

        // Default the permissions
        this._canDelete = typeof (this._props.canDelete) === "boolean" ? this._props.canDelete : true;
        this._canEdit = typeof (this._props.canEdit) === "boolean" ? this._props.canEdit : true;
        this._canView = typeof (this._props.canView) === "boolean" ? this._props.canView : true;

        // Create the element
        this._el = document.createElement("div");
        this._props.el ? this._props.el.appendChild(this._el) : null;

        // Generate the table properties
        this.generateTableProps();

        // Render the component
        this.render();
    }

    /** The data table. */
    private _dt: DataTable = null;
    get DataTable(): DataTable { return this._dt; }

    // Can delete documents
    private _canDelete = true;
    get CanDelete(): boolean { return this._canDelete; }

    // Can edit documents
    private _canEdit = true;
    get CanEdit(): boolean { return this._canEdit; }

    // Can view documents
    private _canView = true;
    get CanView(): boolean { return this._canView; }

    // The files
    get Files(): Types.SP.File[] {
        // Default the files to the root folder
        let files = this.RootFolder.Files.results

        // Parse the folders
        let formsFl = false;
        for (let i = 0; i < this.RootFolder.Folders.results.length; i++) {
            let folder: Types.SP.FolderOData = this.RootFolder.Folders.results[i] as any;

            // Skip the internal forms folder
            if (!formsFl && folder.Name == "Forms") {
                // Set the flag and skip this folder
                formsFl = true;
                continue;
            }

            // Append files
            files = files.concat(folder.Files.results);
        }

        // Return the files
        return files;
    }

    // The navigation component
    private _navbar: Components.INavbar = null;
    get Navigation(): Components.INavbar { return this._navbar; }

    // The navigation element
    get NavigationElement(): HTMLElement { return this.Navigation.el; }

    // The root folder of the library
    private _rootFolder: Types.SP.FolderOData = null;
    get RootFolder(): Types.SP.FolderOData { return this._rootFolder; }

    // The table element
    get TableElement(): HTMLElement { return this.DataTable.el; }

    // The template files
    private _templatesFiles: Types.SP.File[] = null;
    get TemplateFiles(): Types.SP.File[] { return this._templatesFiles; }

    // The template folders
    private _templateFolders: Types.SP.Folder[] = null;
    get TemplateFolders(): Types.SP.Folder[] { return this._templateFolders; }

    /**
     * Copies a file to a folder to the library
     * @param item The dropdown item containing the file/folder to copy.
     */
    private copyFile(item: Components.IDropdownItem) {
        // Show a loading dialog
        LoadingDialog.setHeader("Initializing the Transfer");
        LoadingDialog.setBody("Copying the file(s) to the workspace...");
        LoadingDialog.show();

        // See if this is a folder
        if (item.data.Files) {
            // Parse the files
            let folder = item.data as Types.SP.FolderOData;
            Helper.Executor(folder.Files.results, file => {
                // Return a promise
                return new Promise(resolve => {
                    // Get the file contents
                    Web().getFileByServerRelativeUrl(file.ServerRelativeUrl).content().execute(data => {
                        // Copy the file
                        List(this._props.listName).RootFolder().Files().add(file.Name, true, data).execute(resolve, resolve);
                    });
                });
            }).then(() => {
                // Close the dialog
                LoadingDialog.hide();

                // Refresh the page
                this.refresh();
            });
        } else {
            // Copy the file
            let file = item.data as Types.SP.File;
            file.content().execute(data => {
                // Copy the file
                List(this._props.listName).RootFolder().Files().add(file.Name, true, data).execute(() => {
                    // Close the dialog
                    LoadingDialog.hide();

                    // Refresh the page
                    this.refresh();
                });
            });
        }
    }

    // Generates the template files/folders dropdown items
    private generateItems() {
        let items: Components.IDropdownItem[] = [];

        // Parse the template folders
        for (let i = 0; i < this.TemplateFolders.length; i++) {
            let folder = this.TemplateFolders[i];

            // Skip the internal forms folder
            if (folder.Name == "Forms") { continue; }

            // Add a dropdown item
            items.push({
                text: folder.Name,
                data: folder,
                value: folder.ServerRelativeUrl,
                onClick: item => { this.copyFile(item); }
            });
        }

        // Parse the template files
        for (let i = 0; i < this.TemplateFiles.length; i++) {
            let file = this.TemplateFiles[i];

            // Add a dropdown item
            items.push({
                text: file.Name,
                data: file,
                value: file.ServerRelativeUrl,
                onClick: item => { this.copyFile(item); }
            });
        }

        // Return the dropdown items
        return items;
    }

    // Generates the table properties
    private generateTableProps() {
        // Default the indexes for the type and action buttons
        let idxType = 0;
        let idxActions = 8;

        // See if custom columns are defined
        if (this._props.table && this._props.table.columns) {
            // Clear the indexes
            idxActions = -1;
            idxType = -1;

            // Parse the columns
            for (let i = 0; i < this._props.table.columns.length; i++) {
                // Update the index based on the value
                switch (this._props.table.columns[i].name) {
                    // Actions
                    case "Actions":
                        idxActions = i;
                        break;

                    // Type
                    case "Type":
                        idxType = i;
                        break;
                }
            }
        }

        // See if the column definitions are not defined
        let columnDefs = this._props.table ? this._props.table.dtProps : null;
        if (columnDefs == null) {
            columnDefs = [];

            // Add the default options for the Actions and Type
            idxActions >= 0 ? columnDefs.push({ targets: idxActions, orderable: false, searchable: false }) : null;
            idxType >= 0 ? columnDefs.push({ targets: idxType, searchable: false }) : null;
        }

        // Create the table properties
        this._tblProps = {
            el: null,
            dtProps: this._props.table && this._props.table.dtProps ? this._props.table.dtProps : {
                dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
                columnDefs,
                createdRow: function (row, data, index) {
                    jQuery('td', row).addClass('align-middle');
                },
                // Add some classes to the dataTable elements
                drawCallback: function (settings) {
                    let api = new jQuery.fn.dataTable.Api(settings) as any;
                    jQuery(api.context[0].nTable).removeClass('no-footer');
                    jQuery(api.context[0].nTable).addClass('tbl-footer');
                    jQuery(api.context[0].nTable).addClass('table-striped');
                    jQuery(api.context[0].nTableWrapper).find('.dataTables_info').addClass('text-center');
                    jQuery(api.context[0].nTableWrapper).find('.dataTables_length').addClass('pt-2');
                    jQuery(api.context[0].nTableWrapper).find('.dataTables_paginate').addClass('pt-03');
                },
                headerCallback: function (thead, data, start, end, display) {
                    jQuery('th', thead).addClass('align-middle');
                },
                // Order by the 1st column by default; ascending
                order: [[1, "asc"]]
            },
            columns: this._props.table && this._props.table.columns ? this._props.table.columns : [
                {
                    name: "Type",
                    title: "Type",
                },
                {
                    name: "Name",
                    title: "Name"
                },
                {
                    name: "Title",
                    title: "Description"
                },
                {
                    name: "FileSize",
                    title: "File Size"
                },
                {
                    name: "Created",
                    title: "Created"
                },
                {
                    name: "Author",
                    title: "Created By"
                },
                {
                    name: "Modified",
                    title: "Modified"
                },
                {
                    name: "ModifiedBy",
                    title: "Modified By"
                },
                {
                    className: "text-end text-nowrap",
                    name: "Actions",
                    title: ""
                }
            ]
        };

        // Parse the columns
        Helper.Executor(this._tblProps.columns, col => {
            let customEvent = col.onRenderCell;

            // See if this is the type column
            if (col.name == "Type") {
                // Set the event to render an icon
                col.onRenderCell = (el, col, file: Types.SP.File) => {
                    // Render the file
                    this.renderFileIcon(el, file);

                    // Set the filter/sort value
                    el.setAttribute("data-filter", getFileExt(file.Name));
                    el.setAttribute("data-sort", getFileExt(file.Name));

                    // Call the custom event
                    customEvent ? customEvent(el, col, file) : null;
                };
            }
            // Else, see if this is the file size
            else if (col.name == "FileSize") {
                // Set the event to render the size
                col.onRenderCell = (el, col, file: Types.SP.File) => {
                    // Render the file size value
                    el.innerHTML = formatBytes(file.Length);

                    // Set the sort value
                    el.setAttribute("data-sort", file.Length.toString());

                    // Call the custom event
                    customEvent ? customEvent(el, col, file) : null;
                }
            }
            // Else, see if this is a date/time field
            else if (col.name == "Created" || col.name == "Modified") {
                // Set the event to render the size
                col.onRenderCell = (el, col, file: Types.SP.File) => {
                    // Render the date/time value
                    let value = col.name == "Created" ? file.TimeCreated : file.TimeLastModified;
                    el.innerHTML = formatTimeValue(value);

                    // Set the date/time filter/sort values
                    el.setAttribute("data-filter", moment(value).format("dddd MMMM DD YYYY"));
                    el.setAttribute("data-sort", value);

                    // Call the custom event
                    customEvent ? customEvent(el, col, file) : null;
                }
            }
            // Else, see if this is a user field
            else if (col.name == "Author" || col.name == "ModifiedBy") {
                // Set the event to render the size
                col.onRenderCell = (el, col, file: Types.SP.File) => {
                    // Render the Person field Title
                    el.innerHTML = (file[col.name] ? file[col.name]["Title"] : null) || "";

                    // Call the custom event
                    customEvent ? customEvent(el, col, file) : null;
                }
            }
            // Else, see if this is the "actions" buttons
            else if (col.name == "Actions") {
                // Set the event to render the size
                col.onRenderCell = (el, col, file: Types.SP.File) => {
                    // Render the action buttons
                    this.renderActionButtons(el, file);

                    // Call the custom event
                    customEvent ? customEvent(el, col, file) : null;
                }
            } else {
                // Set the default render event
                col.onRenderCell = (el, col, file: Types.SP.FileOData) => {
                    // Set the value
                    el.innerHTML = file[col.name] || file.ListItemAllFields.FieldValuesAsText[col.name] || "";

                    // Call the custom event
                    customEvent ? customEvent(el, col, file) : null;
                }
            }
        });
    }

    // Determines if the document can be viewed in office online servers
    private isWopi(file: Types.SP.File) {
        switch (getFileExt(file.Name)) {
            // Excel
            case "csv":
            case "doc":
            case "docx":
            case "dot":
            case "dotx":
            case "pot":
            case "potx":
            case "pps":
            case "ppsx":
            case "ppt":
            case "pptx":
            case "xls":
            case "xlsx":
            case "xlt":
            case "xltx":
                return true;
            // Default
            default: {
                return false;
            }
        }
    }

    // Loads the data
    private load(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let web = Web();

            // Clear the properties
            this._rootFolder = null;
            this._templateFolders = null;
            this._templatesFiles = null;

            // See if the templates library was set
            if (this._props.templatesUrl) {
                // Load the files and folders
                web.getFolderByServerRelativeUrl(this._props.templatesUrl).query({
                    Expand: ["Folders/Files", "Files"]
                }).execute(folder => {
                    // Set the template files
                    this._templatesFiles = folder.Files.results;

                    // Set the template folders
                    this._templateFolders = folder.Folders.results;
                }, reject);
            }

            // Set the query
            let query: Types.IODataQuery = this._props.query || {};
            query.Expand = (query.Expand ? query.Expand : []).concat([
                "Folders/Files", "Folders/Files/Author",
                "Folders/Files/ListItemAllFields", "Folders/Files/ListItemAllFields/FieldValuesAsText",
                "Folders/Files/ModifiedBy", "Files", "Files/Author", "Files/ListItemAllFields",
                "Files/ListItemAllFields/FieldValuesAsText", "Files/ModifiedBy"
            ]);

            // See if we are targeting a document set folder
            if (this._props.docSetId) {
                web.Lists(this._props.listName).Items(this._props.docSetId).Folder().query(query).execute(folder => {
                    // Set the root folder
                    this._rootFolder = folder;
                }, reject);
            } else {
                // Load library information
                web.Lists(this._props.listName).RootFolder().query(query).execute(folder => {
                    // Set the root folder
                    this._rootFolder = folder;
                }, reject);
            }

            // Wait for the requests to complete
            web.done(() => {
                // Resolve the request
                resolve();
            });
        });
    }

    // Renders the component
    private render(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the data
            this.load().then(() => {
                // Render the navigation
                this.renderNavigation();

                // Render the table
                this.renderTable();

                // Call the render event
                this._props.onRendered ? this._props.onRendered() : null;

                // Resolve the promise
                resolve();
            }, reject);
        });
    }

    // Renders the file actions
    private renderActionButtons(el: HTMLElement, file: Types.SP.File) {
        // Create a span to wrap the icons in
        let span = document.createElement("span");
        span.className = "bg-white d-inline-flex ms-2 rounded";
        let spanEdit = span.cloneNode() as HTMLSpanElement;
        let spanProps = span.cloneNode() as HTMLSpanElement;
        let spanDownload = span.cloneNode() as HTMLSpanElement;
        let spanDel = span.cloneNode() as HTMLSpanElement;
        spanDel.classList.add("me-1");

        // Add the icons
        el.appendChild(span);
        el.appendChild(spanEdit);
        el.appendChild(spanProps);
        el.appendChild(spanDownload);
        el.appendChild(spanDel);

        // Render the buttons
        span.appendChild(this.generateButton(ActionButtonTypes.View, file));
        spanEdit.appendChild(this.generateButton(ActionButtonTypes.Edit, file));
        spanProps.appendChild(this.generateButton(ActionButtonTypes.Properties, file));
        spanDownload.appendChild(this.generateButton(ActionButtonTypes.Download, file));
        spanDel.appendChild(this.generateButton(ActionButtonTypes.Delete, file));
    }

    private generateButton(btnType: number, file: Types.SP.File): HTMLElement {
        let isWopi = this.isWopi(file);

        // Render the button based on the type
        switch (btnType) {
            case ActionButtonTypes.Delete:
                return Components.Tooltip({
                    content: "Delete",
                    btnProps: {
                        // Render the icon button
                        className: "p-1",
                        iconType: x,
                        iconSize: 24,
                        isDisabled: !this.CanDelete,
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            if (this.CanDelete) {
                                // Confirm we want to delete the item
                                if (confirm("Are you sure you want to delete this document?")) {
                                    // Display a loading dialog

                                    LoadingDialog.setHeader("Deleting Document");
                                    LoadingDialog.setBody("Deleting Document: " + file.Name + ". This will close afterwards.");
                                    LoadingDialog.show();
                                    // Delete the document

                                    Web().getFileByServerRelativeUrl(file.ServerRelativeUrl).delete().execute(
                                        // Success
                                        () => {
                                            // close dialog
                                            LoadingDialog.hide();
                                            // Refresh the page                         
                                            this.refresh();
                                        },
                                        // Error
                                        err => {
                                            // TODO
                                        }
                                    );
                                }
                            }
                        }
                    },
                }).el;

            case ActionButtonTypes.Download:
                return Components.Tooltip({
                    content: "Download",
                    btnProps: {
                        // Render the icon button
                        className: "p-1",
                        iconType: fileEarmarkArrowDown,
                        iconSize: 24,
                        isDisabled: !this.CanView,
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            if (this.CanView) {
                                window.open(ContextInfo.webServerRelativeUrl + "/_layouts/15/download.aspx?SourceUrl=" + file.ServerRelativeUrl, "_blank");
                            }
                        }
                    },
                }).el;

            case ActionButtonTypes.Edit:
                return Components.Tooltip({
                    content: "Edit",
                    btnProps: {
                        // Render the icon button
                        className: "p-1",
                        iconType: inputCursorText,
                        iconSize: 24,
                        isDisabled: (!isWopi || !this.CanEdit),
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            if (isWopi && this.CanEdit) {
                                // Open the file in a new tab
                                window.open(ContextInfo.webServerRelativeUrl + "/_layouts/15/WopiFrame.aspx?sourcedoc=" + file.ServerRelativeUrl + "&action=edit");
                            }
                        }
                    },
                }).el;

            case ActionButtonTypes.Properties:
                return Components.Tooltip({
                    content: "Properties",
                    btnProps: {
                        // Render the icon button
                        className: "p-1",
                        iconType: layoutTextSidebar,
                        iconSize: 24,
                        isDisabled: !this.CanEdit && !this.CanView,
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            // Set the item form properties
                            ItemForm.ListName = this._props.listName;
                            ItemForm.UseModal = false;

                            // Ensure the user can edit the item
                            if (this.CanEdit) {
                                // Define the properties
                                let editProps: IItemFormEditProps = {
                                    ...(this._props.onItemFormEditing || {}),
                                    ...{
                                        // Set the item id
                                        itemId: file.ListItemAllFields["Id"],

                                        // Set the edit form properties
                                        onCreateEditForm: props => {
                                            // Set the rendering event
                                            props.onControlRendering = (ctrl, field) => {
                                                if (field.InternalName == "FileLeafRef") {
                                                    // Validate the name of the file
                                                    ctrl.onValidate = (ctrl, results) => {
                                                        // Ensure the value is less than 128 characters
                                                        if (results.value?.length > 128) {
                                                            // Return an error message
                                                            results.invalidMessage = "The file name must be less than 128 characters.";
                                                            results.isValid = false;
                                                        }

                                                        // Return the results
                                                        return results;
                                                    }
                                                }
                                            }

                                            // See if a custom event exists
                                            if (this._props.onItemFormEditing && this._props.onItemFormEditing.onCreateEditForm) {
                                                // Return the properties
                                                return this._props.onItemFormEditing.onCreateEditForm(props);
                                            }

                                            // Return the properties
                                            return props;
                                        },

                                        // Update the footer
                                        onSetFooter: (el) => {
                                            let updateBtn = el.querySelector('[role="group"]').firstChild as HTMLButtonElement;
                                            updateBtn.classList.remove("btn-outline-primary");
                                            updateBtn.classList.add("btn-primary");

                                            // See if a custom event exists
                                            if (this._props.onItemFormEditing && this._props.onItemFormEditing.onSetFooter) {
                                                // Execute the event
                                                this._props.onItemFormEditing.onSetFooter(el);
                                            }
                                        },

                                        // Update the header
                                        onSetHeader: (el) => {
                                            // Update the header
                                            el.querySelector("h5").innerHTML = "Properties";

                                            // See if a custom event exists
                                            if (this._props.onItemFormEditing && this._props.onItemFormEditing.onSetHeader) {
                                                // Execute the event
                                                this._props.onItemFormEditing.onSetHeader(el);
                                            }
                                        },

                                        // Refresh the view when updates occur
                                        onUpdate: (item) => {
                                            // See if a custom event exists
                                            if (this._props.onItemFormEditing && this._props.onItemFormEditing.onUpdate) {
                                                // Execute the event
                                                this._props.onItemFormEditing.onUpdate(item);
                                            } else {
                                                // Refresh the data table
                                                this.refresh();
                                            }
                                        }
                                    }
                                };

                                // Show the edit form
                                ItemForm.edit(editProps);
                            } else {
                                // View the form
                                ItemForm.view({
                                    ...(this._props.onItemFormViewing || {}),
                                    ...{
                                        itemId: file.ListItemAllFields["Id"]
                                    }
                                });
                            }
                        }
                    }
                }).el;

            case ActionButtonTypes.View:
                return Components.Tooltip({
                    content: "View",
                    btnProps: {
                        // Render the icon button
                        className: "img-flip-x p-1",
                        iconType: front,
                        iconSize: 24,
                        isDisabled: !this.CanView,
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            if (this.CanView) {
                                // Open the file in a new tab
                                window.open(isWopi ? ContextInfo.webServerRelativeUrl + "/_layouts/15/WopiFrame.aspx?sourcedoc=" + file.ServerRelativeUrl + "&action=view" : file.ServerRelativeUrl, "_blank");
                            }
                        }
                    },
                }).el;
        }
    }

    // Renders the file icon
    private renderFileIcon(el: HTMLElement, file: Types.SP.File) {
        // Render the icon wrapper
        let span = document.createElement("span");
        span.className = "text-muted";
        el.appendChild(span);

        // Render the icon
        let size = 28;
        switch (getFileExt(file.Name)) {
            // Power BI
            case "pbix":
                span.appendChild(fileEarmarkBarGraph(size));
                span.title = "Power BI Report";
                break;
            // Binary
            case "bin":
            case "blg":
            case "dat":
            case "dmg":
            case "dmp":
            case "log":
            case "pbi":
                span.appendChild(fileEarmarkBinary(size));
                span.title = "Binary File";
                break;
            // Code
            case "asp":
            case "aspx":
            case "css":
            case "hta":
            case "htm":
            case "html":
            case "js":
            case "json":
            case "mht":
            case "mhtml":
            case "scss":
            case "xml":
            case "yaml":
                span.appendChild(fileEarmarkCode(size));
                span.title = "Code";
                break;
            // Excel
            case "csv":
            case "ods":
            case "xls":
            case "xlsx":
            case "xlt":
            case "xltx":
                span.appendChild(fileEarmarkExcel(size));
                span.title = "Excel Spreadsheet";
                break;
            // Image
            case "ai":
            case "bmp":
            case "eps":
            case "gif":
            case "heic":
            case "heif":
            case "jpe":
            case "jpeg":
            case "jpg":
            case "png":
            case "psd":
            case "svg":
            case "tif":
            case "tiff":
            case "webp":
                span.appendChild(fileEarmarkImage(size));
                span.title = "Image";
                break;
            // Audio
            case "aac":
            case "aiff":
            case "alac":
            case "flac":
            case "m4a":
            case "m4p":
            case "mka":
            case "mp3":
            case "mp4a":
            case "wav":
            case "wma":
                span.appendChild(fileEarmarkMusic(size));
                span.title = "Audio File";
                break;
            // PDF
            case "pdf":
                span.appendChild(fileEarmarkPdf(size));
                span.title = "Adobe PDF";
                break;
            // PowerPoint
            case "odp":
            case "pot":
            case "potx":
            case "pps":
            case "ppsx":
            case "ppt":
            case "pptx":
                span.appendChild(fileEarmarkPpt(size));
                span.title = "PowerPoint Presentation";
                break;
            // Media
            case "avi":
            case "flv":
            case "m2ts":
            case "mkv":
            case "mov":
            case "m4p":
            case "m4v":
            case "mp4":
            case "mpe":
            case "mpeg":
            case "mpg":
            case "mpv":
            case "qt":
            case "ts":
            case "vob":
            case "webm":
            case "wmv":
                span.appendChild(fileEarmarkPlay(size));
                span.title = "Media File";
                break;
            // Rich Text
            case "rtf":
                span.appendChild(fileEarmarkRichtext(size));
                span.title = "Rich Text";
                break;
            // Database
            case "ldf":
            case "mdb":
            case "mdf":
                span.appendChild(fileEarmarkSpreadsheet(size));
                span.title = "Database";
                break;
            // Text
            case "md":
            case "text":
            case "txt":
                span.appendChild(fileEarmarkText(size));
                span.title = "Text";
                break;
            // Word
            case "doc":
            case "docx":
            case "dot":
            case "dotx":
            case "odt":
            case "wpd":
                span.appendChild(fileEarmarkWord(size));
                span.title = "Word Document";
                break;
            // Compressed
            case "7z":
            case "cab":
            case "gz":
            case "iso":
            case "rar":
            case "tgz":
            case "zip":
                span.appendChild(fileEarmarkZip(size));
                span.title = "Compressed Folder";
                break;
            // Default
            default: {
                span.appendChild(fileEarmark(size));
                span.title = "File";
            }
        }
    }

    // Renders the navigation
    private renderNavigation() {
        let itemsEnd: Components.INavbarItem[] = [];

        // See if templates exist
        if (this._props.templatesUrl) {
            // Add the item
            itemsEnd.push({
                text: "Templates",
                className: "btn btn-sm btn-outline-secondary",
                iconSize: 20,
                iconType: bookmarkPlus,
                isButton: true,
                items: this.generateItems(),
                onRender: (el) => {
                    el.classList.add("bg-white");
                    el.querySelector("svg").style.margin = "0 0.25rem 0.1rem -0.25rem";
                },
                onMenuRendering: props => {
                    // Update the placement
                    props.options.offset = [7, 0];

                    // Return the properties
                    return props;
                }
            });
        }

        // Add the upload button
        itemsEnd.push({
            text: "Upload",
            onRender: (el, item) => {
                // Clear the existing button
                el.innerHTML = "";
                // Create a span to wrap the icon in
                let span = document.createElement("span");
                span.className = "bg-white d-inline-flex ms-2 rounded";
                el.appendChild(span);

                // Render a tooltip
                Components.Tooltip({
                    el: span,
                    content: item.text,
                    btnProps: {
                        // Render the icon button
                        className: "p-1",
                        iconType: fileEarmarkArrowUp,
                        iconSize: 24,
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            // Show the file upload dialog
                            Helper.ListForm.showFileDialog().then(fileInfo => {
                                // Show a loading dialog
                                LoadingDialog.setHeader("Uploading File");
                                LoadingDialog.setBody("Saving the file you selected. Please wait...");
                                LoadingDialog.show();
                                // Upload the file to the objective folder
                                List(this._props.listName).RootFolder().Files().add(fileInfo.name, true, fileInfo.data).execute(
                                    // Success
                                    file => {
                                        // Hide the dialog
                                        LoadingDialog.hide();

                                        // Refresh the page
                                        this.refresh();
                                    },

                                    // Error
                                    err => {
                                        // Hide the dialog
                                        LoadingDialog.hide();
                                    }
                                )
                            });
                        }
                    },
                });
            }
        });

        // Set the default properties
        let navProps: Components.INavbarProps = {
            el: this._el,
            brand: "Documents View",
            itemsEnd,
            searchBox: this._props.enableSearch ? {
                hideButton: true,
                onChange: value => {
                    // Search the data table
                    this._dt.search(value);
                },
                onSearch: value => {
                    // Search the data table
                    this._dt.search(value);
                }
            } : null
        };

        // Call the rendering event
        this._props.onNavigationRendering ? this._props.onNavigationRendering(navProps) : null;

        // Create the navbar
        this._navbar = Components.Navbar(navProps);

        /* Fix the padding on the left & right of the nav */
        this._navbar.el.querySelector("div.container-fluid").classList.add("ps-75");
        this._navbar.el.querySelector("div.container-fluid").classList.add("pe-2");

        // See if we are showing the filter
        if (this._props.onShowFilter) {
            // Render the filter icon
            let icon = document.createElement("div");
            icon.classList.add("filter-icon");
            icon.classList.add("nav-link");
            icon.classList.add("text-dark");
            icon.style.cursor = "pointer";
            icon.appendChild(filterSquare());
            this._props.onShowFilter ? icon.addEventListener("click", this._props.onShowFilter as any) : null;
            this._navbar.el.firstElementChild.appendChild(icon);

            // Call the render event
            this._props.onFilterRendered ? this._props.onFilterRendered(icon) : null;
        }

        // Call the render events
        this._props.onSearchRendered ? this._props.onSearchRendered(this._navbar.el.querySelector("input[type='search']")) : null;
        this._props.onNavigationRendered ? this._props.onNavigationRendered(this.Navigation) : null;
    }

    // Renders the datatable with the file information
    private renderTable() {
        // Create the element
        let el = document.createElement("div");
        this._el.appendChild(el);
        this._tblProps.el = el;

        // Set the data
        this._tblProps.rows = this.Files;

        // Call the rendering event
        this._props.table && this._props.table.onRendering ? this._props.table.onRendering(this._tblProps) : null;

        // Render the table
        this._dt = new DataTable(this._tblProps);

        // Call the rendered event
        this._tblProps.onRendered ? this._tblProps.onRendered(el, this._dt.datatable) : null;
    }

    /** Public Methods */

    // Generates an action button
    generateActionButton(btnType: number, file: Types.SP.File): HTMLElement {
        // Return the button
        return this.generateButton(btnType, file);
    }

    // Refreshes the documents
    refresh() {
        // Show a loading dialog
        LoadingDialog.setHeader("Reloading Workspace");
        LoadingDialog.setBody("Reloading the workspace data. This will close afterwards.");
        LoadingDialog.show();

        // Clear the element
        while (this._el.firstChild) { this._el.removeChild(this._el.firstChild); }

        // Render the component
        this.render().then(() => {
            // Call the event
            this._props.onRefresh ? this._props.onRefresh() : null;

            // Hide the dialog
            LoadingDialog.hide();
        });
    }

    // Searches the data table
    search(value: string) {
        // Search the table data
        this._dt.search(value);
    }
}