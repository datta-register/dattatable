import { ContextInfo, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import {
    ItemForm, IItemFormCreateProps, IItemFormEditProps, IItemFormViewProps,
    CanvasForm, Modal, getContextInfo
} from ".";

/** List Properties */
export interface IListProps<T = Types.SP.ListItem> {
    camlQuery?: string;
    itemQuery?: Types.IODataQuery;
    listId?: string;
    listName?: string;
    viewName?: string;
    onInitError?: (...args) => void;
    onInitialized?: () => void;
    onItemLoading?: (item?: T) => void;
    onItemsLoaded?: (items?: T[]) => void;
    onLoadFormError?: (...args) => void;
    onRefreshItems?: (items?: T[]) => void | PromiseLike<any>;
    onResetForm?: () => void;
    webUrl?: string;
}

/**
 * List
 */
export class List<T = Types.SP.ListItem> {
    /** Public Properties */

    // CAML query
    private _camlQuery: string = null;
    get CAMLQuery(): string { return this._camlQuery; }

    // Flag to determine if the user can add/edit items to the list
    get CanAddEditItems(): boolean {
        return this.hasPermissions([
            SPTypes.BasePermissionTypes.AddListItems,
            SPTypes.BasePermissionTypes.EditListItems
        ]);
    }

    // Flag to determine if the user can delete items to the list
    get CanDeleteItems(): boolean {
        return this.hasPermissions(SPTypes.BasePermissionTypes.DeleteListItems);
    }

    // Flag to determine if the user can view items to the list    
    get CanViewItems(): boolean {
        return this.hasPermissions(SPTypes.BasePermissionTypes.ViewListItems);
    }

    // Reference to the edit form
    get EditForm() { return ItemForm.EditForm; }

    // Reference to the edit forms, if tabs are used
    get EditForms() { return ItemForm.EditForms; }

    // List Content Types Information
    private _listContentTypes: Types.SP.ContentTypeOData[] = null;
    get ListContentTypes(): Types.SP.ContentTypeOData[] { return this._listContentTypes; }

    // List Fields Information
    private _listFields: Types.SP.Field[] = null;
    get ListFields(): Types.SP.Field[] { return this._listFields; }

    // List Information
    private _listInfo: Types.SP.List = null;
    get ListInfo(): Types.SP.List { return this._listInfo; }

    // List Settings Url
    get ListSettingsUrl(): string { return `${this.WebUrl}/${ContextInfo.layoutsUrl}/listedit.aspx?List=${this.ListInfo.Id}`; }

    // List Url
    private _listUrl: string = null;
    get ListUrl(): string { return this._listUrl; }

    // List Views Information
    private _listViews: Types.SP.ViewOData[] = null;
    get ListViews(): Types.SP.ViewOData[] { return this._listViews; }

    // Items
    private _items: T[] = null;
    get Items(): T[] { return this._items; }

    // List Id
    private _listId: string = null;
    get ListId(): string { return this._listId; }

    // List Name
    private _listName: string = null;
    get ListName(): string { return this._listName; }

    // OData query
    private _odata: Types.IODataQuery = null;
    get OData(): Types.IODataQuery { return this._odata; }

    // Reference to the display form
    get ViewForm() { return ItemForm.DisplayForm; }

    // Reference to the display forms, if tabs are used
    get ViewForms() { return ItemForm.DisplayForms; }

    // View name
    private _viewName: string = null;
    get ViewName(): string { return this._viewName; }

    // View Xml
    private _viewXml: string = null;
    get ViewXml(): string { return this._viewXml; }

    // Web Url
    private _webUrl: string = null;
    get WebUrl(): string { return this._webUrl; }

    /** Private Properties */

    // Base Permissions
    private _basePermissions: Types.SP.BasePermissions = null;

    // Error event when loading the items fail
    private _onInitError: (...args) => void = null;

    // Event triggered when the component is initialized
    private _onInitialized?: () => void = null;

    // Items loading event
    private _onItemLoading?: (item?: T) => void = null;

    // Items loaded event
    private _onItemsLoaded?: (items?: T[]) => void = null;

    // Error event when loading a form
    private _onLoadFormError: (...args) => void = null;

    // Refresh event when refreshing the items
    private _onRefreshItems?: (items?: T[]) => void | PromiseLike<any> = null;

    // Event triggered when the form is reset
    private _onResetForm: () => void = null;

    // The request digest value of the target site
    private _requestDigest: string = null;

    // Constructor
    constructor(props: IListProps<T>) {
        // Save the properties
        this._camlQuery = props.camlQuery;
        this._listId = props.listId;
        this._listName = props.listName;
        this._odata = props.itemQuery;
        this._onInitError = props.onInitError;
        this._onInitialized = props.onInitialized;
        this._onItemLoading = props.onItemLoading;
        this._onItemsLoaded = props.onItemsLoaded;
        this._onLoadFormError = props.onLoadFormError;
        this._onRefreshItems = props.onRefreshItems;
        this._onResetForm = props.onResetForm;
        this._viewName = props.viewName;
        this._webUrl = props.webUrl || ContextInfo.webServerRelativeUrl;

        // Set the context information
        getContextInfo(props.webUrl).then((requestDigest) => {
            this._requestDigest = requestDigest;

            // Load the list information
            this.init().then(() => {
                // Load the items
                this.loadItems().then(() => {
                    // Call the event
                    this._onInitialized ? this._onInitialized() : null;

                    // Call the items loaded event
                    this._onItemsLoaded ? this._onItemsLoaded(this.Items) : null;
                }, (...args) => {
                    // Call the init error event
                    this._onInitError ? this._onInitError(...args) : null;
                });
            }, (...args) => {
                // Call the init error event
                this._onInitError ? this._onInitError(...args) : null;
            });
        }, (...args) => {
            // Call the init error event
            this._onInitError ? this._onInitError(...args) : null;
        });
    }

    // Clears the modal or canvas
    private clear(props: IItemFormCreateProps | IItemFormEditProps | IItemFormViewProps) {
        // Ensure we are not rendering to a specified element
        if (props.elForm == null) {
            // Clear the modal/form
            let useModal = typeof (props.useModal) === "boolean" ? props.useModal : ItemForm.UseModal;
            useModal ? Modal.clear() : CanvasForm.clear();
        }

        // Set the list name
        ItemForm.ListName = this.ListName;

        // Call the event
        this._onResetForm ? this._onResetForm() : null;
    }

    // Reference to the create item method
    createItem(values): PromiseLike<T> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Create the item
            this.ListInfo.Items().add(values).execute(resolve as any, reject);
        });
    }

    // Deletes the item and removes it from the internal items array
    deleteItem(itemId: number): PromiseLike<void> {
        // Parse the items
        for (let i = 0; i < this.Items.length; i++) {
            let item = this.Items[i];

            // See if this is the target item
            if (item["Id"] == itemId || item["ID"] == itemId) {
                // Remove the item
                this.Items.splice(i, 1);
                break;
            }
        }

        // Return a promise
        return new Promise((resolve, reject) => {
            // Delete the item
            this.ListInfo.Items(itemId).delete().execute(resolve, reject);
        });
    }

    // Displays the edit form
    editForm(props: IItemFormEditProps) {
        // Clear the modal/canvas
        this.clear(props);

        // Display the form
        ItemForm.edit(props).then(null, this._onLoadFormError);
    }

    // Gets the changes for a list item's version history
    getChanges(id: number, defaultFields: string[] = []): PromiseLike<any> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let changes = {};

            // Get the item
            let item = this.getItem(id);
            if (item) {
                // Get the versions for this item
                item["Versions"]().execute(versions => {
                    // Parse the versions
                    for (let i = 0; i < versions.results.length; i++) {
                        let version = versions.results[i];
                        let prevVersion = versions.results[i + 1];
                        let versionId = version.VersionLabel;

                        // Set the change object
                        changes[versionId] = {};

                        // Set the default properties
                        for (let i = 0; i < defaultFields.length; i++) {
                            let defaultField = defaultFields[i];

                            // Set the default value
                            changes[versionId][defaultField] = version[defaultField];
                        }

                        // Parse the list fields
                        for (let j = 0; j < this.ListFields.length; j++) {
                            let field = this.ListFields[j];
                            let value = version[field.InternalName] || version[field.InternalName + "Id"];
                            let prevValue = prevVersion ? prevVersion[field.InternalName] || prevVersion[field.InternalName + "Id"] : null;

                            // Skip the default properties
                            if (defaultFields.indexOf(field.InternalName) >= 0) { continue; }

                            // Skip functions
                            if (typeof (value) === "function") { continue; }

                            // Try to convert the values to a string
                            try {
                                if (JSON.stringify(value) != JSON.stringify(prevValue)) {
                                    // Append the change
                                    changes[versionId][field.InternalName] = value;
                                }
                            }
                            // Skip this property on error
                            catch { }
                        }
                    }

                    // Resolve the request
                    resolve(changes);
                }, reject);
            }
        });
    }

    // Gets a list field by internal or title
    getField(name: string): Types.SP.Field {
        let titleField = null;

        // Parse the fields
        for (let i = 0; i < this.ListFields.length; i++) {
            let field = this.ListFields[i];

            // Try to find the field by internal name first
            if (field.InternalName == name) {
                // Return the field if we match by internal
                return field;
            }
            // Else, see if the title field matches and save a reference
            else if (field.Title == name) {
                // Save the field reference
                titleField = field;
            }
        }

        // Return the title field if it was found
        return titleField;
    }

    // Gets a list field by internal or title
    getFieldById(id: string = ""): Types.SP.Field {
        let fieldId = id.replace(/{|}/g, '');

        // Parse the fields
        for (let i = 0; i < this.ListFields.length; i++) {
            let field = this.ListFields[i];

            // See if this is the target field
            if (field.Id == fieldId) {
                // Return the field if we match by internal
                return field;
            }
        }

        // Not found
        return null;
    }

    // Gets a list item by id
    getItem(id: number): T {
        // Parse the items
        for (let i = 0; i < this.Items.length; i++) {
            let item = this.Items[i];

            // See if this is the target item
            if (item["Id"] == id || item["ID"] == id) {
                // Return the item
                return item;
            }
        }

        // Not found
        return null;
    }

    // Determines if the user has permissions to the list
    hasPermissions(permissions: number | number[]): boolean {
        // See if the user has permissions
        return Helper.hasPermissions(this._basePermissions, permissions);
    }

    // Initializes the list component
    private init(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Set the web
            let web = Web(this.WebUrl, {
                disableCache: true,
                requestDigest: this._requestDigest
            });

            // Get the list
            let list = this.ListId ? web.Lists().getById(this.ListId) : web.Lists(this.ListName);

            // Query the list content types
            list.execute(list => {
                // Set the list id/name
                this._listId = list.Id;
                this._listName = list.Title;

                // Save the list information
                this._listInfo = list as any;
            }, (...args) => {
                // Reject the request
                reject(...args);
            });

            // Get the user permissions for this list
            list.query({ Expand: ["EffectiveBasePermissions"] }).execute(list => {
                this._basePermissions = list.EffectiveBasePermissions;
            });

            // Get the root folder
            list.RootFolder().execute(folder => {
                // Set the url
                this._listUrl = folder.ServerRelativeUrl;
            }, true);

            // Query the content types
            list.ContentTypes().query({
                Expand: ["FieldLinks", "Fields", "Parent"]
            }).execute(cts => {
                // Save the content types
                this._listContentTypes = cts.results;
            }, true);

            // Query the list fields
            list.Fields().execute(fields => {
                // Save the fields
                this._listFields = fields.results;
            }, true);

            // Query the list views
            list.Views().query({
                Expand: ["ViewFields"]
            }).execute(views => {
                // Save the views
                this._listViews = views.results;
            }, true);

            // Wait for the requests to complete
            list.done(() => {
                // See if the view name is specified
                if (this.ViewName) {
                    // Parse the views
                    for (let i = 0; i < this.ListViews.length; i++) {
                        let view = this.ListViews[i];

                        if (view.Title.toLowerCase() == this.ViewName.toLowerCase()) {
                            // Set the view xml
                            this._viewXml = view.ListViewXml;
                        }
                    }
                }

                // Resolve the request
                resolve();
            });
        });
    }

    // Loads the items
    private loadItem(itemId: number, query: Types.IODataQuery = this.OData): PromiseLike<T> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Query the items
            Web(this.WebUrl, {
                disableCache: true,
                requestDigest: this._requestDigest
            }).Lists(this.ListName).Items(itemId).query({
                Custom: query ? query.Custom : null,
                Expand: query ? query.Expand : null,
                Select: query ? query.Select : null
            }).execute(newItem => {
                let itemFound = false;

                // Parse the items
                for (let i = 0; i < this._items.length; i++) {
                    let item = this._items[i] as Types.SP.ListItem;

                    // See if this is the item
                    if (itemId == item.Id) {
                        // Call the event
                        this._onItemLoading ? this._onItemLoading(newItem as T) : null;

                        // Replace the item
                        this._items[i] = newItem as T;

                        // Set the flag and break from the loop
                        itemFound = true;
                        break;
                    }
                }

                // See if the item wasn't found
                if (!itemFound) {
                    // Append the item
                    this._items.push(newItem as T);
                }

                // Resolve the request
                resolve(newItem as T);
            }, (...args) => {
                // Reject the request
                reject(...args);
            }, true);
        });
    }

    // Loads the items
    private loadItems(query?: Types.IODataQuery): PromiseLike<T[]> {
        // See if the caml query exists
        if (this.CAMLQuery) {
            return this.loadItemsByCAMLQuery();
        }
        // Else, see if the view xml exists
        else if (this.ViewXml) {
            // Get the items by the view name
            return this.loadItemsByView();
        } else {
            // Get the items by odata
            return this.loadItemsByQuery(query);
        }
    }

    // Loads the items by CAML query
    private loadItemsByCAMLQuery(): PromiseLike<T[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // See if the items already exist
            if (this._items) { resolve(this._items); return; }

            // Query the items
            Web(this.WebUrl, {
                disableCache: true,
                requestDigest: this._requestDigest
            }).Lists(this.ListName).getItemsByQuery(this.CAMLQuery).execute(items => {
                // See if the event exists
                if (this._onItemLoading) {
                    // Parse the items
                    for (let i = 0; i < items.results.length; i++) {
                        // Call the event
                        this._onItemLoading(items.results[i] as T);
                    }
                }

                // Save the items
                this._items = items.results as any;

                // Resolve the request
                resolve(this._items);
            }, (...args) => {
                // Reject the request
                reject(...args);
            }, true);
        });
    }

    // Loads the items by odata query
    private loadItemsByQuery(query: Types.IODataQuery = this.OData): PromiseLike<T[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // See if the items exist
            if (this._items) { resolve(this._items); return; }

            // Query the items
            Web(this.WebUrl, { requestDigest: this._requestDigest }).Lists(this.ListName).Items().query(query).execute(items => {
                // See if the event exists
                if (this._onItemLoading) {
                    // Parse the items
                    for (let i = 0; i < items.results.length; i++) {
                        // Call the event
                        this._onItemLoading(items.results[i] as T);
                    }
                }

                // Save the items
                this._items = items.results as any;

                // Resolve the request
                resolve(this._items);
            }, (...args) => {
                // Reject the request
                reject(...args);
            }, true);
        });
    }

    // Loads the items by view name
    private loadItemsByView(): PromiseLike<T[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // See if the items already exist
            if (this._items) { resolve(this._items); return; }

            // Query the items
            Web(this.WebUrl, {
                disableCache: true,
                requestDigest: this._requestDigest
            }).Lists(this.ListName).getItems(this.ViewXml).execute(items => {
                // See if the event exists
                if (this._onItemLoading) {
                    // Parse the items
                    for (let i = 0; i < items.results.length; i++) {
                        // Call the event
                        this._onItemLoading(items.results[i] as T);
                    }
                }

                // Save the items
                this._items = items.results as any;

                // Resolve the request
                resolve(this._items);
            }, (...args) => {
                // Reject the request
                reject(...args);
            }, true);
        });
    }

    // Displays the new form
    newForm(props: IItemFormCreateProps) {
        // Clear the modal/canvas
        this.clear(props);

        // Display the form
        ItemForm.create(props).then(null, this._onLoadFormError);
    }

    // Refresh the data
    refresh(query: Types.IODataQuery = this.OData): PromiseLike<T[]> {
        // Clear the items
        this._items = null;

        // Return a promise
        return new Promise((resolve, reject) => {
            // Refresh the context information
            this.refreshContextInfo().then(() => {
                // Load the data
                this.loadItems(query).then((items) => {
                    // Call the event
                    let callback = this._onRefreshItems ? this._onRefreshItems(items) : null;
                    if (callback && typeof (callback.then) === "function") {
                        // Wait for the request to complete
                        callback.then(resolve, reject);
                    } else {
                        // Resolve the request
                        resolve(items);
                    }
                }, reject);
            }, reject);
        });
    }

    // Refreshes the context information
    private refreshContextInfo(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // See if the request digest is set
            if (this._requestDigest) {
                // Get the context information
                getContextInfo(this.WebUrl).then(requestDigest => {
                    // Save the value and resolve the request
                    this._requestDigest = requestDigest;
                    resolve();
                }, reject);
            } else {
                // Resolve the request
                resolve();
            }
        });
    }

    // Refresh the data
    refreshItem(itemId: number, query: Types.IODataQuery = this.OData): PromiseLike<T> {
        // Load the item
        return this.loadItem(itemId, query);
    }

    // Saves the item with an option to bypass validation
    save(bypassValidation?: boolean): PromiseLike<T> {
        // Save the item
        return ItemForm.save({ bypassValidation });
    }

    // Updates an item
    updateItem(itemId: number, values): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Create the item
            this.ListInfo.Items(itemId).update(values).execute(resolve, reject);
        });
    }

    // Displays the view form
    viewForm(props: IItemFormViewProps) {
        // Clear the modal/canvas
        this.clear(props);

        // Display the form
        ItemForm.view(props).then(null, this._onLoadFormError);
    }
}