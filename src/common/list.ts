import { ContextInfo, Types, Web } from "gd-sprest-bs";
import {
    ItemForm, IItemFormCreateProps, IItemFormEditProps, IItemFormViewProps,
    CanvasForm, Modal, getContextInfo
} from ".";

/** List Properties */
export interface IListProps<T = Types.SP.ListItem> {
    camlQuery?: string;
    itemQuery?: Types.IODataQuery;
    listName: string;
    viewName?: string;
    onInitError?: (...args) => void;
    onInitialized?: () => void;
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

    // Error event when loading the items fail
    private _onInitError: (...args) => void = null;

    // Event triggered when the component is initialized
    private _onInitialized?: () => void = null;

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
        this._listName = props.listName;
        this._odata = props.itemQuery;
        this._onInitError = props.onInitError;
        this._onInitialized = props.onInitialized;
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
    private clear() {
        // Clear the modal/form
        ItemForm.UseModal ? Modal.clear() : CanvasForm.clear();

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

    // Displays the edit form
    editForm(props: IItemFormEditProps) {
        // Clear the modal/canvas
        this.clear();

        // Display the form
        ItemForm.edit(props).then(null, this._onLoadFormError);
    }

    // Gets a list field by internal or title
    getField(name: string) {
        // Parse the fields
        for (let i = 0; i < this.ListFields.length; i++) {
            let field = this.ListFields[i];

            // See if this is the target field
            if (field.InternalName == name || field.Title == name) {
                // Return the field
                return field;
            }
        }

        // Not found
        return null;
    }

    // Initializes the list component
    private init(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let list = Web(this.WebUrl, { requestDigest: this._requestDigest }).Lists(this.ListName);

            // Query the list content types
            list.execute(list => {
                // Save the list information
                this._listInfo = list as any;
            }, (...args) => {
                // Reject the request
                reject(...args);
            });

            // Get the root folder
            list.RootFolder().execute(folder => {
                // Set the url
                this._listUrl = folder.ServerRelativeUrl;
            }, true);

            // Query the content types
            list.ContentTypes().query({
                Expand: ["FieldLinks", "Fields"]
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
            Web(this.WebUrl, { requestDigest: this._requestDigest }).Lists(this.ListName).Items(itemId).query({
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
    private loadItems(query?: Types.IODataQuery) {
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
            Web(this.WebUrl, { requestDigest: this._requestDigest }).Lists(this.ListName).getItemsByQuery(this.CAMLQuery).execute(items => {
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
            Web(this.WebUrl, { requestDigest: this._requestDigest }).Lists(this.ListName).getItems(this.ViewXml).execute(items => {
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
        this.clear();

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
        this.clear();

        // Display the form
        ItemForm.view(props).then(null, this._onLoadFormError);
    }
}