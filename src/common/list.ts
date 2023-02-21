import { ContextInfo, Helper, Types, Web } from "gd-sprest-bs";
import {
    ItemForm, IItemFormCreateProps, IItemFormEditProps, IItemFormViewProps,
    CanvasForm, Modal
} from "../common";

/** List Properties */
export interface IListProps {
    itemQuery?: Types.IODataQuery;
    listName: string;
    onItemsLoaded?: () => void;
    onLoadItemsError?: (...args) => void;
    onLoadFormError?: (...args) => void;
    webUrl?: string;
}

/**
 * List
 */
export class List<T = Types.SP.ListItem> {
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

    // Web Url
    private _webUrl: string = null;
    get WebUrl(): string { return this._webUrl; }

    // Items loaded event
    private _onItemsLoaded?: () => void = null;

    // Error event when loading a form
    private _onLoadFormError: (...args) => void = null;

    // Error event when loading the items fail
    private _onLoadItemsError: (...args) => void = null;

    // Constructor
    constructor(props: IListProps) {
        // Save the properties
        this._listName = props.listName;
        this._odata = props.itemQuery;
        this._onItemsLoaded = props.onItemsLoaded;
        this._onLoadFormError = props.onLoadFormError;
        this._onLoadItemsError = props.onLoadItemsError;
        this._webUrl = props.webUrl || ContextInfo.webServerRelativeUrl;

        // Load the data
        this.load().then(items => {
            // Set the items
            this._items = items;

            // Call the items loaded event
            this._onItemsLoaded ? this._onItemsLoaded() : null;
        });
    }

    // Clears the modal or canvas
    private clear() {
        // Clear the modal/form
        ItemForm.UseModal ? Modal.clear() : CanvasForm.clear();

        // Set the list name
        ItemForm.ListName = this.ListName;
    }

    // Loads the items
    load(query: Types.IODataQuery = this.OData): PromiseLike<T[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let list = Web(this.WebUrl).Lists(this.ListName);

            // See if the items exist
            if (this._items) { return this._items; }

            // Query the list content types
            list.execute(list => {
                // Save the list information
                this._listInfo = list;
            }, reject);

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
                Expand: ["FieldLinks", "Fields"]
            }).execute(views => {
                // Save the views
                this._listViews = views.results;
            }, true);

            // Query the items
            list.Items().query(query).execute(items => {
                // Save the items
                this._items = items.results as any;

                // Resolve the request
                resolve(this._items);
            }, (...args) => {
                // Call the event
                this._onLoadItemsError ? this._onLoadItemsError(...args) : null;

                // Reject the request
                reject(...args);
            }, true);
        });
    }

    // Displays the new form
    createItem(props: IItemFormCreateProps) {
        // Clear the modal/canvas
        this.clear();

        // Display the form
        ItemForm.create(props).then(null, this._onLoadFormError);
    }

    // Displays the edit form
    editItem(props: IItemFormEditProps) {
        // Clear the modal/canvas
        this.clear();

        // Display the form
        ItemForm.edit(props).then(null, this._onLoadFormError);
    }

    // Refresh the data
    refresh(query: Types.IODataQuery = this.OData): PromiseLike<T[]> {
        // Clear the items
        this._items = null;

        // Load the data
        return this.load(query);
    }

    // Displays the view form
    viewItem(props: IItemFormViewProps) {
        // Clear the modal/canvas
        this.clear();

        // Display the form
        ItemForm.view(props).then(null, this._onLoadFormError);
    }
}