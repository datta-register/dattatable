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
    // Form Information
    private _formInfo: Helper.IListFormResult = null;
    get FormInfo(): Helper.IListFormResult { return this._formInfo; }

    // Items
    private _items: T[] = null;
    get Items(): T[] { return this._items; }

    // List Name
    private _listName: string = null;
    get ListName(): string { return this._listName; }

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
        this._onItemsLoaded = props.onItemsLoaded;
        this._onLoadFormError = props.onLoadFormError;
        this._onLoadItemsError = props.onLoadItemsError;
        this._webUrl = props.webUrl || ContextInfo.webServerRelativeUrl;

        // Load the data
        this.load(props.itemQuery).then(items => {
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
    load(query?: Types.IODataQuery): PromiseLike<T[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // See if the items exist
            if (this._items) { return this._items; }

            // Query the items
            Web(this.WebUrl).Lists(this.ListName).Items().query(query).execute(items => {
                // Resolve the request
                resolve(items.results as any);
            }, (...args) => {
                // Call the event
                this._onLoadItemsError ? this._onLoadItemsError(...args) : null;

                // Reject the request
                reject(...args);
            });
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

    // Reset the items
    reset(query?: Types.IODataQuery): PromiseLike<T[]> {
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