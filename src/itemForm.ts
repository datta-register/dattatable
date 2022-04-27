import { Components, Helper, SPTypes } from "gd-sprest-bs";
import { CanvasForm, LoadingDialog, Modal } from "./common";

/** Create Item Properties */
export interface IItemFormCreateProps {
    info?: Helper.IListFormResult;
    onCreateEditForm?: (props: Components.IListFormEditProps) => Components.IListFormEditProps;
    onFormButtonsRendering?: (buttons: Components.IButtonProps[]) => Components.IButtonProps[];
    onGetListInfo?: (props: Helper.IListFormProps) => Helper.IListFormProps;
    onSave?: (values: any) => any | PromiseLike<any>;
    onSetFooter?: (el: HTMLElement) => void;
    onSetHeader?: (el: HTMLElement) => void;
    onUpdate?: (item?: any) => void;
    onValidation?: (values?: any) => boolean | PromiseLike<boolean>;
    useModal?: boolean;
    webUrl?: string;
}

/** Edit Item Properties */
export interface IItemFormEditProps {
    info?: Helper.IListFormResult;
    itemId: number;
    onCreateEditForm?: (props: Components.IListFormEditProps) => Components.IListFormEditProps;
    onFormButtonsRendering?: (buttons: Components.IButtonProps[]) => Components.IButtonProps[];
    onGetListInfo?: (props: Helper.IListFormProps) => Helper.IListFormProps;
    onSave?: (values: any) => any | PromiseLike<any>;
    onSetFooter?: (el: HTMLElement) => void;
    onSetHeader?: (el: HTMLElement) => void;
    onUpdate?: (item?: any) => void;
    onValidation?: (values?: any) => boolean | PromiseLike<boolean>;
    useModal?: boolean;
    webUrl?: string;
}

/** View Item Properties */
export interface IItemFormViewProps {
    info?: Helper.IListFormResult;
    itemId: number;
    onCreateViewForm?: (props: Components.IListFormDisplayProps) => Components.IListFormDisplayProps;
    onFormButtonsRendering?: (buttons: Components.IButtonProps[]) => Components.IButtonProps[];
    onGetListInfo?: (props: Helper.IListFormProps) => Helper.IListFormProps;
    onSetFooter?: (el: HTMLElement) => void;
    onSetHeader?: (el: HTMLElement) => void;
    useModal?: boolean;
    webUrl?: string;
}

/**
 * Item Form
 */
export class ItemForm {
    private static _onCreateEditForm: (props: Components.IListFormEditProps) => Components.IListFormEditProps = null;
    private static _onCreateViewForm: (props: Components.IListFormDisplayProps) => Components.IListFormDisplayProps = null;
    private static _onFormButtonsRendering: (buttons: Components.IButtonProps[]) => Components.IButtonProps[] = null;
    private static _onGetListInfo: (props: Helper.IListFormProps) => Helper.IListFormProps = null;
    private static _onSetFooter?: (el: HTMLElement) => void = null;
    private static _onSetHeader?: (el: HTMLElement) => void = null;
    private static _onSave: (values: any) => any | PromiseLike<any> = null;
    private static _onValidation: (values?: any) => boolean | PromiseLike<boolean> = null;
    private static _updateEvent: Function = null;

    // Auto Close Flag
    static set AutoClose(value: boolean) {
        // Update the flag
        this.UseModal ? Modal.setAutoClose(value) : CanvasForm.setAutoClose(value);
    }

    // Display Form
    private static _displayForm: Components.IListFormDisplay = null;
    static get DisplayForm(): Components.IListFormDisplay { return this._displayForm; }

    // Edit Form
    private static _editForm: Components.IListFormEdit = null;
    static get EditForm(): Components.IListFormEdit { return this._editForm; }

    // Form Information
    private static _info = null;
    static get FormInfo(): Helper.IListFormResult { return this._info; }

    // Form Modes
    private static _controlMode: number = null;
    static get IsDisplay(): boolean { return this._controlMode == SPTypes.ControlMode.Display; }
    static get IsEdit(): boolean { return this._controlMode == SPTypes.ControlMode.Edit; }
    static get IsNew(): boolean { return this._controlMode == SPTypes.ControlMode.New; }

    // List name
    private static _listName: string = null;
    static get ListName(): string { return this._listName; }
    static set ListName(value: string) { this._listName = value; }

    // Flag to use a modal or canvas (default)
    private static _useModal: boolean = false;
    static get UseModal(): boolean { return this._useModal; }
    static set UseModal(value: boolean) { this._useModal = value; }

    /** Public Methods */

    // Closes the item form
    static close() {
        this._useModal ? Modal.hide() : CanvasForm.hide();
    }

    // Creates a new task
    static create(props: IItemFormCreateProps = {}) {
        // Set the properties
        this._controlMode = SPTypes.ControlMode.New;
        this._info = props.info;
        this._onCreateEditForm = props.onCreateEditForm;
        this._onFormButtonsRendering = props.onFormButtonsRendering;
        this._onGetListInfo = props.onGetListInfo;
        this._onSave = props.onSave;
        this._onSetFooter = props.onSetFooter;
        this._onSetHeader = props.onSetHeader;
        this._onValidation = props.onValidation;
        this._updateEvent = props.onUpdate;
        typeof (props.useModal) === "boolean" ? this._useModal = props.useModal : false;

        // Load the item
        this.load(props.webUrl);
    }

    // Edits a task
    static edit(props: IItemFormEditProps) {
        // Set the properties
        this._controlMode = SPTypes.ControlMode.Edit;
        this._info = props.info;
        this._onCreateEditForm = props.onCreateEditForm;
        this._onFormButtonsRendering = props.onFormButtonsRendering;
        this._onGetListInfo = props.onGetListInfo;
        this._onSave = props.onSave;
        this._onSetFooter = props.onSetFooter;
        this._onSetHeader = props.onSetHeader;
        this._onValidation = props.onValidation;
        this._updateEvent = props.onUpdate;
        typeof (props.useModal) === "boolean" ? this._useModal = props.useModal : false;

        // Load the form
        this.load(props.webUrl, props.itemId);
    }

    // Loads the form information
    static loadFormInfo(props: Helper.IListFormProps): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Load the form info
            Helper.ListForm.create(props).then(info => {
                // Save the information
                this._info = info;

                // Resolve the request
                resolve();
            });
        });
    }

    // Views the task
    static view(props: IItemFormViewProps) {
        // Set the properties
        this._controlMode = SPTypes.ControlMode.Display;
        this._info = props.info;
        this._onCreateViewForm = props.onCreateViewForm;
        this._onFormButtonsRendering = props.onFormButtonsRendering;
        this._onGetListInfo = props.onGetListInfo;
        this._onSetFooter = props.onSetFooter;
        this._onSetHeader = props.onSetHeader;
        typeof (props.useModal) === "boolean" ? this._useModal = props.useModal : false;

        // Load the form
        this.load(props.webUrl, props.itemId);
    }

    /** private static Methods */

    // Load the form information
    private static load(webUrl?: string, itemId?: number) {
        // Clear the forms
        this._displayForm = null;
        this._editForm = null;

        // Show a loading dialog
        LoadingDialog.setHeader("Loading the Item");
        LoadingDialog.setBody("This will close after the form is loaded...");
        LoadingDialog.show();

        // Load the form information
        ((): PromiseLike<void> => {
            // Return a promise
            return new Promise(resolve => {
                // See if the info already exists
                if (this._info) {
                    // Resolve the request
                    resolve();
                } else {
                    // Set the list form properties
                    let listProps: Helper.IListFormProps = {
                        listName: this.ListName,
                        itemId,
                        webUrl
                    };

                    // Call the event
                    listProps = this._onGetListInfo ? this._onGetListInfo(listProps) : listProps;

                    // Load the form info
                    Helper.ListForm.create(listProps).then(info => {
                        // Save the information
                        this._info = info;

                        // Resolve the request
                        resolve();
                    });
                }
            });
        })().then(() => {
            // Set the header
            (this._useModal ? Modal : CanvasForm).setHeader('<h5 class="m-0">' + (this._info.item ? this._info.item.Title : "Create Item") + '</h5>');

            // Call the header event
            this._onSetHeader ? this._onSetHeader(this._useModal ? Modal.HeaderElement : CanvasForm.HeaderElement) : null;

            // Render the form based on the type
            if (this.IsDisplay) {
                let el = document.createElement("div");
                let props: Components.IListFormDisplayProps = {
                    el,
                    info: this._info,
                    rowClassName: "mb-3"
                };

                // Call the event if it exists
                props = this._onCreateViewForm ? this._onCreateViewForm(props) : props;

                // Render the display form
                this._displayForm = Components.ListForm.renderDisplayForm(props);

                /* Remove the bottom margin from the last row of the form */
                (this._displayForm.el.lastChild as HTMLElement).classList.remove("mb-3");

                // Render the form buttons
                let elButtons = document.createElement("div");
                el.appendChild(elButtons);

                // Add styling if not using a modal
                if (!this._useModal) {
                    elButtons.classList.add("float-end");
                    elButtons.style.padding = "1rem 0";
                }

                // Append the create/update button
                this._useModal ? Modal.setFooter(elButtons) : el.appendChild(elButtons);

                // Call the item form button rendering event
                let formButtons: Components.IButtonProps[] = [];
                formButtons = this._onFormButtonsRendering ? this._onFormButtonsRendering(formButtons) : formButtons;

                // Render the form buttons
                formButtons && formButtons.length > 0 ? Components.ButtonGroup({
                    el: elButtons,
                    buttons: formButtons
                }) : null;

                // Call the footer event
                this._onSetFooter ? this._onSetFooter(this._useModal ? Modal.FooterElement : elButtons) : null;

                // Update the body
                (this._useModal ? Modal : CanvasForm).setBody(el);
            } else {
                let el = document.createElement("div");
                let props: Components.IListFormEditProps = {
                    el,
                    info: this._info,
                    rowClassName: "mb-3",
                    controlMode: this.IsNew ? SPTypes.ControlMode.New : SPTypes.ControlMode.Edit
                };

                // Call the event if it exists
                props = this._onCreateEditForm ? this._onCreateEditForm(props) : props;

                // Render the edit form
                this._editForm = Components.ListForm.renderEditForm(props);

                /* Remove the bottom margin from the last row of the form */
                (this._editForm.el.lastChild as HTMLElement).classList.remove("mb-3");

                // Render the form buttons
                let elButtons = document.createElement("div");

                // Add styling if not using a modal
                if (!this._useModal) {
                    elButtons.classList.add("float-end");
                    elButtons.style.padding = "1rem 0";
                }

                // Append the create/update button
                this._useModal ? Modal.setFooter(elButtons) : el.appendChild(elButtons);

                // Call the item form button rendering event
                let formButtons: Components.IButtonProps[] = [{
                    text: this.IsNew ? "Create" : "Update",
                    type: Components.ButtonTypes.OutlinePrimary,
                    onClick: () => { this.save(this._editForm); }
                }];
                formButtons = this._onFormButtonsRendering ? this._onFormButtonsRendering(formButtons) : formButtons;

                // Render the form buttons
                formButtons && formButtons.length > 0 ? Components.ButtonGroup({
                    el: elButtons,
                    buttons: formButtons
                }) : null;

                // Call the footer event
                this._onSetFooter ? this._onSetFooter(this._useModal ? Modal.FooterElement : elButtons) : null;

                // Update the body
                (this._useModal ? Modal : CanvasForm).setBody(el);
            }

            // Close the dialog
            LoadingDialog.hide();

            // Show the form
            (this._useModal ? Modal : CanvasForm).show();
        });
    }

    // Saves the edit form
    static save(form: Components.IListFormEdit = this._editForm) {
        // Validate the form
        this.validate(form).then(
            // Success
            () => {
                // Display a loading dialog
                LoadingDialog.setHeader("Saving the Item");
                LoadingDialog.setBody((this.IsNew ? "Creating" : "Updating") + " the Item");
                LoadingDialog.show();

                // Saves the item
                let saveItem = (values) => {
                    // Save the item
                    form.save(values).then(item => {
                        // Call the update event
                        this._updateEvent ? this._updateEvent(item) : null;

                        // Close the dialogs
                        (this._useModal ? Modal : CanvasForm).hide();
                        LoadingDialog.hide();
                    });
                }

                // Call the save event
                let values = form.getValues();
                values = this._onSave ? this._onSave(values) : values;

                // See if the onSave event returned a promise
                if (values && typeof (values.then) === "function") {
                    // Wait for the promise to complete
                    values.then(values => {
                        // Save the item
                        saveItem(values);
                    });
                } else {
                    // Save the item
                    saveItem(values);
                }
            },
            // Error
            () => {
                // Do Nothing
            }
        );
    }

    // Validates the form
    private static validate(form: Components.IListFormEdit): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let isValid = form.isValid();

            // Display a loading dialog
            LoadingDialog.setHeader("Validation");
            LoadingDialog.setBody("Validating the form...");
            LoadingDialog.show();

            // Ensure it's valid
            if (!isValid) {
                // Close the dialog
                LoadingDialog.hide();

                // Reject the request
                reject();
                return;
            }

            // Call the validation event
            let returnVal: any = this._onValidation ? this._onValidation(form.getValues()) : null;
            if (returnVal && typeof (returnVal.then) === "function") {
                // Wait for the promise to complete
                returnVal.then(isValid => {
                    // Close the dialog
                    LoadingDialog.hide();

                    // Resolve the request
                    isValid ? resolve() : reject();
                });
            } else {
                if (typeof (returnVal) === "boolean") {
                    // Update the flag
                    isValid = returnVal;
                }

                // Close the dialog
                LoadingDialog.hide();

                // Resolve the request
                isValid ? resolve() : reject();
            }
        });
    }
}