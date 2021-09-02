import { Components, Helper, SPTypes } from "gd-sprest-bs";
import { CanvasForm, LoadingDialog, Modal } from "./common";

/**
 * Item Form
 */
class _ItemForm {
    private _onCreateEditForm: (props: Components.IListFormEditProps) => Components.IListFormEditProps = null;
    private _onCreateViewForm: (props: Components.IListFormDisplayProps) => Components.IListFormDisplayProps = null;
    private _onFormButtonsRendering: (buttons: Components.IButtonProps[]) => Components.IButtonProps[] = null;
    private _onGetListInfo: (props: Helper.IListFormProps) => Helper.IListFormProps = null;
    private _onSetFooter?: (el: HTMLElement) => void = null;
    private _onSetHeader?: (el: HTMLElement) => void = null;
    private _onSave: (values: any) => any | PromiseLike<any> = null;
    private _onValidation: (values?: any) => boolean | PromiseLike<boolean> = null;
    private _updateEvent: Function = null;

    // Auto Close Flag
    private set AutoClose(value: boolean) {
        // Update the flag
        this.UseModal ? Modal.setAutoClose(value) : CanvasForm.setAutoClose(value);
    }

    // Display Form
    private _displayForm: Components.IListFormDisplay = null;
    get DisplayForm(): Components.IListFormDisplay { return this._displayForm; }

    // Edit Form
    private _editForm: Components.IListFormEdit = null;
    get EditForm(): Components.IListFormEdit { return this._editForm; }

    // Form Information
    private _info = null;
    get FormInfo(): Helper.IListFormResult { return this._info; }

    // Form Modes
    private _controlMode: number = null;
    get IsDisplay(): boolean { return this._controlMode == SPTypes.ControlMode.Display; }
    get IsEdit(): boolean { return this._controlMode == SPTypes.ControlMode.Edit; }
    get IsNew(): boolean { return this._controlMode == SPTypes.ControlMode.New; }

    // List name
    private _listName: string = null;
    get ListName(): string { return this._listName; }
    set ListName(value: string) { this._listName = value; }

    // Flag to use a modal or canvas (default)
    private _useModal: boolean = false;
    get UseModal(): boolean { return this._useModal; }
    set UseModal(value: boolean) { this._useModal = value; }

    /** Public Methods */

    // Closes the item form
    close() {
        this._useModal ? Modal.hide() : CanvasForm.hide();
    }

    // Creates a new task
    create(props: {
        onCreateEditForm?: (props: Components.IListFormEditProps) => Components.IListFormEditProps;
        onFormButtonsRendering?: (buttons: Components.IButtonProps[]) => Components.IButtonProps[];
        onGetListInfo?: (props: Helper.IListFormProps) => Helper.IListFormProps;
        onSave?: (values: any) => any | PromiseLike<any>;
        onSetFooter?: (el: HTMLElement) => void;
        onSetHeader?: (el: HTMLElement) => void;
        onUpdate?: (item?: any) => void;
        onValidation?: (values?: any) => boolean | PromiseLike<boolean>;
        useModal?: boolean;
    } = {}) {
        // Set the properties
        this._controlMode = SPTypes.ControlMode.New;
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
        this.load();
    }

    // Edits a task
    edit(props: {
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
    }) {
        // Set the properties
        this._controlMode = SPTypes.ControlMode.Edit;
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
        this.load(props.itemId);
    }

    // Views the task
    view(props: {
        itemId: number;
        onCreateViewForm?: (props: Components.IListFormDisplayProps) => Components.IListFormDisplayProps;
        onFormButtonsRendering?: (buttons: Components.IButtonProps[]) => Components.IButtonProps[];
        onGetListInfo?: (props: Helper.IListFormProps) => Helper.IListFormProps;
        onSetFooter?: (el: HTMLElement) => void;
        onSetHeader?: (el: HTMLElement) => void;
        useModal?: boolean;
    }) {
        // Set the properties
        this._controlMode = SPTypes.ControlMode.Display;
        this._onCreateViewForm = props.onCreateViewForm;
        this._onFormButtonsRendering = props.onFormButtonsRendering;
        this._onGetListInfo = props.onGetListInfo;
        this._onSetFooter = props.onSetFooter;
        this._onSetHeader = props.onSetHeader;
        typeof (props.useModal) === "boolean" ? this._useModal = props.useModal : false;

        // Load the form
        this.load(props.itemId);
    }

    /** Private Methods */

    // Load the form information
    private load(itemId?: number) {
        // Clear the forms
        this._displayForm = null;
        this._editForm = null;

        // Show a loading dialog
        LoadingDialog.setHeader("Loading the Item");
        LoadingDialog.setBody("This will close after the form is loaded...");
        LoadingDialog.show();

        // Set the list form properties
        let listProps: Helper.IListFormProps = {
            listName: this.ListName,
            itemId
        };

        // Call the event
        listProps = this._onGetListInfo ? this._onGetListInfo(listProps) : listProps;

        // Load the form information
        Helper.ListForm.create(listProps).then(info => {
            // Save the form information
            this._info = info;

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

    // Saves the form
    private save(form: Components.IListFormEdit) {
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
                    Components.ListForm.saveItem(this._info, values).then(item => {
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
    private validate(form: Components.IListFormEdit): PromiseLike<void> {
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
export const ItemForm = new _ItemForm();
