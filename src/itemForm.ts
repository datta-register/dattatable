import { Components, Helper, SPTypes } from "gd-sprest-bs";
import { CanvasForm, LoadingDialog, Modal } from "./common";

/** Tab */
export interface IItemFormTab {
    title: string;
    fields: string[];
    onFormRendered?: (form?: Components.IListFormDisplay | Components.IListFormEdit) => void;
    onRendered?: (el?: HTMLElement, item?: Components.IListGroupItem) => void;
    onRendering?: (item?: Components.IListGroupItem) => object;
}

/** Tab Information */
export interface IItemFormTabInfo {
    isVertical?: boolean;
    tabs: IItemFormTab[];
}

/** Create Item Properties */
export interface IItemFormCreateProps {
    info?: Helper.IListFormResult;
    onCreateEditForm?: (props: Components.IListFormEditProps) => Components.IListFormEditProps;
    onFormButtonsRendering?: (buttons: Components.IButtonProps[]) => Components.IButtonProps[];
    onGetListInfo?: (props: Helper.IListFormProps) => Helper.IListFormProps;
    onSave?: (values: any) => object | PromiseLike<object>;
    onSetFooter?: (el: HTMLElement) => void;
    onSetHeader?: (el: HTMLElement) => void;
    onShowForm?: (form?: CanvasForm | Modal) => void;
    onUpdate?: (item?: any) => void;
    onValidation?: (values?: any) => boolean | PromiseLike<boolean>;
    tabInfo?: IItemFormTabInfo;
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
    onSave?: (values: any) => object | PromiseLike<object>;
    onSetFooter?: (el: HTMLElement) => void;
    onSetHeader?: (el: HTMLElement) => void;
    onShowForm?: (form?: CanvasForm | Modal) => void;
    onUpdate?: (item?: any) => void;
    onValidation?: (values?: any) => boolean | PromiseLike<boolean>;
    tabInfo?: IItemFormTabInfo;
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
    onShowForm?: (form?: CanvasForm | Modal) => void;
    tabInfo?: IItemFormTabInfo;
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
    private static _onShowForm?: (form?: CanvasForm | Modal) => void = null;
    private static _onSave: (values: any) => any | PromiseLike<any> = null;
    private static _onValidation: (values?: any) => boolean | PromiseLike<boolean> = null;
    private static _tabInfo: IItemFormTabInfo = null;
    private static _updateEvent: Function = null;

    // Auto Close Flag
    static set AutoClose(value: boolean) {
        // Update the flag
        this.UseModal ? Modal.setAutoClose(value) : CanvasForm.setAutoClose(value);
    }

    // Display Form
    private static _displayForms: Components.IListFormDisplay[] = null;
    static get DisplayForm(): Components.IListFormDisplay { return this._displayForms[0]; }
    static get DisplayForms(): Components.IListFormDisplay[] { return this._displayForms; }

    // Edit Form
    private static _editForms: Components.IListFormEdit[] = null;
    static get EditForm(): Components.IListFormEdit { return this._editForms[0]; }
    static get EditForms(): Components.IListFormEdit[] { return this._editForms; }

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

    // List Form Tabs
    private static _tabs: Components.IListGroup = null;
    static get Tabs(): Components.IListGroup { return this._tabs; }

    // Flag to use a modal or canvas (default)
    private static _useModal: boolean = false;
    static get UseModal(): boolean { return this._useModal; }
    static set UseModal(value: boolean) { this._useModal = value; }

    // Sets the size of the modal or canvas
    static setSize(value: number) {
        // Set the modal type or offcanvas size
        this._useModal ? Modal.setType(value) : CanvasForm.setSize(value);
    }

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
        this._onShowForm = props.onShowForm;
        this._onValidation = props.onValidation;
        this._tabInfo = props.tabInfo;
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
        this._onShowForm = props.onShowForm;
        this._onValidation = props.onValidation;
        this._tabInfo = props.tabInfo;
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
        this._onShowForm = props.onShowForm;
        this._tabInfo = props.tabInfo;
        typeof (props.useModal) === "boolean" ? this._useModal = props.useModal : false;

        // Load the form
        this.load(props.webUrl, props.itemId);
    }

    /** Private Methods */

    // Load the form information
    private static load(webUrl?: string, itemId?: number) {
        // Clear the forms
        this._displayForms = [];
        this._editForms = [];

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
            // Render the form
            this.renderForm();
        });
    }

    // Renders the display form
    private static renderDisplayForm(el?: HTMLElement, tab?: IItemFormTab) {
        // Ensure the element exists
        el = el || document.createElement("div");

        // Set the form properties
        let displayAttachments = tab ? tab.fields && tab.fields.indexOf("Attachments") >= 0 : true;
        let props: Components.IListFormDisplayProps = {
            el,
            displayAttachments,
            info: this._info,
            rowClassName: "mb-3",
            includeFields: tab ? tab.fields : null,
            onFormRendered: form => {
                /* Remove the bottom margin from the last row of the form */
                (form.el.lastChild as HTMLElement).classList.remove("mb-3");
            }
        };

        // Call the event if it exists
        props = this._onCreateViewForm ? this._onCreateViewForm(props) : props;

        // Render the display form
        let form = Components.ListForm.renderDisplayForm(props);
        this._displayForms.push(form);

        // Call the event if it exists
        tab && tab.onFormRendered ? tab.onFormRendered(form) : null;

        // Return the form element
        return el;
    }

    // Renders the edit form
    private static renderEditForm(el?: HTMLElement, tab?: IItemFormTab) {
        // Ensure the element exists
        el = el || document.createElement("div");

        // Set the form properties
        let displayAttachments = tab ? tab.fields && tab.fields.indexOf("Attachments") >= 0 : true;
        let props: Components.IListFormEditProps = {
            el,
            displayAttachments,
            info: this._info,
            rowClassName: "mb-3",
            controlMode: this.IsNew ? SPTypes.ControlMode.New : SPTypes.ControlMode.Edit,
            includeFields: tab ? tab.fields : null,
            onFormRendered: form => {
                /* Remove the bottom margin from the last row of the form */
                (form.el.lastChild as HTMLElement).classList.remove("mb-3");
            }
        };

        // Call the event if it exists
        props = this._onCreateEditForm ? this._onCreateEditForm(props) : props;

        // Render the edit form
        let form = Components.ListForm.renderEditForm(props);
        this._editForms.push(form);

        // Call the event if it exists
        tab && tab.onFormRendered ? tab.onFormRendered(form) : null;

        // Return the form element
        return el;
    }

    // Renders the footer
    private static renderFooter() {
        // Render the form buttons
        let elButtons = document.createElement("div");

        // Add styling if not using a modal
        if (!this._useModal) {
            elButtons.classList.add("float-end");
            elButtons.style.padding = "1rem 0";
        }

        // Append the create/update button
        this._useModal ? Modal.setFooter(elButtons) : CanvasForm.BodyElement.appendChild(elButtons);

        // See if we are rendering a display form
        if (this.IsDisplay) {
            // Call the item form button rendering event
            let formButtons: Components.IButtonProps[] = [];
            formButtons = this._onFormButtonsRendering ? this._onFormButtonsRendering(formButtons) : formButtons;

            // Render the form buttons
            formButtons && formButtons.length > 0 ? Components.ButtonGroup({
                el: elButtons,
                buttons: formButtons
            }) : null;
        } else {
            // Call the item form button rendering event
            let formButtons: Components.IButtonProps[] = [{
                text: this.IsNew ? "Create" : "Update",
                type: Components.ButtonTypes.OutlinePrimary,
                onClick: () => { this.save(); }
            }];
            formButtons = this._onFormButtonsRendering ? this._onFormButtonsRendering(formButtons) : formButtons;

            // Render the form buttons
            formButtons && formButtons.length > 0 ? Components.ButtonGroup({
                el: elButtons,
                buttons: formButtons
            }) : null;
        }

        // Call the footer event
        this._onSetFooter ? this._onSetFooter(this._useModal ? Modal.FooterElement : elButtons) : null;
    }

    // Renders the form
    private static renderForm() {
        // Clear the form
        (this._useModal ? Modal : CanvasForm).clear();

        // Set the header
        (this._useModal ? Modal : CanvasForm).setHeader('<h5 class="m-0">' + (this._info.item ? this._info.item.Title : "Create Item") + '</h5>');

        // Call the header event
        this._onSetHeader ? this._onSetHeader(this._useModal ? Modal.HeaderElement : CanvasForm.HeaderElement) : null;

        // See if we are rendering tabs
        if (this._tabInfo) {
            // Render the tabs
            this.renderTabs();
        } else {
            // Render the form based on the type
            let elForm = this.IsDisplay ? this.renderDisplayForm() : this.renderEditForm();

            // Update the body
            (this._useModal ? Modal : CanvasForm).setBody(elForm);
        }

        // Render the footer
        this.renderFooter();

        // Call the event
        this._onShowForm ? this._onShowForm(this._useModal ? Modal : CanvasForm) : null;

        // Close the dialog
        LoadingDialog.hide();

        // Show the form
        (this._useModal ? Modal : CanvasForm).show();
    }

    // Generates a tab
    private static renderTabs(): Element {
        let tabs: Components.IListGroupItem[] = [];

        // Parse the tabs to render
        for (let i = 0; i < this._tabInfo.tabs.length; i++) {
            let tabInfo = this._tabInfo.tabs[i];

            // Generate the tab
            let tab: Components.IListGroupItem = {
                data: tabInfo,
                isActive: i == 0,
                tabName: tabInfo.title,
                onRender: (el, item) => {
                    let tab = item.data as IItemFormTab;

                    // Render the form
                    this.IsDisplay ? this.renderDisplayForm(el, tab) : this.renderEditForm(el, tab);

                    // Call the event
                    tab.onRendered ? tab.onRendered(el, item) : null;
                }
            };

            // Call the rendering event
            tab = tabInfo.onRendering ? tabInfo.onRendering(tab) : tab;
            if (tab) {
                // Append the tab
                tabs.push(tab);
            }
        }

        // Render the tabs
        this._tabs = Components.ListGroup({
            el: this.UseModal ? Modal.BodyElement : CanvasForm.BodyElement,
            colWidth: this._tabInfo.isVertical ? 4 : 12,
            isTabs: true,
            isHorizontal: this._tabInfo.isVertical != true,
            items: tabs
        });

        // Return the tabs element
        return this._tabs.el;
    }

    // Saves the edit form
    static save(form?: Components.IListFormEdit) {
        let values = {};

        // Default the form
        let forms = form ? [form] : this._editForms;

        // Display a loading dialog
        LoadingDialog.setHeader("Validation");
        LoadingDialog.setBody("Validating the form...");
        LoadingDialog.show();

        // Validate the forms
        let isValid = true;
        let counter = 0;
        Helper.Executor(forms, form => {
            // Update the values
            values = { ...values, ...form.getValues() }

            // See if the form is not valid
            if (!form.isValid()) {
                // Set the flag
                isValid = false;
            }

            // See if tabs exist
            if (this.Tabs) {
                // Get the tab
                let tab = this.Tabs.el.querySelectorAll(".list-group-item")[counter++] as HTMLAnchorElement;
                if (tab) {
                    // Clear the class name
                    tab.classList.remove("is-valid");
                    tab.classList.remove("is-invalid");

                    // Set the class name
                    tab.classList.add(isValid ? "is-valid" : "is-invalid");
                }
            }
            // Set the tab class
        }).then(
            // Success
            () => {
                // Call the custom validation event
                this.validate(values).then(
                    // Valid
                    () => {
                        // See if the form(s) are valid
                        if (isValid) {
                            // Update the loading dialog
                            LoadingDialog.setHeader("Saving the Item");
                            LoadingDialog.setBody((this.IsNew ? "Creating" : "Updating") + " the Item");

                            // Saves the item
                            let saveItem = (values) => {
                                // Save the item
                                forms[0].save(values).then(item => {
                                    // Call the update event
                                    this._updateEvent ? this._updateEvent(item) : null;

                                    // Close the dialogs
                                    (this._useModal ? Modal : CanvasForm).hide();
                                    LoadingDialog.hide();
                                });
                            }

                            // Call the save event
                            values = this._onSave ? this._onSave(values) : values;

                            // See if the onSave event returned a promise
                            if (values && typeof (values["then"]) === "function") {
                                // Wait for the promise to complete
                                values["then"](values => {
                                    // Save the item
                                    saveItem(values);
                                });
                            } else {
                                // Save the item
                                saveItem(values);
                            }
                        } else {
                            // Close the dialog
                            LoadingDialog.hide();
                        }
                    },

                    // Not Valid
                    () => {
                        // Close the dialog
                        LoadingDialog.hide();
                    }
                );
            }
        );
    }

    // Validates the form
    private static validate(values: any): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let isValid = true;

            // Call the validation event
            let returnVal: any = this._onValidation ? this._onValidation(values) : null;
            if (returnVal && typeof (returnVal.then) === "function") {
                // Wait for the promise to complete
                returnVal.then(isValid => {
                    // Resolve the request
                    isValid ? resolve() : reject();
                });
            } else {
                if (typeof (returnVal) === "boolean") {
                    // Update the flag
                    isValid = returnVal;
                }

                // Resolve the request
                isValid ? resolve() : reject();
            }
        });
    }
}