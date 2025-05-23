import { Components, Helper, SPTypes, Web } from "gd-sprest-bs";
import { CanvasForm, LoadingDialog, Modal, getContextInfo } from ".";

/** Tab */
export interface IItemFormTab {
    title: string;
    fields?: string[];
    excludeFields?: string[];
    isReadOnly?: boolean;
    onCreateForm?: (props: Components.IListFormDisplayProps | Components.IListFormEditProps) => Components.IListFormDisplayProps | Components.IListFormEditProps;
    onFormRendered?: (form?: Components.IListFormDisplay | Components.IListFormEdit) => void;
    onRendered?: (el?: HTMLElement, item?: Components.IListGroupItem) => void;
    onRendering?: (item?: Components.IListGroupItem) => object;
    onValidation?: (values?: any) => boolean;
}

/** Tab Information */
export interface IItemFormTabInfo {
    isVertical?: boolean;
    onClick?: (el?: HTMLElement, item?: Components.IListGroupItem) => void;
    tabs: IItemFormTab[];
}

/** Create Item Properties */
export interface IItemFormCreateProps {
    elForm?: HTMLElement;
    info?: Helper.IListFormResult;
    onCreateEditForm?: (props: Components.IListFormEditProps) => Components.IListFormEditProps;
    onFormButtonsRendering?: (buttons: Components.IButtonProps[]) => Components.IButtonProps[];
    onFormRendered?: (form?: Components.IListFormEdit) => void;
    onGetListInfo?: (props: Helper.IListFormProps) => Helper.IListFormProps;
    onResetForm?: () => void;
    onSave?: (values: any) => object | PromiseLike<object>;
    onSaveError?: (err: any) => void;
    onSetFooter?: (el: HTMLElement) => void;
    onSetHeader?: (el: HTMLElement) => void;
    onShowForm?: (form?: CanvasForm | Modal) => void;
    onUpdate?: (item?: any) => void;
    onValidation?: (values?: any, isValid?: boolean) => boolean | PromiseLike<boolean>;
    tabInfo?: IItemFormTabInfo;
    useModal?: boolean;
    webUrl?: string;
}

/** Edit Item Properties */
export interface IItemFormEditProps {
    elForm?: HTMLElement;
    info?: Helper.IListFormResult;
    itemId: number;
    onCreateEditForm?: (props: Components.IListFormEditProps) => Components.IListFormEditProps;
    onFormButtonsRendering?: (buttons: Components.IButtonProps[]) => Components.IButtonProps[];
    onFormRendered?: (form?: Components.IListFormEdit) => void;
    onGetListInfo?: (props: Helper.IListFormProps) => Helper.IListFormProps;
    onResetForm?: () => void;
    onSave?: (values: any) => object | PromiseLike<object>;
    onSaveError?: (err: any) => void;
    onSetFooter?: (el: HTMLElement) => void;
    onSetHeader?: (el: HTMLElement) => void;
    onShowForm?: (form?: CanvasForm | Modal) => void;
    onUpdate?: (item?: any) => void;
    onValidation?: (values?: any, isValid?: boolean) => boolean | PromiseLike<boolean>;
    tabInfo?: IItemFormTabInfo;
    useModal?: boolean;
    webUrl?: string;
}

/** Save Item Properties */
export interface IItemFormSaveProps {
    bypassValidation?: boolean;
    checkItemVersion?: boolean;
    form?: Components.IListFormEdit;
}

/** View Item Properties */
export interface IItemFormViewProps {
    elForm?: HTMLElement;
    info?: Helper.IListFormResult;
    itemId: number;
    onCreateViewForm?: (props: Components.IListFormDisplayProps) => Components.IListFormDisplayProps;
    onFormButtonsRendering?: (buttons: Components.IButtonProps[]) => Components.IButtonProps[];
    onFormRendered?: (form?: Components.IListFormDisplay) => void;
    onGetListInfo?: (props: Helper.IListFormProps) => Helper.IListFormProps;
    onResetForm?: () => void;
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
    private static _elForm: HTMLElement = null;
    private static _onCreateEditForm: (props: Components.IListFormEditProps) => Components.IListFormEditProps = null;
    private static _onCreateViewForm: (props: Components.IListFormDisplayProps) => Components.IListFormDisplayProps = null;
    private static _onFormButtonsRendering: (buttons: Components.IButtonProps[]) => Components.IButtonProps[] = null;
    private static _onFormRendered?: (form?: Components.IListFormDisplay | Components.IListFormEdit) => void = null;
    private static _onGetListInfo: (props: Helper.IListFormProps) => Helper.IListFormProps = null;
    private static _onResetForm: () => void = null;
    private static _onSetFooter?: (el: HTMLElement) => void = null;
    private static _onSetHeader?: (el: HTMLElement) => void = null;
    private static _onShowForm?: (form?: CanvasForm | Modal) => void = null;
    private static _onSave: (values: any) => any | PromiseLike<any> = null;
    private static _onSaveError: (err: any) => void = null;
    private static _onValidation: (values?: any, isValid?: boolean) => boolean | PromiseLike<boolean> = null;
    private static _requestDigest: string = null;
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
    private static _info: Helper.IListFormResult = null;
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
    static create(props: IItemFormCreateProps = {}): PromiseLike<void> {
        // Clear the properties
        this.clearProps();

        // Set the properties
        this._controlMode = SPTypes.ControlMode.New;
        this._elForm = props.elForm;
        this._info = props.info;
        this._onCreateEditForm = props.onCreateEditForm;
        this._onFormButtonsRendering = props.onFormButtonsRendering;
        this._onFormRendered = props.onFormRendered;
        this._onGetListInfo = props.onGetListInfo;
        this._onResetForm = props.onResetForm;
        this._onSave = props.onSave;
        this._onSaveError = props.onSaveError;
        this._onSetFooter = props.onSetFooter;
        this._onSetHeader = props.onSetHeader;
        this._onShowForm = props.onShowForm;
        this._onValidation = props.onValidation;
        this._tabInfo = props.tabInfo;
        this._updateEvent = props.onUpdate;
        typeof (props.useModal) === "boolean" ? this._useModal = props.useModal : false;

        // Load the item
        return this.load(props.webUrl);
    }

    // Edits a task
    static edit(props: IItemFormEditProps): PromiseLike<void> {
        // Clear the properties
        this.clearProps();

        // Set the properties
        this._controlMode = SPTypes.ControlMode.Edit;
        this._elForm = props.elForm;
        this._info = props.info;
        this._onCreateEditForm = props.onCreateEditForm;
        this._onFormButtonsRendering = props.onFormButtonsRendering;
        this._onFormRendered = props.onFormRendered;
        this._onGetListInfo = props.onGetListInfo;
        this._onResetForm = props.onResetForm;
        this._onSave = props.onSave;
        this._onSaveError = props.onSaveError;
        this._onSetFooter = props.onSetFooter;
        this._onSetHeader = props.onSetHeader;
        this._onShowForm = props.onShowForm;
        this._onValidation = props.onValidation;
        this._tabInfo = props.tabInfo;
        this._updateEvent = props.onUpdate;
        typeof (props.useModal) === "boolean" ? this._useModal = props.useModal : false;

        // Load the form
        return this.load(props.webUrl, props.itemId);
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

    // Renders the footer for a form
    static renderFooter(elForm: HTMLElement) {
        // Render the form footer
        this.renderFormFooter(elForm);
    }

    // Views the task
    static view(props: IItemFormViewProps): PromiseLike<void> {
        // Clear the properties
        this.clearProps();

        // Set the properties
        this._controlMode = SPTypes.ControlMode.Display;
        this._elForm = props.elForm;
        this._info = props.info;
        this._onCreateViewForm = props.onCreateViewForm;
        this._onFormButtonsRendering = props.onFormButtonsRendering;
        this._onFormRendered = props.onFormRendered;
        this._onGetListInfo = props.onGetListInfo;
        this._onSetFooter = props.onSetFooter;
        this._onSetHeader = props.onSetHeader;
        this._onShowForm = props.onShowForm;
        this._tabInfo = props.tabInfo;
        typeof (props.useModal) === "boolean" ? this._useModal = props.useModal : false;

        // Load the form
        return this.load(props.webUrl, props.itemId);
    }

    /** Private Methods */

    // Clears the properties
    private static clearProps() {
        this._elForm = null;
        this._onCreateEditForm = null;
        this._onCreateViewForm = null;
        this._onFormButtonsRendering = null;
        this._onFormRendered = null;
        this._onGetListInfo = null;
        this._onResetForm = null;
        this._onSetFooter = null;
        this._onSetHeader = null;
        this._onShowForm = null;
        this._onSave = null;
        this._onSaveError = null;
        this._onValidation = null;
        this._requestDigest = null;
        this._tabInfo = null;
        this._updateEvent = null;
    }

    // Load the form information
    private static load(webUrl?: string, itemId?: number): PromiseLike<void> {
        // Clear the forms
        this._displayForms = [];
        this._editForms = [];

        // Show a loading dialog
        LoadingDialog.setHeader("Loading the Item");
        LoadingDialog.setBody("This will close after the form is loaded...");
        LoadingDialog.show();

        // Return a promise
        return new Promise((resolve, reject) => {
            // Code to run after the data is loaded
            let onComplete = () => {
                // Render the form
                this.renderForm();

                // Resolve the request
                resolve(null);
            }

            // Get the context information
            getContextInfo(webUrl).then(requestDigest => {
                // Set the request digest
                this._requestDigest = requestDigest;

                // See if the info already exists
                if (this._info) {
                    // Form information is already loaded
                    onComplete();
                } else {
                    // Set the list form properties
                    let listProps: Helper.IListFormProps = {
                        listName: this.ListName,
                        itemId,
                        requestDigest,
                        webUrl
                    };

                    // Call the event
                    listProps = this._onGetListInfo ? this._onGetListInfo(listProps) : listProps;

                    // Load the form info
                    Helper.ListForm.create(listProps).then(info => {
                        // Save the information
                        this._info = info;

                        // Form information is already loaded
                        onComplete();
                    }, reject);
                }
            });
        });
    }

    // Renders the display form
    private static renderDisplayForm(el?: HTMLElement, tab?: IItemFormTab) {
        // Ensure the element exists
        el = el || document.createElement("div");

        // Set the form properties
        let displayAttachments = tab ? tab.fields && tab.fields.indexOf("Attachments") >= 0 : this.FormInfo.attachments != null;
        let props: Components.IListFormDisplayProps = {
            el,
            displayAttachments,
            excludeFields: tab ? tab.excludeFields : null,
            includeFields: tab ? tab.fields : null,
            info: this._info,
            rowClassName: "mb-3"
        };

        // See if we are rendering a tab
        if (tab && tab.onCreateForm) {
            // Call the event
            props = tab.onCreateForm(props);
        }

        // Call the event if it exists
        props = this._onCreateViewForm ? this._onCreateViewForm(props) : props;

        // Override the form rendered event
        let customEvent = props.onFormRendered;
        props.onFormRendered = listForm => {
            /* Remove the bottom margin from the last row of the form */
            (listForm.el.lastChild as HTMLElement).classList.remove("mb-3");

            // Call the event if it exists
            tab && tab.onFormRendered ? tab.onFormRendered(form) : null;

            // Call the custom event if it exists
            customEvent ? customEvent(listForm) : null;
        }

        // Render the display form
        let form = Components.ListForm.renderDisplayForm(props);
        this._displayForms.push(form);

        // Return the form element
        return el;
    }

    // Renders the edit form
    private static renderEditForm(el?: HTMLElement, tab?: IItemFormTab) {
        // Ensure the element exists
        el = el || document.createElement("div");

        // Set the form properties
        let displayAttachments = tab ? tab.fields && tab.fields.indexOf("Attachments") >= 0 : this.FormInfo.attachments != null;
        let props: Components.IListFormEditProps = {
            el,
            displayAttachments,
            excludeFields: tab ? tab.excludeFields : null,
            includeFields: tab ? tab.fields : null,
            info: this._info,
            rowClassName: "mb-3",
            controlMode: this.IsNew ? SPTypes.ControlMode.New : SPTypes.ControlMode.Edit
        };

        // See if we are rendering a tab
        if (tab && tab.onCreateForm) {
            // Call the event
            props = tab.onCreateForm(props);
        }

        // Call the event if it exists
        props = this._onCreateEditForm ? this._onCreateEditForm(props) : props;

        // Override the form rendered event
        let customEvent = props.onFormRendered;
        props.onFormRendered = listForm => {
            /* Remove the bottom margin from the last row of the form */
            (listForm.el.lastChild as HTMLElement).classList.remove("mb-3");

            // Call the event if it exists
            tab && tab.onFormRendered ? tab.onFormRendered(form) : null;

            // Call the custom event if it exists
            customEvent ? customEvent(listForm) : null;
        }

        // Render the edit form
        let form = Components.ListForm.renderEditForm(props);

        // Set the data object to store the tab information
        form.data = tab;

        // Append the form
        this._editForms.push(form);

        // Return the form element
        return el;
    }

    // Renders the footer
    private static renderFormFooter(elFooter?: HTMLElement) {
        // Render the form buttons
        let elButtons = document.createElement("div");

        // See if we are rendering to the canvas
        if (elFooter == null && !this._useModal) {
            // Add the classes
            elButtons.classList.add("d-flex");
            elButtons.classList.add("justify-content-end");
            elButtons.classList.add("my-2");
        }

        // See if we are using a custom element
        if (elFooter) {
            // Append the buttons
            elFooter.appendChild(elButtons);
        } else {
            // Append the create/update button
            this._useModal ? Modal.setFooter(elButtons) : CanvasForm.BodyElement.appendChild(elButtons);
        }

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
                onClick: () => {
                    // Save the form
                    this.save().then(() => { }, () => { });
                }
            }];
            formButtons = this._onFormButtonsRendering ? this._onFormButtonsRendering(formButtons) : formButtons;

            // Render the form buttons
            formButtons && formButtons.length > 0 ? Components.ButtonGroup({
                el: elButtons,
                buttons: formButtons
            }) : null;
        }

        // Call the footer event
        this._onSetFooter ? this._onSetFooter(elFooter || (this._useModal ? Modal.FooterElement : elButtons)) : null;
    }

    // Renders the form
    private static renderForm() {
        // See if we are not rendering to a custom element
        if (this._elForm == null) {
            // Clear the form
            (this._useModal ? Modal : CanvasForm).clear();

            // Call the event
            this._onResetForm ? this._onResetForm() : null;

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

            // Render the form footer
            this.renderFormFooter();

            // Call the event
            this._onShowForm ? this._onShowForm(this._useModal ? Modal : CanvasForm) : null;

            // Show the form
            (this._useModal ? Modal : CanvasForm).show();
        } else {
            // See if we are rendering tabs
            if (this._tabInfo) {
                // Render the tabs
                this.renderTabs();
            } else {
                // Render the form based on the type
                let elForm = this.IsDisplay ? this.renderDisplayForm() : this.renderEditForm();

                // Copy the elements
                for (let i = 0; i < elForm.children.length; i++) {
                    // Append the element
                    this._elForm.appendChild(elForm.children[i]);
                }
            }
        }

        // Call the rendered event
        this._onFormRendered ? this._onFormRendered(this.IsDisplay ? this.DisplayForm : this.EditForm) : null;

        // Close the dialog
        LoadingDialog.hide();
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
                onClick: this._tabInfo.onClick,
                onRender: (el, item) => {
                    let tab = item.data as IItemFormTab;

                    // Render the form
                    this.IsDisplay || tab.isReadOnly ? this.renderDisplayForm(el, tab) : this.renderEditForm(el, tab);

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
            el: this._elForm || (this.UseModal ? Modal.BodyElement : CanvasForm.BodyElement),
            colWidth: this._tabInfo.isVertical ? 4 : 12,
            isTabs: true,
            isHorizontal: this._tabInfo.isVertical != true,
            items: tabs
        });

        // Return the tabs element
        return this._tabs.el;
    }

    // Saves the edit form
    static save(props: IItemFormSaveProps = {}): PromiseLike<any> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let values = {};

            // Default the form
            let forms = props.form ? [props.form] : this._editForms;
            let defaultForm = forms[0];

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

                // See if this form has attachments
                if (form.hasAttachments()) {
                    // Set the default form
                    defaultForm = form;
                }

                // See if we are bypassing validation
                if (props.bypassValidation == null || props.bypassValidation != true) {
                    // Validate the form
                    let tabIsValid = form.isValid();

                    // See if we are using tabs and an event exists
                    let tabInfo = form.data as IItemFormTab;
                    if (tabInfo && tabInfo.onValidation) {
                        // Call the event
                        tabIsValid = tabInfo.onValidation(values);
                    }

                    // See if the form is not valid
                    if (!tabIsValid) {
                        // Set the flag
                        isValid = false;
                    }

                    // See if tabs exist
                    if (this.Tabs && tabInfo) {
                        // Get the tab
                        let tab = this.Tabs.el.querySelector(`.list-group-item[data-tab-title='${tabInfo.title}']`) as HTMLAnchorElement;
                        if (tab) {
                            // Clear the class name
                            tab.classList.remove("is-valid");
                            tab.classList.remove("is-invalid");

                            // Set the class name
                            tab.classList.add(tabIsValid ? "is-valid" : "is-invalid");
                        }
                    }
                }
            }).then(
                // Success
                () => {
                    // Call the custom validation event
                    this.validate(values, isValid).then(
                        // Valid
                        isValid => {
                            // See if the form(s) are valid
                            if (isValid || props.bypassValidation == true) {
                                // Update the loading dialog
                                LoadingDialog.setHeader("Saving the Item");
                                LoadingDialog.setBody((this.IsNew ? "Creating" : "Updating") + " the Item");

                                // Saves the item
                                let saveItem = (values, retryFl?: boolean) => {
                                    // If values is null then do nothing. This would happen when the
                                    // onSave event wants to cancel the save and do something custom
                                    if (values) {
                                        // Save the item
                                        defaultForm.save(values, props.checkItemVersion).then(item => {
                                            // Call the update event
                                            this._updateEvent ? this._updateEvent(item) : null;

                                            // Close the dialogs
                                            (this._useModal ? Modal : CanvasForm).hide();
                                            LoadingDialog.hide();

                                            // Resolve the request
                                            resolve(item);
                                        }, err => {
                                            // See if we have already retried to save the item
                                            if (retryFl) {
                                                // Log
                                                console.error("[List Form] Unable to get the context for web.");

                                                // Call the event
                                                this._onSaveError ? this._onSaveError(err) : null;

                                                // Reject the request
                                                reject();
                                            } else {
                                                // Try to get the error message
                                                try {
                                                    let errorMessage: string = JSON.parse(err.response).error.message.value;

                                                    // See if the page timed out
                                                    if (errorMessage.indexOf("The security validation for this page is invalid") == 0) {
                                                        // Set the loading dialog
                                                        LoadingDialog.setBody("The page timed out, retrying the request...");

                                                        // Refresh the request digest
                                                        defaultForm.refreshRequestDigest().then(() => {
                                                            // Try to save the item again
                                                            saveItem(values, true);
                                                        });
                                                    } else {
                                                        // Call the event
                                                        this._onSaveError ? this._onSaveError(err) : null;

                                                        // Reject the request
                                                        reject();
                                                    }
                                                } catch {
                                                    // Call the event
                                                    this._onSaveError ? this._onSaveError(err) : null;

                                                    // Reject the request
                                                    reject();
                                                }
                                            }
                                        });
                                    }
                                }

                                // Call the save event and ensure values exist
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

                            // Reject the request
                            reject();
                        }
                    );
                }
            );
        });
    }

    // Validates the form
    private static validate(values: any, isValid: boolean): PromiseLike<boolean> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Call the validation event
            let returnVal: any = this._onValidation ? this._onValidation(values, isValid) : null;
            // See if it's a promise function
            if (returnVal && typeof (returnVal.then) === "function") {
                // Wait for the promise to complete
                returnVal.then(returnVal => {
                    // Resolve the request
                    isValid = typeof (returnVal) === "boolean" ? returnVal : isValid;
                    isValid ? resolve(isValid) : reject();
                });
            } else {
                // Update the flag
                isValid = typeof (returnVal) === "boolean" ? returnVal : isValid;

                // Resolve the request
                isValid ? resolve(isValid) : reject();
            }
        });
    }
}