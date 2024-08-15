import { ContextInfo, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { LoadingDialog } from "./loadingDialog";
import { List } from "./list";

// Create Lookup List Data
export interface ICreateLookupData {
    lookupData: ILookupData[];
    showDialog?: boolean;
    webUrl: string;
}

// List Configuration
export interface IListConfig {
    cfg: Helper.ISPConfigProps;
    lookupFields: Types.SP.FieldLookup[];
}

// List Configuration Properties
export interface IListConfigProps {
    showDialog?: boolean;
    srcList: Types.SP.List;
    srcWebUrl: string;
}

// Lookup List Data
export interface ILookupData {
    field: string;
    list: string;
    items: string[];
}

// Generate Lookup Data Properties
export interface IGenerateLookupDataProps {
    lookupFields: Types.SP.FieldLookup[];
    showDialog?: boolean;
    srcList: Types.SP.List;
    srcWebUrl: string;
}

// Lookup List Validation Properties
export interface IValidateLookupsProps {
    cfg: Helper.ISPConfigProps;
    dstUrl: string;
    lookupFields: Types.SP.FieldLookup[];
    showDialog?: boolean;
    srcList: Types.SP.List;
    srcWebUrl: string;
}

/**
 * List Configuration Generator
 */
export class ListConfig {
    // Internal Fields
    private static InternalFields = [
        "ContentType", "TaxCatchAll", "TaxCatchAllLabel", "Title"
    ]
    // Generates the lookup list data
    static createLookupListData(props: ICreateLookupData): PromiseLike<void> {
        // See if we are showing a loading dialog
        if (props.showDialog) {
            // Show a loading dialog
            LoadingDialog.setHeader("Lookup List Data");
            LoadingDialog.setBody("Updating the lookup list data...");
            LoadingDialog.show();
        }

        // Return a promise
        return new Promise((resolve, reject) => {
            // Ensure data exists
            if (props.lookupData == null) {
                // Hide the dialog
                props.showDialog ? LoadingDialog.hide() : null;
                resolve(null);
                return;
            }

            // Update the dialog
            props.showDialog ? LoadingDialog.setBody("Getting the web context information...") : null;

            // Get the web context info
            ContextInfo.getWeb(props.webUrl).execute(contextInfo => {
                // Parse the lookup list data
                Helper.Executor(props.lookupData, lookupData => {
                    // Update the dialog
                    props.showDialog ? LoadingDialog.setBody("Getting the list: " + lookupData.list) : null;

                    // Return a promise
                    return new Promise(resolve => {
                        // Update the dialog
                        props.showDialog ? LoadingDialog.setBody("Getting the current list data for: " + lookupData.list) : null;

                        // Get the current list items
                        Web(props.webUrl).Lists(lookupData.list).Items().query({ Top: 5000, GetAllItems: true }).execute(currItems => {
                            let dstList = Web(props.webUrl, { requestDigest: contextInfo.GetContextWebInformation.FormDigestValue }).Lists(lookupData.list);

                            // Update the dialog
                            props.showDialog ? LoadingDialog.setBody("Importing the list data: " + lookupData.list) : null;

                            // Parse the list items
                            for (let i = 0; i < lookupData.items.length; i++) {
                                let addItem = true;
                                let value = lookupData.items[i];

                                // See if the items exists
                                currItems.results.find(a => {
                                    // See if the item exists
                                    if (a[lookupData.field] == value) {
                                        // Set the flag
                                        addItem = false;
                                    }
                                });

                                // See if we are adding the item
                                if (addItem) {
                                    // Create the item
                                    let dstItem = {};
                                    dstItem[lookupData.field] = value;
                                    dstList.Items().add(dstItem).batch(item => {
                                        // Log
                                        console.log("[" + lookupData.list + "] Item added: " + item[lookupData.field]);
                                    });
                                }
                            }

                            // Execute the batch job
                            dstList.execute(() => {
                                // Check the next list
                                resolve(null);
                            });
                        }, () => {
                            // Error getting the list data
                            console.error("Error getting the destination list: " + lookupData.list);

                            // Check the next lookup field
                            resolve(null);
                        });
                    });
                }).then(() => {
                    // Hide the dialog
                    props.showDialog ? LoadingDialog.hide() : null;

                    // Resolve the request
                    resolve();
                }, reject);
            });
        });
    }

    // Generates the list configuration
    static generate(props: IListConfigProps): PromiseLike<IListConfig> {
        // See if we are showing a loading dialog
        if (props.showDialog) {
            // Show a loading dialog
            LoadingDialog.setHeader("Loading List");
            LoadingDialog.setBody("Getting the source list configuration...");
            LoadingDialog.show();
        }

        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the list information
            var list = new List({
                listName: props.srcList.Title,
                webUrl: props.srcWebUrl,
                itemQuery: { Filter: "Id eq 0" },
                onInitError: () => {
                    // Reject the request
                    reject("Error loading the list information. Please check your permission to the source list.");
                },
                onInitialized: () => {
                    let calcFields: Types.SP.Field[] = [];
                    let fields: { [key: string]: boolean } = {};
                    let lookupFields: Types.SP.FieldLookup[] = [];
                    let mmsFields: { [key: string]: Types.SP.Field } = {};

                    // Update the loading dialog
                    LoadingDialog.setBody("Analyzing the list information...");

                    // Create the configuration
                    let cfgProps: Helper.ISPConfigProps = {
                        ListCfg: [{
                            ListInformation: {
                                AllowContentTypes: list.ListInfo.AllowContentTypes,
                                BaseTemplate: list.ListInfo.BaseTemplate,
                                ContentTypesEnabled: list.ListInfo.ContentTypesEnabled,
                                Title: props.srcList.Title,
                                Hidden: list.ListInfo.Hidden,
                                NoCrawl: list.ListInfo.NoCrawl
                            },
                            ContentTypes: [],
                            CustomFields: [],
                            ViewInformation: []
                        }]
                    };

                    // Parse the content types
                    for (let i = 0; i < list.ListContentTypes.length; i++) {
                        let ct = list.ListContentTypes[i];

                        // Skip sealed content types
                        if (ct.Sealed) { continue; }

                        // Parse the content type fields
                        let fieldRefs = [];
                        for (let j = 0; j < ct.FieldLinks.results.length; j++) {
                            let fieldLink = ct.FieldLinks.results[j];

                            // Skip hidden fields
                            if (fieldLink.Hidden) { continue; }

                            // Get the field
                            let field: Types.SP.Field = list.getField(fieldLink.Name);

                            // See if this is a lookup field
                            if (field.FieldTypeKind == SPTypes.FieldType.Lookup) {
                                // Ensure this isn't an associated lookup field
                                if ((field as Types.SP.FieldLookup).IsDependentLookup != true) {
                                    // Append the field ref
                                    fieldRefs.push(field.InternalName);
                                }
                            } else {
                                // Append the field ref
                                fieldRefs.push(field.InternalName);
                            }

                            // Skip internal fields
                            if (this.InternalFields.indexOf(field.InternalName) >= 0) { continue; }

                            // See if this is a calculated field
                            if (field.FieldTypeKind == SPTypes.FieldType.Calculated) {
                                // Add the field and continue the loop
                                calcFields.push(field);
                            }
                            // Else, see if this is a lookup field
                            else if (field.FieldTypeKind == SPTypes.FieldType.Lookup) {
                                // Add the field
                                lookupFields.push(field);
                            }
                            // Else, see if this is a MMS field
                            else if (field.TypeDisplayName == "Managed Metadata") {
                                // Add the field
                                mmsFields[field.InternalName] = field;
                            }
                            // Else, ensure the field hasn't been added
                            else if (fields[field.InternalName] == null) {
                                // Add the field information
                                fields[field.InternalName] = true;
                                cfgProps.ListCfg[0].CustomFields.push({
                                    name: field.InternalName,
                                    schemaXml: field.SchemaXml
                                });
                            }
                        }

                        // Add the list content type
                        cfgProps.ListCfg[0].ContentTypes.push({
                            Name: ct.Name,
                            Description: ct.Description,
                            ParentName: ct.Parent.Name,
                            FieldRefs: fieldRefs
                        });
                    }

                    // Parse the views
                    for (let i = 0; i < list.ListViews.length; i++) {
                        let viewInfo = list.ListViews[i];

                        // Skip hidden views
                        if (viewInfo.Hidden) { continue; }

                        // Parse the fields
                        for (let j = 0; j < viewInfo.ViewFields.Items.results.length; j++) {
                            let field = list.getField(viewInfo.ViewFields.Items.results[j]);

                            // Ensure the field exists
                            if (fields[field.InternalName] == null) {
                                // See if this is a calculated field
                                if (field.FieldTypeKind == SPTypes.FieldType.Calculated) {
                                    // Add the field and continue the loop
                                    calcFields.push(field);
                                }
                                // Else, see if this is a lookup field
                                else if (field.FieldTypeKind == SPTypes.FieldType.Lookup) {
                                    // Add the field
                                    lookupFields.push(field);
                                }
                                // Else, see if this is a MMS field
                                else if (field.TypeDisplayName == "Managed Metadata") {
                                    // Add the field
                                    mmsFields[field.InternalName] = field;
                                } else {
                                    // Append the field
                                    fields[field.InternalName] = true;
                                    cfgProps.ListCfg[0].CustomFields.push({
                                        name: field.InternalName,
                                        schemaXml: field.SchemaXml
                                    });
                                }
                            }
                        }

                        // Add the view
                        cfgProps.ListCfg[0].ViewInformation.push({
                            Default: viewInfo.DefaultView,
                            Hidden: viewInfo.Hidden,
                            JSLink: viewInfo.JSLink,
                            MobileDefaultView: viewInfo.MobileDefaultView,
                            MobileView: viewInfo.MobileView,
                            RowLimit: viewInfo.RowLimit,
                            Tabular: viewInfo.TabularView,
                            ViewName: viewInfo.Title,
                            ViewFields: viewInfo.ViewFields.Items.results,
                            ViewQuery: viewInfo.ViewQuery
                        });
                    }

                    // Update the loading dialog
                    LoadingDialog.setBody("Analyzing the lookup fields...");

                    // Parse the lookup fields
                    Helper.Executor(lookupFields, lookupField => {
                        // Skip the field, if it was already added
                        if (fields[lookupField.InternalName]) { return; }

                        // Return a promise
                        return new Promise((resolve) => {
                            // Get the lookup list
                            Web(props.srcWebUrl).Lists().getById(lookupField.LookupList).execute(
                                list => {
                                    // Add the lookup list field
                                    fields[lookupField.InternalName] = true;
                                    cfgProps.ListCfg[0].CustomFields.push({
                                        description: lookupField.Description,
                                        fieldRef: lookupField.PrimaryFieldId,
                                        hidden: lookupField.Hidden,
                                        id: lookupField.Id,
                                        indexed: lookupField.Indexed,
                                        listName: list.Title,
                                        multi: lookupField.AllowMultipleValues,
                                        name: lookupField.InternalName,
                                        readOnly: lookupField.ReadOnlyField,
                                        relationshipBehavior: lookupField.RelationshipDeleteBehavior,
                                        required: lookupField.Required,
                                        showField: lookupField.LookupField,
                                        title: lookupField.Title,
                                        type: Helper.SPCfgFieldType.Lookup
                                    } as Helper.IFieldInfoLookup);

                                    // Check the next field
                                    resolve(null);
                                },

                                err => {
                                    // Broken lookup field, don't add it
                                    console.log("Error trying to find lookup list for field '" + lookupField.InternalName + "' with id: " + lookupField.LookupList);
                                    resolve(null);
                                }
                            )
                        });
                    }).then(() => {
                        // Update the loading dialog
                        LoadingDialog.setBody("Analyzing the calculated fields...");

                        // Parse the calculated fields
                        for (let i = 0; i < calcFields.length; i++) {
                            let calcField = calcFields[i];

                            if (fields[calcField.InternalName] == null) {
                                let parser = new DOMParser();
                                let schemaXml = parser.parseFromString(calcField.SchemaXml, "application/xml");

                                // Get the formula
                                let formula = schemaXml.querySelector("Formula");

                                // Parse the field refs
                                let fieldRefs = schemaXml.querySelectorAll("FieldRef");
                                for (let j = 0; j < fieldRefs.length; j++) {
                                    let fieldRef = fieldRefs[j].getAttribute("Name");

                                    // Ensure the field exists
                                    let field = list.getField(fieldRef);
                                    if (field) {
                                        // Calculated formulas are supposed to contain the display name
                                        // Replace any instance of the internal field w/ the correct format
                                        let regexp = new RegExp(fieldRef, "g");
                                        formula.innerHTML = formula.innerHTML.replace(regexp, "[" + field.Title + "]");
                                    }
                                }

                                // Append the field
                                fields[calcField.InternalName] = true;
                                cfgProps.ListCfg[0].CustomFields.push({
                                    name: calcField.InternalName,
                                    schemaXml: schemaXml.querySelector("Field").outerHTML
                                });
                            }
                        }

                        // Parse the MMS fields
                        for (let fieldName in mmsFields) {
                            let mmsField = mmsFields[fieldName];

                            // Get the schema xml
                            let parser = new DOMParser();
                            let schemaXml = parser.parseFromString(mmsField.SchemaXml, "application/xml");

                            // Remove all properties, except for the TextField property
                            let removeProps = [];
                            let props = schemaXml.querySelector("ArrayOfProperty");
                            for (let j = props.children.length - 1; j >= 0; j--) {
                                // See if this isn't the text field property
                                let prop = props.children[j];
                                if (prop.querySelector("Name").innerHTML == "TextField") {
                                    // Find the hidden text field for this MMS field
                                    let field = list.getFieldById(prop.querySelector("Value").innerHTML);
                                    if (field && fields[field.InternalName] != true) {
                                        // Append the field
                                        fields[field.InternalName] = true;
                                        cfgProps.ListCfg[0].CustomFields.push({
                                            name: field.InternalName,
                                            schemaXml: field.SchemaXml
                                        });
                                    }
                                } else {
                                    // Remove it
                                    removeProps.push(prop);
                                }
                            }

                            // Remove the properties
                            for (let i = 0; i < removeProps.length; i++) {
                                // Remove the property
                                props.removeChild(removeProps[i]);
                            }

                            // Append the field
                            fields[mmsField.InternalName] = true;
                            cfgProps.ListCfg[0].CustomFields.push({
                                name: mmsField.InternalName,
                                schemaXml: schemaXml.querySelector("Field").outerHTML
                            });
                        }

                        // Hide the loading dialog
                        props.showDialog ? LoadingDialog.hide() : null;

                        // Resolve the request
                        resolve({
                            cfg: cfgProps,
                            lookupFields
                        });
                    });
                }
            });
        });
    }

    // Generates the lookup list data
    static generateLookupListData(props: IGenerateLookupDataProps): PromiseLike<ILookupData[]> {
        let lookupListData: ILookupData[] = [];

        // See if we are showing a loading dialog
        if (props.showDialog) {
            // Show a loading dialog
            LoadingDialog.setHeader("Analyzing Lookup Lists");
            LoadingDialog.setBody("Getting the lookup lists data...");
            LoadingDialog.show();
        }

        // Keep track of the lookup lists we are searching for
        let lookupLists: { [key: string]: boolean } = {};

        // Return a promise
        return new Promise((resolve, reject) => {
            // Parse the lookup fields
            Helper.Executor(props.lookupFields, lookupField => {
                // Ensure this lookup isn't to the source list
                if (lookupField.LookupList?.indexOf(props.srcList.Id) >= 0) { return; }

                // See if we have already checked this list
                let listId = (lookupField.LookupList || "").replace(/{|}/g, '');
                if (lookupLists[listId]) { return; }

                // Updated the loading dialog
                props.showDialog ? LoadingDialog.setBody("Getting the lookup list for: " + lookupField.InternalName) : null;

                // Return a promise
                return new Promise((resolve) => {
                    // Get the source list
                    Web(props.srcWebUrl).Lists().getById(lookupField.LookupList).execute(list => {
                        // Get the list items
                        list.Items().query({
                            GetAllItems: true,
                            Top: 5000,
                            OrderBy: ["Id"]
                        }).execute(items => {
                            let values: string[] = [];

                            // Parse the items
                            for (let i = 0; i < items.results.length; i++) {
                                // Add the value
                                let value = items.results[i][lookupField.LookupField];
                                value ? values.push(value) : null;
                            }

                            // Save the item data
                            lookupListData.push({
                                field: lookupField.LookupField,
                                items: values.sort(),
                                list: list.Title
                            });

                            // Set the flag
                            lookupLists[listId] = true;

                            // Check the next list
                            resolve(null);
                        });
                    }, () => {
                        // Check the next list
                        resolve(null);
                    });
                });
            }).then(() => {
                // Resolve the request
                resolve(lookupListData);
            }, reject);
        });
    }

    // Validates the lookup fields
    static validateLookups(props: IValidateLookupsProps): PromiseLike<Helper.ISPConfigProps> {
        // See if we are showing a loading dialog
        if (props.showDialog) {
            // Show a loading dialog
            LoadingDialog.setHeader("Validating Lookup Lists");
            LoadingDialog.setBody("Validating the required lookup lists have been added...");
            LoadingDialog.show();
        }

        // Keep track of the lookup lists we are searching for
        let lookupLists: { [key: string]: boolean } = {};

        // Return a promise
        return new Promise((resolve, reject) => {
            // Parse the lookup fields
            Helper.Executor(props.lookupFields, lookupField => {
                // Ensure this lookup isn't to the source list
                if (lookupField.LookupList?.indexOf(props.srcList.Id) >= 0) { return; }

                // See if we have already checked this list
                let listId = (lookupField.LookupList || "").replace(/{|}/g, '');
                if (lookupLists[listId]) { return; }

                // Updated the loading dialog
                props.showDialog ? LoadingDialog.setBody("Getting the lookup list for: " + lookupField.InternalName) : null;

                // Return a promise
                return new Promise((resolve, reject) => {
                    // Get the source list
                    Web(props.srcWebUrl).Lists().getById(lookupField.LookupList).execute(list => {
                        // Updated the loading dialog
                        props.showDialog ? LoadingDialog.setBody("Checking the destination web for: " + list.Title) : null;

                        // Ensure the list exists in the destination
                        Web(props.dstUrl).Lists(list.Title).execute(() => {
                            // Set the flag
                            lookupLists[list.Id] = true;

                            // Check the next list
                            resolve(null);
                        }, () => {
                            // Updated the loading dialog
                            props.showDialog ? LoadingDialog.setBody("Generating the configuration for: " + list.Title) : null;

                            // Generate the lookup list configuration
                            this.generate({
                                srcList: list,
                                srcWebUrl: props.srcWebUrl,
                                showDialog: false
                            }).then(
                                // Success
                                cfg => {
                                    // Set the flag
                                    lookupLists[list.Id] = true;

                                    // Prepend the list configuration
                                    props.cfg.ListCfg.splice(props.cfg.ListCfg.length - 1, 0, cfg.cfg.ListCfg[0]);

                                    // Check the next list
                                    resolve(null);
                                },

                                // Error
                                () => {
                                    // Reject the reqeust
                                    reject("Lookup list '" + list.Title + "' for field '" + lookupField.InternalName + "' does not exist in the configuration. Please add the lists in the appropriate order.");
                                }
                            );
                        });

                    }, () => {
                        // Reject the reqeust
                        reject("Lookup list for field '" + lookupField.InternalName + "' does not exist in the source web. Please review the source list for any issues.");
                    });
                });
            }).then(() => {
                // Resolve the request
                resolve(props.cfg);
            }, reject);
        });
    }
}