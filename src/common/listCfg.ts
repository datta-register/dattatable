import { Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { LoadingDialog } from "./loadingDialog";
import { List } from "./list";

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
                            let fldInfo: Types.SP.Field = list.getField(ct.FieldLinks.results[j].Name);

                            // See if this is a lookup field
                            if (fldInfo.FieldTypeKind == SPTypes.FieldType.Lookup) {
                                // Ensure this isn't an associated lookup field
                                if ((fldInfo as Types.SP.FieldLookup).IsDependentLookup != true) {
                                    // Append the field ref
                                    fieldRefs.push(fldInfo.InternalName);
                                }
                            } else {
                                // Append the field ref
                                fieldRefs.push(fldInfo.InternalName);
                            }

                            // Skip internal fields
                            if (this.InternalFields.indexOf(fldInfo.InternalName) >= 0) { continue; }

                            // See if this is a calculated field
                            if (fldInfo.FieldTypeKind == SPTypes.FieldType.Calculated) {
                                // Add the field and continue the loop
                                calcFields.push(fldInfo);
                            }
                            // Else, see if this is a lookup field
                            else if (fldInfo.FieldTypeKind == SPTypes.FieldType.Lookup) {
                                // Add the field
                                lookupFields.push(fldInfo);
                            }
                            // Else, ensure the field hasn't been added
                            else if (fields[fldInfo.InternalName] == null) {
                                // Add the field information
                                fields[fldInfo.InternalName] = true;
                                cfgProps.ListCfg[0].CustomFields.push({
                                    name: fldInfo.InternalName,
                                    schemaXml: fldInfo.SchemaXml
                                });
                            }
                        }

                        // Add the list content type
                        cfgProps.ListCfg[0].ContentTypes.push({
                            Name: ct.Name,
                            Description: ct.Description,
                            ParentName: ct.Name,
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
                let listId = (lookupField.LookupList || "").replace(/{}/g, '');
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

                                    // Append the list configuration
                                    props.cfg.ListCfg.push(cfg.cfg.ListCfg[0]);

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