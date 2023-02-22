import { Helper, SPTypes } from "gd-sprest-bs";

// Configuration
export const Configuration: Helper.ISPConfig = Helper.SPConfig({
    ListCfg: [{
        ListInformation: {
            BaseTemplate: SPTypes.ListTemplateType.GenericList,
            Title: "Audit Log",
            Hidden: true,
            NoCrawl: true
        },
        TitleFieldIndexed: true,
        ContentTypes: [{
            Name: "Item",
            FieldRefs: [
                "Title",

            ]
        }],
        CustomFields: [
            {
                name: "ParentIId",
                title: "Parent Id",
                description: "The parent identifier linked to this item.",
                type: Helper.SPCfgFieldType.Text,
                indexed: true
            },
            {
                name: "ParentListName",
                title: "Parent List Name",
                description: "The list name linked to this item.",
                type: Helper.SPCfgFieldType.Text,
                indexed: true
            },
            {
                name: "LogComment",
                title: "LogComment",
                description: "The comment to write to the audit log.",
                type: Helper.SPCfgFieldType.Note,
                noteType: SPTypes.FieldNoteType.TextOnly
            } as Helper.IFieldInfoNote,
            {
                name: "LogData",
                title: "Log Data",
                description: "The data to write to the audit log.",
                type: Helper.SPCfgFieldType.Note,
                noteType: SPTypes.FieldNoteType.TextOnly
            } as Helper.IFieldInfoNote,
            {
                name: "LogUser",
                title: "Log User",
                description: "The user who is associated with this log entry.",
                type: Helper.SPCfgFieldType.User
            }
        ]
    }]
});