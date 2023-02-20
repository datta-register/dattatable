import * as Common from "./common";
import { AuditLog } from "./auditLog";
import { Comments } from "./comments";
import { Dashboard } from "./dashboard";
import { ActionButtonTypes } from "./common";
export * from "./common";
export * from "./auditLog";
export * from "./comments";
export * from "./dashboard";

/** Styling */
import "./styles";

/** Global Variable */
window["DattaTable"] = {
    ActionButtonTypes, AuditLog, Common,
    Comments, Dashboard, Documents: Common.Documents,
    ItemForm: Common.ItemForm, List: Common.List
}