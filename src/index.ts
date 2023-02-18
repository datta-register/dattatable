import * as Common from "./common";
export * from "./common";
import { Comments } from "./comments";
export * from "./comments";
import { Dashboard } from "./dashboard";
export * from "./dashboard";
import { ActionButtonTypes } from "./common";

export const Documents = Common.Documents;
export const ItemForm = Common.ItemForm;
export const List = Common.List;

/** Styling */
import "./styles";

/** Global Variable */
window["DattaTable"] = {
    ActionButtonTypes, Common, Comments,
    Dashboard, Documents, ItemForm: Common.ItemForm,
    List: Common.List
}