import * as Common from "./common";
export * from "./common";
import { Comments } from "./comments";
export * from "./comments";
import { Dashboard } from "./dashboard";
export * from "./dashboard";
import { ActionButtonTypes, Documents } from "./documents";
export * from "./documents";
import { ItemForm } from "./itemForm";
export * from "./itemForm";

/** Styling */
import "./styles";

/** Global Variable */
window["DattaTable"] = {
    ActionButtonTypes, Common, Comments, Dashboard, Documents, ItemForm
}