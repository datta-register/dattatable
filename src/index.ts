import * as Common from "./common";
export * from "./common";
import { Dashboard } from "./dashboard";
export * from "./dashboard";
import { ActionButtonTypes, Documents } from "./documents";
export * from "./documents";
import { ItemForm } from "./itemForm";
export * from "./itemForm";
import { ItemFormTabs } from "./itemFormTabs";
export * from "./itemFormTabs";

/** Styling */
import "./styles";

/** Global Variable */
window["DattaTable"] = {
    ActionButtonTypes, Common, Dashboard, Documents, ItemForm//, ItemFormTabs
}