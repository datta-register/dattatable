import * as Common from "./common";
export * from "./common";
import * as Dashboard from "./dashboard";
export * from "./dashboard";
import { ItemForm } from "./itemForm";
export * from "./itemForm";

/** Styling */
import "./styles";

/** Global Variable */
window["DattaTable"] = {
    Common, Dashboard, ItemForm
}