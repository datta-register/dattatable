/** Library */
export * from "./lib";

/** Styling */
import "./styles";

/** Global Variable */
import { DattaTable } from "./lib";
window["DattaTable"] = DattaTable;

/** Apply the Current Theme */
import { ThemeManager } from "gd-sprest-bs";
ThemeManager.load(true);