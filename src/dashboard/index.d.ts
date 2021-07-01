import { Components } from "gd-sprest-bs";
import { IFilterItem } from "./filter";

// Dashboard
export interface IDashboardProps {
    columns: Components.ITableColumn[];
    el: HTMLElement;
    footer?: {
        items?: Components.INavbarItem[];
        itemsEnd?: Components.INavbarItem[];
    };
    filters?: IFilterItem[];
    header?: {
        title?: string;
    }
    navigation?: {
        title?: string;
        items?: Components.INavbarItem[];
        itemsEnd?: Components.INavbarItem[];
    };
    rows?: any[];
}