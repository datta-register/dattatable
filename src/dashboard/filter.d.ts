import { Components } from "gd-sprest-bs";

/**
 * Filter Item
 */
export interface IFilterItem {
    header: string;
    items: Components.ICheckboxGroupItem[];
    onFilter?: (value: string) => void;
}