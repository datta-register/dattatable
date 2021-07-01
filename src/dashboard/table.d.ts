import { Components } from "gd-sprest-bs";

/**
 * Data Table
 */
export interface IDataTable {
    filter: (idx: number, value?: string) => void;
    refresh: (rows: any[]) => void;
    search: (value?: string) => void;
}

/**
 * Properties
 */
export interface IDataTableProps {
    columns: Components.ITableColumn[];
    el: HTMLElement;
    rows?: any[];
}