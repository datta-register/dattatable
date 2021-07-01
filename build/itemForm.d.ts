import { Components } from "gd-sprest-bs";
/**
 * Item Form
 */
declare class _ItemForm {
    private _displayForm;
    private _editForm;
    private _updateEvent;
    get form(): Components.IListFormDisplay | Components.IListFormEdit;
    private _listName;
    get ListName(): string;
    set ListName(value: string);
    /** Public Methods */
    create(onUpdate?: Function): void;
    edit(itemId: number, onUpdate?: () => void): void;
    view(itemId: number, onUpdate?: () => void): void;
    /** Private Methods */
    private load;
    private save;
}
export declare const ItemForm: _ItemForm;
export {};
