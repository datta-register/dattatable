export const ItemForm: IItemForm;
export interface IItemForm {
    ListName: string;

    create(onUpdate?: Function);
    edit(itemId: number, onUpdate?: Function);
    view(itemId: number);
}