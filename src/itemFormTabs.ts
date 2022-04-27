import { Components } from "gd-sprest-bs";
import { ItemForm, IItemFormCreateProps, IItemFormEditProps, IItemFormViewProps } from "./itemForm";

export interface IItemFormTabInfo {
    name: string;
    fields: string[];
}

export interface IItemFormTabsProps {
    createProps: IItemFormCreateProps;
    editProps?: IItemFormEditProps;
    el: HTMLElement;
    isVertical?: boolean;
    tabs: IItemFormTabInfo[]
    viewProps?: IItemFormViewProps;
}

/**
 * Item Form w/ Tabs
 */
export class ItemFormTabs {
    // Properties
    private _props: IItemFormTabsProps = null;

    // Horizontal Tabs
    private _horizontalTabs: Components.INav = null;

    // Vertical Tabs
    private _verticalTabs: Components.IListGroup = null;

    // Item Form reference
    private _itemForms: ItemForm[] = null;
    get ItemForms(): ItemForm[] { return this._itemForms; }
    ItemFormByTab(name: string) {
        // Parse the tab
    }

    // Constructor
    constructor(props: IItemFormTabsProps) {
        // Save the properties
        this._props = props;

        // Render the component
        this.render();
    }

    // Generates a tab
    private generateTab(tabInfo: IItemFormTabInfo, isVertical: boolean): Components.IListGroupItem | Components.INavLinkProps {
        // See if this is a vertical tab
        if (isVertical) {
            // Generate a list group item tab
            return {
                data: tabInfo.fields,
                tabName: tabInfo.name,
                onRender: (el) => {
                    // Render the form component
                    // TODO
                }
            } as Components.IListGroupItem;
        }

        // Generate a nav link tab
        return {
            data: tabInfo.fields,
            title: tabInfo.name,
            onRenderTab: (el, item) => {
                // Render the form component
                // TODO
            }
        } as Components.INavLinkProps
    }

    // Renders the component
    private render() {
        // Create the item form
        ItemForm.create()

        // See if we are rendering vertical tabs
        let isVertical = this._props.isVertical ? true : false;
        if (isVertical) {
            let tabs: Components.IListGroupItem[] = [];

            // Parse the tabs
            for (let i = 0; i < this._props.tabs.length; i++) {
                // Add the tab
                tabs.push(this.generateTab(this._props.tabs[i], isVertical) as Components.IListGroupItem);
            }

            // Render the tabs
            Components.ListGroup({
                el: this._props.el,
                isTabs: true,
                items: tabs
            });
        } else {
            let tabs: Components.INavLinkProps[] = [];

            // Parse the tabs
            for (let i = 0; i < this._props.tabs.length; i++) {
                // Add the tab
                tabs.push(this.generateTab(this._props.tabs[i], isVertical) as Components.INavLinkProps);
            }

            // Render the tabs
            Components.Nav({
                el: this._props.el,
                isPills: true,
                isTabs: true,
                items: tabs
            });
        }
    }
}