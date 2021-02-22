import { BaseClientSideWebPart, IPropertyPaneConfiguration } from '@microsoft/sp-webpart-base';
export interface ITaskListWebPartProps {
    listTitle: string;
    listViewName: string;
    viewMode: string;
    applyButton: any;
    context: any;
}
export default class TaskListWebPart extends BaseClientSideWebPart<ITaskListWebPartProps> {
    private _id;
    private _componentElement;
    private _koListTitle;
    private _koListViewName;
    private taskLists;
    private taskListsDropdownDisabled;
    private listViews;
    private listViewMode;
    private listViewDropdownDisabled;
    private ApplyChanges();
    private loadLists();
    private loadItems();
    protected onPropertyPaneConfigurationStart(): void;
    protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void;
    /**
     * Shouter is used to communicate between web part and view model.
     */
    private _shouter;
    private _self;
    /**
     * Initialize the web part.
     */
    protected onInit(): Promise<void>;
    render(): void;
    private _createComponentElement(tagName);
    private _registerComponent(tagName);
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
