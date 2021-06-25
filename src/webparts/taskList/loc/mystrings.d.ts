declare interface ITaskListWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  DueDate: string;
  Goedkeuring: string;
  Opmerking: string;
  Kostenplaats: string;
  StatusWPLU: string;
  TerugkoppelingCO: string;
  AssignedTo: string;
}

declare module 'TaskListWebPartStrings' {
  const strings: ITaskListWebPartStrings;
  export = strings;
}
