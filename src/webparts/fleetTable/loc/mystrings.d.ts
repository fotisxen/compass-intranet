declare interface IFleetTableWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ListTitleFieldLabel: string;
  SortByFieldLabel: string;
  SortDescendingFieldLabel: string;
  MaxRowsFieldLabel: string;
}

declare module 'FleetTableWebPartStrings' {
  const strings: IFleetTableWebPartStrings;
  export = strings;
}
