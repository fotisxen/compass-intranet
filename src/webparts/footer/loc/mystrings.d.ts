declare interface IFooterWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  CompanyNameFieldLabel: string;
  StockApiUrlFieldLabel: string;
  EuronextApiUrlFieldLabel: string;
  AlignmentGroupName: string;
  ContentOffsetXFieldLabel: string;
}

declare module 'FooterWebPartStrings' {
  const strings: IFooterWebPartStrings;
  export = strings;
}
