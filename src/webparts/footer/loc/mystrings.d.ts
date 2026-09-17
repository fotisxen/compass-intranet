declare interface IFooterWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  CompanyNameFieldLabel: string;
  StockApiUrlFieldLabel: string;
}

declare module 'FooterWebPartStrings' {
  const strings: IFooterWebPartStrings;
  export = strings;
}
