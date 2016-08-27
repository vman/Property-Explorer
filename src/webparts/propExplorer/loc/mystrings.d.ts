declare interface IStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  DescriptionFieldLabel: string;
  ReactiveFieldLabel: string;
}

declare module 'mystrings' {
  const strings: IStrings;
  export = strings;
}
