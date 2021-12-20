declare interface IPersonalEmailWebPartStrings {
  PropertyPaneDescription: string;
  Error: string;
  Loading: string;
  NewEmail: string;
  NoMessages: string;
  NumOfMessagesToShow: string;
  ViewAll: string;
}

declare module 'PersonalEmailWebPartStrings' {
  const strings: IPersonalEmailWebPartStrings;
  export = strings;
}
