import { WebPartContext } from "@microsoft/sp-webpart-base";  

export interface ICascadingDropdownDemoProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;  
  list: string,  
  fields: string[]  
}
