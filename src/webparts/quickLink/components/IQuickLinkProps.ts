import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IQuickLinkProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;


  SiteURL:string;
  listName:string;
  context:WebPartContext;
  viewType:boolean;
  NumbrtofColumnsToShow:any;
  componentTitle:string;
  emptyMessage:string;
}
