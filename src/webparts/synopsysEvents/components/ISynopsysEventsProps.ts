import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISynopsysEventsProps {
  webpartTitle: string;
  webpartLabel: string;
  siteURL: string;
  seeAllURL: string;
  eventListName: string;
  context?: WebPartContext;
}
