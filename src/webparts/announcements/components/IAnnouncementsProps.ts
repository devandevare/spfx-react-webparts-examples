import { WebPartContext } from '@microsoft/sp-webpart-base';
export interface IAnnouncementsProps {
  webpartTitle: string;
  webpartLabel: string;
  listTitle: string;
  context?: WebPartContext;
  siteURL: string;
  seeAllURL: string;
}
