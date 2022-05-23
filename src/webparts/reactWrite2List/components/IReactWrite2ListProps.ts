import { WebPartContext } from '@microsoft/sp-webpart-base'; 

export interface IReactWrite2ListProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: any | null;
}
