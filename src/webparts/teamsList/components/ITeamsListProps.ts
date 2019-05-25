import { MSGraphClient } from '@microsoft/sp-http';
export interface ITeamsListProps {
  graphClient: MSGraphClient;
  userEmail:string;
}
export interface ITeams{
  name: string;
  desc:string;
  member:boolean;
  owner:boolean;
  joinLink: string;
}

