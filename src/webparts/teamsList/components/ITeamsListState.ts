import { ITeams } from './ITeamsListProps';
export interface ITeamsListState {
  teamList:ITeams[];
  error: string;
  loading: boolean;
  font:number;
  theme:any;

}
