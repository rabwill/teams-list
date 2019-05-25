import * as React from 'react';
import { ITeamsListProps, ITeams } from './ITeamsListProps';
import { ITeamsListState } from './ITeamsListState';
import * as q from 'q';
import {
  DocumentCard,
  DocumentCardTitle,IDocumentCardStyles
} from 'office-ui-fabric-react/lib/DocumentCard';
import {
  PrimaryButton,
  TeamsComponentContext,
  ConnectedComponent,
  Panel,
  PanelBody,
  PanelHeader,
  PanelFooter,
  Surface

} from 'msteams-ui-components-react';
import styles from './TeamsList.module.scss';
export default class TeamsList extends React.Component<ITeamsListProps, ITeamsListState> {
  constructor(props: ITeamsListProps) {
    super(props);
    this.state = {
      teamList: [],
      error: null,
      loading: true,
      font: 14,
      theme: null,
    
    };
  }


  /*map the given teams to suitable properties
  */
  private async _processTeamsList(allTeams: any[]): Promise<void> {
    const teamsNames: any[] = await allTeams.map(async w => {
      return {
        name: w.displayName,
        desc:w.description,
        member: await this.isMember(w.id),
        owner: await this.isOwner(w.id),
        joinLink: await this.getJoinLink(w.id)
      };
    });
    /* resolve all promises
    */
    var promisedRequests = await Promise.all(teamsNames);
    this.setState({
      teamList: promisedRequests,
      loading: false
    });
  }
  /*create a join link
  */
  private async getJoinLink(teamId): Promise<any> {
    var p = new Promise<any>((resolve, reject) => {
      //get a link similar to join link from the General channel webUrl
      this.props.graphClient.api(`/teams/${teamId}/channels?$filter=displayName eq 'General'`)
        .get((err, res: any) => {
          if (err) {
            resolve("/");
          }
          // Check if a response was retrieved
          if (res && res.value && res.value.length > 0) {
            //replace the necessary parts of the link to make it a join link
            resolve(res.value[0].webUrl.replace("General", "conversations").replace("https://teams.microsoft.com/l/channel/", "https://teams.microsoft.com/l/team/"));
          } else {
            resolve("/");
          }
        });
    });
    return p;
  }
  /*check if user is a member of the team
  */
  private async isMember(teamId): Promise<any> {
    var p = new Promise<any>((resolve, reject) => {
      this.props.graphClient.api(`/groups/${teamId}/members`)
        .get((err, res: any) => {
          if (err) {
            resolve(false);
          }
          // Check if a response was retrieved
          if (res && res.value && res.value.length > 0) {
            res.value.filter(item => {
              return item.userPrincipalName === this.props.userEmail;
            }).length > 0 ? resolve(true) : resolve(false);
          } else {
            resolve(false);
          }
        });

    });
    return p;
  }

  /*check if user is an owner of the team
 */
  private async isOwner(teamId): Promise<any> {
    var deffered = q.defer();
    if (this.props.graphClient) {

      this.props.graphClient.api(`/groups/${teamId}/owners`)
        .get((err, res: any) => {
          if (err) {
            deffered.reject(false);
          }
          // Check if a response was retrieved
          if (res && res.value && res.value.length > 0) {
            var newA = res.value.filter(item => {
              return item.userPrincipalName === this.props.userEmail;
            });

            if (newA.length > 0) {
              deffered.resolve(true);
            } else {
              deffered.resolve(false);
            }
          }
        });
    }
    return deffered.promise;
  }
  /* fetch all the public list of teams in the tenant
  */
  private async _fetchListOfTeams() {
    if (this.props.graphClient) {
      this.setState({
        loading: true
      });
      this.props.graphClient
        .api("/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')")
        .version("beta")
        .get(async (err, res: any) => {
          if (err) {
            // Something failed calling the MS Graph
            this.setState({
              error: err.message ? err.message : "Error: something failed in the MS Graph call to get all teams",
              teamList: [],
              loading: false
            });
            return;
          }


          // Check if a response was retrieved
          if (res && res.value && res.value.length > 0) {
            await this._processTeamsList(res.value);
          } else {
            // No sites retrieved
            this.setState({
              loading: false,
              teamList: []
            });
          }
        });
    }
  }

  public componentDidMount(): void {
    this._fetchListOfTeams();
  }
  public render(): React.ReactElement<ITeamsListProps> {

    return (
      <TeamsComponentContext
        fontSize={this.state.font}
        theme={this.state.theme}
      >

        <ConnectedComponent render={(props) => {
          const { context } = props;
          const { rem, font } = context;
          const { sizes, weights } = font;
          const stylesteams = {
            header: { ...sizes.title, ...weights.semibold },
            section: { ...sizes.base, marginTop: rem(1.4), marginBottom: rem(1.4) },
            footer: { ...sizes.xsmall }
          };
          const cardStyles: IDocumentCardStyles = {
            root: { display: 'inline-block', marginRight: 20, marginBottom: 20, width: 320 }
          };
         
          return (<Surface>
            <Panel>
              <PanelHeader>
                <div style={stylesteams.header}> List of all the teams</div>
              </PanelHeader>
              <PanelBody>
                <div>{
                  this.state.teamList && this.state.teamList.length > 0 ? (
                    <div className={styles.listWrapper}>
                      <ul>
                        {
                          this.state.teamList.map(team => (
                            <li key={team.name}>
                            <DocumentCard styles={cardStyles}  onClickHref={team.joinLink}>
                              <DocumentCardTitle title={team.name} shouldTruncate />
                              <DocumentCardTitle title={team.desc} shouldTruncate />
                              
                              <div className={styles.listCard} >
                              {(team.owner || team.member) ? 
                               <PrimaryButton>Click to join
                               </PrimaryButton>
                              : <DocumentCardTitle className={styles.titlecardHeight} title="Already joined" />}
                              </div>
                              </DocumentCard>
                            </li>
                          ))
                        }
                      </ul>
                    </div>
                  ) : (
                      !this.state.loading && (
                        this.state.error ?
                          <span>{this.state.error}</span> :
                          <span>Nothing to display!</span>
                      )
                    )
                }
                </div>
              </PanelBody>
              <PanelFooter>
                <div style={stylesteams.footer}>
                  (C) Rabia Williams
                                  </div>
              </PanelFooter>
            </Panel>
          </Surface>

          );
        }}>
        </ConnectedComponent>
      </TeamsComponentContext >
    );
  }
}
