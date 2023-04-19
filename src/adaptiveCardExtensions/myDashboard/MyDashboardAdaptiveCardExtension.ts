import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { MyDashboardPropertyPane } from './MyDashboardPropertyPane';
import { ICardAudienceItem } from './models/ICardAudeinceItem';
import { graphService } from './services/MSGraphService';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http'

export interface IMyDashboardAdaptiveCardExtensionProps {
  title: string;
  listTitle: string
}

export interface IMyDashboardAdaptiveCardExtensionState {
  cardAudienceItems: ICardAudienceItem[];
}

const CARD_VIEW_REGISTRY_ID: string = 'MyDashboard_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'MyDashboard_QUICK_VIEW';

export default class MyDashboardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IMyDashboardAdaptiveCardExtensionProps,
  IMyDashboardAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: MyDashboardPropertyPane | undefined;

  public async onInit(): Promise<void> {
    this.state = {
      cardAudienceItems: [] 
     };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    const graphClient = await this.context.msGraphClientFactory.getClient("3");

    graphService.Init(graphClient, this.context.pageContext.legacyPageContext.aadUserId);

    // Get Config from SPO
    await this.getAudeinceConfigItemsFromSPOAndUserMembershipFromGraph()

    // get users details from MS Graph
    //return graphService.GetLoggedInUsersDirectoryId()
    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'MyDashboard-property-pane'*/
      './MyDashboardPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.MyDashboardPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }

  private getAudeinceConfigItemsFromSPOAndUserMembershipFromGraph(): Promise<void> {
    if (this.properties.listTitle) {
      return this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}` +
        // https://groverale.sharepoint.com/sites/home/_api/web/lists/getByTitle('DashboardCardsPersonalisation')/items?$filter=(ShowInCard%20eq%201)&$select=Title,CardIdentifier,Mandatory,Audience/Title&$expand=Audience
        `/_api/web/lists/getByTitle('${this.properties.listTitle}')/items?$filter=(ShowInCard%20eq%201)&$select=Title,CardIdentifier,Mandatory,Audience/Title&$expand=Audience`,
        SPHttpClient.configurations.v1
      )
      .then((response) => response.json())
      .then((jsonResponse) => jsonResponse.value.map(
         (item: any) => { 
          
          return {
            title: item.Title, 
            identifier: item.CardIdentifier,
            groupName: item.Audience.Title,
            groupId: "groupId", 
            locked: item.Mandatory,
            isMember: "isMember.toString()"
          }; 
        })
      )
      .then((items) => this.setState(
        { 
          cardAudienceItems: items 
        }))
      .then(() => {
        // Loop through all the items and get the group id and check if the user is a member
        this.state.cardAudienceItems.forEach(async (item) => {
          const group = await graphService.GetGroupWithName(item.groupName);
          
          // group should allways be found as group field is used in SPO
          // if group is not found, set the group id to empty string
          const groupId = group ? group.id : "";

          let isMember = false
          if (groupId !== "") {
            isMember = await graphService.IsLoggedInUserAMember(groupId);
          }

          // method to update list items based on group id and isMember
          this.updateItemListState(item.identifier, groupId, isMember.toString());
        })})
    }

    return Promise.resolve();
  }

  private updateItemListState(identifier: string, groupId: string, isMember: string) : void {
    this.setState({
      cardAudienceItems: this.state.cardAudienceItems.map(obj => {
        if(obj.identifier === identifier) {
          obj.groupId = groupId
          obj.isMember = isMember
        }
        return obj
      })
    })
  }
}
