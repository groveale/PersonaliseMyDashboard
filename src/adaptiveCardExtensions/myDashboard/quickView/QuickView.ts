import { ISPFxAdaptiveCard, BaseAdaptiveCardView, IActionArguments, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'MyDashboardAdaptiveCardExtensionStrings';
import { IMyDashboardAdaptiveCardExtensionProps, IMyDashboardAdaptiveCardExtensionState } from '../MyDashboardAdaptiveCardExtension';
import { graphService } from '../services/MSGraphService';

export interface IQuickViewData {
  subTitle: string;
  title: string;
  cards: any[];
}

export class QuickView extends BaseAdaptiveCardView<
  IMyDashboardAdaptiveCardExtensionProps,
  IMyDashboardAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    return {
      subTitle: strings.SubTitle,
      title: strings.Title,
      cards: this.state.cardAudienceItems
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/DynamicView.json');
  }

  public async onAction(action: IActionArguments): Promise<void> {
    if ((<ISubmitActionArguments>action).type === 'Submit') {
      const submitAction = <ISubmitActionArguments>action;
      const { id } = submitAction.data;
      if (id === 'save' || id === 'reject') {
        if (id === 'save') {
          // save the membership updates
          this.state.cardAudienceItems.forEach(async cardAud => {
            if (submitAction.data.hasOwnProperty(cardAud.identifier)) {
              // membership may has changed
              // check if state is different
              if (cardAud.isMember !== submitAction.data[cardAud.identifier])
              {
                // update the membership
                console.log("Membership changed for " + cardAud.identifier);
                if (submitAction.data[cardAud.identifier] === 'true') {
                  console.log("Adding user to group");
                  await graphService.AddLoggedInUserToGroup(cardAud.groupId);
                  // update the state
                  this.updateItemMembershipState(cardAud.identifier, 'true');
                } else {
                  console.log("Removing user from group");
                  await graphService.RemoveLoggedInUserFromGroup(cardAud.groupId);
                  // update the state
                  this.updateItemMembershipState(cardAud.identifier, 'false');
                }
              } else {
                console.log("Membership unchanged for " + cardAud.identifier);
              }
            }
          });
        }
        else {
          // rerender the card with stored data
        }
        
      }
    }
  }

  private updateItemMembershipState(identifier: string, isMember: string) : void {
    this.setState({
      cardAudienceItems: this.state.cardAudienceItems.map(obj => {
        if(obj.identifier === identifier) {
          obj.isMember = isMember
        }
        return obj
      })
    })
  }
}