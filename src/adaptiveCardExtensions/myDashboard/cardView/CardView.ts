import {
  BasePrimaryTextCardView,
  IPrimaryTextCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'MyDashboardAdaptiveCardExtensionStrings';
import { IMyDashboardAdaptiveCardExtensionProps, IMyDashboardAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../MyDashboardAdaptiveCardExtension';

export class CardView extends BasePrimaryTextCardView<IMyDashboardAdaptiveCardExtensionProps, IMyDashboardAdaptiveCardExtensionState> {
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    return [
      {
        title: strings.QuickViewButton,
        action: {
          type: 'QuickView',
          parameters: {
            view: QUICK_VIEW_REGISTRY_ID
          }
        }
      }
    ];
  }

  public get data(): IPrimaryTextCardParameters {
    return {
      primaryText: strings.PrimaryText,
      description: strings.Description + " Cards: " + this.state.cardAudienceItems.length.toString(),
      title: this.properties.title 
    };
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'QuickView',
        parameters: {
          view: QUICK_VIEW_REGISTRY_ID
        }
    };
  }
}
