import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  BasicCardView,
  IExternalLinkCardAction,
  IQuickViewCardAction
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'TicketsAceAdaptiveCardExtensionStrings';
import {
  ITicketsAceAdaptiveCardExtensionProps,
  ITicketsAceAdaptiveCardExtensionState,
  QUICK_VIEW_REGISTRY_ID
} from '../TicketsAceAdaptiveCardExtension';

export class CardView extends BaseComponentsCardView<
  ITicketsAceAdaptiveCardExtensionProps,
  ITicketsAceAdaptiveCardExtensionState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    const openCount: number = this.state.tickets.filter((t) => t.status === 'open').length;
    return BasicCardView({
      cardBar: {
        componentName: 'cardBar',
        title: this.properties.title
      },
      header: {
        componentName: 'text',
        text: `${openCount} open ticket${openCount === 1 ? '' : 's'}`
      },
      footer: {
        componentName: 'cardButton',
        title: strings.QuickViewButton,
        action: {
          type: 'QuickView',
          parameters: {
            view: QUICK_VIEW_REGISTRY_ID
          }
        }
      }
    });
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