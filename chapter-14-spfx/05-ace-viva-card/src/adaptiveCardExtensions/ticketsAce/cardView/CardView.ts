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
import { countOpenTickets, buildOpenTicketsHeader } from '../ticketsUtils';

export class CardView extends BaseComponentsCardView<
  ITicketsAceAdaptiveCardExtensionProps,
  ITicketsAceAdaptiveCardExtensionState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    const openCount: number = countOpenTickets(this.state.tickets);
    return BasicCardView({
      cardBar: {
        componentName: 'cardBar',
        title: this.properties.title
      },
      header: {
        componentName: 'text',
        text: buildOpenTicketsHeader(openCount)
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