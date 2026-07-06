import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView } from '@microsoft/sp-adaptive-card-extension-base';
import {
  ITicketsAceAdaptiveCardExtensionProps,
  ITicketsAceAdaptiveCardExtensionState,
  ITicket
} from '../TicketsAceAdaptiveCardExtension';

export interface IQuickViewData {
  tickets: ITicket[];
}

export class QuickView extends BaseAdaptiveCardQuickView<
  ITicketsAceAdaptiveCardExtensionProps,
  ITicketsAceAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    return {
      tickets: this.state.tickets
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }
}