import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { TicketsAcePropertyPane } from './TicketsAcePropertyPane';

export interface ITicketsAceAdaptiveCardExtensionProps {
  title: string;
}

export interface ITicket {
  id: string;
  title: string;
  status: 'open' | 'inProgress' | 'done';
}

export interface ITicketsAceAdaptiveCardExtensionState {
  tickets: ITicket[];
}

const CARD_VIEW_REGISTRY_ID: string = 'TicketsAce_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'TicketsAce_QUICK_VIEW';

/**
 * Adaptive Card Extension for Viva Connections: shows a count of open tickets
 * on the dashboard card and a list in the QuickView.
 *
 * Covers the book's chapter 14 "Adaptive Card Extensions (ACE) in detail".
 * The ticket list is seeded with demo data; wire `loadTickets()` to Microsoft
 * Graph (`/me/planner/tasks`, scope `Tasks.Read`) for a real feed.
 */
export default class TicketsAceAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ITicketsAceAdaptiveCardExtensionProps,
  ITicketsAceAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: TicketsAcePropertyPane;

  public onInit(): Promise<void> {
    this.state = {
      tickets: [
        { id: '1', title: 'Review PR #142', status: 'open' },
        { id: '2', title: 'Fix login redirect', status: 'inProgress' },
        { id: '3', title: 'Update dependencies', status: 'open' }
      ]
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'TicketsAce-property-pane'*/
      './TicketsAcePropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.TicketsAcePropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }
}