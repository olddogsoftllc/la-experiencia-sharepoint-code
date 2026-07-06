import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';

const LOG_SOURCE: string = 'ActionsCommandSetCommandSet';

export interface IActionsCommandSetCommandSetProperties {
  /** Prefix shown by the "Notify count" command. */
  notifyPrefix?: string;
}

/**
 * ListView Command Set with two context buttons:
 *   COMMAND_1 "Export selection" — visible only when >= 1 row is selected.
 *   COMMAND_2 "Notify count"     — always visible; shows how many rows are selected.
 *
 * Covers the book's cap14 "Tipos de Extensiones" (Command Set).
 */
export default class ActionsCommandSetCommandSet extends BaseListViewCommandSet<IActionsCommandSetCommandSetProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ActionsCommandSet');

    const exportCommand: Command | undefined = this.tryGetCommand('COMMAND_1');
    if (exportCommand) {
      exportCommand.visible = false; // hidden until rows are selected
    }

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    const selectedCount: number = this.context.listView.selectedRows?.length ?? 0;

    switch (event.itemId) {
      case 'COMMAND_1':
        Dialog.alert(`Export: ${selectedCount} row(s) selected (demo).`).catch(() => {
          /* dialog dismissed */
        });
        break;

      case 'COMMAND_2': {
        const prefix: string = this.properties.notifyPrefix ?? 'Selected';
        Dialog.alert(`${prefix}: ${selectedCount} row(s).`).catch(() => {
          /* dialog dismissed */
        });
        break;
      }

      default:
        throw new Error(`Unknown command: ${event.itemId}`);
    }
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');

    const exportCommand: Command | undefined = this.tryGetCommand('COMMAND_1');
    if (exportCommand) {
      // Visible only when at least one row is selected.
      exportCommand.visible = (this.context.listView.selectedRows?.length ?? 0) > 0;
    }

    // Refresh the command bar so visibility changes take effect.
    this.raiseOnChange();
  }
}