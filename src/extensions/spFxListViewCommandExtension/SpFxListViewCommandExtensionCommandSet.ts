/* eslint-disable @typescript-eslint/no-explicit-any */
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  RowAccessor,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import { IItemAddResult } from '@pnp/sp/items';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISpFxListViewCommandExtensionCommandSetProperties {
  archiveList: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'SpFxListViewCommandExtensionCommandSet';

export default class SpFxListViewCommandExtensionCommandSet extends BaseListViewCommandSet<ISpFxListViewCommandExtensionCommandSetProperties> {
  private _sourceList: string = "";
  private _targetList: string = "";

  private _moveRows = async (selecteRows: readonly RowAccessor[]): Promise<void> => {
    if (selecteRows.length === 0) return;

    // Connect to SharePoint
    const sp: SPFI = spfi().using(SPFx(this.context));
    const sourcelist = sp.web.lists.getByTitle(this._sourceList);
    const targetList = sp.web.lists.getByTitle(this._targetList);

    selecteRows.forEach(async (row) => {
      const itemId = parseInt(row.getValueByName("ID"));

      try {
        // Get item to move
        //const item = await list.items.getById(itemId)(); // use type "any" rather than "IItem"
        const item = await sourcelist.items.getById(itemId).select("Title, Age")();
        const { Title, Age } = item;

        // Get copy item to target list
        const addItemResult: IItemAddResult = await targetList.items.add({ Title, Age });

        if (addItemResult.data.Id) {
          // Move item to recycle
          await sourcelist.items.getById(itemId).recycle();

          // Delete item from source list
          // await sourcelist.items.getById(itemId).delete()
        }

      } catch (error: any) {
        // Handle any errors that occurred during the fetch
        await Dialog.alert(`An error occurred moving '${row.getValueByName("Title")}' from '${this._sourceList}' to '${this._targetList}'`);
        console.error(`An error occurred moving '${row.getValueByName("Title")}' (ID ${itemId}) from '${this._sourceList}' to '${this._targetList}': ${error}`);
      }
    });

    await Dialog.alert("The selected items have been moved successfuly.");
  }

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized SpFxListViewCommandExtensionCommandSet');

    // initial state of the command's visibility
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    compareOneCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    this._sourceList = this.context.listView.list?.title || "";
    this._targetList = this.properties.archiveList || "";

    return Promise.resolve();
  }

  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    switch (event.itemId) {
      case 'COMMAND_1':
        if (this.context.listView.selectedRows && this.context.listView.selectedRows.length > 0) {
          await this._moveRows(this.context.listView.selectedRows);
        }

        break;
      case 'COMMAND_2':
        //await sp.web.List
        Dialog.alert(`${this.properties.sampleTextTwo}`).catch(() => {
          /* handle error */
        });
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');

    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = this.context.listView.selectedRows && this.context.listView.selectedRows.length > 0 || false;
    }

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  }
}
