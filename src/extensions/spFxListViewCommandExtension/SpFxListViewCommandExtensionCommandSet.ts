/* eslint-disable @typescript-eslint/no-explicit-any */
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
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


  private _moveItem = async (itemId: number): Promise<void> => {
    const sourceList: string = this.context.listView.list?.title || "";
    const targetList: string = this.properties.archiveList || "";

    // Connect to SharePoint
    const sp: SPFI = spfi().using(SPFx(this.context));

    try {
      // Get item to move
      //const item = await sp.web.lists.getByTitle(sourceList).items.getById(itemId)(); // use type "any" rather than "IItem"
      const item = await sp.web.lists.getByTitle(sourceList).items.getById(itemId).select("Title, Age")();
      const { Title, Age } = item;

      // Get copy item to target list
      const addItemResult: IItemAddResult = await sp.web.lists.getByTitle(targetList).items.add({ Title, Age });

      //console.log(addItemResult.data.Id)
      if (addItemResult.data.Id) {

        // Delete item from source list
        await sp.web.lists.getByTitle(sourceList).items.getById(itemId).delete();

        // Refresh source list data
        await sp.web.lists.getByTitle(sourceList).items.getAll();
        await Dialog.alert(`Item with title '${Title}' has been archived successfully.`);
      }

    } catch (error: any) {
      // Handle any errors that occurred during the fetch
      console.error(`An error occured moving item ${itemId} from '${sourceList}' to '${targetList}': ${error}`);
    }
  }

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized SpFxListViewCommandExtensionCommandSet');

    // initial state of the command's visibility
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    compareOneCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    switch (event.itemId) {
      case 'COMMAND_1':
        if (this.context.listView.selectedRows?.length === 1) {
          await this._moveItem(parseInt(this.context.listView.selectedRows[0].getValueByName("ID")));
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
      compareOneCommand.visible = this.context.listView.selectedRows?.length === 1;
    }

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  }
}
