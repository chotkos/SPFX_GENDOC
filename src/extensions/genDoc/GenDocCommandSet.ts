import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog'; 

import * as strings from 'GenDocCommandSetStrings';
import { autobind, assign } from 'office-ui-fabric-react';
import CustomPanel, { ICustomPanelProps } from './CustomPanel/CustomPanel';

import * as React from 'react';
import * as ReactDom from 'react-dom';	
import {
  sp
} from "@pnp/sp";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IGenDocCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'GenDocCommandSet';

export default class GenDocCommandSet extends BaseListViewCommandSet<IGenDocCommandSetProperties> {

  private panelPlaceHolder: HTMLDivElement = null;	

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized GenDocCommandSet');

  
    
    // @Prezentacja_1_SPFX_4: This is the place we created for the panel
    this.panelPlaceHolder = document.body.appendChild(document.createElement("div"));	
  

    //return Promise.resolve();
    return super.onInit().then(_ => {      
      sp.setup({spfxContext: this.context});
    });
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // @Prezentacja_1_SPFX_1: Show command only if item is selected
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':   
        // @Prezentacja_1_SPFX_2: get item data from selected item      
        let selectedItem = event.selectedRows[0];	        
        const listItemId = selectedItem.getValueByName('ID') as number;	        
        const title = selectedItem.getValueByName("Title");	        
        this._showPanel(listItemId, title);	        
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private _showPanel(itemId: number, currentTitle: string) {	
    // @Prezentacja_1_SPFX_3: render the panel and show it :)      
    this._renderPanelComponent({	      
      isOpen: true,	      
      currentTitle,	      
      itemId,	      
      listId: this.context.pageContext.list.id.toString(),	      
      onClose: this._dismissPanel	    
    });	  
  }

  @autobind	  
  private _dismissPanel() {
    this._renderPanelComponent({ isOpen: false });
  }

  private _renderPanelComponent(props: any) {	    
    // @Prezentacja_1_SPFX_3: rendering react component :)      

    const element: React.ReactElement<ICustomPanelProps> = React.createElement(CustomPanel, assign({
              onClose: null,	      
              currentTitle: null,	      
              itemId: null,	      
              isOpen: false,	      
              listId: null	    
            }, props));	    

    ReactDom.render(element, this.panelPlaceHolder);	  
  }
}
