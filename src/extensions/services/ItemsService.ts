 import {
    sp
} from "@pnp/sp";
import { ThemeGenerator } from "office-ui-fabric-react";

export default class ItemsService {
    
   itemsListName = 'Gen_Invoice';

   public GetAllFields(): Promise<any>{
        return sp.web.lists.getByTitle(this.itemsListName)
            .fields
            .filter('Hidden eq false') //@Prezentacja_3_PNP_2
            .get()                   
            .then(fields=>{return fields;})
            .catch(e=>{throw e;});        
   }

   public GetListItemById(itemId):Promise<any>{
        return sp.web.lists.getByTitle(this.itemsListName)
            .items
            .getById(parseInt(itemId))
            .get()
            .then(item=>{return item;})
            .catch(e=>{throw e;});  
   }
 
}