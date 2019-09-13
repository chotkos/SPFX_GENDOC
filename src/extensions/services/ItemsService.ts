 import {
    sp
} from "@pnp/sp";

export default class ItemsService {
   //sp.web.
   itemsListName = 'Gen_Invoice';

   public GetAllFields(): Promise<any>{
       return sp.web.lists.getByTitle(this.itemsListName)
            .fields
            .filter('Hidden eq false') //@Prezentacja
            .get()                   
            .then(fields=>{return fields;})
            .catch(e=>{throw e;});        
   }
 
}