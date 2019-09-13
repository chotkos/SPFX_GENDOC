
 import {
     sp
 } from "@pnp/sp";

export default class TemplateService {
    //sp.web.
    templatesListName = 'Templates';

    public GetAllTemplates(): Promise<any>{
        return sp.web.lists.getByTitle(this.templatesListName)
            .items
            .getAll()
            .then(templates=>{return templates;});        
    }

}