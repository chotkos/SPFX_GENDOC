
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

    public UpdateTemplate(model:any):Promise<boolean>{
        return sp.web.lists.getByTitle(this.templatesListName)
        .items
        .getById(model.ID)
        .update(model)
        .then(e => true).catch(error => {
            return error.message;
        });

    }

    public CreateTemplate(model:any):Promise<boolean>{
        return sp.web.lists.getByTitle(this.templatesListName)
        .items
        .add(model)
        .then(e => true).catch(error => {
            return error.message;
        });
    }
}