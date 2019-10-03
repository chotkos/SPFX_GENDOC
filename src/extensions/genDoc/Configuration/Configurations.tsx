import * as React from 'react';
import { TextField, DefaultButton, PrimaryButton, DialogFooter, autobind, Panel, Spinner, SpinnerType, Dropdown, PanelType } from "office-ui-fabric-react";
import IConfigurationState from './IConfigurationState';
import IConfigurationProps from './IConfigurationProps';
import ReactQuill from 'react-quill'; // ES6
import 'react-quill/dist/quill.snow.css'; // ES6
import styles from './Configurations.module.scss'; 
import { Dialog } from '@microsoft/sp-dialog';
import TemplateService from '../../services/TemplateService';
import ItemsService from '../../services/ItemsService';

export default class Configuration extends React.Component<IConfigurationProps, IConfigurationState> {

    
    private templateService = new TemplateService();
    private itemsService = new ItemsService();
    //@Prezentacja_2_Quill_2 Konfiguracja edytora
    modules = {
        toolbar: [
            [{ 'header': [1, 2, 3, 4, 5, 6, false] }],
            [{ 'font': [] }],
            [{ 'align': [] }, { 'direction': 'rtl' }, { 'color': [] }],
            ['table'],
            ['bold', 'italic', 'underline', 'blockquote', 'size'],
            [{ 'list': 'ordered' }, { 'list': 'bullet' }, { 'indent': '-1' }, { 'indent': '+1' }],
            ['link', 'image']
        ],
    };
    
    formats = [
        'header', 'size', 'font', 'align', 'direction', 'color',
        'bold', 'italic', 'underline', 'blockquote',
        'list', 'bullet', 'indent',
        'link', 'image'
    ];

    quillRef=null;

    constructor(props: IConfigurationProps) {
        super(props);
        this.state = {
            template: props.template,
            templateName: props.template ? props.template.Title : '',
            templateContent:props.template ? props.template.Template : '',
            fields:[]
        };    
    
        this.quillRef = React.createRef();

        this.setupFields();
    }

    @autobind
    private setupFields(){
        this.itemsService.GetAllFields().then((fields:any[])=>{
            console.log(fields);
            
            let newFields = fields.map(f=>{
                return {title: f.Title, name:f.InternalName};
            });

            this.setState({fields: newFields});
        });
    }

    @autobind
    private _saveConfig(){
        //@TODO: saving pnp        
        let updateModel = this.state.template;
        updateModel.Title = this.state.templateName;
        updateModel.Template = this.state.templateContent;

        this.templateService.UpdateTemplate(updateModel).then(r=>{if(r){
            this._closePanel();
        }else{
            alert('Failed to save');
        }});        
    }
 
    //@Prezentacja_1_SPFX_7 Zapis templatki :)
    @autobind
    private _createConfig(){  
               
        let createModel = {
            Title : this.state.templateName,
            Template : this.state.templateContent
        }

        //@Prezentacja_3_PNP_1 Przykład użycia PNP.js
        this.templateService.CreateTemplate(createModel).then(r=>{if(r){
            this._closePanel();
        }else{
            alert('Failed to save');
        }});  
    }

    @autobind
    private _closePanel(){
        this.props.onClose();
    }

    @autobind
    private quillChange(newVal){
        this.setState({templateContent: newVal}); 
    }

    @autobind
    private fieldClick(field){ //@Prezentacja_2_Quill_3 dodawanie field tagów
        console.log(this.quillRef, field)
        var range = this.quillRef.current.editor.getSelection();

        if (range) { 
            this.quillRef.current.editor.insertText(range.index, `{{${field.name}}}`, 'tag', `{{${field.name}}}`, 'fieldTag');
        } else {
            //One day I will implement nice alert here :)
        }
    }

    componentWillReceiveProps(nextProps){
        this.setState({
            template:nextProps.template,
            templateName: nextProps.template ? nextProps.template.Title : '',
            templateContent: nextProps.template ? nextProps.template.Template : '',
        });

    }
    
    // @Prezentacja_1_SPFX_6 Rendering configuration screen
    public render(): React.ReactElement<IConfigurationProps> {
        return (                 
            <div className={"ms-Grid "+styles.Configurations}>
                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-md8"> 
                            <TextField 
                                label="Layout name:" 
                                value={this.state.templateName}
                                onChanged={v=>this.setState({templateName:v})}
                            />  
                            <br/>
                            {false && '' /* @Prezentacja_2_Quill_1 Rendering editor */}
                            <ReactQuill 
                                value={this.state.templateContent}
                                onChange={this.quillChange}
                                modules={this.modules} 
                                formats={this.formats}
                                className={styles.quillEdit}   
                                ref={this.quillRef}                         
                            /> 
                    </div>
                    <div className="ms-Grid-col ms-md4">
                        <div className={"ms-Label "+styles.headerLabel}>Fields:</div> 
                        <div className="ms-Grid-row">
                            {this.state.fields.map(f=>{
                                return <div className={styles.inlineBtn +" ms-Grid-col ms-md6"}>
                                            <PrimaryButton 
                                                style={{width:"100%"}}
                                                onClick={(d)=>{this.fieldClick(f)}}
                                                text={f.title} 
                                                iconProps={{iconName:"Add"}} />
                                        </div>
                            })}
                        </div>
                    </div>
                </div>
                <DialogFooter>
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-md12">
                            <DefaultButton text="Printing" onClick={this._closePanel} iconProps={{iconName:'Back'}}/> 
                            {this.state.template && <PrimaryButton text="Save template" onClick={this._saveConfig} iconProps={{iconName:'Save'}} />}
                            {!this.state.template && <PrimaryButton text="Create template" onClick={this._createConfig} iconProps={{iconName:'Save'}} />}
                        </div>
                    </div>
                </DialogFooter>
            </div>                
        );
    }
}