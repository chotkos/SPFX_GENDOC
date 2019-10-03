import * as React from 'react';
import { TextField, DefaultButton, PrimaryButton, DialogFooter, autobind, Panel, Spinner, SpinnerType, Dropdown, PanelType } from "office-ui-fabric-react";
//import { sp } from "@pnp/sp";
import s from './CustomPanel.module.scss'
import Configuration from '../Configuration/Configurations';
import TemplateService from '../../services/TemplateService';
import ReactHtmlParser from 'react-html-parser';
import ReactQuill from 'react-quill'; // ES6
import 'react-quill/dist/quill.snow.css'; // ES6 
// @Prezentacja_4_ReactToPrint_2 ładowanie klas css z edytora quill
import ItemsService from '../../services/ItemsService';
import ReactToPrint from "react-to-print";

export interface ICustomPanelState {
    saving: boolean;
    showConfiguration:boolean,
    allTemplates: any[];
    optionsTemplates:[];
    selectedKey: string;
    selectedTemplate: any;
    currentItem: any;
    filledTemplate: string;
}

export interface ICustomPanelProps {
    onClose: () => void;
    isOpen: boolean;
    currentTitle: string;
    itemId: number;
    listId: string;
}

export default class CustomPanel extends React.Component<ICustomPanelProps, ICustomPanelState> {
    
    private templateService = new TemplateService();
    private itemsService = new ItemsService();

    previewRef=null;

    constructor(props: ICustomPanelProps) {
        super(props);

        this.state = {
            saving: false,
            showConfiguration:false,
            allTemplates:[],
            optionsTemplates:[],
            selectedKey:null,
            selectedTemplate:null,
            currentItem:null,
            filledTemplate:'',
        };

        this.initTemplates();
        this.getListItem();
    }
 
    @autobind
    initTemplates(){
        this.templateService.GetAllTemplates().then(templates=>{

            let options = templates.map((x)=>{return {key: x.ID, text:x.Title};})
            options.push({key:'',text:'New template'});

            this.setState({
                allTemplates: templates,
                optionsTemplates: options,
            })
        });
    }

    @autobind
    getListItem(){
        this.itemsService.GetListItemById(this.props.itemId)
            .then((item):any=>{
                this.setState({currentItem: item});
            });
    }
   

    @autobind
    private _onConfiguration() { 
        this.setState({showConfiguration:true})
    }

    @autobind
    private _hideConfigPanel(){
        this.setState({
            showConfiguration:false,
            selectedKey:null,
            selectedTemplate:null,
        }), this.forceUpdate;
        this.initTemplates();        
    }
 
    @autobind
    private changedTemplate(newValue){
        let selectedTemplate = this.state.allTemplates.filter(x=>{return x.ID == newValue.key})[0];


        this.setState({
            selectedKey: newValue.key, 
            selectedTemplate: selectedTemplate, 
            showConfiguration: false
        }, this.forceUpdate);

        
        this.itemsService.GetAllFields().then((allFields:any[])=>{
                        
            let filledTemplate = selectedTemplate.Template;
            allFields.forEach(field => {

                let fieldValue = this.state.currentItem[field.InternalName];
                let left = '&#123;';
                let right = '&#125;';
                let marker = left+left+field.InternalName+right+right;
                while(filledTemplate.indexOf(marker)!=-1){
                    filledTemplate = filledTemplate.replace(marker, fieldValue);
                }
            });
            this.setState({filledTemplate:filledTemplate})
            
        });
    }

    // @Prezentacja_1_SPFX_5 rendering panel content
    public render(): React.ReactElement<ICustomPanelProps> {
        let { isOpen, currentTitle } = this.props;
        return (
            <Panel 
                type={PanelType.medium}
                className={s.CustomPanel}
                onAbort={this._hideConfigPanel}
                isOpen={isOpen}
                headerText="Print document">                
                <div className="ms-Grid">
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-md12">
                            <Dropdown
                                options={this.state.optionsTemplates}
                                label={"Choose template:"}
                                onChanged={this.changedTemplate}  
                                defaultSelectedKey={''} />
                            <br/>
                        </div>
                    </div>
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-md6">
                            <DefaultButton text="Template Configuration" onClick={this._onConfiguration} iconProps={{iconName:'ConfigurationSolid'}}/>
                        </div>
                        <div className="ms-Grid-col ms-md6" style={{textAlign:"right"} /* @Prezentacja_4_ReactToPrint_1 podgląd elementu*/}>
                            <ReactToPrint
                                trigger={() => <PrimaryButton text="Print" iconProps={{iconName:'Print'}} />}
                                content={() => this.previewRef}
                            />
                        </div>
                    </div> 
                    <br/>
                    {this.state.selectedTemplate!=null &&
                    <div className="ms-Grid-row">
                        <p>Preview:</p>
                        <div className="ql-container ql-snow" ref={(el)=>{this.previewRef=el;} /* @Prezentacja_4_ReactToPrint_2 Wykorzystanie refa aby dostać kontener*/}>
                            <div className="ql-editor">
                                {this.state.selectedTemplate!=null && 
                                    ReactHtmlParser(this.state.filledTemplate)}
                            </div>
                        </div>
                    </div>}
                </div>
                <Panel 
                    type={PanelType.large}
                    isOpen={this.state.showConfiguration}
                    onAbort={this._hideConfigPanel}
                    headerText="Template configuration">
                    <Configuration
                        onClose={this._hideConfigPanel}
                        template={this.state.selectedTemplate}
                    />
                </Panel>
            </Panel>
        );
    }
}