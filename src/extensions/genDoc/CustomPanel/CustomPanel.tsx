import * as React from 'react';
import { TextField, DefaultButton, PrimaryButton, DialogFooter, autobind, Panel, Spinner, SpinnerType, Dropdown, PanelType } from "office-ui-fabric-react";
//import { sp } from "@pnp/sp";
import s from './CustomPanel.module.scss'
import Configuration from '../Configuration/Configurations';
import TemplateService from '../../services/TemplateService';
import ReactHtmlParser from 'react-html-parser';
import ReactQuill from 'react-quill'; // ES6
import 'react-quill/dist/quill.snow.css'; // ES6

export interface ICustomPanelState {
    saving: boolean;
    showConfiguration:boolean,
    allTemplates: any[];
    optionsTemplates:[];
    selectedKey: string;
    selectedTemplate: any;
}

export interface ICustomPanelProps {
    onClose: () => void;
    isOpen: boolean;
    currentTitle: string;
    itemId: number;
    listId: string;
}

export default class CustomPanel extends React.Component<ICustomPanelProps, ICustomPanelState> {

    private editedTitle: string = null;
    private templateService = new TemplateService();

    constructor(props: ICustomPanelProps) {
        super(props);

        this.state = {
            saving: false,
            showConfiguration:false,
            allTemplates:[],
            optionsTemplates:[],
            selectedKey:null,
            selectedTemplate:null,
        };

        this.initTemplates();
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
    private _onConfiguration() {
        //this.props.onClose();
        this.setState({showConfiguration:true})
    }

    @autobind
    private _hideConfigPanel(){
        this.setState({showConfiguration:false});
        this.initTemplates();
    }

    @autobind
    private _onPrint() {
        /*this.setState({ saving: true });
        sp.web.lists.getById(this.props.listId).items.getById(this.props.itemId).update({
            'Title': this.editedTitle
        }).then(() => {
            this.setState({ saving: false });
            this.props.onClose();
        });*/
    }

    @autobind
    private changedTemplate(newValue){
        let selectedTemplate = this.state.allTemplates.filter(x=>{return x.ID == newValue.key})[0];

        this.setState({
            selectedKey: newValue.key, 
            selectedTemplate: selectedTemplate, 
            showConfiguration: false
        }, this.forceUpdate);
    }

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
                        <div className="ms-Grid-col ms-md6" style={{textAlign:"right"}}>
                            <PrimaryButton text="Print" onClick={this._onPrint} iconProps={{iconName:'Print'}} />
                        </div>
                    </div> 
                    <br/>
                    <div className="ms-Grid-row">
                        <div className="ql-container ql-snow">
                            <div className="ql-editor">
                                {this.state.selectedTemplate!=null && ReactHtmlParser(this.state.selectedTemplate.Template)}
                            </div>
                        </div>
                    </div>
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