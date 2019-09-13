import * as React from 'react';
import { TextField, DefaultButton, PrimaryButton, DialogFooter, autobind, Panel, Spinner, SpinnerType, Dropdown, PanelType } from "office-ui-fabric-react";
import IConfigurationState from './IConfigurationState';
import IConfigurationProps from './IConfigurationProps';
import ReactQuill from 'react-quill'; // ES6
import 'react-quill/dist/quill.snow.css'; // ES6
import styles from './Configurations.module.scss'; 
import { Dialog } from '@microsoft/sp-dialog';

export default class Configuration extends React.Component<IConfigurationProps, IConfigurationState> {

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
            templateContent:"",
            fields:[
                {title:'Title', name: 'Title'},                
                {title:'Created By', name: 'CreatedBy'},
                {title:'Price', name: 'Price'},
                {title:'Title', name: 'Title'},
                {title:'Created By', name: 'CreatedBy'},
                {title:'Price', name: 'Price'},
                {title:'Title', name: 'Title'},
                {title:'Created By', name: 'CreatedBy'},
                {title:'Price', name: 'Price'},
                {title:'Title', name: 'Title'},
                {title:'Created By', name: 'CreatedBy'},
                {title:'Price', name: 'Price'},
            ]

        };    
    
        this.quillRef = React.createRef();
    }

    @autobind
    private _saveConfig(){
        //@TODO: saving pnp
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
    private fieldClick(field){ //@Prezentacja: dodawanie field tagów
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
            templateContent:"",
        });

    }
    

    public render(): React.ReactElement<IConfigurationProps> {
        return (                 
            <div className={"ms-Grid "+styles.Configurations}>
                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-md8">
                        <TextField 
                            label="Layout name:" 
                            value={this.state.templateName}
                        />
                        <br/>
                    </div>
                </div>
                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-md8">
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
                        <h2>Fields:</h2> 
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
                            <PrimaryButton text="Save template" onClick={this._saveConfig} iconProps={{iconName:'Save'}} />
                        </div>
                    </div>
                </DialogFooter>
            </div>                
        );
    }
}