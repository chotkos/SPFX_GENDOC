import * as React from 'react';
import { TextField, DefaultButton, PrimaryButton, DialogFooter, autobind, Panel, Spinner, SpinnerType, Dropdown, PanelType } from "office-ui-fabric-react";
//import { sp } from "@pnp/sp";
import s from './CustomPanel.module.scss'
import Configuration from '../Configuration/Configurations';

export interface ICustomPanelState {
    saving: boolean;
    showConfiguration:boolean,
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

    constructor(props: ICustomPanelProps) {
        super(props);
        this.state = {
            saving: false,
            showConfiguration:false
        };
    }

    @autobind
    private _onTitleChanged(title: string) {
        this.editedTitle = title;
    }

    @autobind
    private _onConfiguration() {
        //this.props.onClose();
        this.setState({showConfiguration:true})
    }

    @autobind
    private _hideConfigPanel(){
        this.setState({showConfiguration:false});
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
                                options={[]}
                                label={"Choose template:"}
                            />
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
                </div>
                <Panel 
                    type={PanelType.large}
                    isOpen={this.state.showConfiguration}
                    onAbort={this._hideConfigPanel}
                    headerText="Template configuration">
                    <Configuration
                        onClose={this._hideConfigPanel}
                        template={null}
                    />
                </Panel>
            </Panel>
        );
    }
}