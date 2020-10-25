import * as React from 'react';
//import styles;
import { IWebApiClientProps } from './IWebApiClientProps';
import { escape } from '@microsoft/sp-lodash-subset';

import {
	CommandBar,
	DetailsList,
	ISelection,
	Selection,
	SelectionMode,
	Panel,
	TextField,
	PrimaryButton,
	DefaultButton
} from 'office-ui-fabric-react';

import { ITimeSheet } from '../../../entities/ITimeSheet';
import { TimeSheetsServiceKey, ITimeSheetsService } from '../../../services/TimeSheetsService';
import { ApiConfigServiceKey, IApiConfigService } from '../../../services/ApiConfigService';

import { SPComponentLoader } from '@microsoft/sp-loader';
import { Container,Row,Modal,Button,Toast }from "react-bootstrap";

import CustomDialog from '../../../extensions/CustomDialog';  

export interface IWebApiClientState {
	timeSheets?: ITimeSheet[];
	selectedDocument?: ITimeSheet;
	selection?: ISelection;
	isAdding?: boolean;
	isEditing?: boolean;
	selectedView?: 'All' | 'My';
}

export default class WebApiClient extends React.Component<IWebApiClientProps, IWebApiClientState> {
	private timeSheetsService: ITimeSheetsService;
	private apiConfig: IApiConfigService;
	private authenticated: boolean;

	
	constructor(props: IWebApiClientProps) {
		super(props);
		this.state = {
			timeSheets: [],			
			selectedDocument: null,
			isAdding: false,
			isEditing: false,
			selectedView: 'All',
			selection: new Selection({
				onSelectionChanged: this._onSelectionChanged.bind(this)
			})
		};
	}

	public componentWillMount() {
		this.props.serviceScope.whenFinished(() => {
			this.timeSheetsService = this.props.serviceScope.consume(TimeSheetsServiceKey);
			this.apiConfig = this.props.serviceScope.consume(ApiConfigServiceKey);
			this._loadTimesheets(); //Load all timesheets...
		});
	}

	private _loadTimesheets(stateRefresh?: IWebApiClientState, forceView?: 'All' | 'My') {
		let { selectedView } = this.state;

        let effectiveView = forceView || selectedView;
        
		// Mickion - After being authenticated
		this._executeOrDelayUntilAuthenticated(() => {
			switch (effectiveView) {
				case 'All':
					//Load only time entries created today..
					this.timeSheetsService.getMyTimeSheets().then((docs) => {
						let state = stateRefresh || {};	
						state.timeSheets = docs;
						this.setState(state);
					});

					//The below commented out code is for getting all List entries
					/*this.timeSheetsService.getAllTimeSheets().then((docs) => {
						let state = stateRefresh || {};	
						state.timeSheets = docs;
						this.setState(state);
					});*/
					break;

				/*case 'My': - the other
					// Load My business documents when component is being mounted
					this.businessDocsService.getMyBusinessDocuments().then((docs) => {
						let state = stateRefresh || {};
						state.timeSheets = docs;
						this.setState(state);
					});
					break;*/
			}
		});
	}

	private _executeOrDelayUntilAuthenticated(action: Function): void {
		if (this.authenticated) {
			console.log('Is authenticated');
			action();
		} else {
			console.log('Still not authenticated');
			setTimeout(() => {
				this._executeOrDelayUntilAuthenticated(action);
			}, 1000);
		}
	}

	private _onSelectionChanged() {
		let { selection } = this.state;
		let selectedDocuments = selection.getSelection() as ITimeSheet[];

		console.log("LENGHT: "+ selectedDocuments.length);
		console.log("TRUE MAN: "+ selectedDocuments[0]);
		let selectedDocument = selectedDocuments && selectedDocuments.length == 1 && selectedDocuments[0];

		console.log('_onSelectionChanged SELECTED DOCUMENT: ', selectedDocument);
		this.setState({
			selectedDocument: selectedDocument || null
		});
	}

	private _buildCommands() {
		let { selectedDocument } = this.state;

		const add = {
			key: 'add',
			name: 'Create',
			icon: 'Add',
			onClick: () => this.addNewTimeSheet()
		};

		const edit = {
			key: 'edit',
			name: 'Edit',
			icon: 'Edit',
			onClick: () => this.editCurrentTimeSheet()
		};

		const remove = {
			key: 'remove',
			name: 'Remove',
			icon: 'Remove',
			onClick: () => this.removeCurrentTimeSheet()
		};

		let commands = [ add ];

		if (selectedDocument) {
			commands.push(edit, remove);
		}

		return commands;
	}

	private _buildFarCommands() {
		let { selectedDocument, selectedView } = this.state;

		const views = {
			key: 'views',
			name: selectedView == 'All' ? 'All' : "I'm in charge of",
			icon: 'View',
		};

		let commands = [ views ];

		return commands;
	} 

	public selectView(view: 'All' | 'My') {
		this.setState({
			selectedView: view
		});

		this._loadTimesheets(null, view);
	}

	public addNewTimeSheet() {
		this.setState({
			isAdding: true,			
			selectedDocument: {
				Id: 1,
				Title: 'Override',
				Description: 'Override',
				Category: 'Override',
				Hours: 1,
				Date: new Date(),
				Created: new Date(),
				DayOfWeek: 1,
			} 
		});
	} 

	public editCurrentTimeSheet() {
		let { selectedDocument } = this.state;
		if (!selectedDocument) {
			return;
		}

		this.setState({
			isEditing: true
		});
	}
	
	public removeCurrentTimeSheet() {
		let { selectedDocument } = this.state;
		if (!selectedDocument) {
			return;
		}

		if (confirm('Are you sure you want to remove entry?')) {
			this._executeOrDelayUntilAuthenticated(() => {
				this.timeSheetsService
					.removeTimeSheet(selectedDocument.Id)
					.then(() => {
						alert('Timesheet entry has been removed successfully!');
						this._loadTimesheets();
					})
					.catch((error) => {
						console.log(error);
						alert('Failed to REMOVE entry :Error '+ error);
					});
			});
		}
	} 

	//Set value to of any type
	private onValueChange(property: string, event) {
		const {name,value} = event.target;
		//console.log("NAZO: "+ name +" value "+ value);

		let { selectedDocument } = this.state;
		if (!selectedDocument) {
			//console.log("onValueChange exit!!!")
			return;
		}
		//console.log(property +" **SAVE THIS VALUE** "+ value);
		selectedDocument[property] = value;
	}

	private onApply() {
		let { selectedDocument, isAdding, isEditing } = this.state;
		//console.log("OnApply function..."+ selectedDocument.Title);

		if (isAdding) {
			this._executeOrDelayUntilAuthenticated(() => {
				this.timeSheetsService
					.createTimeSheet(selectedDocument)
					.then(() => {						
						alert('Timesheet entry has been captured successfully!');
						this._loadTimesheets({
							selectedDocument: null,
							isAdding: false,
							isEditing: false
						});
					})
					.catch((error) => {
						//console.log(error);
						//alert('Document CANNOT be created !');
						alert('Failed to CREATE entry :Error '+ error);
					});
			});
		} 
		else if (isEditing) {
			this._executeOrDelayUntilAuthenticated(() => {
				this.timeSheetsService
					.updateTimeSheet(selectedDocument.Id, selectedDocument)
					.then(() => {
						alert('Timesheet entry has been updated successfully!');
						this._loadTimesheets({
							selectedDocument: null,
							isAdding: false,
							isEditing: false
						});
					})
					.catch((error) => {
						//console.log(error);
						alert('Failed to UPTADE entry :Error '+ error);
					});
			});
		} 
	}

	private onCancel() {
		//console.log("onCancel function...");
		this.setState({
			selectedDocument: null,
			isAdding: false,
			isEditing: false
		}); 
	}

	private convertToString(val) {
		return String(val);
	}

	componentDidMount(){
		//Show Welcome message
		const dialog: CustomDialog = new CustomDialog();  
		dialog.show();
	}

	public render(): React.ReactElement<IWebApiClientProps> {
		let { timeSheets, selection, selectedDocument, isAdding, isEditing } = this.state;				

		let todaysDate = new Date(); 
		const daysOftheWeek = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];

		let totalHours = 0;
		let overtimeHrs = 0;
		for(var i=0; i<timeSheets.length; i++){
			totalHours = totalHours + timeSheets[i].Hours;
			
			if(totalHours > 8)
			{
				overtimeHrs = totalHours -8;				
			}
		}
		
		        //Style with bootstrap
        SPComponentLoader.loadCss("https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css");
		return (
			<Container>
				<Row>
					<div><h4>{todaysDate.getDate()}</h4></div>
					<div><h4>{daysOftheWeek[todaysDate.getDay()]}</h4></div>
				</Row>
				<Row>Total Hours Captured: {totalHours}</Row>
				<Row>Overtime Hours Pending Approval: {overtimeHrs}</Row>
				
				<Row>
					<iframe
						src={this.apiConfig.appRedirectUri}
						style={{ display: 'none' }}
						onLoad={() => (this.authenticated = true)}
					/>
				
					<DetailsList
						items={timeSheets}
						columns={[
							/*{
								key: 'id',
								name: 'Id',
								fieldName: 'Id',
								minWidth: 15,
								maxWidth: 30
							},*/
							{
								key: 'Title',
								name: 'Title',
								fieldName: 'Title',
								minWidth: 100,
								maxWidth: 200
							},
							{
								key: 'Description',
								name: 'Description',
								fieldName: 'Description',
								minWidth: 100,
								maxWidth: 200
							},
							{
								key: 'Category',
								name: "Category",
								fieldName: 'Category',
								minWidth: 100,
								maxWidth: 200
							},
							{
								key: 'Hours',
								name: "Hours",
								fieldName: 'Hours',
								minWidth: 100,
								maxWidth: 200
							},
							{
								key: 'Date',
								name: "Date",
								fieldName: 'Date',
								minWidth: 100,
								maxWidth: 200
							}
						]}
						selectionMode={SelectionMode.single}
						selection={selection}
					/>
					{selectedDocument &&
					(isAdding) && (
						<Panel isOpen={true}>
							<TextField
								label="Title"
								onChange={(v) => this.onValueChange('Title', v)}
							/>
							<TextField
								label="Description"
								onChange={(v) => this.onValueChange('Description', v)}
							/>
							<TextField
								label="Category"
								onChange={(v) => this.onValueChange('Category', v)}
							/>							
							<TextField
								label="Hours"
								onChange={(v) => this.onValueChange('Hours', v)}
							/>
						
							<PrimaryButton text="Apply" onClick={() => this.onApply()} />
							<DefaultButton text="Cancel" onClick={() => this.onCancel()} />
						</Panel>
					)} 

					{selectedDocument &&
					(isEditing) && (
						<Panel isOpen={true}>
							<TextField
								label="Title"
								defaultValue={selectedDocument.Title}
								onChange={(v) => this.onValueChange('Title', v)}								
							/>
							<TextField
								label="Description"
								defaultValue={selectedDocument.Description}
								onChange={(v) => this.onValueChange('Description', v)}
							/>
							<TextField
								label="Category"
								defaultValue={selectedDocument.Category}
								onChange={(v) => this.onValueChange('Category', v)}
							/>							
							<TextField
								label="Hours"
								defaultValue={this.convertToString(selectedDocument.Hours)}
								onChange={(v) => this.onValueChange('Hours', v)}
							/>
						
							<PrimaryButton text="Apply" onClick={() => this.onApply()} />
							<DefaultButton text="Cancel" onClick={() => this.onCancel()} />
						</Panel>
					)} 
				</Row>
								
				<CommandBar items={this._buildCommands()}/>
				
			</Container>
		);
	}
	
}