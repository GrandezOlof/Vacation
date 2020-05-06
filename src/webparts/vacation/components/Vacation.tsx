import * as React from 'react';
import styles from './Vacation.module.scss';
import { IVacationProps } from './IVacationProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { sp } from "@pnp/sp";
import { PrimaryButton, Panel, TextField  } from 'office-ui-fabric-react';
import { DatePicker, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { DetailsList, DetailsListLayoutMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Callout, DirectionalHint } from 'office-ui-fabric-react/lib/Callout';
import { Calendar, DayOfWeek } from 'office-ui-fabric-react/lib/Calendar';
import { FocusTrapZone } from 'office-ui-fabric-react/lib/FocusTrapZone';

import { addDays } from 'office-ui-fabric-react/lib/utilities/dateMath/DateMath';

import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { initializeIcons } from '@uifabric/icons';
initializeIcons();

const DayPickerStrings: IDatePickerStrings = {
  months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],

  shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],

  days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],

  shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],

  goToToday: 'Go to today',
  prevMonthAriaLabel: 'Go to previous month',
  nextMonthAriaLabel: 'Go to next month',
  prevYearAriaLabel: 'Go to previous year',
  nextYearAriaLabel: 'Go to next year',
  closeButtonAriaLabel: 'Close date picker'
};

export interface IVacationState{
  semesterList: any;
  currentUser: any;
  adminId: number;
  isAdmin: boolean;
  showPanel: boolean;
  showPanelUpdate: boolean;
  dateToFilter: any;
  Title: String;
  slut: Date;
  start: Date;
  Status: string;
  ItemId: number;
  semesterListHolder:any;
  restrictedDates:any;
  validationMessage: any;
  showAdminList: boolean;
  hideDialog: boolean;
  itemIdStatus: number;
  status: any;
  currentUserId: number;


}

export default class Vacation extends React.Component<IVacationProps,IVacationState, {}> {

  private _columns: IColumn[];
  private options: IDropdownOption[];
  
  constructor(props:IVacationProps, state: IVacationState)
  {
    super(props);
    this.state = {
      semesterList: [],
      currentUser: {},
      showPanel: false,
      showPanelUpdate: false,
      adminId: null,
      isAdmin: false,
      dateToFilter: null,
      slut: null,
      start: null,
      Title: "",
      ItemId: null,
      Status: "",
      semesterListHolder:[],
      restrictedDates:[],
      validationMessage: "",
      showAdminList: true,
      hideDialog: true,
      itemIdStatus: null,
      status: "",
      currentUserId: null,
    };
    this._columns = [
      { key: 'column1', name: 'Namn', fieldName: 'Title', minWidth: 100, maxWidth: 120, isResizable: true },
      { key: 'column2', name: 'StartDatum', minWidth: 60, maxWidth: 80, isResizable: false, onRender: (item) => <span>{item.StartDatum.substr(0,10)}</span> },
      { key: 'column3', name: 'SlutDatum', minWidth: 60, maxWidth: 80, isResizable: false, onRender: (item) => <span>{item.SlutDatum.substr(0,10)}</span>  },
      { key: 'column4', name: 'Status', fieldName: 'Status', minWidth: 50, maxWidth: 60, isResizable: false },
      { key: 'column5', name: 'Ansvarig', minWidth: 60, maxWidth: 80, isResizable: false, onRender: (item) => <span>{item.Ansvarig.Title}</span>  },
      { key: 'column6', name: 'Ta bort', minWidth: 40, maxWidth: 40, isResizable: false, onRender: (item) => this.state.isAdmin === false && item.Title === this.state.currentUser["Title"] && item.Status === "Skapad" || item.Status === "Behandlas" ? <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" ariaLabel="Delete" onClick={() => {this.deleteSemester(item.Id)}} /> : null },
      { key: 'column6', name: 'Uppdatera', minWidth: 50, maxWidth: 60, isResizable: false, onRender: (item) => this.state.isAdmin === false && item.Title === this.state.currentUser["Title"] && item.Status === "Skapad" || item.Status === "Behandlas" ? <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" ariaLabel="Edit" onClick={() => {this._showPanelUpdate(item)}} /> : null },
      { key: 'column8', name: 'Ändra status', minWidth: 20, maxWidth: 30, isResizable: false, onRender: (item) => this.state.isAdmin === true && item.Status === "Skapad" || item.Status === "Behandlas"  ? <PrimaryButton className={styles.button} text="Status" onClick={() => {this._showDialog(item.Id)}} /> : null },
    ];
    this.options = [
      { key: '', text: 'Status', itemType: DropdownMenuItemType.Header},
      { key: 'Beviljad', text: 'Beviljad' },
      { key: 'Behandlas', text: 'Behandlas' },
      { key: 'Avslagen', text: 'Avslagen' }
    ];
  }

  componentDidMount(){
    this.getSemesterList();
    this.getCurrentUser();
  }

  private getCurrentUser = () =>{
    sp.web.currentUser.get()
    .then((result: any) =>{
      this.setState({
        currentUser: result,
        adminId: result["Id"],
        isAdmin: result["IsSiteAdmin"]
      });
    });
  }
  private getSemesterList = ():void => {
    sp.web.currentUser.get()
    .then((result: any) =>{
      console.log(result)
      this.setState({
        currentUser: result,
        adminId: result["Id"],
        currentUserId: result["Id"],
        isAdmin: result["IsSiteAdmin"],

      }, function () {
        console.log(this.state.isAdmin);
        console.log(this.state.adminId);   
      })
    })

    sp.web.lists.getByTitle("Vacation").items.select("*", "Ansvarig/Title", "Ansvarig/Id").expand("Ansvarig").getAll()
    .then((result: any[] )=>{
      console.log(result);
      let semester = result.filter(x => Date.parse(x.StartDatum) >= Date.now().valueOf())
      let currentUserSemester: any[] = result.filter(x => x.AuthorId === this.state.currentUser.Id && new Date(x.StartDatum) > new Date());
      console.log(currentUserSemester);
      console.log(semester);
      console.log(this.state.semesterListHolder);
      this.setState({
        semesterListHolder: result,
        semesterList: semester,
        restrictedDates: [].concat(...currentUserSemester.map( x=> this.getRestrictedDates(x.StartDatum, x.SlutDatum)))
      })
    })
  }

  private _showPanel = (): void => {
    this.setState({
      showPanel: true
      });
  }
  private _showPanelUpdate = (item: any): void =>{
    console.log(item);
    this.setState({
      showPanelUpdate: true,
      Title: item.Title,
      start: new Date(item.StartDatum),
      slut: new Date(item.SlutDatum),
      ItemId: item.Id
      });
  }

  private _hidePanel = (): void => {
    this.setState({
       showPanel: false,
       showPanelUpdate: false,
       validationMessage: ""
       });
  }

  private getFilterDate = (date: Date) => {
    console.log(date.valueOf())
    let test = this.state.semesterListHolder.filter(x=> Date.parse(x.StartDatum) >= date.valueOf())
    this.setState({
      semesterList: test
    })
  }

  private getPeoplePickerItems = (items: any[]) => {
    console.log(items[0].id);
    this.setState({
      adminId: items[0].id
    });
  }

  public getRestrictedDates = (startDatum: string, slutDatum: string) => {
    let currentDate = new Date(startDatum);
    var resDates: Date[] = new Array();
    while(currentDate <= new Date(slutDatum)){
        resDates.push(currentDate);
        currentDate = addDays(currentDate, +1);
    }
    return resDates;
  };

  private testValidation = (startDate:Date , endDate: Date): boolean =>{
    let dateValues = this.state.restrictedDates.map((x:Date) => x.setHours(0,0,0,0));
    console.log(dateValues);
    let valiDate = false;

    while(startDate.valueOf() <= endDate.valueOf() && valiDate === false){
      valiDate = dateValues.includes(startDate.valueOf());
      startDate = addDays(startDate, +1);
    }
    valiDate ? this.setState({validationMessage: "You have already applied for vacation on these dates"}): null;
    console.log(valiDate);
    return valiDate;
  }


  private addSemester = (e:any): void => {
    e.preventDefault();
    console.log(this.state)

    this.testValidation(this.state.start, this.state.slut) === false ?
    sp.web.lists.getByTitle("Vacation").items.add({
      Title: this.state.currentUser["Title"],
      StartDatum: new Date(this.state.start.setHours(8)),
      SlutDatum: new Date(this.state.slut.setHours(17)),
      AnsvarigId: this.state.adminId,
    }).then(() => { 
        this.getSemesterList();
        this._hidePanel();
    })
    :null;
  }

  private updateSemester = (id:number): void =>{
    console.log(id);
    this.testValidation(this.state.start, this.state.slut) === false ?
    sp.web.lists.getByTitle("Vacation").items.getById(id).update({
      Title: this.state.Title,
      StartDatum: new Date(this.state.start.setHours(8)),
      SlutDatum: new Date(this.state.slut.setHours(17)),
    }).then(()=>{
      this.getSemesterList();
      this._hidePanel();
    })
    :null;
  }

  public deleteSemester = (id): void => {
    console.log("hej" + id);
    sp.web.lists.getByTitle("Vacation").items.getById(id).delete()
      .then(() => {
        this.getSemesterList();
      });
  }

  private updateStatus = (id: number): void => {
    sp.web.lists.getByTitle("Vacation").items.getById(id).update({
      Status: this.state.status
    }).then( () => {
      this.getSemesterList();
      this._closeDialog();
    } )
  }

  public render(): React.ReactElement<IVacationProps> {

    let filteredList: any[] = this.state.isAdmin === true ? this.state.semesterList.filter(x => x.AnsvarigId === this.state.currentUserId ) : 
    this.state.semesterList.filter(x => this.state.currentUser["Title"] == x.Title)

    let hideButton = this.state.isAdmin === false ? <PrimaryButton className={styles.button} text="Ny Ansökan" onClick={this._showPanel} />
    : null;

    let text = this.state.isAdmin === false ? 
    <div className={styles.AnställdDiv}>
        <h1>Anställd: {this.state.currentUser["Title"]}</h1>
        <h2>Ansök om Semester</h2>
        </div> : 
        <div className={styles.AdminDiv}>
          <h1>Admin: {this.state.currentUser["Title"]}</h1>
          <h2>Bevilja Semester</h2>
          </div> ;

    return (
      <div className={ styles.vacation }>
        {text}
        <DatePicker
          strings={DayPickerStrings}
          showWeekNumbers={true}
          firstWeekOfYear={1}
          showMonthPickerAsOverlay={true}
          onSelectDate = {this.getFilterDate}
          placeholder="Select a date..."
          ariaLabel="Select a date"
        />
        <br/>
        <DetailsList 
          items={filteredList}
          columns={this._columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}           
          selectionPreservedOnEmptyClick={true}
          ariaLabelForSelectionColumn="Toggle selection"
          ariaLabelForSelectAllCheckbox="Toggle selection for all items"
          checkButtonAriaLabel="Row checkbox"
          />
        <br/>
        {hideButton}
        <Panel
          isOpen={this.state.showPanel}
          closeButtonAriaLabel="Close"
          isLightDismiss={true}
          headerText="Ny ansökan"
          onDismiss={this._hidePanel}
        >
          <TextField label="Namn" readOnly defaultValue={this.state.currentUser["Title"]} />
          <br/>
          <label>Startdatum</label>
          <Calendar
                onSelectDate={this._onSelectDate1}                   
                value={this.state.start}
                strings={DayPickerStrings}
                isMonthPickerVisible={false}
                restrictedDates={this.state.restrictedDates}
                minDate= {new Date()}
              />
          <label>Slutdatum</label>
          <Calendar
                onSelectDate={this._onSelectDate2}                   
                value={this.state.slut}
                strings={DayPickerStrings}
                isMonthPickerVisible={false}
                restrictedDates={this.state.restrictedDates}
                minDate={this.state.start}
              />

          <PeoplePicker
                context={this.props.context}
                titleText="Ansvarig"
                personSelectionLimit={1}
                groupName={"MyDeveleporSite Owners"}
                showtooltip={true}
                isRequired={true}
                disabled={false}
                selectedItems={this.getPeoplePickerItems}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                ensureUser={true}
                resolveDelay={1000}
              />
              <br/>
              <p className={styles.validationMs}>{this.state.validationMessage}</p>
          <PrimaryButton className={styles.button} text="Spara" onClick={this.addSemester} />
        </Panel>
                {/* ------------------Update Panel */}
                <Panel
                isOpen={this.state.showPanelUpdate}
                closeButtonAriaLabel="Close"
                isLightDismiss={true}
                headerText="Ny ansökan"
                onDismiss={this._hidePanel}
              >
                <TextField label="Namn" readOnly defaultValue={this.state.currentUser["Title"]} />
                <br />
                <label>Startdatum</label>
                <Calendar
                  onSelectDate={this._onSelectDate1}
                  value={this.state.start}
                  strings={DayPickerStrings}
                  isMonthPickerVisible={false}
                  minDate= {new Date()}
                  restrictedDates={this.state.restrictedDates}
                  
                />
                <label>Slutdatum</label>
                <Calendar
                  onSelectDate={this._onSelectDate2}
                  value={this.state.slut}
                  strings={DayPickerStrings}
                  isMonthPickerVisible={false}
                  minDate={this.state.start}
                  restrictedDates={this.state.restrictedDates}
                />

                <PeoplePicker
                  context={this.props.context}
                  titleText="Ansvarig"
                  personSelectionLimit={1}
                  groupName={"MyDeveleporSite Owners"}
                  showtooltip={true}
                  isRequired={true}
                  disabled={false}
                  selectedItems={this.getPeoplePickerItems}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  ensureUser={true}
                  resolveDelay={1000}
                />
                <br />
              <p className={styles.validationMs}>{this.state.validationMessage}</p>
          <PrimaryButton className={styles.button} text="Updatera" onClick={() => {this.updateSemester(this.state.ItemId)}} />
        </Panel>
                {/* ---------------Slut Update Panel */}
                <Dialog
                  hidden={this.state.hideDialog}
                  onDismiss={this._closeDialog}
                  dialogContentProps={{
                    type: DialogType.normal,
                    title: 'Ändra status',
                  }}
                >
                  <Dropdown
                    placeholder="Välj status"
                    label=""
                    options={this.options}
                    onChange={this._dropdownChange}
                  />
                <DialogFooter>
            <PrimaryButton onClick={ () => {this.updateStatus(this.state.itemIdStatus) }} text="Spara" />
            <PrimaryButton onClick={this._closeDialog} text="Avbryt" />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }
  private _onSelectDate1 = (date: Date | null | undefined): void => {
    this.setState({ start: date });
  }
  private _onSelectDate2 = (date: Date | null | undefined): void => {
    this.setState({ slut: date });
  } 

  private _showDialog = (id?: number):void =>{
    this.setState({
      hideDialog: false,
      itemIdStatus: id
    })
  }

  private _closeDialog = () => {
    this.setState({
      hideDialog: true
    })
  }
  private _dropdownChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number): void => {
    this.setState({
      status: option.text
    });
  }

}
