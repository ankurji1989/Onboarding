import * as React from 'react';
import styles from './OffboardingCheckList.module.scss';
import { IOffboardingCheckListProps } from './IOffboardingCheckListProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import { SPDataOperations } from '../../../common/SPDataOperations';
import { Checkbox, DatePicker, DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, IDropdown, IDropdownOption, Label, MessageBar, MessageBarType, PrimaryButton, Spinner, Toggle } from 'office-ui-fabric-react';
import * as strings from 'OffboardingCheckListWebPartStrings';

export interface IOffboardingCheckListState {
  employeeData: any[];
  checkList: any[];
  hideDialog: boolean;
  hideConfirmDialog: boolean;
  formSubmit: boolean;
  checkedListTotalCount: number;
  submitRecall: boolean;
  offboardingStatus: string;
  lastWorkingDate?: Date;
  leaveType: string;
  leaveTypeOption: IDropdownOption[];
  offboardingStartDate: any;
  emplayeeName: string;
}

export default class OffboardingCheckList extends React.Component<IOffboardingCheckListProps, IOffboardingCheckListState> {
  private offboardingId: number = null;
  constructor(props) {
    super(props);
    this.getChoicesFromChoiceColumn = this.getChoicesFromChoiceColumn.bind(this);
    this.getCheckListDetails = this.getCheckListDetails.bind(this);
    this.toggleOnChange = this.toggleOnChange.bind(this);
    this.checkBoxOnChange = this.checkBoxOnChange.bind(this);
    this.submitCheckList = this.submitCheckList.bind(this);
    this.submitRecall = this.submitRecall.bind(this);
    this.state = {
      employeeData: [],
      checkList: [],
      hideDialog: true,
      hideConfirmDialog: true,
      formSubmit: false,
      checkedListTotalCount: 0,
      submitRecall: false,
      offboardingStatus: null,
      lastWorkingDate: null,
      leaveType: null,
      leaveTypeOption: [],
      offboardingStartDate: null,
      emplayeeName: null
    }
  }

  public componentDidMount() {
    const queryParms = new UrlQueryParameterCollection(window.location.href);
    const offboardingId: string = queryParms.getValue('offboardID');
    if(offboardingId !== undefined){
      this.offboardingId = Number(offboardingId);
      this.getChoicesFromChoiceColumn();
      this.getEmployeeCheckList(Number(offboardingId));
    }
  }

  public async getChoicesFromChoiceColumn() {
    let leaveTypeOption: IDropdownOption[] = [];
    SPDataOperations.getChoicesFromChoiceColumn(this.props.onboardingList, strings.LeaveTypeColumn).then((choiceData) => {
      if(choiceData.Choices !== undefined && choiceData.Choices.length > 0) {
        choiceData.Choices.map((choiceItem) => {
          leaveTypeOption.push({key: choiceItem, text: choiceItem});
        });
        this.setState({leaveTypeOption: leaveTypeOption});
      }
    });
  }

  public async getEmployeeCheckList(empId: number) {

    SPDataOperations.getListItems(this.props.onboardingList, 'Employee,RegistrationListItemID,OffBoardingCheckList/Id,OffBoardingCheckListNA/Id,EmployeeID1/Title,CheckListCompleted,OffBoardingStatus,LastWorkingDay,LeaveType,OffBoardingStartDate','OffBoardingCheckList,OffBoardingCheckListNA,EmployeeID1',`Id eq ${empId}`).then((offboardingData) => {
      if(offboardingData.length > 0 ) {
        this.getCheckListDetails(offboardingData[0]);
      }
    });
  }

  public getCheckListDetails(employeeData: any) {
    SPDataOperations.getListItems(this.props.checkList, 'Id,Title,Order0,Required,NASlider,ColType', ``, `CheckListStatus eq 1`).then((checkList) => {
      let updatedChecklist = [];
      checkList.sort((item1, item2) => item1.Order0 - item2.Order0).map((listItem) => {
        const checkListObject: any = {
          Id: listItem.Id,
          Title: listItem.Title,
          selected: employeeData.OffBoardingCheckList.filter(item => listItem.Id === item.Id).length > 0 ? true : false,
          notApplicable: employeeData.OffBoardingCheckListNA.filter(item => listItem.Id === item.Id).length > 0 ? true : false,
          required: listItem.Required === true ? true : false,
          colType: (listItem.ColType === null || listItem.ColType === ``) ? null : (listItem.ColType).toLowerCase()
        };
        updatedChecklist.push(checkListObject);
      });

      const checkListCount: number = updatedChecklist.filter(item => item.colType !== 'heading').length;
      let CheckedListCount: number = 0;
      updatedChecklist.map((item) => {
        if (item.selected === true) {
          CheckedListCount++;
        }
        if (item.notApplicable === true) {
          CheckedListCount++
        }
      });

      const lastWorkingDay = employeeData.LastWorkingDay !== null ? new Date(employeeData.LastWorkingDay) : null;
      const leaveType = employeeData.LeaveType !== null ? employeeData.LeaveType : null;
      const emplayeeName = employeeData.EmployeeID1 ? employeeData.EmployeeID1.Title : null;
      this.setState({checkList: updatedChecklist, checkedListTotalCount: CheckedListCount, offboardingStatus: employeeData.OffBoardingStatus, lastWorkingDate: lastWorkingDay, hideDialog: true, hideConfirmDialog: true, leaveType: leaveType, offboardingStartDate: employeeData.OffBoardingStartDate, emplayeeName: emplayeeName});
    });
  }

  public toggleOnChange(item: any) {
    let checkList = this.state.checkList.slice();
    checkList.filter(listItem => listItem.Id === item.Id).map(selectedItem => {
      selectedItem.notApplicable = !item.notApplicable;
      selectedItem.selected = false;
    });
    this.setState({checkList: checkList});
  }

  public checkBoxOnChange(item: any) {
    const {lastWorkingDate, checkList, leaveType} = this.state;
    checkList.filter(listItem => listItem.Id === item.Id).map(selectedItem => {
      selectedItem.selected = !item.selected;
    });
    this.setState({checkList: checkList, lastWorkingDate: item.Title === strings.LastWorkingDateTitle && item.selected === false ? null : lastWorkingDate, leaveType: item.Title === strings.LeaveTypeTitle && item.selected === false ? null : leaveType });
  }

  public async submitCheckList(confirmDialog: boolean) {
    const {checkList, lastWorkingDate, leaveType, offboardingStartDate} = this.state;
    let selectedCheckList: any[] = [];
    let checkListNA: any[] = [];
    let checkListCount: number = checkList.filter(item => item.colType !=='heading').length;
    let selectedCheckListCount: number = 0;
    let lastWorkingDaySelected: boolean = true;
    let leaveTypeSelected: boolean = true;

    checkList.map((item) => {
      if(item.Title === strings.LastWorkingDateTitle && item.selected && lastWorkingDate === null) {
        lastWorkingDaySelected = false;
      }
      if(item.Title === strings.LeaveTypeTitle && item.selected && leaveType === null) {
        leaveTypeSelected = false;
      }
      if (item.selected === true) {
        selectedCheckList.push(item.Id);
        selectedCheckListCount++;
      }
      if (item.notApplicable === true) {
        checkListNA.push(item.Id);
        selectedCheckListCount++;
      }
    });

    if(!lastWorkingDaySelected) {
      alert('Please select the last working date');
    } else if(!leaveTypeSelected) {
      alert('Please select the leaver type');
    } else {
      let updateObject: any = {OffBoardingCheckListId:{'__metadata': { type: 'Collection(Edm.Int32)' },'results': selectedCheckList}, OffBoardingCheckListNAId: {'__metadata': { type: 'Collection(Edm.Int32)' },'results': checkListNA}};

      if(checkListCount === selectedCheckListCount && confirmDialog === false) {
        this.setState({hideConfirmDialog: false});
      } else {
        this.setState({hideDialog: false});
        if(checkListCount === selectedCheckListCount){
          updateObject.OffBoardingStatus = 'Completed';
          updateObject.OffBoardingCompletionDate = new Date();
        } else {
          updateObject.OffBoardingStatus = 'In Progress';
        }
        if(offboardingStartDate === null) {
          updateObject.OffBoardingStartDate = new Date();
        }
        if(lastWorkingDate !== null) {
          updateObject.LastWorkingDay = lastWorkingDate;
        }
        if(leaveType !== null) {
          updateObject.LeaveType = leaveType;
        }
        await SPDataOperations.updateListItem(this.props.onboardingList, this.offboardingId, updateObject);
        this.getEmployeeCheckList(this.offboardingId);
        this.setState({formSubmit: true});
        setTimeout(() => {
          this.setState({formSubmit: false});
        }, 5000);
      }
    }
  }

  public async submitRecall() {
    this.setState({hideDialog: false});
    let updateObject: any = {OffBoardingCheckListId:{'__metadata': { type: 'Collection(Edm.Int32)' },'results': []}, OffBoardingCheckListNAId: {'__metadata': { type: 'Collection(Edm.Int32)' },'results': []}};
    updateObject.LeaveType = null;
    updateObject.LastWorkingDay = null;
    updateObject.OffBoardingStartDate = null;
    updateObject.OffBoardingCompletionDate = null;
    updateObject.OffBoardingStatus = 'Recalled';
    await SPDataOperations.updateListItem(this.props.onboardingList, this.offboardingId, updateObject);
    this.getEmployeeCheckList(this.offboardingId);
  }

  public render(): React.ReactElement<IOffboardingCheckListProps> {
    const {checkList, formSubmit, hideConfirmDialog, hideDialog, checkedListTotalCount, submitRecall, offboardingStatus, lastWorkingDate, leaveType, leaveTypeOption, emplayeeName} = this.state;
    return (
      <div className={ styles.offboardingCheckList }>
        <div className={styles.tableContainer}>
            <div className={styles.userDetail}><span>User Name:</span> {emplayeeName}</div>
            <table className={styles.checkListTable}>
              <tr>
                <th>N/A?</th>
                <th>Checklist</th>
                <th></th>
              </tr>
              {checkList.map((item) =>{
                return (<tr>
                  {item.colType === 'heading' ?
                    <td colSpan={3} className={styles.checklistHeading}><Label>{item.Title}</Label></td>
                  :
                  <td>
                    {item.colType !== 'heading' && item.notApplicable &&
                      <Toggle label='' inlineLabel onText='Yes' offText='No' onChange={() => this.toggleOnChange(item)} defaultChecked={item.notApplicable} disabled={item.required} title={item.required ? 'Checklist Required' : null }  />
                    }
                    {item.colType !== 'heading' && !item.notApplicable &&
                      <Toggle label='' inlineLabel onText='Yes' offText='No' onChange={() => this.toggleOnChange(item)} defaultChecked={item.notApplicable} disabled={item.required} title={item.required ? 'Checklist Required' : null }  />
                    }
                  </td>
                  }
                  {item.colType !== 'heading' &&
                  <td>
                    <Checkbox label={item.Title + (item.required ? '*' : '')} onChange={() => this.checkBoxOnChange(item)} checked={item.selected} disabled={item.notApplicable} />
                  </td>
                  }
                  {item.colType !== 'heading' &&
                  <td>
                    {item.colType === 'date' &&
                      <div>
                        <DatePicker
                          value={lastWorkingDate}
                          disabled={item.selected ? false : true}
                          isRequired={item.selected ? true : false}
                          showMonthPickerAsOverlay={true}
                          minDate={new Date()}
                          placeholder="Select last working date..."
                          ariaLabel="Select last working date"
                          onSelectDate={date => this.setState({ lastWorkingDate: date })}
                        ></DatePicker>
                      </div>
                    }
                    {item.colType === 'dropdown' &&
                      <Dropdown
                        placeholder="Select Leaver Type"
                        options={item.selected ? leaveTypeOption : []}
                        disabled={item.selected ? false : true}
                        defaultSelectedKey={leaveType}
                        required={item.selected ? true : false}
                        onChange={(event, ddvalue) => this.setState({leaveType: ddvalue.text})}
                      />
                    }
                  </td>
                  }
                </tr>);
              })}
            </table>
            <div className={styles.footerButtons}>
              {formSubmit === true &&
                <MessageBar messageBarType={MessageBarType.success}  onDismiss={() => this.setState({formSubmit:false})} dismissButtonAriaLabel="Close">Checklist items are saved successfully.</MessageBar>
              }
              <DefaultButton text={'Recall'} onClick={() => this.setState({submitRecall: true, hideConfirmDialog: false})} disabled={(checkedListTotalCount === 0 || offboardingStatus === 'Closed') ? true : false}></DefaultButton>
              <PrimaryButton text={'Submit'} onClick={() => this.submitCheckList(false)} disabled={(offboardingStatus === 'Closed') ? true : false} style={{marginLeft: '10px'}} />
            </div>
          </div>

          <Dialog
          hidden={hideConfirmDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Alert!'
          }}
        >
          {submitRecall ?
            <p>All offboarding activities will be set to default. Press ‘Ok’ to continue or ‘Cancel’ to go back.</p>
          :
            <p>You are going to submit all the checklist. Please confirm?</p>
          }
          <DialogFooter>
          {submitRecall ?
            <PrimaryButton text="Ok" onClick={() => this.submitRecall()} />
          :
            <PrimaryButton text="Submit" onClick={() => this.submitCheckList(true)} />
          }
            <DefaultButton text="Cancel" onClick={() => this.setState({hideConfirmDialog: true})} />
          </DialogFooter>
        </Dialog>

        <Dialog
          hidden={hideDialog}
          dialogContentProps={{
            type: DialogType.normal
          }}
        >
          <Spinner label="Please wait..." />
        </Dialog>

      </div>
    );
  }
}
