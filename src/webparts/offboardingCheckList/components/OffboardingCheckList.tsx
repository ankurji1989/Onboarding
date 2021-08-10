import * as React from 'react';
import styles from './OffboardingCheckList.module.scss';
import { IOffboardingCheckListProps } from './IOffboardingCheckListProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import { SPDataOperations } from '../../../common/SPDataOperations';
import { Checkbox, DatePicker, DefaultButton, Dialog, DialogFooter, DialogType, MessageBar, MessageBarType, PrimaryButton, Spinner, Toggle } from 'office-ui-fabric-react';

export interface IOffboardingCheckListState {
  employeeData: any[];
  checkList: any[];
  buttonDisabled: boolean;
  hideDialog: boolean;
  hideConfirmDialog: boolean;
  formSubmit: boolean;
  checkedListTotalCount: number;
  submitRecall: boolean;
  offboardingStatus: string;
  lastWorkingDate?: Date;
}

export default class OffboardingCheckList extends React.Component<IOffboardingCheckListProps, IOffboardingCheckListState> {
  private offboardingId: number = null;
  constructor(props) {
    super(props);
    this.getCheckListDetails = this.getCheckListDetails.bind(this);
    this.toggleOnChange = this.toggleOnChange.bind(this);
    this.checkBoxOnChange = this.checkBoxOnChange.bind(this);
    this.submitCheckList = this.submitCheckList.bind(this);
    this.submitRecall = this.submitRecall.bind(this);
    this.state = {
      employeeData: [],
      checkList: [],
      buttonDisabled: false,
      hideDialog: true,
      hideConfirmDialog: true,
      formSubmit: false,
      checkedListTotalCount: 0,
      submitRecall: false,
      offboardingStatus: null,
      lastWorkingDate: null
    }
  }

  public componentDidMount() {
    const queryParms = new UrlQueryParameterCollection(window.location.href);
    const offboardingId: string = queryParms.getValue('offboardID');
    if(offboardingId !== undefined){
      this.offboardingId = Number(offboardingId);
      this.getEmployeeCheckList(Number(offboardingId));
    }
  }
  
  public async getEmployeeCheckList(empId: number) {
    SPDataOperations.getListItems(this.props.onboardingList, 'Employee,RegistrationListItemID,OffBoardingCheckList/Id,OffBoardingCheckListNA/Id,EmployeeID1/Id,CheckListCompleted,OffBoardingStatus,LastWorkingDay','OffBoardingCheckList,OffBoardingCheckListNA,EmployeeID1',`Id eq ${empId}`).then((offboardingData) => {
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
          colType: listItem.ColType
        };
        updatedChecklist.push(checkListObject);
      });

      const checkListCount: number = updatedChecklist.length;
      let CheckedListCount: number = 0;
      updatedChecklist.map((item) => {
        if (item.selected === true) {
          CheckedListCount++;
        }
        if (item.notApplicable === true) {
          CheckedListCount++
        }
      });

      const buttonDisabled: boolean = checkListCount === CheckedListCount ? true : false;
      const lastWorkingDay = employeeData.LastWorkingDay !== null ? new Date(employeeData.LastWorkingDay) : null;
      this.setState({checkList: updatedChecklist, buttonDisabled: buttonDisabled, checkedListTotalCount: CheckedListCount, offboardingStatus: employeeData.OffBoardingStatus, lastWorkingDate: lastWorkingDay, hideDialog: true, hideConfirmDialog: true});
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
    const {lastWorkingDate, checkList} = this.state;
    checkList.filter(listItem => listItem.Id === item.Id).map(selectedItem => {
      selectedItem.selected = !item.selected;
    });
    this.setState({checkList: checkList, lastWorkingDate: item.Title === 'Last Working Date' && item.selected === false ? null : lastWorkingDate });
  }

  public async submitCheckList(confirmDialog: boolean) {
    const {checkList, lastWorkingDate} = this.state;
    let selectedCheckList: any[] = [];
    let checkListNA: any[] = [];
    let checkListCount: number = checkList.length;
    let selectedCheckListCount: number = 0;
    let lastWorkingDaySelected: boolean = true;

    checkList.map((item) => {
      if(item.Title === 'Last Working Date' && item.selected && lastWorkingDate === null) {
        lastWorkingDaySelected = false;
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

    if(lastWorkingDaySelected) {
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
        if(this.state.lastWorkingDate !== null) {
          updateObject.LastWorkingDay = this.state.lastWorkingDate;
        }
        await SPDataOperations.updateListItem(this.props.onboardingList, this.offboardingId, updateObject);
        this.getEmployeeCheckList(this.offboardingId);
        this.setState({formSubmit: true});  
        setTimeout(() => {
          this.setState({formSubmit: false});  
        }, 5000);
      }
    } else {
      alert('Please select the last working date');
    }
  }

  public async submitRecall() {
    this.setState({hideDialog: false});
    let updateObject: any = {};
    updateObject.OffBoardingStatus = 'Recalled';
    await SPDataOperations.updateListItem(this.props.onboardingList, this.offboardingId, updateObject);
    this.getEmployeeCheckList(this.offboardingId);
  }
  
  public render(): React.ReactElement<IOffboardingCheckListProps> {
    const {checkList, buttonDisabled, formSubmit, hideConfirmDialog, hideDialog, checkedListTotalCount, submitRecall, offboardingStatus, lastWorkingDate} = this.state;
    console.log(lastWorkingDate);
    return (
      <div className={ styles.offboardingCheckList }>
        <div className={styles.tableContainer}>
            <table className={styles.checkListTable}>
              <tr>
                <th>NA?</th>
                <th>Checklist</th>
                <th></th>
              </tr>
              {checkList.map((item) =>{
                return (<tr>
                  <td><Toggle label='' inlineLabel onText='Yes' offText='No' onChange={() => this.toggleOnChange(item)} defaultChecked={item.notApplicable} disabled={item.required} title={item.required ? 'Checklist Required' : null }  /></td>
                  <td>
                    <Checkbox label={item.Title + (item.required ? '*' : '')} onChange={() => this.checkBoxOnChange(item)} checked={item.selected} disabled={item.notApplicable} />
                  </td>
                  <td>
                    {item.colType === 'Date' &&
                      <div>
                        <DatePicker
                          value={lastWorkingDate}
                          disabled={item.selected ? false : true}
                          isRequired={item.selected ? true : false}
                          showMonthPickerAsOverlay={true}
                          placeholder="Select last working date..."
                          ariaLabel="Select last working date"
                          onSelectDate={date => this.setState({ lastWorkingDate: date })}
                        ></DatePicker>
                      </div>
                    }
                  </td>
                </tr>);
              })}
            </table>
            <div className={styles.footerButtons}>
              {formSubmit === true &&
                <MessageBar messageBarType={MessageBarType.success}  onDismiss={() => this.setState({formSubmit:false})} dismissButtonAriaLabel="Close">Checklist items are saved successfully.</MessageBar>
              }
              <DefaultButton text={'Recall'} onClick={() => this.setState({submitRecall: true, hideConfirmDialog: false})} disabled={(checkedListTotalCount === 0 || offboardingStatus === 'Recalled' || offboardingStatus === 'Completed') ? true : false}></DefaultButton>
              <PrimaryButton text={'Submit'} onClick={() => this.submitCheckList(false)} disabled={(buttonDisabled || offboardingStatus === 'Recalled') ? true : false} style={{marginLeft: '10px'}} />
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
