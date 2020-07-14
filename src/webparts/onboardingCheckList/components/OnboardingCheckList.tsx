import * as React from 'react';
import styles from './OnboardingCheckList.module.scss';
import { IOnboardingCheckListProps } from './IOnboardingCheckListProps';
import { escape, get } from '@microsoft/sp-lodash-subset';
import { SPDataOperations } from '../../../common/SPDataOperations';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import { Toggle, Checkbox, PrimaryButton, Dialog, DialogType, Spinner, DialogFooter, DefaultButton, Button, ActionButton, Link, MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

export interface IOnboardingCheckListState {
  employeeData: any[];
  checkList: any[];
  hideDialog: boolean;
  hideEmailDialog: boolean;
  hideConfirmDialog: boolean;
  registrationListItemID: number;
  userIdCreation: boolean;
  filePickerResult:any;
  buttonDisabled: boolean;
}

export default class OnboardingCheckList extends React.Component<IOnboardingCheckListProps, IOnboardingCheckListState> {
  private onboardingId: number = null;
  private userLoginName: string = null;
  private checkListCompleted: boolean = false;
  constructor(props) {
    super(props);
    this.toggleOnChange = this.toggleOnChange.bind(this);
    this.checkBoxOnChange = this.checkBoxOnChange.bind(this);
    this.submitCheckList = this.submitCheckList.bind(this);
    this.updateUserEmail = this.updateUserEmail.bind(this);
    this.getPeoplePickerItems = this.getPeoplePickerItems.bind(this);
    this.uploadAttachment = this.uploadAttachment.bind(this);
    this.state = {
      employeeData: [],
      checkList: [],
      hideDialog: true,
      hideEmailDialog: true,
      hideConfirmDialog: true,
      registrationListItemID: 0,
      userIdCreation: false,
      filePickerResult: null,
      buttonDisabled: false
    }
  }

  public componentDidMount() {
    let queryParms = new UrlQueryParameterCollection(window.location.href);
    const onboardingId: string = queryParms.getValue('onboardID');
    if(onboardingId !== undefined){
      this.onboardingId = Number(onboardingId);
      this.getEmployeeCheckList(Number(onboardingId));
    }
  }
  
  public async getEmployeeCheckList(empId: number) {
    let allAttachment: any[] = [];
    let getAttachement: any[] = await SPDataOperations.getAttachment(this.props.onboardingList, empId);
    getAttachement.map((file) => {
      allAttachment.push({fileNo: Number(file.FileName.split('.')[0]), fileName: file.FileName, filePath: file.ServerRelativeUrl});
    });
    SPDataOperations.getListItems(this.props.onboardingList, 'Employee,RegistrationListItemID,OnBoardingCheckList/Id,OnBoardingCheckListNA/Id,EmployeeID1/Id,CheckListCompleted','OnBoardingCheckList,OnBoardingCheckListNA,EmployeeID1',`Id eq ${empId}`).then((onboardingData) => {
      if(onboardingData.length > 0 ) {
        this.checkListCompleted = onboardingData[0].CheckListCompleted === 'Complete' ? true : false;
        this.getCheckListDetails(onboardingData[0], allAttachment);
      }
    });
  }

  public getCheckListDetails(employeeData: any, attachment: any[]) {
    SPDataOperations.getListItems(this.props.checkList, 'Id,Title,attachmentRequired,Order0,Required', ``, `CheckListStatus eq 1`).then((checkList) => {
      let updatedChecklist = [];
      checkList.sort((item1, item2) => item1.Order0 - item2.Order0).map((listItem) => {
        const attachmentFile = attachment.filter(file => file.fileNo === listItem.Id);
        const checkListObject: any = {
          Id: listItem.Id,
          Title: listItem.Title,
          selected: employeeData.OnBoardingCheckList.filter(item => listItem.Id === item.Id).length > 0 ? true : false,
          notApplicable: employeeData.OnBoardingCheckListNA.filter(item => listItem.Id === item.Id).length > 0 ? true : false,
          attachmentRequired: listItem.attachmentRequired === true ?  true : false,
          attachmentFile: attachmentFile.length > 0 ? attachmentFile[0] : null,
          required: listItem.Required === true ? true : false
        };
        updatedChecklist.push(checkListObject);
      });
      const userIdCreation = (employeeData.EmployeeID1 !== null && employeeData.EmployeeID1 !== undefined) ? true : false;
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
      let buttonDisabled: boolean = checkListCount === CheckedListCount ? true : false;
      this.setState({checkList: updatedChecklist, employeeData: employeeData, registrationListItemID: employeeData.RegistrationListItemID, userIdCreation: userIdCreation, buttonDisabled: buttonDisabled});
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
    let checkList = this.state.checkList.slice();
    checkList.filter(listItem => listItem.Id === item.Id).map(selectedItem => {
      selectedItem.selected = !item.selected;
    });
    this.setState({checkList: checkList});
  }

  public async submitCheckList(confirmDialog: boolean) {
    let checkList = this.state.checkList.slice();
    let selectedCheckList: any[] = [];
    let checkListNA: any[] = [];
    let checkListCount: number = checkList.length;
    let selectedCheckListCount: number = 0;
    checkList.map((item) => {
      if (item.selected === true) {
        selectedCheckList.push(item.Id);
        selectedCheckListCount++;
      }
      if (item.notApplicable === true) {
        checkListNA.push(item.Id);
        selectedCheckListCount++;
      }
    });

    let buttonDisabled: boolean = false;
    let updateObject: any = {OnBoardingCheckListId:{'__metadata': { type: 'Collection(Edm.Int32)' },'results': selectedCheckList}, OnBoardingCheckListNAId: {'__metadata': { type: 'Collection(Edm.Int32)' },'results': checkListNA}};

    if(checkListCount === selectedCheckListCount && confirmDialog === false) {
      this.setState({hideConfirmDialog: false});
    } else {
      this.setState({hideDialog: false});
      if(checkListCount === selectedCheckListCount){
        updateObject.OnBoardingStatus = 'Completed';
        updateObject.CheckListCompleted = 'Complete';
        updateObject.CompletedDate = new Date();
        updateObject.Flag = 1;
        buttonDisabled = true;
      } else {
        updateObject.OnBoardingStatus = 'In Progress';
      }
      await SPDataOperations.updateListItem(this.props.onboardingList, this.onboardingId, updateObject);
      this.setState({hideDialog: true, hideConfirmDialog: true, buttonDisabled: buttonDisabled});
    }

  }

  public getPeoplePickerItems(items: any[]) {
    const uLoginName: string = items.length > 0 ? items[0].loginName : null;
    this.userLoginName = uLoginName;
  }

  public async updateUserEmail() {
    this.setState({hideDialog: false});
    if(this.userLoginName !== null) {
      const userDetail: any = await SPDataOperations.getUserID(this.userLoginName);
      await SPDataOperations.updateListItem(this.props.onboardingList, this.onboardingId, {EmployeeID1Id: userDetail.Id});
      await SPDataOperations.updateListItem(this.props.registrationList, this.state.registrationListItemID, {EmployeeID1Id: userDetail.Id});
      this.setState({hideEmailDialog: true, hideDialog: true, userIdCreation: true});
    }
  }

  public uploadAttachment(fileNo: number): void {
    this.setState({hideDialog:false});
    let file: any = document.getElementById("fileUpload");
    let fileName: string = null;
    if(file) {
      file = file.files[0];
      SPDataOperations.addAttachment(this.props.onboardingList, this.onboardingId, fileNo, file);
    }
  }

  public deleteAttachment(fileName: string) {
    let confirmDelete = confirm('Are you sure, you want to delete the attachment?');
    if (confirmDelete) {
      SPDataOperations.deleteAttachment(this.props.onboardingList, this.onboardingId, fileName);
    }
  }


  public render(): React.ReactElement<IOnboardingCheckListProps> {
    const {checkList, hideDialog, hideEmailDialog, hideConfirmDialog, userIdCreation, buttonDisabled} = this.state;
    return (
      <div className={ styles.onboardingCheckList }>
        <div className={styles.tableContainer}>

          {checkList.length > 0 &&
          <div>
            <table className={styles.checkListTable}>
              <tr>
                <th>NA?</th>
                <th>Checklist</th>
                <th>Attachment</th>
              </tr>
              {checkList.map((item) =>{
                return (<tr>
                  <td><Toggle label='' inlineLabel onText='Yes' offText='No' onChange={() => this.toggleOnChange(item)} defaultChecked={item.notApplicable} disabled={item.required} title={item.required ? 'Checklist Required' : null }  /></td>
                  <td>
                    {item.Id === 17 && userIdCreation === false ?
                      <DefaultButton text={item.Title} onClick={() => this.setState({hideEmailDialog: false})} disabled={item.notApplicable} />
                    :
                    <div>
                      <Checkbox label={item.Title + (item.required ? '*' : '')} onChange={() => this.checkBoxOnChange(item)} checked={item.selected} disabled={item.notApplicable} />
                      {item.attachmentRequired === true && item.attachmentFile === null && item.notApplicable !== true ?
                      <div className={styles.disabled} onClick={() => alert('Please upload the attachment first to check this option!')}>&nbsp;</div>
                      :
                      ``
                      }
                    </div>
                    }
                  </td>
                  <td>
                    {item.attachmentRequired === true ?
                    <div className={styles.actionButton}>
                      {item.attachmentFile === null ?
                        <div>
                          {this.checkListCompleted !== true &&
                            <label><input type='file' id='fileUpload' name='fileUpload' style={{display:'none'}} onChange={() => this.uploadAttachment(item.Id)} disabled={item.notApplicable} accept='image/*,.pdf'/> Upload</label>
                          }
                          </div>
                        :
                        <div>
                          <Link target='_blank' data-interception='off' href={item.attachmentFile.filePath}>View</Link>
                          &nbsp;|&nbsp;
                          <Link onClick={() => this.deleteAttachment(item.attachmentFile.fileName)} disabled={this.state.buttonDisabled}>Delete</Link>
                        </div>
                      }
                      
                    </div>
                     : 
                     ``
                     }
                  </td>
                  </tr>)
              })}
            </table>
            <div className={styles.footerButtons}>
              <MessageBar messageBarType={MessageBarType.error}>* Items are mandatory</MessageBar>
              <PrimaryButton text="Submit" onClick={() => this.submitCheckList(false)} disabled={buttonDisabled} />
            </div>
          </div>
          }
        </div>
        <Dialog
          hidden={hideEmailDialog}
          onDismiss={() => this.setState({hideEmailDialog: true})}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Update P&G EMail'
          }}
          containerClassName={styles.emailDialogContainer}
        >
          <div>
            <p>Please select the user to complete the checklist.</p>

            <PeoplePicker
              context={this.props.context}
              titleText=''
              personSelectionLimit={1}
              isRequired={true}
              selectedItems={this.getPeoplePickerItems}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000} />
          </div>
          <DialogFooter>
            <PrimaryButton text="Update" onClick={() => this.updateUserEmail()} />
            <DefaultButton text="Cancel" onClick={() => this.setState({hideEmailDialog: true})} />
          </DialogFooter>
        </Dialog>

        <Dialog
          hidden={hideConfirmDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Alert!'
          }}
        >
          <p>You are going to submit all the checklist. Please onfirm?</p>
          <DialogFooter>
            <PrimaryButton text="Submit" onClick={() => this.submitCheckList(true)} />
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
