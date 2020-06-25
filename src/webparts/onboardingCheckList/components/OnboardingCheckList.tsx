import * as React from 'react';
import styles from './OnboardingCheckList.module.scss';
import { IOnboardingCheckListProps } from './IOnboardingCheckListProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPDataOperations } from '../../../common/SPDataOperations';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import { Toggle, Checkbox, PrimaryButton, Dialog, DialogType, Spinner } from 'office-ui-fabric-react';

export interface IOnboardingCheckListState {
  employeeData: any[];
  checkList: any[];
  hideDialog: boolean;
}
export default class OnboardingCheckList extends React.Component<IOnboardingCheckListProps, IOnboardingCheckListState> {
  private onboardingId: number = null;
  constructor(props) {
    super(props);
    this.toggleOnChange = this.toggleOnChange.bind(this);
    this.checkBoxOnChange = this.checkBoxOnChange.bind(this);
    this.submitCheckList = this.submitCheckList.bind(this);
    this.state = {
      employeeData: [],
      checkList: [],
      hideDialog: true
    }
  }

  public componentDidMount() {

    let queryParms = new UrlQueryParameterCollection(window.location.href);
    const onboardingId: string = queryParms.getValue('onboardID');
    if(onboardingId !== undefined){
      this.onboardingId = Number(onboardingId);
      this.getEmployeeCheckList(onboardingId);
    }
  }
  
  public getEmployeeCheckList(empId: string) {
    SPDataOperations.getListItems(this.props.onboardingList, 'Employee,OnBoardingCheckList/Id,OnBoardingCheckListNA/Id','OnBoardingCheckList,OnBoardingCheckListNA',`Id eq ${empId}`).then((onboardingData) => {
      if(onboardingData.length > 0 ) {
        this.getCheckListDetails(onboardingData[0]);
      }
    });
  }

  public getCheckListDetails(employeeData: any) {
    SPDataOperations.getListItems(this.props.checkList, 'Id,Title', ``, ``).then((checkList) => {
      console.log(employeeData);
      let updatedChecklist = [];
      checkList.map((listItem) => {
        const checkListObject: any = {
          Id: listItem.Id,
          Title: listItem.Title,
          selected: employeeData.OnBoardingCheckList.filter(item => listItem.Id === item.Id).length > 0 ? true : false,
          isApplicable: employeeData.OnBoardingCheckListNA.filter(item => listItem.Id === item.Id).length > 0 ? true : false
        };
        updatedChecklist.push(checkListObject);
      });
      this.setState({checkList: updatedChecklist, employeeData: employeeData});
    });
  }

  public toggleOnChange(item: any) {
    let checkList = this.state.checkList.slice();
    checkList.filter(listItem => listItem.Id === item.Id).map(selectedItem => {
      selectedItem.isApplicable = !item.isApplicable;
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

  public async submitCheckList() {
    this.setState({hideDialog: false});
    let checkList = this.state.checkList.slice();
    let selectedCheckList: any[] = [];
    let checkListNA: any[] = [];
    checkList.map((item) => {
      if (item.selected === true) {
        selectedCheckList.push(item.Id);
      }
      if (item.isApplicable === true) {
        checkListNA.push(item.Id);
      }
    });

    await SPDataOperations.updateListItem(this.props.onboardingList, this.onboardingId, {OnBoardingCheckListId: {'__metadata': { type: 'Collection(Edm.Int32)' },'results': selectedCheckList}, OnBoardingCheckListNAId: {'__metadata': { type: 'Collection(Edm.Int32)' },'results': checkListNA}} );
    this.setState({hideDialog: true});
  }

  public render(): React.ReactElement<IOnboardingCheckListProps> {
    const {checkList, hideDialog} = this.state;
    return (
      <div className={ styles.onboardingCheckList }>
        <div className={styles.tableContainer}>
          <table className={styles.checkListTable}>
            <tr>
              <th>NA?</th>
              <th>Checklist</th>
            </tr>
            {checkList.map((item) =>{
              return (<tr>
                <td><Toggle label='' inlineLabel onText='Yes' offText='No' onChange={() => this.toggleOnChange(item)} defaultChecked={item.isApplicable}  /></td>
                <td><Checkbox label={item.Title} onChange={() => this.checkBoxOnChange(item)} checked={item.selected} disabled={item.isApplicable} /></td>
                </tr>)
            })}
          </table>
          <PrimaryButton text="Submit" onClick={this.submitCheckList} />
        </div>
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
