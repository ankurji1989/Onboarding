import * as React from 'react';
import styles from './OffboardingCheckList.module.scss';
import { IOffboardingCheckListProps } from './IOffboardingCheckListProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import { SPDataOperations } from '../../../common/SPDataOperations';

export default class OffboardingCheckList extends React.Component<IOffboardingCheckListProps, {}> {
  private onboardingId: number = null;
  private checkListCompleted: boolean = false;

  public componentDidMount() {
    const queryParms = new UrlQueryParameterCollection(window.location.href);
    const onboardingId: string = queryParms.getValue('offboardID');
    if(onboardingId !== undefined){
      this.onboardingId = Number(onboardingId);
      this.getEmployeeCheckList(Number(onboardingId));
    }
  }
  
  public async getEmployeeCheckList(empId: number) {
    SPDataOperations.getListItems(this.props.onboardingList, 'Employee,RegistrationListItemID,OnBoardingCheckList/Id,OnBoardingCheckListNA/Id,EmployeeID1/Id,CheckListCompleted','OnBoardingCheckList,OnBoardingCheckListNA,EmployeeID1',`Id eq ${empId}`).then((onboardingData) => {
      if(onboardingData.length > 0 ) {
        this.checkListCompleted = onboardingData[0].CheckListCompleted === 'Complete' ? true : false;
        this.getCheckListDetails(onboardingData[0]);
      }
    });
  }

  public getCheckListDetails(employeeData: any) {
    console.log(employeeData);
  }
  
  public render(): React.ReactElement<IOffboardingCheckListProps> {
    return (
      <div className={ styles.offboardingCheckList }>
        <div className={styles.tableContainer}>
            <table className={styles.checkListTable}>
              <tr>
                <th>NA?</th>
                <th>Checklist</th>
                <th>Attachment</th>
              </tr>
              <tr>
                <td>d</td>
                <td>df</td>
                <td>df</td>
              </tr>
            </table>
          </div>
      </div>
    );
  }
}
