import * as React from 'react';
import styles from './TempTraining.module.scss';
import { ITempTrainingProps } from './ITempTrainingProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Checkbox, DefaultButton, Dialog, DialogFooter, DialogType, Link, MessageBar, MessageBarType, PrimaryButton, Spinner } from 'office-ui-fabric-react';
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';
import { DisplayMode } from '@microsoft/sp-core-library';
import { SPDataOperations } from '../../../common/SPDataOperations';
import * as moment from 'moment';
export interface ITempTrainingState {
  currentUser: string;
  allTraining: any[],
  selectedTraining: number[],
  openDialogBox: boolean,
  isLoading: boolean,
  agreement: boolean,
  trainingModule: any[]
}

export default class TempTraining extends React.Component<ITempTrainingProps, ITempTrainingState> {
  constructor(props) {
    super(props);
 
    this.state = {
      currentUser: null,
      allTraining:[],
      selectedTraining: [],
      openDialogBox: false,
      isLoading:false,
      agreement: false,
      trainingModule: []
    };
 
    this.onConfigure = this.onConfigure.bind(this);
    this._onChange = this._onChange.bind(this);
    this._onChangeAgreement = this._onChangeAgreement.bind(this);
  }

  public componentDidMount() {
    if(this.props.tempTrainingUserList !== undefined){
      this.getTemporaryTrainings();
    }
  }
 
  public componentDidUpdate(prevProps: ITempTrainingProps) {
    if (prevProps.tempTrainingUserList !== this.props.tempTrainingUserList) {
      this.getTemporaryTrainings();
    }
  }

  private async getTemporaryTrainings(){
    let trainingModule: any[] = this.state.trainingModule;
    this.setState({isLoading: true});
    const userDetail: any = await SPDataOperations.getLoggedInUserDetails(this.props.context.pageContext);
    const today = moment(new Date()).format("YYYY-MM-DD");
    const currentDate = today+'T00:00:00.000Z';
    SPDataOperations.getListItems(this.props.tempTrainingUserList,'Id,Title,StartDate,DueDate,Status,Training/CalcModule,Training/Title,Training/TrainingPath,Employee/Id,Employee/EMail','Training,Employee',`Employee/EMail eq '${this.props.context.pageContext.user.email}' and DueDate ge datetime'${currentDate}'`).then((allTrainigs) => {
      allTrainigs.map((item) => {
        if(trainingModule.indexOf(item.Training.CalcModule) === -1){
          trainingModule.push(item.Training.CalcModule);
        }
      });
      this.setState({currentUser: userDetail.Title, allTraining: allTrainigs, isLoading: false, selectedTraining: [], agreement: allTrainigs.filter((training) => training.Status !== 'Completed').length > 0 ? false : true, trainingModule: trainingModule});
    });
  }

  public _onChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean) {
    let itemID: number = +ev.currentTarget.getAttribute("aria-label");
    let selectedTraining:number[] = this.state.selectedTraining;
    if(isChecked && selectedTraining.indexOf(itemID) === -1){
      selectedTraining.push(itemID);
    } else {
      selectedTraining = selectedTraining.filter(function(item) {return item !== itemID});
    }
    this.setState({selectedTraining:selectedTraining});
  }

  private _onChangeAgreement(ev: React.FormEvent<HTMLElement>, isChecked: boolean){
      this.setState({agreement:isChecked});
  }

  private async updateTrainingStatus(){
    this.setState({isLoading: true, openDialogBox: false});
    const selectedTraining = this.state.selectedTraining;

    await Promise.all(selectedTraining.map(async (itemId) => {
      const jsonData = {
        CompletedDate: new Date(),
        Status: 'Completed'
      }
      await SPDataOperations.updateListItem(this.props.tempTrainingUserList, itemId, jsonData);
    }));
    await this.getTemporaryTrainings();
    
  }

  private onConfigure(): void {
    this.props.context.propertyPane.open();
  }

  public render(): React.ReactElement<ITempTrainingProps> {
    const {currentUser, allTraining, selectedTraining, agreement, trainingModule} = this.state;

    if (this.props.configured) {
    return (
      <div className={ styles.tempTraining }>
        <div><h3>User: {currentUser}</h3></div>

        {trainingModule.map((module) => {
          return(<div className={styles.module}>
            <div className={styles.moduleHeading}><h5>{module}</h5></div>
            <div className={styles.subModule}>
                <table>
                  <tr>
                    <th>#</th>
                    <th>Training(s)</th>
                    <th>Start Date</th>
                    <th>Due Date</th>
                    <th>Status</th>
                  </tr>
                  {allTraining.filter((trainingModule) => trainingModule.Training.CalcModule === module).map((training) => {
                    return(<tr>
                      <td>
                        {training.Status === 'Completed' ?
                          <Checkbox key={training.Id} ariaLabel={training.Id} defaultChecked={true} disabled={true} />
                        :
                          <Checkbox key={training.Id} ariaLabel={training.Id} defaultChecked={selectedTraining.indexOf(training.Id) > -1 ? true : false} disabled={training.Status === 'Completed' ? true : false} onChange={this._onChange} />
                        }
                        </td>
                      <td><Link data-interception="off"  target="_Blank" href={training.Training.TrainingPath}>{training.Training.Title}</Link></td>
                      <td><span>{moment(new Date(training.StartDate)).format("DD-MM-YYYY")}</span></td>
                      <td><span>{moment(new Date(training.DueDate)).format("DD-MM-YYYY")}</span></td>
                      <td><span>{training.Status}</span></td>
                    </tr>);
                  })}
                </table>
            </div>
          </div>);
        })}

        {trainingModule.length > 0 ?
        <div>
          <div className={styles.bottomRow}>
            <div>
              <Checkbox defaultChecked={agreement} disabled={allTraining.filter((training) => training.Status !== 'Completed').length > 0 ? false : true} onChange={this._onChangeAgreement} label={this.props.agreementText}></Checkbox>
            </div>
            <PrimaryButton text="Save" onClick={() => this.setState({openDialogBox: true})} disabled={selectedTraining.length > 0 && agreement ? false : true}></PrimaryButton>
          </div>
          <Dialog
            hidden={!this.state.openDialogBox}
            dialogContentProps={{
              type: DialogType.normal,
              title: 'Confirm!'
            }}
          >
            <div>
              <p>Are you sure you want to update the trainings?</p>
            </div>
            <DialogFooter>
              <PrimaryButton onClick={() => {this.updateTrainingStatus()}} text="OK" />
              <DefaultButton onClick={() => { this.setState({openDialogBox: false})}}  text="Cancel" />
            </DialogFooter>
          </Dialog>
          </div>
            :
          <div>
            <MessageBar messageBarType={MessageBarType.warning}>You don't have any pending training.</MessageBar>
            <Dialog
              hidden={!this.state.isLoading}
            >
              <Spinner label="Please wait..." />
            </Dialog>
          </div>
        }
      </div>
    );
    } else {
      return (
        <Placeholder iconName='Edit'
          iconText='Configure your web part'
          description='Please configure the web part.'
          buttonLabel='Configure'
          hideButton={this.props.displayMode === DisplayMode.Read}
          onConfigure={this.onConfigure} />
      );
    }
  }
}
