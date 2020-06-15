import * as React from 'react';
import styles from './Training.module.scss';
import { ITrainingProps } from './ITrainingProps';
import { SPDataOperations } from '../../../common/SPDataOperations';
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';
import { DisplayMode } from '@microsoft/sp-core-library';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { MessageBar, MessageBarType, Link} from 'office-ui-fabric-react';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
export interface ITrainingState{
  module: any;
  subModule:any;
  allTraining:any;
  selectedIDs:any;
  allTrainingId:any;
  selectedTrainingId:any;
  isFilterOpen:string;
  isClose:boolean;
  assessmentStatus:boolean;
  assessmentModule: any;
  assessmentParm:boolean;
  isLoading:boolean
}
export default class Training extends React.Component<ITrainingProps, ITrainingState> {
 
  constructor(props) {
    super(props);
 
    this.state = {
      module:[],
      subModule:[],
      allTraining:[],
      selectedIDs:[],
      allTrainingId:[],
      selectedTrainingId:'',
      isFilterOpen: 'none',
      isClose:true,
      assessmentStatus:false,
      assessmentModule:{},
      assessmentParm:true,
      isLoading:false
    };
 
    this.onConfigure = this.onConfigure.bind(this);
    this._onChange = this._onChange.bind(this);
    this.saveDraftVersion = this.saveDraftVersion.bind(this);
    this.checkCondition = this.checkCondition.bind(this);
    this.toggleFilters = this.toggleFilters.bind(this);
  }
 
  public componentDidMount() {
    this.renderTrainigModule();
  }
 
  public componentDidUpdate(prevProps: ITrainingProps) {
    /* Render updated topics when the selected subject property value is updated in the web part*/
    if (prevProps.selectedList !== this.props.selectedList || prevProps.userAssessment !== this.props.userAssessment) {
      this.renderTrainigModule();
    }
  }
 
  public renderTrainigModule(){
    if(this.props.userTraining){
      SPDataOperations.LOADSubModuleData(this.props.selectedList,this.props.context.pageContext.user.email,this.props.userTraining).then((allTrainigs) => {
        this.setState({ module: allTrainigs.module,subModule: allTrainigs.subModule,allTraining: allTrainigs.trainingData ,selectedIDs:allTrainigs.selectedTraining, allTrainingId:allTrainigs.trainingIds, selectedTrainingId:allTrainigs.selectedTraining.join(",")});
      });
    }
 
    SPDataOperations.GetAssessmentStatus(this.props.userAssessment,this.props.context.pageContext.user.email).then((assessment) =>{
      let assessmentStatus:boolean = false;
      if(assessment.assessmentStatus !== 'Pass' && assessment.attemptId !== 0){
        assessmentStatus = true;
      }
      if(assessment.assessmentStatus==='Fail' && assessment.totalAttempt===3){
        assessmentStatus = false;
      }
      this.setState({assessmentStatus:assessmentStatus,assessmentModule:assessment});
    });
 
    let queryParms = new UrlQueryParameterCollection(window.location.href);
    let assessmentParm:any = queryParms.getValue("assessment");
    if(assessmentParm === true || assessmentParm === 'true'){
      this.setState({assessmentParm:false});
    }
  }
 
  private onConfigure(): void {
    this.props.context.propertyPane.open();
  }
 
  public _onChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean) {
    let itemID:any = +ev.currentTarget.getAttribute("aria-label");
    let stateIDs:any = this.state.selectedIDs;
    if(isChecked && stateIDs.indexOf(itemID) === -1){
      stateIDs.push(itemID);
    } else {
      stateIDs = stateIDs.filter(function(item) {return item !== itemID});
    }
    this.setState({selectedIDs:stateIDs});
  }
 
  public toggleFilters(ev) {
    let nameid:any = ev;
    if(this.state.isFilterOpen === nameid){
      this.setState({isFilterOpen: ''})
    } else {
      this.setState({isFilterOpen: nameid})
    }
  }
 
  public checkCondition(){
    let selectedTrainingIDs:any = this.state.selectedTrainingId.split(",").map(Number);
    let trainingID:any = this.state.allTrainingId;
    let currentModule = [];
 
    Object.keys(trainingID).map((module) =>{
      let trainingIDs = trainingID[module];
      const singleFound = trainingIDs.some(r=> selectedTrainingIDs.indexOf(r) >= 0);
      const allFound = trainingIDs.every(v => selectedTrainingIDs.includes(v));
      const notFound = trainingIDs.every(v => !selectedTrainingIDs.includes(v));
      const returnVal = selectedTrainingIDs.length === 0 ? this.state.module[0] : notFound ? module : (singleFound && !allFound) ? module : '';
      if(returnVal !== ''){
        currentModule.push(returnVal);
      }
    });
    return currentModule[0];
  }
 
  public saveDraftVersion(flag?:any){
    this.setState({isLoading:true});
    let selectedTrainingIDs:any = this.state.selectedIDs;
    let trainingID:any = this.state.allTrainingId;
    let currentModule:any = this.checkCondition();
    let ModuleStatus:any = "";
    Object.keys(trainingID).map((module) =>{
      let trainingIDs = trainingID[module];
      const allFound = trainingIDs.every(v => selectedTrainingIDs.includes(v));
      if(allFound && currentModule === module){
        ModuleStatus = module;
      }
    });
 
    if(ModuleStatus != "" && flag === 1){
      this.setState({isClose: false});
    }
 
    if((flag === 1 && ModuleStatus ==="") || flag === 2){
      SPDataOperations.UpdateTrainings(this.props.userTraining, this.state.selectedIDs, this.props.context,ModuleStatus, this.props.userAssessment).then((allTrainigs) => {
      });
    }
    this.setState({isLoading:false});
  }
 
  public render(): React.ReactElement<ITrainingProps> {
    //console.log(this.state);
    let enableModule = this.checkCondition();
    let assessmentStatus = this.state.assessmentStatus;
    let assessmentModule = this.state.assessmentModule.assessmentAllData;
    let completedAssessment = 'Pass';
    if (this.props.configured) {
    return (<div className={styles.training}>
      {(assessmentStatus === true) &&
        <MessageBar
        messageBarType={MessageBarType.warning}
        isMultiline={false}
        >
          Please complete the assessment to enable the next trainings
        </MessageBar>
      }
      {this.state.module.map((module) =>{
        if(assessmentStatus===true && assessmentModule.Title === module){
          completedAssessment = (assessmentModule.AssessmentStatus==='' || assessmentModule.AssessmentStatus===null) ? "Pending" : assessmentModule.AssessmentStatus;
        }else if(assessmentModule.Title !== module && enableModule === module || assessmentModule.Attempt === 3 && enableModule === module){
          completedAssessment = "Not Started";
        }
        return (<div className={styles.module}>
          <div className={styles.moduleHeading}><h5 onClick={() =>this.toggleFilters(module)}>{module}
          <span>Assessment Status: <label style={{color: completedAssessment==='Pass'?'green':completedAssessment==='Fail'?'red':completedAssessment==='Pending'?'#ffbf00':''}}>{completedAssessment}</label></span>
          </h5></div>
          <div className={styles.subModule} style={{display: (this.state.isFilterOpen === module || enableModule === module) ? '':'none'}}>
          {this.state.subModule[module].map((submodule) =>{
            return (<table>
                <tr>
                  <th style={{width:'24px'}}>#</th>
                  <th>{submodule}</th>
                  <th style={{width:'82px'}}>Status</th>
                </tr>
              {this.state.allTraining[submodule].map((training) =>{
                if(training.Module === module && training.SubModule === submodule){
                return(<tr>
                  <td>
                    {this.state.selectedIDs.indexOf(training.Id) === -1 &&
                    <Checkbox key={training.Id} disabled={enableModule !== module || assessmentStatus === true} ariaLabel={training.Id} onChange={this._onChange} />
                    }
                    {this.state.selectedIDs.indexOf(training.Id) > -1 &&
                    <Checkbox key={training.Id} disabled={enableModule !== module || assessmentStatus === true} ariaLabel={training.Id} defaultChecked onChange={this._onChange} />
                    }
                    </td>
                  <td>
                    <Link data-interception="off" disabled={enableModule !== module || assessmentStatus === true } target="_Blank" href={training.TrainingPath.Url}>{training.Title}</Link></td>
                  <td>
                  {this.state.selectedIDs.indexOf(training.Id) === -1 &&
                    <span>Pending</span>
                    }
                    {this.state.selectedIDs.indexOf(training.Id) > -1 &&
                    <span>Completed</span>
                    }
                  </td>
                </tr>)
                }
              })}
              </table>)
          })}
          </div>
        </div>);
      })}
    <div className={styles.footerButtons}>
    <Dialog
        hidden={this.state.assessmentParm}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: 'Module Submitted',
          closeButtonAriaLabel: 'Close',
          subText: this.props.moduleCompletionMsg
        }}
        containerClassName={styles.alertdialogContainer}
    >
        <DialogFooter>
          <PrimaryButton href={this.props.URLForYes} text="Yes" onClick={() => { this.setState({assessmentParm: true});}} />
          <DefaultButton href={this.props.URLForNo} text="No" onClick={() => { this.setState({assessmentParm: true});}} />
    </DialogFooter>
    </Dialog>
 
    <Dialog
      hidden={this.state.isClose}
      dialogContentProps={{
        type: DialogType.largeHeader,
        title: 'Alert!'
      }}
      containerClassName={styles.alertdialogContainer}
    >
      <div>
        <div dangerouslySetInnerHTML={{__html: this.props.moduleSubmittionMsg }} />
      </div>
      <DialogFooter>
        <PrimaryButton onClick={() => {this.saveDraftVersion(2)}} text="OK" disabled={this.state.isLoading} />
        <DefaultButton onClick={() => { this.setState({isClose: true})}}  text="Cancel" />
      </DialogFooter>
    </Dialog>
    <PrimaryButton iconProps={{iconName:"Draft"}} text="Save" disabled={(assessmentStatus === true || enableModule === undefined) ? true : false} onClick={() => {this.saveDraftVersion(1)}} />
      </div>
    </div>);
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