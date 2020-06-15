import * as React from 'react';
import styles from './AssessmentHome.module.scss';
import { IAssessmentHomeProps } from './IAssessmentHomeProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPDataOperations } from '../../../common/SPDataOperations';
import { PrimaryButton } from 'office-ui-fabric-react';
export interface IAssessmentHomeState{
  moduleAssessment:any[];
  assessmentStatus:string;
}
export default class AssessmentHome extends React.Component<IAssessmentHomeProps, IAssessmentHomeState> {
  constructor(props) {
    super(props);
 
    this.state = {
      moduleAssessment:[],
      assessmentStatus:''
    };
  }
 
  public componentDidMount() {
    SPDataOperations.LOADCurrentUserAssessment(this.props.selectedList, this.props.assessmentList,1,this.props.context.pageContext.user.email,this.props.userAssessmentList).then((assessment) => {
      this.setState({moduleAssessment:assessment.assessmentData});
      if(assessment.assessmentData.length===0){
        this.setState({assessmentStatus:'You have no assessment pending!'});
      }
    });
  }
  public render(): React.ReactElement<IAssessmentHomeProps> {
    return (
      <div className={ styles.assessmentHome }>
        <div className={ styles.container }>
          <div className={ styles.row }>
              {this.state.moduleAssessment.length > 0 &&
                <PrimaryButton href={this.props.description}>Start Assessment</PrimaryButton>
              }
              {(this.state.moduleAssessment.length === 0 && this.state.assessmentStatus !=='') &&
               <div style={{textAlign:"center"}}>
               <img style={{width:'auto'}} src="/sites/ROOT/RootAssets/Images/Yay.jpg" />
               <h2>{this.state.assessmentStatus}</h2>
             </div>
              }
          </div>
        </div>
      </div>
    );
  }
}