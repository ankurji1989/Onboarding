import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DisplayMode } from '@microsoft/sp-core-library';
 
export interface IAssessmentProps {
  context: WebPartContext;
  description: string;
  userTrainingList: string;
  displayMode: DisplayMode;
  configured: boolean;
  assessmentList:string;
  totalQuestion:any;
  passingScore:any;
  userAssessmentList: string;
  URLAssessmentHome:string;
}