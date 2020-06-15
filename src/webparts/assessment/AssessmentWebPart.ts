import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, PropertyPaneDropdown} from '@microsoft/sp-webpart-base';
import * as strings from 'AssessmentWebPartStrings';
import Assessment from './components/Assessment';
import { IAssessmentProps } from './components/IAssessmentProps';
import { sp } from '@pnp/sp';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
 
export interface IAssessmentWebPartProps {
  description: string;
  userTrainingList: string;
  assessmentList:string;
  totalQuestion:any;
  passingScore:any;
  userAssessmentList: string;
  URLAssessmentHome:string;
}
 
export default class AssessmentWebPart extends BaseClientSideWebPart<IAssessmentWebPartProps> {
  protected onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context
    });
    if (this.properties.totalQuestion === undefined) {
      this.properties.totalQuestion = '20';
    }
    if (this.properties.passingScore === undefined) {
      this.properties.passingScore = '75';
    }
    if (this.properties.URLAssessmentHome === undefined) {
      this.properties.URLAssessmentHome = window.location.href;
    }
    return Promise.resolve();
  }
  public render(): void {
    const element: React.ReactElement<IAssessmentProps > = React.createElement(
      Assessment,
      {
        context: this.context,
        userTrainingList: this.properties.userTrainingList,
        displayMode: this.displayMode,
        configured: (this.properties.userTrainingList && this.properties.assessmentList) ? true : false,
        assessmentList:this.properties.assessmentList,
        totalQuestion: this.properties.totalQuestion,
        passingScore: this.properties.passingScore,
        userAssessmentList: this.properties.userAssessmentList,
        URLAssessmentHome:this.properties.URLAssessmentHome
      }
    );
 
    ReactDom.render(element, this.domElement);
  }
 
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
 
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
 
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldListPicker('userTrainingList', {
                  label: 'Select user training list',
                  selectedList: this.properties.userTrainingList,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                }),
                PropertyFieldListPicker('assessmentList', {
                  label: 'Select assessment master list',
                  selectedList: this.properties.assessmentList,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId2'
                }),
                PropertyPaneDropdown('totalQuestion',{
                  label:"Set total question in an assessment",
                  options:[
                    { key: '10', text: '10'},
                    { key: '15', text: '15'},
                    { key: '20', text: '20'},
                    { key: '25', text: '25'}
                  ],
                  selectedKey: '20'
                }),
                PropertyPaneDropdown('passingScore',{
                  label:"Set passing score",
                  options:[
                    { key: '50', text: '50%'},
                    { key: '60', text: '60%'},
                    { key: '75', text: '75%'},
                    { key: '80', text: '80%'}
                  ],
                  selectedKey: '75'
                }),
                PropertyFieldListPicker('userAssessmentList', {
                  label: 'Select a user Assessment List',
                  selectedList: this.properties.userAssessmentList,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                }),
                PropertyPaneTextField('URLAssessmentHome', {
                  label: 'URL for OK Message',
                  value: this.properties.URLAssessmentHome
                })
              ]
            }
          ]
        }
      ]
    };
  }
}