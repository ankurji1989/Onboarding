import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
 
import * as strings from 'AssessmentHomeWebPartStrings';
import AssessmentHome from './components/AssessmentHome';
import { IAssessmentHomeProps } from './components/IAssessmentHomeProps';
import { sp } from '@pnp/sp';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
export interface IAssessmentHomeWebPartProps {
  description: string;
  lists: string;
  assessmentList:string;
  userAssessmentList: string;
}
 
export default class AssessmentHomeWebPart extends BaseClientSideWebPart <IAssessmentHomeWebPartProps> {
  protected onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context
    });
    return Promise.resolve();
  }
  public render(): void {
    const element: React.ReactElement<IAssessmentHomeProps> = React.createElement(
      AssessmentHome,
      {
        description: this.properties.description,
        context: this.context,
        selectedList: this.properties.lists,
        assessmentList:this.properties.assessmentList,
        userAssessmentList: this.properties.userAssessmentList
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
                PropertyFieldListPicker('lists', {
                  label: 'Select user training list',
                  selectedList: this.properties.lists,
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
                PropertyPaneTextField('description', {
                  label: "Assessment URL",
                  value:""
                }),
                PropertyFieldListPicker('userAssessmentList', {
                  label: 'Select a userAssessment List',
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
              ]
            }
          ]
        }
      ]
    };
  }
}