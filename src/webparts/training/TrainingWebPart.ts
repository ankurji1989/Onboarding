import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'TrainingWebPartStrings';
import Training from './components/Training';
import { ITrainingProps } from './components/ITrainingProps';
import { sp } from '@pnp/sp';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
 
export interface ITrainingWebPartProps {
  description: string;
  lists: string;
  userAssessment: string;
  userTraining: string;
  moduleSubmittionMsg: string;
  moduleCompletionMsg: string;
  URLForYes: string;
  URLForNo: string;
}
 
export default class TrainingWebPart extends BaseClientSideWebPart <ITrainingWebPartProps> {
 
  protected onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context
    });
    if (this.properties.moduleSubmittionMsg === undefined) {
      this.properties.moduleSubmittionMsg = '<p>Click "OK", to submit the module (Link to documents will be disabled post submission)</p><p>Click "Cancel", to return to module training.</p>';
    }
    if (this.properties.moduleCompletionMsg === undefined) {
      this.properties.moduleCompletionMsg = 'Would you like to take the assessment now?';
    }
    if (this.properties.URLForYes === undefined) {
      this.properties.URLForYes = window.location.href;
    }
    if (this.properties.URLForNo === undefined) {
      this.properties.URLForNo = window.location.href;
    }
    return Promise.resolve();
  }
 
  public render(): void {
    const element: React.ReactElement<ITrainingProps> = React.createElement(
      Training,
      {
        context: this.context,
        selectedList: this.properties.lists,
        displayMode: this.displayMode,
        configured: (this.properties.lists && this.properties.userAssessment && this.properties.userTraining) ? true : false,
        userAssessment: this.properties.userAssessment,
        userTraining: this.properties.userTraining,
        moduleSubmittionMsg: this.properties.moduleSubmittionMsg,
        moduleCompletionMsg: this.properties.moduleCompletionMsg,
        URLForYes: this.properties.URLForYes,
        URLForNo: this.properties.URLForNo
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
                  label: 'Select a training master list',
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
                PropertyFieldListPicker('userAssessment', {
                  label: 'Select a UserAssessment',
                  selectedList: this.properties.userAssessment,
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
                PropertyFieldListPicker('userTraining', {
                  label: 'Select a UserTraining',
                  selectedList: this.properties.userTraining,
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
                PropertyPaneTextField('moduleSubmittionMsg', {
                  label: 'Module submittion message',
                  value: this.properties.moduleSubmittionMsg
                }),
                PropertyPaneTextField('moduleCompletionMsg', {
                  label: 'Module completion message.',
                  value: this.properties.moduleCompletionMsg
                }),
                PropertyPaneTextField('URLForYes', {
                  label: 'URL for Yes',
                  value: this.properties.URLForYes
                }),
                PropertyPaneTextField('URLForNo', {
                  label: 'URL for No',
                  value: this.properties.URLForNo
                })
              ]
            }
          ]
        }
      ]
    };
  }
}