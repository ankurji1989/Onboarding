import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'TempTrainingWebPartStrings';
import TempTraining from './components/TempTraining';
import { ITempTrainingProps } from './components/ITempTrainingProps';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

export interface ITempTrainingWebPartProps {
  tempTrainingUserList: string;
}

export default class TempTrainingWebPart extends BaseClientSideWebPart <ITempTrainingWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ITempTrainingProps> = React.createElement(
      TempTraining,
      {
        context: this.context,
        displayMode: this.displayMode,
        configured: (this.properties.tempTrainingUserList !== undefined) ? true : false,
        tempTrainingUserList: this.properties.tempTrainingUserList
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
                PropertyFieldListPicker('tempTrainingUserList', {
                  label: 'Select training user list',
                  selectedList: this.properties.tempTrainingUserList,
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
