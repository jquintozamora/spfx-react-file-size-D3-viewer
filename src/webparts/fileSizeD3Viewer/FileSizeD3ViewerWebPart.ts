import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'fileSizeD3ViewerStrings';
import FileSizeD3Viewer from './components/FileSizeD3Viewer';
import { IFileSizeD3ViewerProps } from './components/IFileSizeD3ViewerProps';
import { IFileSizeD3ViewerWebPartProps } from './IFileSizeD3ViewerWebPartProps';

export default class FileSizeD3ViewerWebPart extends BaseClientSideWebPart<IFileSizeD3ViewerWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IFileSizeD3ViewerProps > = React.createElement(
      FileSizeD3Viewer,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
