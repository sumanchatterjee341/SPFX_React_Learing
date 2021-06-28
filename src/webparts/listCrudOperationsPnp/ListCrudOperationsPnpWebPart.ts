import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ListCrudOperationsPnpWebPartStrings';
import ListCrudOperationsPnp from './components/ListCrudOperationsPnp';
import { IListCrudOperationsPnpProps } from './components/IListCrudOperationsPnpProps';

export interface IListCrudOperationsPnpWebPartProps {
  ListTitle: string;
}

export default class ListCrudOperationsPnpWebPart extends BaseClientSideWebPart <IListCrudOperationsPnpWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IListCrudOperationsPnpProps> = React.createElement(
      ListCrudOperationsPnp,
      {
        ListTitle:this.properties.ListTitle,
        Context:this.context,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
