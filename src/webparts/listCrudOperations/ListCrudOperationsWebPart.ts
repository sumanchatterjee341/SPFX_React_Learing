import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "ListCrudOperationsWebPartStrings";
import ListCrudOperations from "./components/ListCrudOperations";
import { IListCrudOperationsProps } from "./components/IListCrudOperationsProps";

import { Validation } from "./Services/Validations";

export interface IListCrudOperationsWebPartProps {
  ListTitle: string;
  ListUrl: string;
  Validation: boolean;
  Lists: IPropertyPaneDropdownOption[];
}

export default class ListCrudOperationsWebPart extends BaseClientSideWebPart<IListCrudOperationsWebPartProps> {
  public _Validations: Validation;
  constructor() {
    super();
    this._Validations = new Validation();
  }

  public render(): void {
    const element: React.ReactElement<IListCrudOperationsProps> =
      React.createElement(ListCrudOperations, {
        ListTitle: this.properties.ListTitle,
        Context: this.context,
      });

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
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("ListTitle", {
                  label: strings.ListTitleFieldLabel,
                  onGetErrorMessage:
                    this._Validations.TextFieldValidation.bind(this),
                }),
                PropertyPaneTextField("ListUrl", {
                  label: strings.ListUrlFieldLabel,
                }),
                PropertyPaneCheckbox("Validation", {
                  text: strings.ValidationFieldLabel,
                  checked: false,
                }),
                PropertyPaneDropdown("Lists", {
                  label: strings.ListsDropDownLabel,
                  options: [
                    {
                      key: "----Select Your List----",
                      text: "----Select Your List----",
                    },
                    { key: "List 1", text: "List 1" },
                    { key: "List 2", text: "List 2" },
                    { key: "List 3", text: "List 3" },
                  ],
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
