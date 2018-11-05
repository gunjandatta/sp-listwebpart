import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ListWebPartWebPart.module.scss';
import * as strings from 'ListWebPartWebPartStrings';

import { ContextInfo, List, SPTypes, Types } from "gd-sprest";
import { Components } from "gd-sprest-bs";

export interface IListWebPartWebPartProps {
  description: string;
}

export default class ListWebPartWebPart extends BaseClientSideWebPart<IListWebPartWebPartProps> {

  public render(): void {
    // Set the page context
    ContextInfo.setPageContext(this.context.pageContext);

    // Render a loading dialog
    let progress = Components.Progress({
      el: this.domElement,
      isAnimated: true,
      isStriped: true,
      size: 100,
      label: "Loading the List Data"
    });

    // Get the demo list
    List("Demo List")
      // Get the items
      .Items()
      // Set th equery
      .query({
        OrderBy: ["Title"]
      })
      // Execute the request
      .execute(items => {
        // Remove the progress bar
        this.domElement.removeChild(progress.el);

        // Render the table
        this.renderTable(this.domElement, "Demo List", items.results);
      });
  }

  // Method to render the table
  protected renderTable(el: HTMLElement, listName: string, items: Array<Types.SP.IListItemQueryResult>) {
    // Render the table
    Components.Table({
      el,
      rows: items,
      columns: [
        {
          name: "Actions",
          isHidden: true,
          onRenderCell: (el, col, item: Types.SP.IListItemQueryResult) => {
            // Render an edit button
            Components.Button({
              el,
              isSmall: true,
              text: "Edit",
              type: Components.ButtonTypes.Secondary,
              onClick: (btn) => {
                // Get the form element
                let elForm = document.querySelector("#dlg-listform");
                if (elForm == null) {
                  // Create the element
                  elForm = document.createElement("div");
                  elForm.id = "dlg-listform";
                  document.body.appendChild(elForm);
                }

                // Render a list form
                Components.ListFormDialog({
                  el: elForm,
                  controlMode: SPTypes.ControlMode.Edit,
                  itemId: item.Id,
                  listName: listName,
                  visible: true,
                  modalProps: {
                    onClose: () => {
                      // Remove the form element
                      document.body.removeChild(elForm);
                    }
                  }
                });
              }
            });
          }
        },
        {
          name: "Title"
        },
        {
          name: "Description"
        }
      ]
    });
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
