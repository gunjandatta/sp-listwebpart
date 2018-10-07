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
    // Create the table
    let table = document.createElement("table");
    table.className = "table table-striped";
    el.appendChild(table);

    // Create the header
    let header = document.createElement("thead");
    header.innerHTML = [
      "<tr>",
      "<th></th>",
      "<th>Title</th>",
      "<th>Description</th>",
      "</tr>"
    ].join('\n');
    table.appendChild(header);

    // Create the body
    let body = document.createElement("tbody");
    table.appendChild(body);

    // Parse the items
    for (let i = 0; i < items.length; i++) {
      let item = items[i];

      // Create the row
      let row = document.createElement("tr");
      row.innerHTML = [
        '<td></td>',
        '<td>' + item.Title + '</td>',
        '<td>' + item["Description"] + '</td>'
      ].join('\n');
      body.appendChild(row);

      // Render an edit button
      Components.Button({
        data: item,
        el: row.children[0],
        isSmall: true,
        text: "Edit",
        type: Components.ButtonTypes.Secondary,
        onClick: (btn) => {
          let item = btn.data as Types.SP.IListItemResult;

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
