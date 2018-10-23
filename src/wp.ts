import { SPTypes, Types } from "gd-sprest";
import { Components, WebParts } from "gd-sprest-bs";

/**
 * WebPart
 */
export const WebPart = () => {
    // Method to render the table
    let renderTable = (el: HTMLElement, listName: string, items: Array<Types.SP.IListItemQueryResult>) => {
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
                '<td>' + item["Title"] + '</td>',
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

    // Create the list webpart
    WebParts.WPList({
        elementId: "wp-demolist",
        cfgElementId: "wp-demolistcfg",
        onRenderItems: (wpInfo, items) => {
            // Render the table
            renderTable(wpInfo.el, wpInfo.cfg.ListName, items as any);
        }
    });
}