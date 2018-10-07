import { SPTypes, Types } from "gd-sprest";
import { Components, WebParts } from "gd-sprest-bs";

/**
 * WebPart
 */
export const WebPart = () => {
    // Create the list webpart
    WebParts.WPList({
        elementId: "wp-demolist",
        cfgElementId: "wp-demolistcfg",
        onRenderItems: (wpInfo, items) => {
            // Create the table
            let table = document.createElement("table");
            table.className = "table table-striped";
            wpInfo.el.appendChild(table);

            // Create the header
            let header = document.createElement("thead");
            header.innerHTML = [
                "<tr>",
                "<th></th>",
                "<th>Title</th>",
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
                    '<td>' + item.Title + '</td>'
                ].join('\n');
                body.appendChild(row);

                // Render an edit button
                Components.Button({
                    data: item,
                    el: row.children[0],
                    isSmall: true,
                    text: "View",
                    type: Components.ButtonTypes.Secondary,
                    onClick: (btn) => {
                        let item = btn.data as Types.SP.IListItemResult;

                        // Render a list form
                        let form = Components.ListFormDialog({
                            el: document.body,
                            controlMode: SPTypes.ControlMode.Edit,
                            itemId: item.Id,
                            listName: wpInfo.cfg.ListName,
                            webUrl: wpInfo.cfg.WebUrl,
                            modalProps: {
                                onClose: () => {
                                    // Remove the form element
                                    document.body.removeChild(form.el);
                                }
                            }
                        });
                    }
                });
            }
        }
    });
}