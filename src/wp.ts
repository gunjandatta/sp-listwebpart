import { SPTypes, Types } from "gd-sprest";
import { Components, WebParts } from "gd-sprest-bs";

/**
 * WebPart
 */
export const WebPart = () => {
    // Method to render the table
    let renderTable = (el: HTMLElement, listName: string, items: Array<Types.SP.IListItemQueryResult>) => {
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