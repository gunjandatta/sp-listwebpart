import { Helper, List, SPTypes } from "gd-sprest-bs";

/**
 * Configuration
 */
export const Configuration = {
    List: Helper.SPConfig({
        ListCfg: [
            {
                ListInformation: {
                    BaseTemplate: SPTypes.ListTemplateType.GenericList,
                    Title: "Demo List",
                    Description: "Datasource for the sp-listwepbart demo project."
                },
                CustomFields: [
                    {
                        name: "Description",
                        title: "Description",
                        type: Helper.SPCfgFieldType.Note
                    }
                ],
                ViewInformation: [
                    {
                        ViewName: "All Items",
                        ViewFields: [
                            "LinkTitle", "Description"
                        ],
                        ViewQuery: '<OrderBy><FieldRef Name="Title" /></OrderBy>'
                    }
                ]
            }
        ]
    }),

    generateTestData: () => {
        // Log
        console.log("Generating the test data...");

        // Get the demo list
        let list = List("Demo List");

        // Create a loop
        for (let i = 1; i <= 25; i++) {
            // Create an item
            list.Items().add({
                Title: "Test " + i,
                Description: "This is the description for test " + i + "."
            }).execute(true); // Set true to wait for the previous request to be completed
            // Note - If you are developing Online, use the "batch" method instead of "execute"
        }

        // Wait for the requests to complete
        list.done(() => {
            // Log
            console.log("Test data completed successfully.");
        });
    },

    WebPart: Helper.SPConfig({
        WebPartCfg: [
            {
                FileName: "demo-list.webpart",
                Group: "Demo",
                XML: Helper.WebPart.generateScriptEditorXML({
                    chromeType: "None",
                    description: "Displays the demo list data.",
                    title: "Demo List",
                    content: [
                        '&lt;div id="wp-demolist"&gt;&lt;/div&gt;',
                        '&lt;div id="wp-demolistcfg" style="display: none;"&gt;&lt;/div&gt;',
                        '&lt;script type="text/javascript" src="/sites/dev/siteassets/sp-examples/demoListWP.js"&gt;&lt;/script&gt;',
                        '&lt;script type="text/javascript"&gt;SP.SOD.executeOrDelayUntilScriptLoaded(function() { new WPList(); }, \'demoListWP\');&lt;/script&gt;'
                    ].join('\n')
                })
            }
        ]
    })
};