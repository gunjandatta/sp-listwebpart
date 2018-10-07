import { Helper, List, SPTypes } from "gd-sprest";

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

    WebPartCfg: Helper.SPConfig({
        WebPartCfg: [
            {
                FileName: "demo-list.webpart",
                Group: "Demo",
                XML:
                    `<?xml version="1.0" encoding="utf-8"?>
<webParts>
    <webPart xmlns="http://schemas.microsoft.com/WebPart/v3">
        <metaData>
            <type name="Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" />
            <importErrorMessage>$Resources:core,ImportantErrorMessage;</importErrorMessage>
        </metaData>
        <data>
            <properties>
                <property name="Title" type="string">Demo List</property>
                <property name="Description" type="string">Displays the demo list data.</property>
                <property name="ChromeType" type="chrometype">None</property>
                <property name="Content" type="string">
                    &lt;div id="wp-demolist"&gt;&lt;/div&gt;
                    &lt;div id="wp-demolistcfg" style="display: none;"&gt;&lt;/div&gt;
                    &lt;script type="text/javascript"&gt;SP.SOD.executeOrDelayUntilScriptLoaded(function() { new WPList(); }, 'demoListWP');&lt;/script&gt;
                </property>
            </properties>
        </data>
    </webPart>
</webParts>`
            }
        ]
    })
};