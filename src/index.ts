import { ContextInfo, List } from "gd-sprest-bs";
import { Configuration } from "./cfg";
import { WebPart } from "./wp";

/**
 * List WebPart
 */
class ListWebPart {
    // Configuration
    static Configuration = Configuration;

    // Loads the data for SPFx solutions
    static load = (context?) => {
        // Set the page context if it exists
        context ? ContextInfo.setPageContext(context) : null;

        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the list items
            List(Configuration.List._configuration.ListCfg[0].ListInformation.Title).Items().execute(
                // Success
                items => {
                    // Resolve the request
                    resolve(items.results);
                },

                // Error
                reject
            );
        });
    }

    // Render table method for SPFx solutions
    static renderTable = (el, items) => {
        // Render the table
        WebPart.renderTable(el, Configuration.List._configuration.ListCfg[0].ListInformation.Title, items);
    };

    // Constructor
    constructor(el?) {
        // Create an instance of the webpart
        new WebPart(el);
    }
}

/**
 * Global Variable
 */
window["WPList"] = ListWebPart;

// Notify SharePoint the library is loaded
window["SP"] ? window["SP"].SOD.notifyScriptLoadedAndExecuteWaitingJobs("demoListWP") : null;