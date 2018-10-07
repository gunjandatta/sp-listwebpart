import { Configuration } from "./cfg";
import { WebPart } from "./wp";

/**
 * List WebPart
 */
class ListWebPart {
    // Configuration
    static Configuration = Configuration;

    // Constructor
    constructor() {
        // Create an instance of the webpart
        WebPart();
    }
}

/**
 * Global Variable
 */
window["WPList"] = ListWebPart;

// Notify SharePoint the library is loaded
window["SP"] ? window["SP"].SOD.notifyScriptLoadedAndExecuteWaitingJobs("demoListWP") : null;