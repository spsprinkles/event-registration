import { ContextInfo } from "gd-sprest-bs";

/**
 * Global Constants
 */

// Global Path
let SourceUrl: string = ContextInfo.webServerRelativeUrl + "/SiteAssets/Event-Registration/";

// Updates the strings for SPFx
export const setContext = (context) => {
    // Set the page context
    ContextInfo.setPageContext(context);

    // Update the global path
    SourceUrl = ContextInfo.webServerRelativeUrl + "/SiteAssets/Event-Registration/";

    // Update the values
    Strings.EventRegConfig = SourceUrl + "eventreg-config.json";
    Strings.SolutionUrl = SourceUrl + "index.html";
    Strings.SourceUrl = SourceUrl;
}

// Strings
const Strings = {
    AppElementId: "event-registration",
    GlobalVariable: "EventRegistration",
    Lists: {
        Events: "Events"
    },
    ProjectName: "Event Registration",
    ProjectDescription: "Allows users to sign up for events.",
    EventRegConfig: SourceUrl + "eventreg-config.json",
    SolutionUrl: SourceUrl + "index.html",
    SourceUrl: SourceUrl,
    Version: "0.2",
};
export default Strings;