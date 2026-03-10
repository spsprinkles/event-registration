import { LoadingDialog } from "dattatable";
import { App } from "./app";
import { Configuration } from "./cfg";
import { DataSource } from "./ds";
import { InstallationModal } from "./install";
import { Security } from "./security";
import Strings, { setContext } from "./strings";

// Styling
import "./styles.scss";

// Create the global variable for this solution
const GlobalVariable = {
    Configuration,
    render: (el: HTMLElement, context?) => {
        // See if the page context exists
        if (context) {
            // Set the context
            setContext(context);
        }

        // Initialize the solution
        DataSource.init().then(
            // Success
            () => {
                // Create the application
                new App(el);
            },
            // Error
            () => {
                // Update the loading dialog
                LoadingDialog.setHeader("Error Loading Application");
                LoadingDialog.setBody("Doing some more checks...");

                // See if the user has the correct permissions
                Security.hasPermissions().then(hasPermissions => {
                    // See if the user has permissions
                    if (hasPermissions) {
                        // Show the installation modal
                        InstallationModal.show();
                    }

                    // Hide the dialog
                    LoadingDialog.hide();
                }, () => {
                    // Hide the dialog
                    LoadingDialog.hide();
                });
            }
        );
    }
};

// Update the DOM
window[Strings.GlobalVariable] = GlobalVariable;

// Get the element and render the app if it is found
let elApp = document.querySelector("#" + Strings.AppElementId) as HTMLElement;
if (elApp) {
    // Render the application
    GlobalVariable.render(elApp);
}