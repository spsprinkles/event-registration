import { ItemForm, LoadingDialog, Modal } from "dattatable";
import { Components } from "gd-sprest-bs";
import * as moment from "moment";
import { IEventItem } from "./ds";

/**
 * 
 */
export class EventForms {
    // Cancels an event
    static cancel(eventItem: IEventItem, onRefresh: () => void) {
        // Set the modal header
        Modal.setHeader("Cancel Event");

        // Create the form
        let form = Components.Form({
            controls: [
                {
                    type: Components.FormControlTypes.Readonly,
                    value: "Are you sure you want to cancel the event?"
                },
                {
                    name: "SendEmail",
                    label: "Send Email?",
                    type: Components.FormControlTypes.Switch,
                    value: true
                }
            ]
        });

        // Set the modal body
        Modal.setBody(form.el);

        // Set the modal footer
        Modal.setFooter(Components.ButtonGroup({
            buttons: [
                {
                    text: "Confirm",
                    type: Components.ButtonTypes.Danger,
                    onClick: () => {
                        // Ensure the form is valid
                        if (form.isValid()) {
                            let sendEmail = form.getValues()["SendEmail"];

                            // Close the modal
                            Modal.hide();

                            // Show a loading dialog
                            LoadingDialog.setHeader("Cancelling Event");
                            LoadingDialog.setBody("This dialog will close after the item is updated.");
                            LoadingDialog.show();

                            // Update the item
                            eventItem.update({ IsCancelled: true, SendEmail: sendEmail ? "Cancel" : null }).execute(() => {
                                // Refresh the dashboard
                                onRefresh();

                                // Close the loading dialog
                                LoadingDialog.hide();
                            });
                        }
                    }
                },
                {
                    text: "Cancel",
                    type: Components.ButtonTypes.Secondary,
                    onClick: () => {
                        Modal.hide();
                    }
                }
            ]
        }).el);

        // Display the modal
        Modal.show();
    }

    // Creates an event
    static create(onRefresh: () => void) {
        // Create an item
        ItemForm.create({
            onUpdate: () => { onRefresh(); },
            onCreateEditForm: props => { return this.updateProps(props); },
            onFormButtonsRendering: buttons => { return this.updateFooter(buttons); }
        });

        // Update the modal properties
        Modal.setScrollable(true);
    }

    // Edits the event
    static edit(eventItem: IEventItem, onRefresh: () => void) {
        // Display the edit form
        ItemForm.edit({
            itemId: eventItem.Id,
            onUpdate: () => { onRefresh(); },
            onCreateEditForm: props => { return this.updateProps(props); },
            onFormButtonsRendering: buttons => { return this.updateFooter(buttons); }
        });

        // Update the modal
        Modal.setScrollable(true);
    }

    // Deletes an event
    static delete(eventItem: IEventItem, onRefresh: () => void) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Delete Event");

        // Set the body
        Modal.setBody("Are you sure you wanted to delete the selected event?");

        // Set the type
        Modal.setType(Components.ModalTypes.Medium);

        // Set the footer
        Modal.setFooter(Components.ButtonGroup({
            buttons: [
                {
                    text: "Yes",
                    type: Components.ButtonTypes.Primary,
                    onClick: () => {
                        // Show the loading dialog
                        LoadingDialog.setHeader("Delete Event");
                        LoadingDialog.setBody("Deleting the event");
                        LoadingDialog.show();

                        // Delete the item
                        eventItem.delete().execute(
                            () => {
                                // Refresh the dashboard
                                onRefresh();

                                // Hide the dialog/modal
                                Modal.hide();
                                LoadingDialog.hide();
                            },
                            () => {
                                LoadingDialog.hide();
                            }
                        );
                    },
                },
                {
                    text: "No",
                    type: Components.ButtonTypes.Secondary,
                    onClick: () => {
                        // Hide the modal
                        Modal.hide();
                    },
                },
            ],
        }).el);

        // Show the modal
        Modal.show();
    }

    // Uncancels an event
    static uncancel(eventItem: IEventItem, onRefresh: () => void) {
        // Set the modal header
        Modal.setHeader("Uncancel Event");

        // Create the form
        let form = Components.Form({
            controls: [
                {
                    type: Components.FormControlTypes.Readonly,
                    value: "Are you sure you want to uncancel the event?"
                },
                {
                    name: "SendEmail",
                    label: "Send Email?",
                    type: Components.FormControlTypes.Switch,
                    value: true
                }
            ]
        });

        // Set the modal body
        Modal.setBody(form.el);

        // Set the modal footer
        Modal.setFooter(Components.ButtonGroup({
            buttons: [
                {
                    text: "Confirm",
                    type: Components.ButtonTypes.Primary,
                    onClick: () => {
                        // Ensure the form is valid
                        if (form.isValid()) {
                            let sendEmail = form.getValues()["SendEmail"];

                            // Close the modal
                            Modal.hide();

                            // Show a loading dialog
                            LoadingDialog.setHeader("Uncancelling Event");
                            LoadingDialog.setBody("This dialog will close after the item is updated.");
                            LoadingDialog.show();

                            // Update the item
                            eventItem.update({ IsCancelled: false, SendEmail: sendEmail ? "Uncancel" : null }).execute(() => {
                                // Refresh the dashboard
                                onRefresh();

                                // Close the loading dialog
                                LoadingDialog.hide();
                            });
                        }
                    }
                },
                {
                    text: "Cancel",
                    type: Components.ButtonTypes.Secondary,
                    onClick: () => {
                        Modal.hide();
                    }
                }
            ]
        }).el);

        // Display the modal
        Modal.show();
    }

    // Updates the footer for the new/edit form
    private static updateFooter(buttons: Components.IButtonProps[]) {
        // Update the default button
        buttons[0].type = Components.ButtonTypes.Primary;

        // Add the cancel button
        buttons.push({
            text: "Cancel",
            type: Components.ButtonTypes.Secondary,
            onClick: () => {
                ItemForm.close();
            }
        });

        // Return the buttons
        return buttons;
    }

    // Updates the new/edit form properties
    private static updateProps(props: Components.IListFormEditProps) {
        // Set the rendering event
        props.onControlRendering = (ctrl, field) => {
            // See if this is the start date
            if (field.InternalName == "StartDate") {
                // Add validation
                ctrl.onValidate = (ctrl, results) => {
                    // See if the start date is after the end date
                    let startDate = results.value;
                    let endDate = ItemForm.EditForm.getControl("EndDate").getValue();
                    if (moment(startDate).isAfter(moment(endDate))) {
                        // Update the validation
                        results.isValid = false;
                        results.invalidMessage = "Start Date cannot be after the End Date";
                    }

                    // Return the results
                    return results;
                }
            }
            // Else, see if this is the capacity field
            else if (field.InternalName == "Capacity") {
                // Add validation
                ctrl.onValidate = (ctrl, results) => {
                    // Ensure an item exists
                    let item = ItemForm.FormInfo.item as IEventItem;
                    if (item) {
                        let registeredUsers = item.RegisteredUsers ? item.RegisteredUsers.results.length : 0;

                        // See if the value is less than the # of registered users
                        if (results.value < registeredUsers) {
                            // Set the results
                            results.isValid = false;
                            results.invalidMessage = "The value is less than the current number of registered users.";
                        }
                    }

                    // Return the results
                    return results;
                }
            }
        }

        // Return the properties
        return props;
    }
}