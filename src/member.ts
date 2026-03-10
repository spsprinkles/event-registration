import { LoadingDialog, Modal } from "dattatable";
import { Components } from "gd-sprest-bs";
import { gearWideConnected } from "gd-sprest-bs/build/icons/svgs/gearWideConnected";
import { IEventItem } from "./ds";
import { Registration } from "./registration";

/**
 * Member
 */
export class Member {
    // Renders the event menu
    renderEventMenu(el: HTMLElement, eventItem: IEventItem) {
        // Add the member dropdown
        let ddl = Components.Dropdown({
            el: el,
            className: "eventRegAdmin",
            items: [
                {
                    text: " Send Email to POCs",
                    isDisabled: eventItem.POC == null,
                    onClick: (button) => {
                        this.sendEmail(eventItem);
                    },
                }
            ],
        });

        let elButton = ddl.el.querySelector("button");
        if (elButton) {
            // Update the class
            elButton.classList.add("btn-icon");
            elButton.classList.add("w-100");

            // Append the icon
            elButton.appendChild(gearWideConnected(16));
        }
    }

    // Sends an email to the event members
    private sendEmail(eventItem: IEventItem) {
        // Set the modal header
        Modal.setHeader("Send Email")

        // Create the form
        let form = Components.Form({
            controls: [
                {
                    name: "EmailSubject",
                    label: "Email Subject",
                    required: true,
                    errorMessage: "A subject is required to send an email.",
                    type: Components.FormControlTypes.TextField,
                    value: eventItem.Title
                },
                {
                    name: "EmailBody",
                    label: "Email Body",
                    required: true,
                    errorMessage: "Content is required to send an email.",
                    rows: 10,
                    type: Components.FormControlTypes.TextArea,
                    value: ""
                } as Components.IFormControlPropsTextField
            ]
        });

        // Set the modal body
        Modal.setBody(form.el);

        // Set the modal footer
        Modal.setFooter(Components.ButtonGroup({
            buttons: [
                {
                    text: "Send",
                    type: Components.ButtonTypes.Primary,
                    onClick: () => {
                        // Ensure the form is valid
                        if (form.isValid()) {
                            let values = form.getValues();

                            // Close the modal
                            Modal.hide();

                            // Show a loading dialog
                            LoadingDialog.setHeader("Sending Email");
                            LoadingDialog.setBody("This dialog will close after the email is sent.");
                            LoadingDialog.show();

                            // Parse the pocs
                            let pocs = [];
                            for (let i = 0; i < eventItem.POC.results.length; i++) {
                                // Append the user email
                                pocs.push(eventItem.POC.results[i].EMail);
                            }

                            // Send the email
                            eventItem.update({
                                SendEmail: "Custom",
                                SendEmailInfo: JSON.stringify({
                                    To: pocs,
                                    Body: values["EmailBody"].replace(/\n/g, "<br />"),
                                    Subject: values["EmailSubject"]
                                })
                            }).execute(() => {
                                // Close the loading dialog
                                LoadingDialog.hide();
                            });
                        }
                    }
                },
                {
                    text: "Cancel",
                    type: Components.ButtonTypes.Secondary,
                    onClick: () => { Modal.hide(); }
                }
            ]
        }).el);

        // Display the modal
        Modal.show();
    }
}