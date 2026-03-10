import { InstallationRequired, LoadingDialog, Modal } from "dattatable";
import { Components, Helper } from "gd-sprest-bs";
import { calendarPlus } from "gd-sprest-bs/build/icons/svgs/calendarPlus";
import { gearWideConnected } from "gd-sprest-bs/build/icons/svgs/gearWideConnected";
import * as moment from "moment";
import { Configuration } from "./cfg";
import { IEventItem } from "./ds";
import { EventForms } from "./eventForms";
import { Registration } from "./registration";
import { Security } from "./security";

export class Admin {
  // Generates the navigation items
  generateNavItems(canEditEvent: boolean, onRefresh: () => void): Components.INavbarItem[] {
    let navItems: Components.INavbarItem[] = [];

    // See if this is the admin
    if (Security.IsAdmin) {
      // Add the new event option
      navItems.push({
        className: "btn-primary",
        isDisabled: !canEditEvent,
        text: " NEW EVENT",
        isButton: true,
        onClick: () => {
          // Create an event
          EventForms.create(onRefresh);
        },
        iconType: calendarPlus,
        iconSize: 18,
      });

      // Add the manage groups option
      navItems.push({
        className: "btn-primary",
        isDisabled: !canEditEvent,
        text: "MANAGE GROUPS",
        isButton: true,
        items: [
          {
            text: "Managers",
            onClick: () => {
              // Show the manager's group
              window.open(Security.ManagersUrl, "_blank");
            },
          },
          {
            text: "Members",
            onClick: () => {
              // Show the member's group
              window.open(Security.MembersUrl, "_blank");
            },
          },
        ],
      });

      // Add an option to manage the application
      navItems.push({
        className: "btn-primary",
        text: "MANAGE APP",
        isButton: true,
        onClick: () => {
          // Display a loading dialog
          LoadingDialog.setHeader("Analyzing the Assets");
          LoadingDialog.setBody("Checking the SharePoint assets.");
          LoadingDialog.show();

          // Determine if an install is required
          InstallationRequired.requiresInstall({ cfg: Configuration }).then(() => {
            // Hide the dialog
            LoadingDialog.hide();

            // Show the installation dialog
            InstallationRequired.showDialog();
          });
        }
      });
    }

    // Return the nav items
    return navItems;
  }

  // Manages the waitlist users
  private manageWaitlist(eventItem: IEventItem, onRefresh: () => void) {
    // Set the modal header
    Modal.setHeader("Manage Waitlist")

    // Parse the waitlisted users and generate the checkboxes
    let items: Components.ICheckboxGroupItem[] = [];
    let users = eventItem.WaitListedUsers ? eventItem.WaitListedUsers.results : [];
    for (let i = 0; i < users.length; i++) {
      let user = users[i];

      // Add the item
      items.push({
        data: user,
        label: user.Title
      });
    }

    // Create the form
    let form = Components.Form({
      controls: [
        {
          name: "Users",
          label: "Users",
          items,
          required: true,
          errorMessage: "A user is required.",
          type: Components.FormControlTypes.MultiCheckbox,
          onValidate: (ctrl, results) => {
            // See if users being added exceed the capacity
            if (eventItem.RegisteredUsers.results.length + results.value.length > eventItem.Capacity) {
              // Update the flag
              results.isValid = false;
              results.invalidMessage = "The selected users will exceed the capacity of the event.";
            }

            // Return the results
            return results;
          }
        } as Components.IFormControlPropsMultiCheckbox,
        {
          name: "SendEmail",
          label: "Send Email?",
          type: Components.FormControlTypes.Switch
        }
      ]
    });

    // Set the modal body
    Modal.setBody(form.el);

    // Set the modal footer
    Modal.setFooter(Components.ButtonGroup({
      buttons: [
        {
          text: "Delete",
          type: Components.ButtonTypes.Danger,
          onClick: () => {
            let formValues = form.getValues();
            let sendEmail = formValues["SendEmail"];

            // Close the modal
            Modal.hide();

            // Show a loading dialog
            LoadingDialog.setHeader("Removing User(s)");
            LoadingDialog.setBody("This dialog will close after the user(s) are removed.");
            LoadingDialog.show();

            // Parse the waitlisted users
            let usersToRemove: Components.ICheckboxGroupItem[] = formValues["Users"];
            let waitlistedUsers = eventItem.WaitListedUsersId ? eventItem.WaitListedUsersId.results : [];
            for (let i = 0; i < usersToRemove.length; i++) {
              let userId = usersToRemove[i].data.Id;

              // Find the user
              let idx = waitlistedUsers.indexOf(userId);

              // Remove the user
              waitlistedUsers.splice(idx, 1);
            }

            // Update the item
            eventItem.update({
              "WaitListedUsersId": { results: waitlistedUsers }
            }).execute(() => {
              // Refresh the dashboard
              onRefresh();

              // See if we are sending an email
              if (sendEmail) {
                // Parse the users to delete
                let usersToAdd: Components.ICheckboxGroupItem[] = formValues["Users"];
                Helper.Executor(usersToAdd, user => {
                  // Return a promise
                  return new Promise((resolve, reject) => {
                    // Send an email
                    Registration.sendMail(eventItem, user.data.Id, false, false).then(() => {
                      // Resolve the request
                      resolve(null);
                    }, reject);
                  }).then(() => {
                    // Close the dialog
                    LoadingDialog.hide();
                  });
                });
              } else {

              }

              // Close the dialog
              LoadingDialog.hide();
            });
          }
        },
        {
          text: "Register",
          type: Components.ButtonTypes.Primary,
          onClick: () => {
            // Ensure the form is valid
            if (form.isValid()) {
              let formValues = form.getValues();
              let sendEmail = formValues["SendEmail"];

              // Close the modal
              Modal.hide();

              // Show a loading dialog
              LoadingDialog.setHeader("Registering User(s)");
              LoadingDialog.setBody("This dialog will close after the user(s) are registered.");
              LoadingDialog.show();

              // Parse the waitlisted users
              let usersToAdd: Components.ICheckboxGroupItem[] = formValues["Users"];
              let registeredUsers = eventItem.RegisteredUsersId ? eventItem.RegisteredUsersId.results : [];
              let waitlistedUsers = [];
              for (let i = 0; i < items.length; i++) {
                let userId = items[i].data.Id;

                // Parse the users to add
                let registerUser = false;
                for (let j = 0; j < usersToAdd.length; j++) {
                  let user = usersToAdd[j];

                  // See if this user is being registered
                  if (user.data.Id == userId) {
                    // Set the flag
                    registerUser = true;
                    break;
                  }
                }

                // See if we are registering the user
                if (registerUser) {
                  // Append the registered user
                  registeredUsers.push(userId);
                } else {
                  // Append the waitlisted user
                  waitlistedUsers.push(userId);
                }
              }

              // Update the item
              eventItem.update({
                "RegisteredUsersId": { results: registeredUsers },
                "WaitListedUsersId": { results: waitlistedUsers }
              }).execute(() => {
                // Refresh the dashboard
                onRefresh();

                // See if we are sending an email
                if (sendEmail) {
                  Helper.Executor(usersToAdd, user => {
                    // Return a promise
                    return new Promise((resolve, reject) => {
                      // Send an email
                      Registration.sendMail(eventItem, user.data.Id, false, true).then(() => {
                        // Resolve the request
                        resolve(null);
                      }, reject);
                    }).then(() => {
                      // Close the dialog
                      LoadingDialog.hide();
                    });
                  });
                } else {
                  // Close the dialog
                  LoadingDialog.hide();
                }
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

  // Registers a user
  private registerUser(eventItem: IEventItem, onRefresh: () => void) {
    // Set the modal header
    Modal.setHeader("Register User")

    // Create the form
    let form = Components.Form({
      controls: [
        {
          name: "User",
          label: "User",
          required: true,
          errorMessage: "A user is required.",
          type: Components.FormControlTypes.PeoplePicker,
          onValidate: (ctrl, result) => {
            // Parse the current POCs
            let users = eventItem.RegisteredUsersId ? eventItem.RegisteredUsersId.results : [];
            for (let i = 0; i < users.length; i++) {
              // See if the user is already registered
              if (users[i] == result.value[0].Id) {
                // User is already registered
                result.isValid = false;
                result.invalidMessage = "User is already registered.";
              }
            }

            // Return the result
            return result;
          }
        },
        {
          name: "SendEmail",
          label: "Send Email?",
          type: Components.FormControlTypes.Switch
        }
      ]
    });

    // Set the modal body
    Modal.setBody(form.el);

    // Set the modal footer
    Modal.setFooter(Components.ButtonGroup({
      buttons: [
        {
          text: "Register",
          type: Components.ButtonTypes.Primary,
          onClick: () => {
            // Ensure the form is valid
            if (form.isValid()) {
              let formValues = form.getValues();

              // Close the modal
              Modal.hide();

              // Show a loading dialog
              LoadingDialog.setHeader("Registering User");
              LoadingDialog.setBody("This dialog will close after the user is registered.");
              LoadingDialog.show();

              // Append the user
              let userId = formValues["User"][0].Id;
              let value = eventItem.RegisteredUsersId ? eventItem.RegisteredUsersId.results : [];
              value.push(userId);
              let values = {
                "RegisteredUsersId": { results: value }
              };

              // See if the user is waitlisted
              let waitlist = eventItem.WaitListedUsersId ? eventItem.WaitListedUsersId.results : [];
              for (let i = 0; i < waitlist.length; i++) {
                if (waitlist[i] == userId) {
                  // Remove the user
                  waitlist.splice(i, 1);

                  // Update the field value
                  values["WaitListedUsersId"] = { results: waitlist };
                  break;
                }
              }

              // Update the item
              eventItem.update(values).execute(() => {
                // Refresh the dashboard
                onRefresh();

                // See if we are sending an email
                if (formValues["SendEmail"]) {
                  // Update the loading dialog
                  LoadingDialog.setHeader("Sending Email");
                  LoadingDialog.setBody("This dialog will close after the email is sent.");

                  // Send an email
                  Registration.sendMail(eventItem, userId, true, false).then(() => {
                    // Close the dialog
                    LoadingDialog.hide();
                  });
                } else {
                  // Close the dialog
                  LoadingDialog.hide();
                }
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

  // Renders the event menu
  renderEventMenu(el: HTMLElement, eventItem: IEventItem, canEditEvent: boolean, canDeleteEvent: boolean, onRefresh: () => void) {
    // Add the admin dropdown
    let adminDropdown = Components.Dropdown({
      el: el,
      className: "eventRegAdmin",
      items: [
        {
          text: eventItem.IsCancelled ? " Uncancel" : " Cancel",
          onClick: (button) => {
            // See if the event is cancelled
            if (eventItem.IsCancelled) {
              EventForms.uncancel(eventItem, onRefresh);
            } else {
              EventForms.cancel(eventItem, onRefresh);
            }
          },
        },
        {
          isDisabled: eventItem.IsCancelled || !canEditEvent,
          text: " Edit",
          onClick: (button) => {
            EventForms.edit(eventItem, onRefresh);
          },
        },
        {
          isDisabled: !canDeleteEvent,
          text: " Delete",
          onClick: (button) => {
            EventForms.delete(eventItem, onRefresh);
          },
        },
        {
          text: " Manage Waitlist",
          isDisabled: eventItem.IsCancelled || eventItem.WaitListedUsersId == null,
          onClick: (button) => {
            this.manageWaitlist(eventItem, onRefresh);
          },
        },
        {
          text: " Register User",
          isDisabled: eventItem.IsCancelled || Registration.isFull(eventItem),
          onClick: (button) => {
            this.registerUser(eventItem, onRefresh);
          },
        },
        {
          text: " Send Email",
          onClick: (button) => {
            this.sendEmail(eventItem);
          },
        },
        {
          text: " Unregister User",
          isDisabled: eventItem.IsCancelled || Registration.isEmpty(eventItem),
          onClick: (button) => {
            this.unregisterUser(eventItem, onRefresh);
          },
        },
        {
          text: " View Roster",
          onClick: (button) => {
            this.viewRoster(eventItem);
          },
        }
      ],
    });

    let elButton = adminDropdown.el.querySelector("button");
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
          name: "EmailRecipients",
          label: "Email Recipients",
          type: Components.FormControlTypes.MultiSwitch,
          required: true,
          errorMessage: "A selection is required to send an email.",
          items: [
            {
              label: "Email POCs",
              isSelected: true
            },
            {
              label: "Email Members",
              isSelected: true
            }
          ]
        } as Components.IFormControlPropsMultiSwitch,
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

              // Determine who we are sending emails to
              let emailPOCs = false;
              let emailMembers = false;
              let emailRecipients = values["EmailRecipients"] as Components.ICheckboxGroupItem[];
              for (let i = 0; i < emailRecipients.length; i++) {
                let emailRecipient = emailRecipients[i];
                if (emailRecipient.label.indexOf("POC") >= 0) { emailPOCs = true; }
                else if (emailRecipient.label.indexOf("Member") >= 0) { emailMembers = true; }
              }

              // See if we are emailing POCs
              let pocs = [];
              if (emailPOCs) {
                // Parse the pocs
                for (let i = 0; i < eventItem.POC.results.length; i++) {
                  // Append the user email
                  pocs.push(eventItem.POC.results[i].EMail);
                }
              }

              // See if we are emailing Members
              let members = [];
              if (emailMembers) {
                // Parse the registered users
                for (let i = 0; i < eventItem.RegisteredUsers.results.length; i++) {
                  // Append the user email
                  members.push(eventItem.RegisteredUsers.results[i].EMail);
                }
              }

              // Initialize the email TO and CC values
              let To: string[], CC: string[] = null;

              // See if we are emailing everyone
              if (emailMembers && emailPOCs) {
                To = members;
                CC = pocs;
              }
              // Else, see if we are only sending mail to the members
              else if (emailMembers && !emailPOCs) {
                To = members;
              }
              // Else, see if we are only sending mail to the pocs
              else if (!emailMembers && emailPOCs) {
                To = pocs;
              }

              // Ensure we are sending an email
              if ((To && To.length > 0) || (CC && CC.length > 0)) {
                // Send the email
                eventItem.update({
                  SendEmail: "Custom",
                  SendEmailInfo: JSON.stringify({
                    To,
                    CC,
                    Body: values["EmailBody"].replace(/\n/g, "<br />"),
                    Subject: values["EmailSubject"]
                  })
                }).execute(() => {
                  // Close the loading dialog
                  LoadingDialog.hide();
                });
              }
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

  // Unregisters a user
  private unregisterUser(eventItem: IEventItem, onRefresh: () => void) {
    // Set the modal header
    Modal.setHeader("Unregister User")

    // Parse the current users and generate the checkboxes
    let items: Components.ICheckboxGroupItem[] = [];
    let users = eventItem.RegisteredUsers ? eventItem.RegisteredUsers.results : [];
    for (let i = 0; i < users.length; i++) {
      let user = users[i];

      // Add the item
      items.push({
        data: user,
        label: user.Title
      });
    }

    // Create the form
    let form = Components.Form({
      controls: [
        {
          name: "Users",
          label: "Users",
          items,
          required: true,
          errorMessage: "A user is required.",
          type: Components.FormControlTypes.MultiCheckbox
        } as Components.IFormControlPropsMultiCheckbox
      ]
    });

    // Set the modal body
    Modal.setBody(form.el);

    // Set the modal footer
    Modal.setFooter(Components.ButtonGroup({
      buttons: [
        {
          text: "Unregister",
          type: Components.ButtonTypes.Danger,
          onClick: () => {
            // Ensure the form is valid
            if (form.isValid()) {
              // Close the modal
              Modal.hide();

              // Show a loading dialog
              LoadingDialog.setHeader("Unregistering User(s)");
              LoadingDialog.setBody("This dialog will close after the user(s) are unregistered.");
              LoadingDialog.show();

              // Parse the users to remove
              let newUsers = [];
              let usersToRemove = form.getValues()["Users"] as Components.ICheckboxGroupItem[];
              let registeredUsers = eventItem.RegisteredUsersId ? eventItem.RegisteredUsersId.results : [];
              for (let i = 0; i < registeredUsers.length; i++) {
                let userId = registeredUsers[i];

                // Parse the users to remove
                let removeFl = false;
                for (let j = 0; j < usersToRemove.length; j++) {
                  // See if this user is being removed
                  if (userId == usersToRemove[j].data.Id) {
                    removeFl = true;
                    break;
                  }
                }

                // See if this user is not being removed
                if (!removeFl) {
                  // Remove the user
                  newUsers.push(userId);
                }
              }

              // Update the item
              eventItem.update({
                "RegisteredUsersId": { results: newUsers }
              }).execute(() => {
                // Refresh the dashboard
                onRefresh();

                // Close the dialog
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

  // Displays the roster in a modal
  private viewRoster(eventItem: IEventItem) {
    //insert the print style css
    let modalCSS = [
      "@media print {",
      "\tbody * { visibility: hidden; }",
      "\t#core-modal, #core-modal * { visibility: visible; }",
      "\t#core-modal .modal-content { position: absolute; left: 0; top: 0; }",
      "\t#core-modal .modal-footer, #core-modal .modal-footer *, #core-modal .modal-header, #core-modal .modal-header * { display: none; }",
      "}"
    ].join('\n');

    // Create the style element, if it does not exist
    if (document.getElementById("eventRegPrintStyle") == null) {
      let elStyle = document.createElement("style");
      elStyle.id = "eventRegPrintStyle";
      elStyle.innerHTML = modalCSS;
      document.head.appendChild(elStyle);
    }

    // Get the POCs
    let pocString = "";
    let pocs = ((eventItem.POC ? eventItem.POC.results : null) || []).sort((a, b) => {
      if (a.Title < b.Title) { return -1; }
      if (a.Title > b.Title) { return 1; }
      return 0;
    });
    for (let i = 0; i < pocs.length; i++) {
      if (i > 0) pocString += "<br/>";
      pocString += pocs[i].Title;
    }

    // Get the list of registered usernames
    let usersRegistered = ((eventItem.RegisteredUsers ? eventItem.RegisteredUsers.results : null) || []).sort((a, b) => {
      if (a.Title < b.Title) { return -1; }
      if (a.Title > b.Title) { return 1; }
      return 0;
    });
    let usersTable = `<tr><td colspan='4'>No registrations were found for the event</td></tr>`;
    for (let x = 0; x < usersRegistered.length; x++) {
      if (x == 0) usersTable = "";
      usersTable += `<tr><td>${x + 1}</td><td>${usersRegistered[x].Title}</td><td></td><td></td></tr>`;
    }

    // Create the Jumbotron text
    let jumbotronContent = `
      <div class='table-responsive' id='print-view'>
        <table class='table table-striped'>
          <thead class='thead-dark'>
            <tr><th colspan='2' style='text-align:center'><h5>${eventItem.Title}</h5></th></tr>
          </thead>
          <tbody>
            <tr><td><strong>Start Date</strong></td><td> ${moment(eventItem.StartDate).format("MM-DD-YYYY HH:mm")}</td></tr>
            <tr><td><strong>End Date</strong></td><td>${moment(eventItem.EndDate).format("MM-DD-YYYY HH:mm")}</td></tr>
            <tr><td><strong>Location</strong></td><td>${eventItem.Location}</td></tr>
            <tr><td><strong>POC</strong></td><td>${pocString}</td></tr>
            </tbody>
        </table>
        <table class='table table-striped table-hover table-bordered'>
          <thead class='thead-dark'>
            <tr><th colspan='4' style='text-align:center'>Registered Users</th></tr>
            <tr>
            <th scope='col'>#</th>
            <th scope='col'>NAME</th>
            <th scope='col'>TIME IN</th>
            <th scope='col'>SIGNATURE</th>
            </tr>
          </thead>
          <tbody>
            ${usersTable}
          </tbody> 
        </table>
      </div>`;

    // Clear the modal
    Modal.clear();

    // Set the type
    Modal.setType(Components.ModalTypes.Large);

    // Set the header
    Modal.setHeader(eventItem.Title + " - Event Registrations");

    // Set the body
    Modal.setBody(Components.Jumbotron({
      title: eventItem.Title,
      content: jumbotronContent,
    }).el);

    // Set the footer
    Modal.setFooter(Components.ButtonGroup({
      buttons: [
        {
          text: "Print",
          type: Components.ButtonTypes.Primary,
          onClick: () => {
            // Show the print dialog
            window.print();
          },
        },
        {
          text: "Close",
          type: Components.ButtonTypes.Secondary,
          onClick: () => {
            // Close the modal
            Modal.hide();
          },
        },
      ]
    }).el);

    // Show the modal
    Modal.show();
  }
}