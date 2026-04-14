import { ListSecurity, ListSecurityDefaultGroups } from "dattatable";
import { ContextInfo, SPTypes, Types } from "gd-sprest-bs";
import { DataSource } from "./ds";
import Strings from "./strings";

/**
 * Security
 * Code related to the security groups the user belongs to.
 */
export class Security {
    private static _listSecurity: ListSecurity;

    // Current User
    static get CurrentUser(): Types.SP.User { return this._listSecurity.CurrentUser; }

    // Admin
    private static _isAdmin: boolean = false;
    static get IsAdmin(): boolean { return this._isAdmin; }
    private static _adminGroup: Types.SP.GroupOData;
    static get AdminGroup(): Types.SP.GroupOData { return this._adminGroup; }
    static get ManagersUrl(): string { return ContextInfo.webServerRelativeUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + this.AdminGroup.Id; }

    // Members
    private static _memberGroup: Types.SP.GroupOData;
    static get MemberGroup(): Types.SP.GroupOData { return this._memberGroup; }
    static get MembersUrl(): string { return ContextInfo.webServerRelativeUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + this.MemberGroup.Id; }

    // Visitors
    private static _visitorGroup: Types.SP.GroupOData;
    static get VisitorGroup(): Types.SP.GroupOData { return this._visitorGroup; }

    // Initializes the class
    static init(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let groups: Types.SP.GroupCreationInformation[] = [];

            // See if we are using a custom owner group
            let ownersGroupName = ListSecurityDefaultGroups.Owners;
            if (DataSource.Configuration.adminGroupName) {
                ownersGroupName = DataSource.Configuration.adminGroupName;
                groups.push({
                    AllowMembersEditMembership: false,
                    Title: DataSource.Configuration.adminGroupName,
                    Description: "Owners of the event registration system.",
                    OnlyAllowMembersViewMembership: false
                });
            }

            // See if we are using a custom members group
            let membersGroupName = ListSecurityDefaultGroups.Members;
            if (DataSource.Configuration.membersGroupName) {
                membersGroupName = DataSource.Configuration.membersGroupName;
                groups.push({
                    AllowMembersEditMembership: false,
                    Title: DataSource.Configuration.membersGroupName,
                    Description: "Members of the event registration system.",
                    OnlyAllowMembersViewMembership: false
                });
            }

            // Create the security component
            this._listSecurity = new ListSecurity({
                groups,
                listItems: [
                    {
                        listName: Strings.Lists.Events,
                        groupName: ownersGroupName,
                        permission: SPTypes.RoleType.Administrator
                    },
                    {
                        listName: Strings.Lists.Events,
                        groupName: membersGroupName,
                        permission: SPTypes.RoleType.Contributor
                    },
                    {
                        listName: Strings.Lists.Events,
                        groupName: ListSecurityDefaultGroups.Visitors,
                        permission: SPTypes.RoleType.Reader
                    }
                ],
                onGroupsLoaded: () => {
                    // Set the groups
                    this._adminGroup = this._listSecurity.getGroup(ownersGroupName);
                    this._memberGroup = this._listSecurity.getGroup(membersGroupName);
                    this._visitorGroup = this._listSecurity.getGroup(ListSecurityDefaultGroups.Visitors);

                    // Set the user flags
                    this._isAdmin = this._listSecurity.isInGroup(ContextInfo.userId, ownersGroupName);

                    // Ensure the groups exist
                    if (this._adminGroup && this._memberGroup && this._visitorGroup) {
                        // Resolve the request
                        resolve();
                    } else {
                        // Reject the request
                        reject();
                    }
                }
            });
        });
    }

    static hasPermissions(): PromiseLike<boolean> {
        // See if the user has permissions
        return this._listSecurity.checkUserPermissions();
    }

    // Displays the security group configuration
    static show(onComplete: () => void) {
        // Create the groups
        this._listSecurity.show(true, onComplete);
    }
}