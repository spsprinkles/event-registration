import { ListSecurity, ListSecurityDefaultGroups } from "dattatable";
import { ContextInfo, SPTypes, Types } from "gd-sprest-bs";
import { DataSource } from "./ds";
import Strings from "./strings";

/**
 * Security
 * Code related to the security groups the user belongs to.
 */
export class Security {
    private static _listSecurity: ListSecurity = null;

    // Current User
    static get CurrentUser(): Types.SP.User { return this._listSecurity.CurrentUser; }

    // Admin
    private static _isAdmin: boolean = false;
    static get IsAdmin(): boolean { return this._isAdmin; }
    private static _adminGroup: Types.SP.GroupOData = null;
    static get AdminGroup(): Types.SP.GroupOData { return this._adminGroup; }
    static get ManagersUrl(): string { return ContextInfo.webServerRelativeUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + this.AdminGroup.Id; }

    // Members
    private static _memberGroup: Types.SP.GroupOData = null;
    static get MemberGroup(): Types.SP.GroupOData { return this._memberGroup; }
    static get MembersUrl(): string { return ContextInfo.webServerRelativeUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + this.MemberGroup.Id; }

    // Visitors
    private static _visitorGroup: Types.SP.GroupOData = null;
    static get VisitorGroup(): Types.SP.GroupOData { return this._visitorGroup; }

    // Initializes the class
    static init(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            this._listSecurity = new ListSecurity({
                listItems: [
                    {
                        listName: Strings.Lists.Events,
                        groupName: DataSource.Configuration.adminGroupName || ListSecurityDefaultGroups.Owners,
                        permission: SPTypes.RoleType.Administrator
                    },
                    {
                        listName: Strings.Lists.Events,
                        groupName: DataSource.Configuration.membersGroupName || ListSecurityDefaultGroups.Members,
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
                    this._adminGroup = this._listSecurity.getGroup(ListSecurityDefaultGroups.Owners);
                    this._memberGroup = this._listSecurity.getGroup(ListSecurityDefaultGroups.Members);
                    this._visitorGroup = this._listSecurity.getGroup(ListSecurityDefaultGroups.Visitors);

                    // Set the user flags
                    this._isAdmin = this._listSecurity.isInGroup(ContextInfo.userId, ListSecurityDefaultGroups.Owners);

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