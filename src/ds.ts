import { List } from "dattatable";
import { Components, Types, Web } from "gd-sprest-bs";
import * as moment from "moment";
import { Security } from "./security";
import Strings from "./strings";

// Event Item
export interface IEventItem extends Types.SP.ListItemOData {
    Description: string;
    StartDate: string;
    EndDate: string;
    EventStatus: { results: string };
    IsCancelled: boolean;
    Location: string;
    OpenSpots: string;
    Capacity: number;
    POC: {
        results: {
            EMail: string;
            Id: number;
            Title: string;
        }[]
    };
    POCId: { results: number[] };
    RegisteredUsers: {
        results: {
            EMail: string;
            Id: number;
            Title: string;
        }[]
    };
    RegisteredUsersId: { results: number[] };
    WaitListedUsers: {
        results: {
            EMail: string;
            Id: number;
            Title: string;
        }[]
    };
    WaitListedUsersId: { results: number[] };
}

// Configuration
export interface IConfiguration {
    adminGroupName?: string;
    headerImage?: string;
    headerTitle?: string;
    hideAddToCalendarColumn?: boolean;
    hideHeader?: boolean;
    membersGroupName?: string;
}
/**
 * Data Source
 */
export class DataSource {
    // Filter Set
    private static _filterSet: boolean = false;
    static get FilterSet(): boolean { return this._filterSet; }
    static SetFilter(filterSet: boolean) {
        this._filterSet = filterSet;
    }

    // List
    private static _list: List<IEventItem> = null;
    static get List(): List<IEventItem> { return this._list; }

    // Events
    static get Events(): IEventItem[] {
        // See if we are filtering for active items
        if (this.FilterSet) {
            let activeEvents: IEventItem[] = [];
            let today = moment();

            // Parse the events
            this.List.Items.forEach((event) => {
                let startDate = event.StartDate;

                // See if this event is active
                if (moment(startDate).isAfter(today)) {
                    // Add the event
                    activeEvents.push(event);
                }
            });

            // Return the active events
            return activeEvents;
        }

        // Return all of the events
        return this.List.Items;
    }
    static loadEvents(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let today = moment().toISOString();

            // Load the list
            this._list = new List({
                listName: Strings.Lists.Events,
                onInitError: () => { reject(); },
                onInitialized: () => { resolve(); },
                itemQuery: Security.IsAdmin ? {
                    Expand: ["AttachmentFiles", "POC", "RegisteredUsers", "WaitListedUsers"],
                    GetAllItems: true,
                    OrderBy: ["StartDate asc"],
                    Top: 5000,
                    Select: [
                        "*", "POC/Id", "POC/Title", "POC/EMail",
                        "RegisteredUsers/Id", "RegisteredUsers/Title", "RegisteredUsers/EMail",
                        "WaitListedUsers/Id", "WaitListedUsers/Title", "WaitListedUsers/EMail"
                    ]
                } : {
                    Expand: ["AttachmentFiles", "POC", "RegisteredUsers", "WaitListedUsers"],
                    Filter: `StartDate ge '${today}'`,
                    GetAllItems: true,
                    OrderBy: ["StartDate asc"],
                    Top: 5000,
                    Select: [
                        "*", "POC/Id", "POC/Title",
                        "RegisteredUsers/Id", "RegisteredUsers/Title", "RegisteredUsers/EMail",
                        "WaitListedUsers/Id", "WaitListedUsers/Title", "WaitListedUsers/EMail"
                    ]
                }
            });
        });
    }

    // Status Filters
    private static _statusFilters: Components.ICheckboxGroupItem[] = [{
        label: "Show inactive events",
        type: Components.CheckboxGroupTypes.Switch,
        isSelected: false
    }];
    static get StatusFilters(): Components.ICheckboxGroupItem[] { return this._statusFilters; }

    // Loads the list data
    static init(): PromiseLike<any> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the optional configuration file
            this.loadConfiguration().then(() => {
                // Initialize the security
                Security.init().then(() => {
                    // Initialize the solution
                    Promise.all([
                        // Load the events
                        this.loadEvents()
                    ]).then(resolve, reject);
                }, reject);
            });
        });
    }

    // Configuration
    private static _cfg: IConfiguration = null;
    static get Configuration(): IConfiguration { return this._cfg; }
    static loadConfiguration(): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Get the current web
            Web().getFileByServerRelativeUrl(Strings.EventRegConfig).content().execute(
                // Success
                file => {
                    // Convert the string to a json object
                    let cfg = null;
                    try { cfg = JSON.parse(String.fromCharCode.apply(null, new Uint8Array(file))); }
                    catch { cfg = {}; }

                    // Set the configuration
                    this._cfg = cfg;

                    // Resolve the request
                    resolve();
                },

                // Error
                () => {
                    // Set the configuration to nothing
                    this._cfg = {} as any;

                    // Resolve the request
                    resolve();
                }
            );
        });
    }
}