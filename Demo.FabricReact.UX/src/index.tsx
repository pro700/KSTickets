import "core-js/modules/es6.array.iterator.js";
import "core-js/modules/es6.array.from.js";
import "whatwg-fetch";
import "es6-map/implement";

import 'raf/polyfill';

import 'core-js/es6/map';
import 'core-js/es6/set';
import "core-js/es6/promise";

//import * as moment from 'moment';
//import 'moment/locale/uk';

// These are the references to the react library
import * as React from "react";
import * as ReactDOM from "react-dom";

import 'jquery';

// This is how you import the components you need from the Office Fabric React Framework
import { Button, List } from 'office-ui-fabric-react';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Callout } from 'office-ui-fabric-react/lib/Callout';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { getId } from 'office-ui-fabric-react/lib/Utilities';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
//import 'office-ui-fabric-react/src/components/List/examples/List.Basic.Example.scss';


import { sp, Web, ItemAddResult, SPBatch } from '@pnp/sp';
//import { SiteUser } from '@pnp/sp/src/siteusers';

import './index.scss';
import { mergeStyleSets, FontWeights, DefaultPalette, getTheme } from 'office-ui-fabric-react/lib/Styling';
import { CurrentUser, SiteUser } from "@pnp/sp/src/siteusers";

import { TicketTable, TicketTableProps } from "./TicketTable";

/**
 * TicketsList Component
 */

/*
interface TicketListProps {
}

interface TicketListState {
    ID: string,
    Title: string,
    user?: any,
    items?: any[],
    isCalloutVisible?: boolean,
    calloutTarget?: Element, 
    detailsCalloutItem?: any,
    Message?: string
}

export class TicketList extends React.Component<TicketListProps, TicketListState> {

    private _menuButtonElement = React.createRef<HTMLDivElement>();
    // Use getId() to ensure that the callout label and description IDs are unique on the page.
    // (It's also okay to use plain strings without getId() and manually ensure their uniqueness.)
    private _labelId: string = getId('callout-label');
    private _descriptionId: string = getId('callout-description');

    constructor(props: TicketListProps) {
        super(props);
        this.state = { ID: "Loading", Title: "Loading...", isCalloutVisible: false };
        this.populate();
    }

    render(): JSX.Element {
        const theme = getTheme();
        const state = this.state.items;

        const styles = mergeStyleSets({
            tableClass: {
                border: '2px solid white',
                selectors: {
                    // selector for child
                    '& th': {
                        background: "#ffe6eb",
                        border: '2px solid white',
                        padding: 4
                    },
                    '& td': {
                        background: theme.palette.themeLighterAlt,
                        border: '2px solid white',
                        padding: 4
                    },
                    '& tr.booking td': {
                    },
                    '& td.free': {
                        color: theme.palette.green,
                        'font-weight': 'bold',
                        padding: 4
                    },
                    '& td.ordered': {
                        padding: 2
                    }
                }
            },

            tableClass1: {
                selectors: {
                    '& td': {
                        background: "#fff4e3",
                        border: '1px solid white',
                        padding: 4
                    }
                }
            },

            buttonArea: {
                verticalAlign: 'top',
                display: 'inline-block',
                textAlign: 'center'
            },

            buttonClass: {
                background: theme.palette.themeLighter,
                margin: 0,
                padding: 4,
                height: null,
                "min-width": 0,
                "border-radius": 4
            },

            buttonClassDel: {
                background: theme.palette.red,
                color: theme.palette.white
            },

            rootClass: {
                background: DefaultPalette.themeLighterAlt,
            },

            itemClass: {
                background: DefaultPalette.themeLight,
                padding: 4,
                selectors: {
                    ':hover': {
                        background: DefaultPalette.themeLight
                    },
                    ':focus': {
                        background: DefaultPalette.themeLight
                    },
                    ':active': {
                        background: DefaultPalette.themeLight
                    }
                }
            }
        });
        if (this.state.ID == "table") {
            return (
                <div>
                    { this.state.Message &&
                        <p>{this.state.Message}</p>
                    }
                    <table className={styles.tableClass}>
                        <tbody>
                            <tr key="0">
                                <th>Час</th>
                                <th>Назва</th>
                                <th></th>
                                <th>Доступно</th>
                                <th>Замовлено</th>
                                <th>Коментар</th>
                                <th>etag</th>
                            </tr>
                            {this.state.items.map(item => {
                                var Ordered = 0;
                                var Free = item.play.Seats;
                                item.bookings.map(booking => {
                                    Free -= booking.Seats;
                                    Ordered += booking.Seats;
                                });

                                return (
                                    <tr key={item.play.ID}>
                                        <td>{(new Date(item.play.DateTime)).format('dd.MM.yyyy')}</td>
                                        <td><a href={item.play.Link}>{item.play.Title}</a></td>
                                        <td>
                                            {(Free > 0) &&
                                                <DefaultButton className={styles.buttonClass}
                                                    id={item.ID}
                                                    onClick={this._onOrderTicketsClicked}
                                                    text="Замовити" />
                                            }
                                        </td>
                                        <td align="center" className='free'>{(Free > 0) && Free}</td>
                                        <td align="center" className='ordered'>
                                            { (Ordered > 0) &&
                                                <DefaultButton className={styles.buttonClass + " orderDetailsButton"}
                                                    id={item.ID}
                                                    onClick={this._onShowMenuClicked}
                                                    text={Ordered + " ..."}
                                                    title='Деталі замовлень' />
                                            }
                                        </td>
                                        <td>{item.play.Comments}</td>
                                        <td>{item.play.etag}</td>
                                    </tr>
                                );
                            })}
                        </tbody>
                    </table>
                    { this.state.isCalloutVisible &&
                        <Callout
                            className="ms-CalloutExample-callout"
                            ariaLabelledBy={this._labelId}
                            ariaDescribedBy={this._descriptionId}
                            role="alertdialog"
                            gapSpace={0}
                            target={this.state.calloutTarget}
                            onDismiss={this._onCalloutDismiss}
                            setInitialFocus={true}
                            hidden={!this.state.isCalloutVisible}
                        >
                            <div>
                                {this.state.detailsCalloutItem && this.state.detailsCalloutItem.bookings.length > 0 && (
                                    <table className={styles.tableClass + ' ' + styles.tableClass1}>
                                        <tbody> {
                                            this.state.detailsCalloutItem.bookings.map(booking => {
                                                return (
                                                    <tr className="booking">
                                                        <td>{booking.WhoBooked.Title}</td>
                                                        <td>{booking.Status}</td>
                                                        <td>{booking.Seats}</td>
                                                        <td>{
                                                            booking.WhoBooked.ID == this.state.user["Id"] &&
                                                            <DefaultButton className={styles.buttonClass + " " + styles.buttonClassDel}
                                                                title="Видалити замовлення"
                                                                id={booking.ID}
                                                                onClick={this._onDeleteBookingClicked}
                                                                text=" Х " />
                                                        }
                                                        </td>
                                                    </tr>
                                                );
                                            })
                                        }
                                        </tbody>
                                    </table>
                                )}
                            </div>
                        </Callout>
                    }
                </div>

            );
        }
        else if (this.state.ID == "error"){
            return <p>Error:{this.state.Title}</p>;
        }
        else {
            return <p>{this.state.Title}</p>;
        }
    }

    private _onDeleteBookingClicked = (event): void => {
        let target = event.target;
        sp.web.lists.getByTitle("Booking").items.getById(target.id)
            .delete()
            .then(res => {
                let newitems = JSON.parse(JSON.stringify(this.state.items));
                newitems = newitems.map(item => {
                    item.bookings = item.bookings.filter(booking => {
                        return ":" + booking.ID !== ":" + target.id;
                    });
                    return item;
                });
                this.setState({
                    items: newitems,
                    isCalloutVisible: false,
                    calloutTarget: null,
                    detailsCalloutItem: null
                });
            });
    }

    private _onOrderTicketsClicked = (event): void => {
        this.state.items.forEach(item => {
            if (item.play.ID == event.target.id) {
                sp.web.lists.getByTitle("Plays").items.getById(item.play.ID)
                    .update({ Title: item.Title }, item.etag)
                    .then(updated_item => {
                        // etag ok!
                        sp.web.currentUser.get()
                            .then((user: CurrentUser) => {
                                sp.web.lists.getByTitle("Booking").items
                                    .add({ PlayId: item.play.ID, Seats: 2, WhoBookedId: user["Id"]})
                                    .then((result: ItemAddResult) => {
                                        this.populate();
                                    });
                            })
                    })
            }
        });
    }

    private _onShowMenuClicked = (event): void => {
        this.state.items.forEach(item => {
            if (item.ID == event.target.id) {
                this.setState({
                    isCalloutVisible: !this.state.isCalloutVisible,
                    calloutTarget: event.target,
                    detailsCalloutItem: item
                });
            }
        });
    };

    private _onCalloutDismiss = (event): void => {
        this.setState({
            isCalloutVisible: false,
            calloutTarget: event.target,
            detailsCalloutItem: null
        });
    };


    private _onRenderCell(item: JSX.Element, index: number | undefined): JSX.Element {
        return item;
    }

    //setUsernameState = (event) => {
    //    this.setState({ userName: event.target.value });
    //}

    private populate() {

        let today = new Date();

        sp.web.currentUser.get().then((user: CurrentUser) => {
            sp.web.lists.getByTitle("Plays").items
                .select("ID", "Title", "Seats", "DateTime", "Link", "Comments")
                .filter("DateTime ge datetime'" + today.toISOString() + "'")
                .orderBy("DateTime", false)
                .get()
                .then((plays: any[]) => {

                    var promises: Promise<any[]>[] = plays.map(play => {
                        return sp.web.lists.getByTitle("Booking").items
                            .select("ID", "Title", "Play/ID", "WhoBooked/ID", "WhoBooked/Title", "Seats", "Status", "Notes", "GivenAway", "Author/ID", "Author/Title")
                            .expand("WhoBooked", "Play", "Author")
                            .filter("Play/ID eq " + play["ID"])
                            .get();
                    });

                    Promise.all(promises).then((allbookings: any[][]) => {
                        var items: any[] = plays.map((play: any, i: number) => {
                            var bookings: any[] = allbookings[i].map((booking: any) => {
                                return {
                                    ID: booking.ID,
                                    Status: booking.Status,
                                    Notes: booking.Notes,
                                    GivenAway: booking.GivenAway,
                                    Seats: booking.Seats,
                                    WhoBooked: {
                                        ID: booking.WhoBooked.ID,
                                        Title: booking.WhoBooked.Title,
                                    },
                                    etag: booking["odata.etag"]
                                };
                            });
                            return {
                                ID: play.ID,
                                Title: play.Title,
                                play: {
                                    ID: play.ID,
                                    DateTime: play.DateTime,
                                    Title: play.Title,
                                    Link: play.Link,
                                    Seats: play.Seats,
                                    Comments: play.Comments,
                                    etag: play["odata.etag"]
                                },
                                bookings: bookings
                            };
                        });

                        this.setState({ ID: "table", Title: "Список вистав", user: user, items: items });

                    }).catch((err: any) => {
                        this.setState({ ID: "error", Title: "populate err 2: " + JSON.stringify(err), user: user });
                    });

                })
                .catch((err: object) => {
                    this.setState({ ID: "error", Title: "populate err 1: " + JSON.stringify(err), user: user });
                });
       });  

    }
}
*/

$("span#DeltaPlaceHolderPageTitleInTitleArea").text("Квитки в театр");



// Get the "main" element
let target = document.querySelector("#tickets_conainer");
if (target) {

    ReactDOM.render(<TicketTable/>, target);

    //var user: SiteUser = sp.web.getUserById(0);
    //var web = new Web("http://intranet.kyivstar.ua");

    /*
    sp.web.siteUsers.getByEmail("Sergiy.Sokol@kyivstar.net").get()
        .then(user => {
            sp.profiles.getPropertiesFor(user.LoginName)
                .then(result => {
                    var props = result.UserProfileProperties;
                    var propValue = "";
                    props.forEach(function (prop) {
                        propValue += prop.Key + " - " + prop.Value + "<br/>";
                    });
                    document.getElementById("tickets_conainer").innerHTML = "user=" + JSON.stringify(user) + "<br/>" + propValue;
                })
                .catch(function (err) {
                    document.getElementById("tickets_conainer").innerHTML = "Error: " + err;
                });
        });
        */

    //sp.profiles.myProperties.get().then(function (result) {
    //    var props = result.UserProfileProperties;
    //    var propValue = "";
    //    props.forEach(function (prop) {
    //        propValue += prop.Key + " - " + prop.Value + "<br/>";
    //    });
    //    document.getElementById("tickets_conainer").innerHTML = propValue;
    //}).catch(function (err) {
    //    console.log("Error: " + err);
    //});
    

        
}


/*

export interface EmployeeRowProps {
    fields: {
        index: number,
        id: string,
        Title: string,
        EMail: string,
        LoginName: string,
    }
}

export interface EmployeeRowState {
    Department: string,
    SPSDepartment: string,
    JobTitle: string,
    SPSJobTitle: string,
    OrganizationalStructure: string,
    PreferredName: string,
    Manager: string,
    SPSBirthday: string,
    EmployeeType: string,
    BranchName: string
}

interface PromiseItem {
    index: number,
    promise: Promise<any>
}

var promiseItems: PromiseItem[] = [];

//var last_proceeding_promise: Promise<any> = null;
//var proceeding_promises_count: number = 0; 

//let batch1: SPBatch = sp.web.createBatch();
//var batchCounter1 = 0;

//setTimeout(function () {
//    batch1.execute();
//}, 100);

export class EmployeeRow extends React.Component<EmployeeRowProps, EmployeeRowState> {
    constructor(props: EmployeeRowProps) {
        super(props);
        // Init state
        this.state = {
            Department: "",
            SPSDepartment: "",
            JobTitle: "",
            SPSJobTitle: "",
            OrganizationalStructure: "",
            PreferredName: "",
            Manager: "",
            SPSBirthday: "",
            EmployeeType: "",
            BranchName: ""
        };
    }

    render() {
        return (
            <tr key={this.props.fields.index}>
                <td>{this.props.fields.index}</td>
                <td>{this.props.fields.id}</td>
                <td>{this.props.fields.Title}</td>
                <td>{this.props.fields.EMail}</td>
                <td>{this.props.fields.LoginName}</td>
                <td>{this.state.Department}</td>
                <td>{this.state.SPSDepartment}</td>
                <td>{this.state.JobTitle}</td>
                <td>{this.state.SPSJobTitle}</td>
                <td>{this.state.PreferredName}</td>
                <td>{this.state.Manager}</td>
                <td>{this.state.SPSBirthday}</td>
                <td>{this.state.EmployeeType}</td>
                <td>{this.state.BranchName}</td>
                <td>{this.state.OrganizationalStructure}</td>
            </tr>
        );
    }


    onPrevResolved(err: string) {

        let promise: Promise<any> = sp.profiles.getPropertiesFor(this.props.fields.LoginName);

        // push
        promiseItems.push({ index: this.props.fields.index, promise: promise })

        promise.then(result => {
            // pop
            promiseItems = promiseItems.filter(item => { return item.index !== this.props.fields.index; });

            var state: EmployeeRowState = {
                Department: "",
                SPSDepartment: "",
                JobTitle: "",
                SPSJobTitle: "",
                OrganizationalStructure: "",
                PreferredName: "",
                Manager: "",
                SPSBirthday: "",
                EmployeeType: "",
                BranchName: ""
            };

            if (result.UserProfileProperties) {
                result.UserProfileProperties.map(prop => {
                    if (false) { }
                    else if (prop.Key == "Department") { state.Department = prop.Value; }
                    else if (prop.Key == "SPS-Department") { state.SPSDepartment = prop.Value; }
                    else if (prop.Key == "Title") { state.JobTitle = prop.Value; }
                    else if (prop.Key == "SPS-JobTitle") { state.SPSJobTitle = prop.Value; }
                    else if (prop.Key == "OrganizationalStructure") { state.OrganizationalStructure = prop.Value; }
                    else if (prop.Key == "PreferredName") { state.PreferredName = prop.Value; }
                    else if (prop.Key == "Manager") { state.Manager = prop.Value; }
                    else if (prop.Key == "SPS-Birthday") { state.SPSBirthday = prop.Value; }
                    else if (prop.Key == "EmployeeType") { state.EmployeeType = prop.Value; }
                    else if (prop.Key == "BranchName") { state.BranchName = prop.Value; }
                });
            }

            this.setState(state);

        })
        .catch(err => {
            //proceeding_promises_count--;

            this.setState({ OrganizationalStructure: "login=" + this.props.fields.LoginName + "<br/> error=" + err });
        });
    }

    componentDidMount() {

        if (promiseItems.length > 1000) {

            let p1 = promiseItems.pop();
            promiseItems.push(p1);

            p1.promise
                .then(() => {
                    if (promiseItems.length > 1000) {
                        this.componentDidMount();
                    }
                    else {
                        this.onPrevResolved("");
                    }
                })
                .catch((err) => {
                    if (promiseItems.length > 1000) {
                        this.componentDidMount();
                    }
                    else {
                        this.onPrevResolved(err);
                    }
                });
        }
        else {
            this.onPrevResolved("");
        }
    }
}

export interface ETProps {

}

export interface ETState {
    rows: {
        index: number,
        id: string,
        Title: string,
        EMail: string,
        LoginName: string
    }[]
}

export class EmployeeTable extends React.Component<ETProps, ETState> {
    constructor(props: ETProps) {
        super(props);
        // Init state
        this.state = { rows: [] };
    }

    render(): JSX.Element {
        return (
            <div>
                <p>
                    EmployeeTable:
                </p>
                <table className="AllBordersTable">
                    <tbody>
                        <tr key={-1}>
                            <th>index</th>
                            <th>id</th>
                            <th>Title</th>
                            <th>EMail</th>
                            <th>LoginName</th>
                            <th>Department</th>
                            <th>SPSDepartment</th>
                            <th>JobTitle</th>
                            <th>SPSJobTitle</th>
                            <th>PreferredName</th>
                            <th>Manager</th>
                            <th>SPSBirthday</th>
                            <th>EmployeeType</th>
                            <th>BranchName</th>
                            <th>OrganizationalStructure</th>
                        </tr>
                        {this.state.rows.map((row: any, index: number) => {
                            return <EmployeeRow fields={{ index: row.index, id: row.id, Title: row.Title, EMail: row.EMail, LoginName: row.LoginName }} />;
                        })}
                    </tbody>
                </table>
            </div>
        );
    }

    componentDidMount() {

        sp.web.siteUsers.get()
            .then((users: any[]) => {
                var state: ETState = { rows: [] };
                users.forEach((user: any, index: number) => {
                    if (index > 0 && index < 80000) {
                        state.rows.push({
                            index: index,
                            id: user.Id,
                            Title: user.Title,
                            EMail: user.Email,
                            LoginName: user.LoginName
                        });
                    }
                });

                this.setState(state);

            });
    }
}

//ReactDOM.render(<EmployeeTable/>, target);

*/