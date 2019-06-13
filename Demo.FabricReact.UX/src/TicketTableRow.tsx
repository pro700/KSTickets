
// These are the references to the react library
import * as React from "react";
import * as ReactDOM from "react-dom";

import 'jquery';

import { TicketTableParams, TicketTableMode } from "./TicketTable";

import { sp, Web, ItemAddResult, ItemUpdateResult, EmailProperties } from '@pnp/sp';
import { CurrentUser } from "@pnp/sp/src/siteusers";
import { PermissionKind, BasePermissions } from "@pnp/sp";

import { Spinner, SpinnerSize, DefaultButton, Callout, getId } from "office-ui-fabric-react";
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react';
import { TooltipHost, ITooltipHostProps } from 'office-ui-fabric-react/lib/Tooltip';


import { mergeStyleSets, FontWeights, DefaultPalette, getTheme } from 'office-ui-fabric-react/lib/Styling';
//import { initializeIcons } from 'office-ui-fabric-react/icons';


import './TicketTable.scss';
//import { Promise, resolve } from "bluebird";
import { string } from "prop-types";

export enum TicketTableRowMouseEventCode {
    showBookings
}

export interface TicketTableRowBooking {
    ID: string,
    Status: string,
    Notes: string,
    Seats: number,
    WhoBooked: {
        ID: string,
        Title: string,
        Name: string,
        EMail: string,
        OrganizationalStructure?: string
    },
    etag: string
}

export interface TicketTableRowProps {
    key: string,
    isHeader?: boolean,
    user: CurrentUser,
    bookingPerm: BasePermissions,
    playsPerm: BasePermissions,
    params: TicketTableParams,
    fields?: {
        id: number,
        DateTime: string,
        Title: string,
        Link: string,
        Seats: number,
        Comments: string,
        etag: string
    },
    rowRef?: React.RefObject<TicketTableRow>,
    BookingChangedCallback?: (playId: string) => {},
    StatusChangedCallback?: (playId: string) => {},
    OrderAttemptFailed?: (playId: string, attempt: number) => {},
    MessageCallback?: (msg: string, err: string, scs: string) => {},
    isMangerForm?: boolean,
    mode?: TicketTableMode,
    bookingArr: TicketTableRowBooking[]
}

export interface TicketTableRowState {
    bookingArr: TicketTableRowBooking[],
    isBookingsCalloutVisible: boolean,
    loading?: boolean,
    calloutTarget?: any
    error?: string,
    etag: string,
    currCalloutWhoBookedName?: string,
    whoBookedInfo?: any,
    whoBookedCalloutTarget?: any
}

const theme = getTheme();
const styles = mergeStyleSets({

    tableClassBookings: {
        selectors: {
            '& td': {
                padding: 2
            }
        }
    },

    buttonClassDel: {
        background: theme.palette.red,
        margin: 0,
        padding: 4,
        height: null,
        "min-width": 0,
        "border-radius": 4,
        color: theme.palette.white
    }
});

var intranetWeb = new Web('http://intranet.kyivstar.ua');

export class TicketTableRow extends React.Component<TicketTableRowProps, TicketTableRowState> {

    private calloutRef = React.createRef<Callout>();
    private tdOrderdRef = React.createRef<HTMLTableDataCellElement>();
    private _hostId: string = getId('tooltipHost');
    private _hostId1: string = getId('tooltipHost');

    constructor(props: TicketTableRowProps) {
        super(props);
        var etag = props.fields ? props.fields.etag : "";
        // Init state
        this.state = { bookingArr: props.bookingArr, loading: false, etag: etag, isBookingsCalloutVisible: false };
    }

    render(): JSX.Element {
        if (this.props.isHeader) {
            return (
                <tr key={this.props.key} >
                    <th></th>
                    <th>Дата</th>
                    <th>Назва</th>
                    <th>Замовити</th>
                    <th>Доступно</th>
                    <th>Замовлено</th>
                </tr>
            );

        }
        var Ordered = 0;
        var Rejected = 0;
        var Free = this.props.fields.Seats;
        this.state.bookingArr.map(booking => {
            if (booking.Status !== "Відхилено") {
                Free -= booking.Seats;
                Ordered += booking.Seats;
            }
            else {
                Rejected += booking.Seats;
            }
        });

        return (
            <tr key={this.props.key}>
                <td>
                    {this.state.error &&
                        <TooltipHost content={this.state.error} id={this._hostId} calloutProps={{ gapSpace: 0 }}>
                            <DefaultButton className="RowErrorButton" text={"!"} />
                        </TooltipHost>
                    }
                </td>
                <td>{(new Date(this.props.fields.DateTime)).format('dd-MM-yyyy HH:mm')}</td>
                <td><a href={this.props.fields.Link}>{this.props.fields.Title}</a></td>
                <td>
                    {
                        this.state.loading && <Spinner size={SpinnerSize.medium} /> ||
                        /*Free > 0 &&*/ (new Date(this.props.fields.DateTime)) >= (new Date()) &&
                        <DefaultButton className="CommonButton OrderButton"
                            id={this.props.fields.id.toString()}
                            onClick={() => { return this._onOrderTicketsClicked(1); }}
                            text="Замовити" />
                    }
                </td>
                <td className='free'>
                        {
                            this.state.loading && <span> </span> || Free > 0 && <div className='freetickets'>{Free}</div>
                        }
                   
                </td>
                <td className='ordered' ref={this.tdOrderdRef}>
                    {
                        this.state.loading && <span> </span> ||
                        Ordered > 0 && Ordered + Rejected > 0 &&
                        <>
                            <div>
                                <DefaultButton className="CommonButton OrderButton"
                                    id={this.props.fields.id.toString()}
                                    onClick={() => {
                                        this.setState({ isBookingsCalloutVisible: !this.state.isBookingsCalloutVisible })
                                    }}
                                    text={Ordered + " ..."}
                                    title='Деталі замовлень' />
                            </div>
                            { this.renderBookingsCallout(Ordered, Rejected, Free) }
                        </>
                    }
                </td>
            </tr>
        );
    }

    renderBookingsCallout(Ordered, Rejected, Free): JSX.Element {
        return this.state.isBookingsCalloutVisible && (
            <Callout
                ref={this.calloutRef}
                role="alertdialog"
                gapSpace={0}
                target={this.tdOrderdRef.current}
                setInitialFocus={true}
                hidden={!this.state.isBookingsCalloutVisible}
                onDismiss={() => {
                    this.setState({ isBookingsCalloutVisible: false })
                }}
                className="TicketTableCallout"
            >
                <div>
                    {this.state.bookingArr && this.state.bookingArr.filter(booking => { return booking.Status != "Відхилено"; }).length > 0 && (
                        <table className={styles.tableClassBookings + " BookingTable"}>
                            <tbody> {
                                this.state.bookingArr.filter(booking => { return booking.Status != "Відхилено"; }).map((booking, i) => {
                                    return (
                                        <tr key={booking.ID}>
                                            <td>
                                                <TooltipHost
                                                    content={booking.WhoBooked.OrganizationalStructure}
                                                    id={this._hostId1}
                                                    calloutProps={{ gapSpace: 0 }}
                                                    onTooltipToggle={(isVisible) => {
                                                        if (isVisible) {
                                                            var arr = this.state.bookingArr;
                                                            sp.profiles.getPropertiesFor(booking.WhoBooked.Name).then(result => {
                                                                result.UserProfileProperties.forEach(function (prop) {
                                                                    if (prop.Key == "OrganizationalStructure") {
                                                                        arr[i].WhoBooked.OrganizationalStructure = (prop.Value as string).split("\\").reverse().join("; ");
                                                                    }
                                                                });
                                                                this.setState({ bookingArr: arr });
                                                            });
                                                        }
                                                    }}
                                                >
                                                    <div style={{cursor: "pointer"}}> {booking.WhoBooked.Title} </div>
                                                </TooltipHost>
                                            </td>
                                            <td>{booking.Seats}</td>
                                            <td>
                                                { booking.Status
                                                    /* || <Dropdown
                                                        defaultSelectedKey={booking.Status}
                                                        options={[
                                                            { key: 'В очікуванні', text: 'В очікуванні' },
                                                            { key: 'Відхилено', text: 'Відхилено' },
                                                            { key: 'Затверджено', text: 'Затверджено' }
                                                        ]}
                                                        onChange={(event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) => {
                                                            this._onStatusChange(booking, event, option, index);
                                                        }}
                                                        disabled={
                                                            !sp.web.hasPermissions(this.props.bookingPerm, PermissionKind.ApproveItems) ||
                                                            !sp.web.hasPermissions(this.props.bookingPerm, PermissionKind.EditListItems) ||
                                                            (booking.Status == "Відхилено" && (booking.Seats > Free))
                                                        }
                                                    />
                                                    */
                                                }
                                            </td>
                                            <td>{
                                                booking.WhoBooked.ID == this.props.user["Id"] &&
                                                booking.Status == "В очікуванні" &&
                                                <DefaultButton className={styles.buttonClassDel + " DelButton"}
                                                    title={"Видалити замовлення"}
                                                    id={booking.ID}
                                                    onClick={() => { this._onDeleteBookingClicked(booking.ID) }}
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
        );
    }

    componentDidMount() {
    }

    componentDidUpdate(prevProps, prevState) {
    }

    private load(): Promise<TicketTableRowBooking[]> {

        return new Promise<TicketTableRowBooking[]>((resolve, reject) => {
            if (this.props.fields) {
                this.setState({ loading: true });

                sp.web.lists.getByTitle("Booking").items
                    .select("ID", "Title", "Play/ID", "WhoBooked/ID", "WhoBooked/Title", "WhoBooked/Name", "WhoBooked/EMail", "Seats", "Status", "Notes", "Author/ID", "Author/Title")
                    .expand("WhoBooked", "Play", "Author")
                    .filter("Play/ID eq " + this.props.fields.id)
                    .get()
                    .then((bookings: any[]) => {

                        var bookingArr: TicketTableRowBooking[] = bookings.map(booking => {
                            return {
                                ID: booking.ID,
                                Status: booking.Status,
                                Notes: booking.Notes,
                                Seats: booking.Seats,
                                WhoBooked: {
                                    ID: booking.WhoBooked.ID,
                                    Title: booking.WhoBooked.Title,
                                    Name: booking.WhoBooked.Name,
                                    EMail: booking.WhoBooked.EMail,
                                    OrganizationalStructure: ""
                                },
                                etag: booking["odata.etag"]
                            };
                        });

                        this.setState({ bookingArr: bookingArr, loading: false, error: "" });
                        resolve(bookingArr);
                    })
                    .catch(err => {
                        this.setState({ loading: false, error: err, isBookingsCalloutVisible: false });
                        reject(err);
                    });
            }
            else {
                resolve([]);
            }
        });
    }


    private monthAdd(date, month) {
        var temp = date;
        temp = new Date(date.getFullYear(), date.getMonth(), 1);
        temp.setMonth(temp.getMonth() + (month + 1));
        temp.setDate(temp.getDate() - 1);

        if (date.getDate() < temp.getDate()) {
            temp.setDate(date.getDate());
        }

        return temp;
    }

/*
    _onStatusChangeDel(booking: TicketTableRowBooking, new_status: string, new_notes: string) {

        this.setState({ error: "" });
        this.props.MessageCallback("", "", "");

        var months: number = this.props.params.monthCount;

        var p3: Promise<any[]> = new Promise<any[]>((resolve) => { resolve([]); });

        if (booking.Status == "Відхилено" && new_status != "Відхилено") {
            p3 = sp.web.lists.getByTitle("Booking").items
                .select("ID", "Title", "Play/ID", "Play/DateTime", "WhoBooked/ID", "WhoBooked/Title", "WhoBooked/Name", "WhoBooked/EMail", "Seats", "Status", "Notes", "Author/ID", "Author/Title")
                .expand("WhoBooked", "Play", "Author")
                .filter("WhoBooked/ID eq " + this.props.user["Id"] + " and Play/DateTime ge datetime'" + this.monthAdd(new Date(), -months).toISOString() + encodeURI("' and Status ne 'Відхилено'"))
                .get();
        }

        p3.then((items: any[]) => {
            if (items.length > 0) {
                this.setState({ error: "Не можна змінювати статус замовлення. За заданний період (" + months + " місяців) вже є замовлення..." });
                this.props.MessageCallback("", "Не можна змінювати статус замовлення. За заданний період (" + months + " місяців) вже є замовлення...", "");
            }
            else {
                var p1: Promise<ItemUpdateResult> = sp.web.lists.getByTitle("Plays").items.getById(this.props.fields.id)
                    .update({ Title: this.props.fields.Title }, this.state.etag);

                var p2: Promise<ItemUpdateResult> = sp.web.lists.getByTitle("Booking").items.getById(parseInt(booking.ID, 10))
                    .update({ Status: new_status, Notes: new_notes }, booking.etag);

                Promise.all([p1, p2])
                    .then(([res1, res2]: [ItemUpdateResult, ItemUpdateResult]) => {
                        this.setState({ etag: res1.data["odata.etag"] });
                        this.load().then((updatedBookingArr: TicketTableRowBooking[]) => {
                            if (this.state.isBookingsCalloutVisible) {
                                this.setState({ isBookingsCalloutVisible: (updatedBookingArr.length > 0) });
                            }

                            updatedBookingArr.forEach(updatedBooking => {
                                if (updatedBooking.ID == booking.ID) {

                                    if (updatedBooking.Status == "Відхилено") {
                                        // Параметри: %Замовлено, %Вистава, %Статус, %Дата, %Лінк",%Коментар
                                        var body: string = this.props.params.TicketsOrderedEmailText
                                            .replace("%Замовлено", updatedBooking.Seats.toString())
                                            .replace("%Вистава", this.props.fields.Title)
                                            .replace("%Дата", (new Date(this.props.fields.DateTime)).format('dd.MM.yyyy HH:mm'))
                                            .replace("%Статус", updatedBooking.Status)
                                            .replace("%Лінк", this.props.fields.Link)
                                            .replace("%Коментар", updatedBooking.Notes);

                                        sp.utility.sendEmail({
                                            To: [updatedBooking.WhoBooked.EMail],
                                            Subject: "Статус замовлення квитків змінився...",
                                            Body: body
                                        })
                                    }
                                    else {
                                        // Параметри: %Замовлено, %Вистава, %Статус, %Дата, %Лінк",%Коментар
                                        var body: string = this.props.params.StatusChangedRejectedEmailText
                                            .replace("%Замовлено", updatedBooking.Seats.toString())
                                            .replace("%Вистава", this.props.fields.Title)
                                            .replace("%Дата", (new Date(this.props.fields.DateTime)).format('dd.MM.yyyy HH:mm'))
                                            .replace("%Статус", updatedBooking.Status)
                                            .replace("%Лінк", this.props.fields.Link)
                                            .replace("%Коментар", updatedBooking.Notes);

                                        sp.utility.sendEmail({
                                            To: [updatedBooking.WhoBooked.EMail],
                                            Subject: "Статус замовлення квитків змінився...",
                                            Body: body
                                        })
                                    }
                                }
                            });
                        });
                        this.props.StatusChangedCallback(this.props.fields.id.toString());

                    })
                    .catch((err1: any) => {
                        // etag invalid!
                        this.load().then(bookings => {
                            this.setState({ error: "Дані, можливо, змінилися. Спробуйте перезавантажити таблицю!" });
                            this.props.MessageCallback("", "Дані, можливо, змінилися. Спробуйте перезавантажити таблицю!", "");
                        });
                    });
            }
        });
    }
    */

    public _onOrderTicketsClicked = (attempt: number): void => {

        this.setState({ error: "" });
        this.props.MessageCallback("", "", "");

        // etag ok!
        var Ordered = 0;
        var Rejected = 0;
        var Free = this.props.fields.Seats;
        this.state.bookingArr.map(booking => {
            if (booking.Status !== "Відхилено") {
                Free -= booking.Seats;
                Ordered += booking.Seats;
            }
            else {
                Rejected += booking.Seats;
            }
        });
        var to_order = Math.min(Free, this.props.params.ticketCount);

        var months: number = this.props.params.monthCount;

        sp.web.lists.getByTitle("Booking").items
            .select("ID", "Title", "Play/ID", "Play/DateTime", "WhoBooked/ID", "WhoBooked/Title", "WhoBooked/Name", "WhoBooked/EMail", "Seats", "Status", "Notes", "Author/ID", "Author/Title")
            .expand("WhoBooked", "Play", "Author")
            .filter("WhoBooked/ID eq " + this.props.user["Id"] + " and Play/DateTime ge datetime'" + this.monthAdd(new Date(), -months).toISOString() + encodeURI("' and Status ne 'Відхилено'"))
            .get()
            .then((items: any[]) => {
                if (items.length > 0) {
                    this.setState({ error: "За заданний період (" + months + " місяців) вже є замовлення..." });
                    this.props.MessageCallback("", "За заданний період (" + months + " місяців) вже є замовлення...", "");
                }
                else {

                        // Ok/ No orders in period 
                    sp.web.lists.getByTitle("Plays").items.getById(this.props.fields.id)
                        .update({ Title: this.props.fields.Title }, this.state.etag)
                        .then((res1: ItemUpdateResult) => {
                            this.setState({ etag: res1.data["odata.etag"] });
                            sp.web.lists.getByTitle("Booking").items
                                .add({ PlayId: this.props.fields.id, Seats: to_order, WhoBookedId: this.props.user["Id"] })
                                .then((res2: ItemAddResult) => {

                                    this.load();
                                    this.props.BookingChangedCallback(this.props.fields.id.toString());

                                    // Параметри: % Замовлено, % Вистава, % Статус, % Дата, % Лінк", %Замовник
                                    var user_body: string = this.props.params.TicketsOrderedEmailText
                                        .replace("%Замовлено", to_order.toString())
                                        .replace("%Вистава", this.props.fields.Title)
                                        .replace("%Дата", (new Date(this.props.fields.DateTime)).format('dd.MM.yyyy HH:mm'))
                                        .replace("%Статус", "В очікуванні")
                                        .replace("%Лінк", this.props.fields.Link)
                                        .replace("%Замовник", this.props.user["Title"]);

                                    var user_subject: string = this.props.params.TicketsOrderedEmailSubject
                                        .replace("%Замовлено", to_order.toString())
                                        .replace("%Вистава", this.props.fields.Title)
                                        .replace("%Дата", (new Date(this.props.fields.DateTime)).format('dd.MM.yyyy HH:mm'))
                                        .replace("%Статус", "В очікуванні")
                                        .replace("%Лінк", this.props.fields.Link)
                                        .replace("%Замовник", this.props.user["Title"]);

                                    if (!this.props.params.disableNotifications) {
                                        sp.utility.sendEmail({
                                            To: [this.props.user["Email"]],
                                            Subject: user_subject,
                                            Body: user_body
                                        })
                                    }

                                    var manager_body: string = this.props.params.TicketsOrderedManagerEmailText
                                        .replace("%Замовлено", to_order.toString())
                                        .replace("%Вистава", this.props.fields.Title)
                                        .replace("%Дата", (new Date(this.props.fields.DateTime)).format('dd.MM.yyyy HH:mm'))
                                        .replace("%Статус", "В очікуванні")
                                        .replace("%Лінк", this.props.fields.Link)
                                        .replace("%Замовник", this.props.user["Title"]);

                                    var manager_subject: string = this.props.params.TicketsOrderedManagerEmailSubject
                                        .replace("%Замовлено", to_order.toString())
                                        .replace("%Вистава", this.props.fields.Title)
                                        .replace("%Дата", (new Date(this.props.fields.DateTime)).format('dd.MM.yyyy HH:mm'))
                                        .replace("%Статус", "В очікуванні")
                                        .replace("%Лінк", this.props.fields.Link)
                                        .replace("%Замовник", this.props.user["Title"]);

                                    var emails: string[] = this.props.params.managers.map(manager => { return manager["EMail"]; });

                                    if (!this.props.params.disableNotifications) {
                                        sp.utility.sendEmail({
                                            To: emails,
                                            Subject: manager_subject,
                                            Body: manager_body
                                        })
                                    }

                                    this.props.MessageCallback("", "", "Замовлено " + to_order + " квитки на виставу '" + this.props.fields.Title + "'. Дата вистави - " + (new Date(this.props.fields.DateTime)).format('dd.MM.yyyy') + ".");
                                });
                        })
                        .catch((err1: any) => {
                            if (Free > to_order && attempt < 10) {
                                this.props.OrderAttemptFailed(this.props.fields.id.toString(), attempt);
                            }
                            else {
                                // etag invalid!
                                this.load().then(bookings => {
                                    this.setState({ error: "Не вдалося забронювати квитки." });
                                    this.props.MessageCallback("", "Не вдалося забронювати квитки.", "");
                                });
                            }
                        });
                }
            });
    }

    _onDeleteBookingClicked(bookingId: string) {
        this.setState({ error: "" });
        this.props.MessageCallback("", "", "");

        var foundBookings = this.state.bookingArr.filter(booking => { return booking.ID == bookingId });
        if (foundBookings.length > 0) {
            sp.web.lists.getByTitle("Booking").items.getById(parseInt(bookingId, 10))
                .delete(foundBookings[0].etag)
                .then(res => {
                    this.load().then((bookings: any[]) => {
                        if (this.state.isBookingsCalloutVisible) {
                            this.setState({ isBookingsCalloutVisible: (bookings.length > 0) });
                        }
                    });
                    this.props.BookingChangedCallback(this.props.fields.id.toString());
                })
                .catch((err1: any) => {
                    this.setState({ error: "Дані, можливо, змінилися. Спробуйте перезавантажити таблицю!" });
                    this.props.MessageCallback("", "Дані, можливо, змінилися. Спробуйте перезавантажити таблицю!", "");
                });
        }
    }

    //BookingChangedCallback() {
    //    this.load();
    //}

    /*
    private sendEmail(from, to, body, subject): Promise<any> {
        //Get the relative url of the site
        return new Promise<any>((resolve, reject) => {
            var urlTemplate = "http://sp.kyivstar.ua/_api/SP.Utilities.Utility.SendEmail";
            $.ajax({
                contentType: 'application/json',
                url: urlTemplate,
                type: "POST",
                data: JSON.stringify({
                    'properties': {
                        '__metadata': {
                            'type': 'SP.Utilities.EmailProperties'
                        },
                        'From': from,
                        'To': {
                            'results': to
                        },
                        'Body': body,
                        'Subject': subject
                    }
                }),
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "content-type": "application/json;odata=verbose",
                    "X-RequestDigest": $("#__REQUESTDIGEST").val().toString()
                },
                success: resolve,
                error: reject
            });
        });

    }
    */


}