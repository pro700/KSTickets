"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    }
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
require("core-js/modules/es6.array.iterator.js");
require("core-js/modules/es6.array.from.js");
require("whatwg-fetch");
require("es6-map/implement");
require("raf/polyfill");
require("core-js/es6/map");
require("core-js/es6/set");
require("core-js/es6/promise");
//import * as moment from 'moment';
//import 'moment/locale/uk';
// These are the references to the react library
var React = require("react");
var ReactDOM = require("react-dom");
var Label_1 = require("office-ui-fabric-react/lib/Label");
var Button_1 = require("office-ui-fabric-react/lib/Button");
var Callout_1 = require("office-ui-fabric-react/lib/Callout");
var Utilities_1 = require("office-ui-fabric-react/lib/Utilities");
//import 'office-ui-fabric-react/src/components/List/examples/List.Basic.Example.scss';
var sp_1 = require("@pnp/sp");
//import { SiteUser } from '@pnp/sp/src/siteusers';
require("./index.scss");
var Styling_1 = require("office-ui-fabric-react/lib/Styling");
var TicketTable_1 = require("./TicketTable");
/**
 * Demo Component
 */
var Demo = /** @class */ (function (_super) {
    __extends(Demo, _super);
    function Demo() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    // Method to render the component
    Demo.prototype.render = function () {
        return (React.createElement("div", null,
            React.createElement(Label_1.Label, null, "Office Fabric React Demo"),
            React.createElement(Label_1.Label, { disabled: true }, this.props.customMessage)));
    };
    return Demo;
}(React.Component));
exports.Demo = Demo;
/**
 * TicketItem Component
 */
var TicketItem = /** @class */ (function (_super) {
    __extends(TicketItem, _super);
    function TicketItem() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    // Method to render the component
    TicketItem.prototype.render = function () {
        return (React.createElement("div", null,
            React.createElement("h3", null, "\u0421\u043F\u0438\u0441\u043E\u043A \u043A\u0432\u0438\u0442\u043A\u0456\u0432")));
    };
    return TicketItem;
}(React.Component));
exports.TicketItem = TicketItem;
var TicketList = /** @class */ (function (_super) {
    __extends(TicketList, _super);
    function TicketList(props) {
        var _this = _super.call(this, props) || this;
        _this._menuButtonElement = React.createRef();
        // Use getId() to ensure that the callout label and description IDs are unique on the page.
        // (It's also okay to use plain strings without getId() and manually ensure their uniqueness.)
        _this._labelId = Utilities_1.getId('callout-label');
        _this._descriptionId = Utilities_1.getId('callout-description');
        _this._onDeleteBookingClicked = function (event) {
            var target = event.target;
            sp_1.sp.web.lists.getByTitle("Booking").items.getById(target.id)
                .delete()
                .then(function (res) {
                var newitems = JSON.parse(JSON.stringify(_this.state.items));
                newitems = newitems.map(function (item) {
                    item.bookings = item.bookings.filter(function (booking) {
                        return ":" + booking.ID !== ":" + target.id;
                    });
                    return item;
                });
                _this.setState({
                    items: newitems,
                    isCalloutVisible: false,
                    calloutTarget: null,
                    detailsCalloutItem: null
                });
            });
        };
        _this._onOrderTicketsClicked = function (event) {
            _this.state.items.forEach(function (item) {
                if (item.play.ID == event.target.id) {
                    sp_1.sp.web.lists.getByTitle("Plays").items.getById(item.play.ID)
                        .update({ Title: item.Title }, item.etag)
                        .then(function (updated_item) {
                        // etag ok!
                        sp_1.sp.web.currentUser.get()
                            .then(function (user) {
                            sp_1.sp.web.lists.getByTitle("Booking").items
                                .add({ PlayId: item.play.ID, Seats: 2, WhoBookedId: user["Id"] })
                                .then(function (result) {
                                _this.populate();
                            });
                        });
                    });
                }
            });
        };
        _this._onShowMenuClicked = function (event) {
            _this.state.items.forEach(function (item) {
                if (item.ID == event.target.id) {
                    _this.setState({
                        isCalloutVisible: !_this.state.isCalloutVisible,
                        calloutTarget: event.target,
                        detailsCalloutItem: item
                    });
                }
            });
        };
        _this._onCalloutDismiss = function (event) {
            _this.setState({
                isCalloutVisible: false,
                calloutTarget: event.target,
                detailsCalloutItem: null
            });
        };
        _this.state = { ID: "Loading", Title: "Loading...", isCalloutVisible: false };
        _this.populate();
        return _this;
    }
    TicketList.prototype.render = function () {
        var _this = this;
        var theme = Styling_1.getTheme();
        var state = this.state.items;
        var styles = Styling_1.mergeStyleSets({
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
                    '& tr.booking td': {},
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
                background: Styling_1.DefaultPalette.themeLighterAlt,
            },
            itemClass: {
                background: Styling_1.DefaultPalette.themeLight,
                padding: 4,
                selectors: {
                    ':hover': {
                        background: Styling_1.DefaultPalette.themeLight
                    },
                    ':focus': {
                        background: Styling_1.DefaultPalette.themeLight
                    },
                    ':active': {
                        background: Styling_1.DefaultPalette.themeLight
                    }
                }
            }
        });
        if (this.state.ID == "table") {
            return (React.createElement("div", null,
                this.state.Message &&
                    React.createElement("p", null, this.state.Message),
                React.createElement("table", { className: styles.tableClass },
                    React.createElement("tbody", null,
                        React.createElement("tr", { key: "0" },
                            React.createElement("th", null, "\u0427\u0430\u0441"),
                            React.createElement("th", null, "\u041D\u0430\u0437\u0432\u0430"),
                            React.createElement("th", null),
                            React.createElement("th", null, "\u0414\u043E\u0441\u0442\u0443\u043F\u043D\u043E"),
                            React.createElement("th", null, "\u0417\u0430\u043C\u043E\u0432\u043B\u0435\u043D\u043E"),
                            React.createElement("th", null, "\u041A\u043E\u043C\u0435\u043D\u0442\u0430\u0440"),
                            React.createElement("th", null, "etag")),
                        this.state.items.map(function (item) {
                            var Ordered = 0;
                            var Free = item.play.Seats;
                            item.bookings.map(function (booking) {
                                Free -= booking.Seats;
                                Ordered += booking.Seats;
                            });
                            return (React.createElement("tr", { key: item.play.ID },
                                React.createElement("td", null, (new Date(item.play.DateTime)).format('dd.MM.yyyy')),
                                React.createElement("td", null,
                                    React.createElement("a", { href: item.play.Link }, item.play.Title)),
                                React.createElement("td", null, (Free > 0) &&
                                    React.createElement(Button_1.DefaultButton, { className: styles.buttonClass, id: item.ID, onClick: _this._onOrderTicketsClicked, text: "\u0417\u0430\u043C\u043E\u0432\u0438\u0442\u0438" })),
                                React.createElement("td", { align: "center", className: 'free' }, (Free > 0) && Free),
                                React.createElement("td", { align: "center", className: 'ordered' }, (Ordered > 0) &&
                                    React.createElement(Button_1.DefaultButton, { className: styles.buttonClass + " orderDetailsButton", id: item.ID, onClick: _this._onShowMenuClicked, text: Ordered + " ...", title: '\u0414\u0435\u0442\u0430\u043B\u0456 \u0437\u0430\u043C\u043E\u0432\u043B\u0435\u043D\u044C' })),
                                React.createElement("td", null, item.play.Comments),
                                React.createElement("td", null, item.play.etag)));
                        }))),
                this.state.isCalloutVisible &&
                    React.createElement(Callout_1.Callout, { className: "ms-CalloutExample-callout", ariaLabelledBy: this._labelId, ariaDescribedBy: this._descriptionId, role: "alertdialog", gapSpace: 0, target: this.state.calloutTarget, onDismiss: this._onCalloutDismiss, setInitialFocus: true, hidden: !this.state.isCalloutVisible },
                        React.createElement("div", null, this.state.detailsCalloutItem && this.state.detailsCalloutItem.bookings.length > 0 && (React.createElement("table", { className: styles.tableClass + ' ' + styles.tableClass1 },
                            React.createElement("tbody", null,
                                " ",
                                this.state.detailsCalloutItem.bookings.map(function (booking) {
                                    return (React.createElement("tr", { className: "booking" },
                                        React.createElement("td", null, booking.WhoBooked.Title),
                                        React.createElement("td", null, booking.Status),
                                        React.createElement("td", null, booking.Seats),
                                        React.createElement("td", null, booking.WhoBooked.ID == _this.state.user["Id"] &&
                                            React.createElement(Button_1.DefaultButton, { className: styles.buttonClass + " " + styles.buttonClassDel, title: "\u0412\u0438\u0434\u0430\u043B\u0438\u0442\u0438 \u0437\u0430\u043C\u043E\u0432\u043B\u0435\u043D\u043D\u044F", id: booking.ID, onClick: _this._onDeleteBookingClicked, text: " \u0425 " }))));
                                }))))))));
        }
        else if (this.state.ID == "error") {
            return React.createElement("p", null,
                "Error:",
                this.state.Title);
        }
        else {
            return React.createElement("p", null, this.state.Title);
        }
    };
    TicketList.prototype._onRenderCell = function (item, index) {
        return item;
    };
    //setUsernameState = (event) => {
    //    this.setState({ userName: event.target.value });
    //}
    TicketList.prototype.populate = function () {
        var _this = this;
        var today = new Date();
        sp_1.sp.web.currentUser.get().then(function (user) {
            sp_1.sp.web.lists.getByTitle("Plays").items
                .select("ID", "Title", "Seats", "DateTime", "Link", "Comments")
                .filter("DateTime ge datetime'" + today.toISOString() + "'")
                .orderBy("DateTime", false)
                .get()
                .then(function (plays) {
                var promises = plays.map(function (play) {
                    return sp_1.sp.web.lists.getByTitle("Booking").items
                        .select("ID", "Title", "Play/ID", "WhoBooked/ID", "WhoBooked/Title", "Seats", "Status", "Notes", "GivenAway", "Author/ID", "Author/Title")
                        .expand("WhoBooked", "Play", "Author")
                        .filter("Play/ID eq " + play["ID"])
                        .get();
                });
                Promise.all(promises).then(function (allbookings) {
                    var items = plays.map(function (play, i) {
                        var bookings = allbookings[i].map(function (booking) {
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
                    _this.setState({ ID: "table", Title: "Список вистав", user: user, items: items });
                }).catch(function (err) {
                    _this.setState({ ID: "error", Title: "populate err 2: " + JSON.stringify(err), user: user });
                });
            })
                .catch(function (err) {
                _this.setState({ ID: "error", Title: "populate err 1: " + JSON.stringify(err), user: user });
            });
        });
    };
    return TicketList;
}(React.Component));
exports.TicketList = TicketList;
// Get the "main" element
var target = document.querySelector("#tickets_conainer");
if (target) {
    ReactDOM.render(React.createElement(TicketTable_1.TicketTable, null), target);
    //var user: SiteUser = sp.web.getUserById(0);
    //var web = new Web("http://intranet.kyivstar.ua");
    /*
    sp.web.siteUsers.get()
        .then((users: any[]) => {
            return <table>
                <tbody>
                    { users.map(user => {
                        return <tr><td>{user.Id}</td><td>{user.Title}</td><td>{user.Email}</td><td>{user.LoginName}</td></tr>;
                    })}
                </tbody>
            </table>;
        })
        .catch((err: object) => {
            return <div>err={JSON.stringify(err)}</div>;
        })
        .then(html => {
            ReactDOM.render(<div><Demo customMessage="SiteUsers:" /> {html} </div>, target);
        });
        */
    // , "DateTime", "Link", "Seats", "Comments", "Type", "WhoBooking", "WhoCancelBooking", "Period", "Idx", "_UIVersionString"
}
