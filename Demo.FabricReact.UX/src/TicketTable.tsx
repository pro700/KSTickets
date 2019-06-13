import * as React from "react";
import * as ReactDOM from "react-dom";

import 'jquery';
//import 'microsoft-ajax';
//import 'sharepoint';


import { sp, Web, ItemAddResult, ItemUpdateResult } from '@pnp/sp';
import { CurrentUser } from "@pnp/sp/src/siteusers";
import { PermissionKind, BasePermissions } from "@pnp/sp";


import { TicketTableRow, TicketTableRowProps, TicketTableRowMouseEventCode, TicketTableRowBooking } from "./TicketTableRow";
import { MessageBar, MessageBarType, initializeIcons, Stack, DefaultButton, PrimaryButton, getId, mergeStyleSets, getTheme, Dropdown, IDropdownOption, TextField, Label, CompactPeoplePicker, IPersonaProps, IBasePicker, IBasePickerSuggestionsProps, PersonaPresence, Pivot, PivotItem, IRefObject, Link, Dialog, DialogFooter, DialogType, SearchBox, Text } from "office-ui-fabric-react";
import { IDatePicker, DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { SpinButton } from 'office-ui-fabric-react/lib/SpinButton';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { CommandBar, ICommandBarItemProps } from 'office-ui-fabric-react/lib/CommandBar';
import { CommandBarButton, IButtonProps, IconButton } from 'office-ui-fabric-react/lib/Button';
import { DirectionalHint } from 'office-ui-fabric-react/lib/Callout';


import './TicketTable.scss';

//require('sp-init');
//require('microsoft-ajax');
//require('sp-runtime');
//require('sharepoint');

// Register icons and pull the fonts from the default SharePoint cdn.
initializeIcons();

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


export enum TicketTableRoles {
    user,
    manager,
    administartor
};

const DayPickerStrings: IDatePickerStrings = {
    months: ['Січень', 'Лютий', 'Березень', 'Квітень', 'Травень', 'Червень', 'Липень', 'Серпень', 'Вересень', 'Жовтень', 'Листопад', 'Грудень'],
    shortMonths: ['Січ', 'Лют', 'Бер', 'Кві', 'Тра', 'Чер', 'Лип', 'Сер', 'Вер', 'Жов', 'Лис', 'Гру'],
    days: ['Неділя', 'Понеділок', 'Вівторок', 'Середа', 'Четверг', 'П\'ятниця', 'Субота'],
    shortDays: ['Нд', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'],
    goToToday: 'Сьогодні',
    prevMonthAriaLabel: 'Попередній місяць',
    nextMonthAriaLabel: 'Наступний місяць',
    prevYearAriaLabel: 'Попередній рік',
    nextYearAriaLabel: 'Наступний рік',
    closeButtonAriaLabel: 'Закрити'
};

const suggestionProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: 'Запропонований вибір',
    mostRecentlyUsedHeaderText: 'Запропоновані контакти',
    noResultsFoundText: 'Результати не здайдено',
    loadingText: 'Завантаження...',
    showRemoveButtons: true,
    suggestionsAvailableAlertText: 'People Picker Suggestions available',
    suggestionsContainerAriaLabel: 'Запропоновані контакти'
};

export interface TicketTableParams {
    id: number,
    publishDate: Date,
    monthCount: number,
    ticketCount: number,
    managers?: any[],
    StatusChangedEmailSubject?: string,
    StatusChangedRejectedEmailSubject?: string,
    TicketsOrderedEmailSubject?: string,
    TicketsOrderedManagerEmailSubject?: string,
    StatusChangedEmailText?: string,
    StatusChangedRejectedEmailText?: string,
    TicketsOrderedEmailText?: string,
    TicketsOrderedManagerEmailText?: string,
    TempEmailSubject1?: string,
    TempEmailSubject2?: string,
    TempEmailSubject3?: string,
    TempEmailSubject4?: string,
    TempEmailText1?: string,
    TempEmailText2?: string,
    TempEmailText3?: string,
    TempEmailText4?: string,
    TempPublishDate?: Date,
    managersId?: { results: any[] },
    disableNotifications?: boolean
}

export interface TicketTableMyBookings {
    "Id",
    "Title",
    "Play",
    "WhoBooked",
    "Seats",
    "Status",
    "Notes",
    "Author"
}

export interface TicketTableProps {
}

export enum TicketTableMode { plays=1, allbookings=2 }

export interface TicketTableState {
    Message?: string;
    SuccessMessage?: string;
    ErrorMessage?: string;
    user?: CurrentUser,
    bookingPerm?: BasePermissions,
    playsPerm?: BasePermissions,
    RowPropsCollection: TicketTableRowProps[];
    RowPropsCollectionArch: TicketTableRowProps[];
    showPanel: boolean,
    params: TicketTableParams,
    myBookings: TicketTableMyBookings[],
    allBookings: TicketTableMyBookings[],
    isManagerForm: boolean,
    mode: TicketTableMode,
    dialogBookingId?: any,
    dialogBookingOrgStruct?: string,
    dialogBookingSeats?: number,
    dialogBookingStatus?: string,
    dialogBookingNotes?: string,
    filter: string,
    allbookingsFromDate: Date,
    playsFromDate: Date
}

export class TicketTable extends React.Component<TicketTableProps, TicketTableState> {

    private _dafaultParams: TicketTableParams;
    private _publishDatePickerRef: IRefObject<IDatePicker>;
    private _spinButtonMonthRef = React.createRef<SpinButton>();
    private _spinButtonTicketRef = React.createRef<SpinButton>();
    private _ticketSpinnerDivId = getId("spinnerDiv");
    private _monthSpinnerDivId = getId("spinnerDiv");
    private _publishDateDivId = getId("DatePickerDiv");

    //private _panelPublishDate;
    private _panelTicketCount;
    private _panelMonthCount;
    private tdOrderdRef = React.createRef<HTMLTableDataCellElement>();
    private _peoplePicker = React.createRef<IBasePicker<IPersonaProps>>();


    constructor(props: TicketTableProps) {
        super(props);
        // Init state
        //this._panelPublishDate = new Date();
        this._panelTicketCount = 2;
        this._panelMonthCount = 6;
        this._dafaultParams = {
            id: -1,
            publishDate: new Date(),
            monthCount: 6,
            ticketCount: 2,
            disableNotifications: false,
            managers: []
        };
        this.state = {
            RowPropsCollection: [],
            RowPropsCollectionArch: [],
            showPanel: false,
            params: this._dafaultParams,
            myBookings: [],
            allBookings: [],
            isManagerForm: false,
            mode: TicketTableMode.plays,
            filter: "",
            allbookingsFromDate: new Date(),
            playsFromDate: new Date()
        };
    }

    render(): JSX.Element {
        return (
            <div className="TicketTableDiv">
                { this.state.playsPerm && sp.web.hasPermissions(this.state.playsPerm, PermissionKind.ManagePermissions) && 
                    <div>
                        <CommandBar items={this.getCommandBarItems()} styles={{ root: { marginBottom: '10px' } }}/>

                    <Panel
                        isOpen={this.state.showPanel}
                        type={PanelType.smallFixedFar}
                        onDismiss={this._hidePanel}
                        headerText="Параметри"
                        closeButtonAriaLabel="Close"
                    //onRenderFooterContent={this._onRenderFooterContent}
                    >
                        <Stack padding={10} gap={10}>
                            <div style={{ width: '150px' }} id={this._publishDateDivId}>
                                <DatePicker
                                    label="Дата публікації"
                                    firstDayOfWeek={DayOfWeek.Monday}
                                    isRequired={false}
                                    allowTextInput={true}
                                    strings={DayPickerStrings}
                                    value={this.state.params.publishDate}
                                    formatDate={(date?: Date) => { return date ? date.format('dd.MM.yyyy HH:mm') : ""; }}
                                    parseDateFromString={(str: string): Date => { return this._parseDateFromString(str); }}
                                    //onSelectDate={(date: Date) => {
                                    //    const { id, publishDate, monthCount, ticketCount } = this.state.params;
                                    //    this.setState({ params: { id, ticketCount, monthCount, publishDate: date } });
                                    //}}
                                    //onChange={(event: React.FormEvent<HTMLElement>) => {
                                    //    this.setState({ Message: $("input", event.currentTarget).attr("value") })
                                    //}}
                                />
                            </div>
                            <div style={{ width: '250px' }} id={this._ticketSpinnerDivId}>
                                <SpinButton
                                    ref={this._spinButtonTicketRef}
                                    value={this.state.params.ticketCount.toString()}
                                    label={'Кількість квитків на 1 замовлення:'}
                                    min={1}
                                    max={10}
                                    step={1}
                                />
                            </div>
                            <div style={{ width: '250px' }} id={this._monthSpinnerDivId}>
                                <SpinButton
                                    ref={this._spinButtonMonthRef}
                                    value={this.state.params.monthCount.toString()}
                                    label={'Кількість місяців на 1 замовлення:'}
                                    min={1}
                                    max={24}
                                    step={1}
                                />
                            </div>
                            <div>
                                <Label> Менеджери-адресати листів</Label>
                                <CompactPeoplePicker
                                    onResolveSuggestions={ this._onPeopleFilterChanged }
                                    //onEmptyInputFocus={this._returnMostRecentlyUsed}
                                    //getTextFromItem={this._getTextFromItem}
                                    className={'ms-PeoplePicker'}
                                    defaultSelectedItems={this.state.params.managers.map(manager => { return this._manager2persona(manager); })}
                                    key={'list'}
                                    pickerSuggestionsProps={suggestionProps}
                                    //onRemoveSuggestion={this._onRemoveSuggestion}
                                    //onValidateInput={this._validateInput}
                                    inputProps={{
                                        onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'),
                                        onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'),
                                        'aria-label': 'People Picker'
                                    }}
                                    componentRef={this._peoplePicker}
                                    resolveDelay={300}
                                />
                            </div>

                            <div style={{ width: '250px' }}>
                                <Checkbox label="Відмінити нотифікації" checked={this.state.params.disableNotifications} onChange={(ev, checked: boolean) => {
                                    //const params = this.state.params;
                                    //params.disableNotifications = !params.disableNotifications;
                                    //this.setState({ params: params });
                                    this.setState({ params: $.extend(this.state.params, { disableNotifications: !this.state.params.disableNotifications }) });
                                }} />
                                <Label>Параметри для листів: %Замовлено, %Вистава, %Статус, %Дата, %Лінк, %Коментар, %Замовник"</Label>
                            </div>
                            <div style={{ width: '250px' }}>
                                <Pivot>
                                    <PivotItem headerText="Лист 1">
                                        <TextField label="Тема Листа"
                                            defaultValue={this.state.params.TempEmailSubject1}
                                            onChange={(event, newvalue) => {
                                                var params = this.state.params;
                                                params.TempEmailSubject1 = newvalue;
                                                this.setState({ params: params });
                                            }}
                                        />
                                        <TextField label="Лист замовнику при замовленні"
                                            multiline rows={6}
                                            defaultValue={this.state.params.TempEmailText1}
                                            resizable={false}
                                            onChange={(event, newvalue) => {
                                                var params = this.state.params;
                                                params.TempEmailText1 = newvalue;
                                                this.setState({ params: params });
                                            }}
                                        />
                                    </PivotItem>
                                    <PivotItem headerText="Лист 2">
                                        <TextField label="Тема Листа"
                                            defaultValue={this.state.params.TempEmailSubject2}
                                            onChange={(event, newvalue) => {
                                                var params = this.state.params;
                                                params.TempEmailSubject2 = newvalue;
                                                this.setState({ params: params });
                                            }}
                                        />
                                        <TextField label="Лист менеджеру при замовленні"
                                            multiline rows={6}
                                            defaultValue={this.state.params.TempEmailText2}
                                            resizable={false}
                                            onChange={(event, newvalue) => {
                                                var params = this.state.params;
                                                params.TempEmailText2 = newvalue;
                                                this.setState({ params: params });
                                            }}
                                        />
                                    </PivotItem>
                                    <PivotItem headerText="Лист 3">
                                        <TextField label="Тема Листа"
                                            defaultValue={this.state.params.TempEmailSubject3}
                                            onChange={(event, newvalue) => {
                                                var params = this.state.params;
                                                params.TempEmailSubject3 = newvalue;
                                                this.setState({ params: params });
                                            }}
                                        />
                                        <TextField label="Лист замовнику при зміні статуса замовлення"
                                            multiline rows={6}
                                            defaultValue={this.state.params.TempEmailText3}
                                            resizable={false}
                                            onChange={(event, newvalue) => {
                                                var params = this.state.params;
                                                params.TempEmailText3 = newvalue;
                                                this.setState({ params: params });
                                            }}
                                        />
                                    </PivotItem>
                                    <PivotItem headerText="Лист 4">
                                        <TextField label="Тема Листа"
                                            defaultValue={this.state.params.TempEmailSubject4}
                                            onChange={(event, newvalue) => {
                                                var params = this.state.params;
                                                params.TempEmailSubject4 = newvalue;
                                                this.setState({ params: params });
                                            }}
                                        />
                                        <TextField label="Лист замовнику при зміні статуса замовлення на 'Відмовлено'"
                                            multiline rows={6}
                                            defaultValue={this.state.params.TempEmailText4}
                                            resizable={false}
                                            onChange={(event, newvalue) => {
                                                var params = this.state.params;
                                                params.TempEmailText4 = newvalue;
                                                this.setState({ params: params });
                                            }}
                                        />
                                    </PivotItem>
                                </Pivot>

                            </div>
                            <div style={{ marginTop: '50px' }}>
                                <PrimaryButton onClick={this._saveParams} style={{ marginRight: '8px' }}> Save </PrimaryButton>
                                <DefaultButton onClick={this._hidePanel}> Cancel </DefaultButton>
                            </div>
                        </Stack>
                    </Panel>
                    </div>
                }
                {this.state.Message &&
                    <MessageBar className="MessageBar" isMultiline={true} onDismiss={x => { this.setState({ Message: "" }) }} dismissButtonAriaLabel="Close" >
                        {this.state.Message}
                    </MessageBar>
                }
                {this.state.SuccessMessage &&
                    <MessageBar className="MessageBar" messageBarType={MessageBarType.success} isMultiline={true} onDismiss={x => { this.setState({ SuccessMessage: "" }) }} dismissButtonAriaLabel="Close" >
                        {this.state.SuccessMessage}
                    </MessageBar>
                }
                {this.state.ErrorMessage &&
                    <MessageBar className="MessageBar" messageBarType={MessageBarType.error} isMultiline={true} onDismiss={x => { this.setState({ ErrorMessage: "" }) }} dismissButtonAriaLabel="Close" >
                        {this.state.ErrorMessage}
                    </MessageBar>
                }
                {this.state.mode == TicketTableMode.plays &&
                    < div >
                        <div style={{ display: "inline-flex" }}>
                            <h3>Вистави:</h3>
                        </div>
                        <table className="TicketTable">
                        <tbody>
                                <TicketTableRow key="0" isHeader={true} user={this.state.user} bookingPerm={null} playsPerm={null} params={this.state.params} bookingArr={[]}/>
                                {this.state.RowPropsCollection.map(rowProps => {
                                    return <TicketTableRow
                                        BookingChangedCallback={this.BookingChangedCallback.bind(this)}
                                        StatusChangedCallback={this.StatusChangedCallback.bind(this)}
                                        OrderAttemptFailed={this.OrderAttemptFailed.bind(this) }
                                        MessageCallback={this.MessageCallback.bind(this)}
                                        ref={rowProps.rowRef}
                                        key={rowProps.key}
                                        fields={rowProps.fields}
                                        user={this.state.user}
                                        bookingPerm={rowProps.bookingPerm}
                                        playsPerm={rowProps.playsPerm}
                                        params={this.state.params}
                                        isMangerForm={this.state.isManagerForm}
                                        mode={this.state.mode}
                                        bookingArr={rowProps.bookingArr}
                                    />;
                                })}
                            </tbody>
                        </table>
                        { this.state.myBookings.length > 0 &&
                            <div>
                                <br />
                                <br />
                                <h3>Мої замовлення:</h3>
                                <table className="TicketTable">
                                    <tbody>
                                        <tr key={-1}>
                                            <th>Дата</th>
                                            <th>Назва</th>
                                            <th>Замовник</th>
                                            <th>Місць</th>
                                            <th>Статус</th>
                                            <th></th>
                                            <th>Коментар</th>
                                        </tr>
                                        {this.state.myBookings.map((myBooking, index) => {
                                            return (
                                                <tr key={myBooking.Id}>
                                                    <td>{(new Date(myBooking.Play["DateTime"])).format('dd-MM-yyyy HH:mm')}</td>
                                                    <td><a href={myBooking.Play["Link"]}>{myBooking.Play["Title"]}</a></td>
                                                    <td>{myBooking.WhoBooked["Title"]}</td>
                                                    <td>{myBooking.Seats}</td>
                                                    <td>{myBooking.Status}</td>
                                                    <td>{myBooking.WhoBooked["ID"] == this.state.user["Id"] &&
                                                        myBooking.Status == "В очікуванні" &&
                                                        <DefaultButton className={styles.buttonClassDel + " DelButton"}
                                                            title={"Видалити замовлення"}
                                                            id={myBooking.Id}
                                                            onClick={() => { this._onDeleteBookingClicked(myBooking.Id) }}
                                                            text=" Х " />}
                                                    </td>
                                                    <td>{myBooking.Notes}</td>
                                                </tr>);
                                        })}
                                    </tbody>
                                </table>
                            </div>
                        }
                    </div>
                }
                {this.state.mode == TicketTableMode.allbookings &&
                    <div>
                        <div style={{ display: "inline-flex" }}>
                            <h3>Всі замовлення:</h3>
                            <h3 style={{ marginLeft: 50 }}>починаючи з дати вистави</h3>
                            <DatePicker
                                style={{ marginLeft: 10 }}
                                firstDayOfWeek={DayOfWeek.Monday}
                                isRequired={false}
                                allowTextInput={false}
                                strings={DayPickerStrings}
                                formatDate={(date?: Date) => { return date ? date.format('dd.MM.yyyy') : ""; }}
                                parseDateFromString={(str: string): Date => { return this._parseDateFromString(str); }}
                                value={this.state.allbookingsFromDate}
                                onSelectDate={(date: Date) => {
                                    this.setState({ allbookingsFromDate: date });
                                    this.populateAllBookings(date);
                                }}
                            />
                            <SearchBox
                                styles={{ root: { width: 250, marginLeft: 50 } }}
                                placeholder="пошук по назві або замовнику"
                                onChange={newValue => { this.setState({ filter: newValue }); }}
                                onSearch={newValue => { this.setState({ filter: newValue }); }}
                            />
                        </div>

                        <table className="TicketTable">
                            <tbody>
                                <tr key={-1}>
                                    <th>Дата</th>
                                    <th>Назва</th>
                                    <th>Замовник</th>
                                    <th>Місць</th>
                                    <th>Статус</th>
                                    <th>Коментар</th>
                                    <th>Всього</th>
                            </tr>
                            {this.state.allBookings
                                .filter(value => {
                                    return this.state.filter ?
                                        (new String(value.WhoBooked["Title"])).toUpperCase().indexOf(this.state.filter.toUpperCase()) >= 0 ||
                                        (new String(value.Play["Title"])).toUpperCase().indexOf(this.state.filter.toUpperCase()) >= 0 : true;
                                })
                                .map((booking, index) => {
                                    var Totals = this.getTotalsdByPlayID(booking.Play["ID"], parseInt(booking.Play["Seats"], 10));
                                    return (
                                        <tr key={booking.Id}>
                                            <td style={{ background: (new Date(booking.Play["DateTime"]) <= new Date() ? "#eaeaea" : "") }}>{(new Date(booking.Play["DateTime"])).format('dd-MM-yyyy HH:mm')}</td>
                                            <td><a href={booking.Play["Link"]}>{booking.Play["Title"]}</a></td>
                                            <td>{booking.WhoBooked["Title"]}</td>
                                            <td>{booking.Seats}</td>
                                            <td> <Link onClick={() => { this._showDialog(booking) }}>{booking.Status}</Link>
                                                {
                                                    booking.Id == this.state.dialogBookingId &&
                                                    <Dialog
                                                        minWidth={480}
                                                        maxWidth={800}
                                                        hidden={false}
                                                        onDismiss={this._closeDialog}
                                                        dialogContentProps={{
                                                            type: DialogType.largeHeader,
                                                            title: 'Редагування статусу замовлення',
                                                            styles: { header: { wordWrap: "break-word" } }
                                                        }}
                                                        modalProps={{
                                                            isBlocking: false
                                                        }}
                                                    >
                                                        <TextField label="Вистава" disabled defaultValue={booking.Play["Title"]} />
                                                        <TextField label="Замовник" disabled defaultValue={booking.WhoBooked["Title"]} />
                                                        <TextField label="Підрозділ" disabled defaultValue={this.state.dialogBookingOrgStruct} title={this.state.dialogBookingOrgStruct} multiline autoAdjustHeight={true} resizable={false} />
                                                        <Label>Статус замовлення:</Label>
                                                        <Dropdown
                                                            defaultSelectedKey={this.state.dialogBookingStatus}
                                                            options={[
                                                                { key: 'В очікуванні', text: 'В очікуванні' },
                                                                { key: 'Відхилено', text: 'Відхилено' },
                                                                { key: 'Затверджено', text: 'Затверджено' }
                                                            ]}
                                                            onChange={(event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) => {
                                                                this.setState({ dialogBookingStatus: option.text });
                                                            }}
                                                            disabled={(booking.Status == "Відхилено" && booking.Seats > Totals.Free) || new Date(booking.Play["DateTime"]) < new Date()}
                                                        />
                                                        <TextField label="У разі відхилення замовлення причину можна вказати у коментарі:"
                                                            multiline rows={6}
                                                            defaultValue={this.state.dialogBookingNotes}
                                                            resizable={false}
                                                            disabled={(booking.Status == "Відхилено" && booking.Seats > Totals.Free) || new Date(booking.Play["DateTime"]) < new Date()}
                                                            onChange={(event, newvalue) => {
                                                                this.setState({ dialogBookingNotes: newvalue });
                                                            }}
                                                        />
                                                        <DialogFooter>
                                                            <PrimaryButton onClick={() => { this._saveDialog(booking) }} text="Save" />
                                                            <DefaultButton onClick={this._closeDialog} text="Cancel" />
                                                        </DialogFooter>
                                                    </Dialog>
                                                }
                                                
                                            </td>
                                            <td>{booking.Notes}</td>
                                            <td title={booking["ApprovedDetails"]} style={{ cursor: "pointer", textDecoration: "underline" }}>{booking["Approved"]}</td>
                                        </tr>);
                                })}
                            </tbody>
                        </table>
                    </div>
                }
                <div id="TicketTableCalloutDiv">
                </div>

            </div>
        );
    }

    private _showDialog = (booking: TicketTableMyBookings): void => {
        var promiseA:  Promise<any> = sp.profiles.getPropertiesFor(booking.WhoBooked.Name);

        Promise.all([promiseA])
            .then(([result]) => {
                var orgStruct = "";
                var seats = 0;
                result.UserProfileProperties.forEach(function (prop) {
                    if (prop.Key == "OrganizationalStructure") {
                        orgStruct = (prop.Value as string).split("\\").reverse().join("; ");
                    }
                });

                this.setState({ dialogBookingId: booking.Id, dialogBookingOrgStruct: orgStruct, dialogBookingSeats: -1, dialogBookingStatus: booking.Status, dialogBookingNotes: booking.Notes, ErrorMessage: "" });
            })
            .catch(err => {
                this.setState({ dialogBookingId: booking.Id, dialogBookingOrgStruct: "?", dialogBookingSeats: -1, dialogBookingStatus: booking.Status, dialogBookingNotes: booking.Notes, ErrorMessage: err });
            });
    };

    private _closeDialog = (): void => {
        this.setState({ dialogBookingId: -1 });
    };

    private _saveDialog = (booking: TicketTableMyBookings): void => {

        this._closeDialog();

        if (this.state.dialogBookingStatus == "Відхилено" && (!this.state.dialogBookingNotes || this.state.dialogBookingNotes.trim().length == 0)) {
            this.setState({ ErrorMessage: "Потрібно вказати причину відхилення замовлення!" });
        }
        else {
            this.setState({ ErrorMessage: "" });
            this._onStatusChange(booking, this.state.dialogBookingStatus, this.state.dialogBookingNotes);
        }
    };

    private _onPeopleFilterChanged = (filterText: string, currentPersonas: IPersonaProps[], limitResults?: number): IPersonaProps[] | Promise<IPersonaProps[]> => {

        return new Promise<IPersonaProps[]>((resolve, reject) => {
            sp.web.siteUsers
                .select("Id", "Title", "LoginName", "Email")
                .filter(`substringof('${encodeURIComponent(filterText)}',Title) or substringof('${encodeURIComponent(filterText)}',LoginName)`)
                .top(limitResults ? limitResults : 10)
                .get()
                .then((siteUsers: any[]) => {
                    var allPersonas: IPersonaProps[] = siteUsers.map(siteUser => {
                        var Persona: IPersonaProps = {
                            //imageUrl: TestImages.personaFemale,
                            imageInitials: siteUser["Title"].split(" ").splice(0, 2).map((n) => n[0]).join(""),
                            text: siteUser["Title"],
                            //secondaryText: 'Designer',
                            //tertiaryText: 'In a meeting',
                            //optionalText: 'Available at 4:00pm',
                            presence: PersonaPresence.none
                        };
                        Persona["LoginName"] = siteUser["LoginName"];
                        Persona["Email"] = siteUser["Email"];
                        Persona["Id"] = siteUser["Id"];
                        return Persona;
                    })
                    resolve(this._removeDuplicates(allPersonas, currentPersonas));
                })
                .catch(err => {
                    reject(err);
                });
        });
    };

    private _removeDuplicates(personas: IPersonaProps[], possibleDupes: IPersonaProps[]) {
        return personas.filter(persona => !this._listContainsPersona(persona, possibleDupes));
    }

    private _listContainsPersona(persona: IPersonaProps, personas: IPersonaProps[]) {
        if (!personas || !personas.length || personas.length === 0) {
            return false;
        }
        return personas.filter(item => item.text === persona.text).length > 0;
    }

    private _onRenderFooterContent = () => {
        return (
            <div>
                <PrimaryButton onClick={this._saveParams} style={{ marginRight: '8px' }}> Save </PrimaryButton>
                <DefaultButton onClick={this._hidePanel}> Cancel </DefaultButton>
            </div>
        );
    };

    //private _ticketCountChanged = (ticketCount: any) => {
    //    this._dafaultParams.ticketCount = ticketCount;
    //    this.setState({ SuccessMessage: "_ticketCountChanged new this._dafaultParams=" + JSON.stringify(this._dafaultParams) });
    //    return ticketCount;
    //}

    //private _monthCountChanged = (monthCount: any) => {
    //    this._dafaultParams.monthCount = monthCount;
    //    this.setState({ SuccessMessage: "_monthCountChanged new this._dafaultParams=" + JSON.stringify(this._dafaultParams) });
    //    return monthCount;
    //}

    private getCommandBarItems = (): ICommandBarItemProps[] => {
        var commands: ICommandBarItemProps[] = 
         [
            {
                key: 'plays',
                text: "Вистави",
                canCheck: true,
                checked: (this.state.mode == TicketTableMode.plays),
                iconProps: { iconName: 'BulletedList2' },
                onClick: () => { this._setMode(TicketTableMode.plays) },
                ariaLabel: 'plays',
                title: "Вистави"
            },
            {
                key: 'allbookings',
                text: "Всі замовлення",
                canCheck: true,
                checked: (this.state.mode == TicketTableMode.allbookings),
                iconProps: { iconName: 'IssueTracking' },
                onClick: () => { this._setMode(TicketTableMode.allbookings) },
                ariaLabel: 'allbookings',
                title: "Всі замовлення"
            },
            {
                key: 'parameters',
                name: 'Параметри...',
                iconProps: { iconName: 'Settings' },
                ariaLabel: 'parameters',
                onClick: this._showPanel,
                title: "Відкрити панель параметрів"
            },
            {
                key: 'import',
                name: 'Імпорт та редагування...',
                iconProps: { iconName: 'Upload' },
                ariaLabel: 'import',
                href: "/Lists/Plays/Import.aspx",
                title: "Перейти в режим імпорту та редагування"
            }
            ];
        /*
        if ((new String(this.state.user["Title"])).toLowerCase().indexOf("sokol") >= 0 ||
            (new String(this.state.user["Title"])).toLowerCase().indexOf("сокол") >= 0) {
            commands.push({
                key: 'load',
                name: 'Load',
                iconProps: { iconName: 'Upload' },
                ariaLabel: 'load',
                onClick: this._load,
                title: "Load"
            });
        }
        */
        return commands;
    };

    private _load = () => {

        var promise1: Promise<any[]> = sp.web.lists.getByTitle("tb00").items
            .select("ID", "Title", "DateTime", "Link", "Seats", "Comments")
            .orderBy("ID", false)
            .getAll();

        var promise2: Promise<any[]> = sp.web.lists.getByTitle("tb10").items
            .select("ID", "Title", "BookingID", "status", "Created", "Author/ID", "Author/Title", "Author/Name", "Author/EMail", "OData__x0411__x0440__x043e__x043d__x04")
            .expand("Author")
            .orderBy("BookingID", false)
            .getAll();

        Promise.all([promise1, promise2])
            .then(([intranetPlayItems, intranetBookingItems]: [any[], any[]]) => {

                intranetPlayItems.forEach(intranetPlayItem => {

                    var subBookingItems = intranetBookingItems
                        .filter(intranetBookingItem => {
                            return intranetBookingItem["BookingID"] == intranetPlayItem["ID"];
                        });

                    sp.web.lists.getByTitle("plays").items.add({
                        "Title": intranetPlayItem["Link"]["Description"],
                        "Link": intranetPlayItem["Link"]["Url"],
                        "DateTime": this._parseDateFromString(intranetPlayItem["DateTime"]), 
                        "Seats": subBookingItems.length * 2
                    }).then(addResult => {
                        this.setState({ SuccessMessage: this.state.SuccessMessage + "\n\n" + "add ok: DateTime=" + addResult.data["DateTime"] + ", DateTime orig=" + intranetPlayItem["DateTime"]})

                        subBookingItems.forEach(bookingItem => {
                            //this.setState({ SuccessMessage: this.state.SuccessMessage + "\n" + "bookingItem=" + JSON.stringify(bookingItem) });

                            var parts = new String(bookingItem["OData__x0411__x0440__x043e__x043d__x04"]).split("-");
                            //this.setState({ SuccessMessage: this.state.SuccessMessage + "\n" + "parts=" + parts});

                            if (parts.length > 0) {
                                var name = parts[0];

                                //this.setState({ SuccessMessage: this.state.SuccessMessage + ", " + "user name: " + name })

                                sp.web.ensureUser(name)
                                    .then(result => {

                                    //this.setState({ SuccessMessage: this.state.SuccessMessage + ", " + "ensureUser ok: Id=" + result.data.Id })

                                    var userId = result.data.Id;
                                    sp.web.lists.getByTitle("Booking").items
                                        .add({
                                            "Title": intranetPlayItem["Link"]["Description"],
                                            "PlayId": addResult.data["ID"], 
                                            "Status": bookingItem["status"],
                                            "Seats": 2,
                                            "WhoBookedId": userId
                                        }).then(res => {
                                            this.setState({ SuccessMessage: this.state.SuccessMessage + ", " + "add booking ok user: " + result.data.Title })
                                            });
                                    }).catch(err => {
                                        this.setState({ ErrorMessage: this.state.ErrorMessage + ", " + "ensureUser error: name=" + name })
                                    });
                            }
                        });
                    });
                });
            });

        /*
        $.ajax({
            xhrFields: {
                withCredentials: true
            },
            headers: {
                'Authorization': 'Basic ' + btoa('myuser:mypswd')
            },
            url: "http://intranet.kyivstar.ua/Lists/TicketsBooking/export.aspx"
        });

        var plays = [];
        $("#tickets_conainer").load("http://intranet.kyivstar.ua/Lists/TicketsBooking/export.aspx table.ms-listviewtable", function (response, status, xhr) {
            $('#tickets_conainer tr').each(function () {
                var cells = $('td', this);
                plays.push(cells);
            });
            self.setState({ SuccessMessage: "plays= " + JSON.stringify(plays) })

        });

        var booking = [];
        $("#tickets_conainer").load("http://intranet.kyivstar.ua/Lists/BookingTickets/export.aspx table.ms-listviewtable", function (response, status, xhr) {
            $('#tickets_conainer tr').each(function () {
                var cells = $('td', this);
                booking.push(cells);
            });
        });


        */

        /*
        var soapEnv =
            "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'> \
              <soap:Body> \
                <GetListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'> \
                  <listName>TicketsBooking</listName> \
                  <viewName></viewName> \
                  <query></query> \
                  <viewFields></viewFields> \
                  <rowLimit></rowLimit> \
                  <queryOptions></queryOptions> \
                  <webID></webID> \
                </GetListItems> \
              </soap:Body> \
            </soap:Envelope>";

        $.ajax({
            url: "http://intranet.kyivstar.ua/_vti_bin/Lists.asmx",
            crossDomain: true,
            type: "POST",
            dataType: "xml",
            data: soapEnv,
            complete: (res) => {
                self.setState({ SuccessMessage: "res= " + JSON.stringify(res) })
            },
            contentType: "application/xml; charset=\"utf-8\""
        });

*/
        /*
        SP.SOD.executeFunc('/_layouts/15/init.js', '$_global_init', function () {
            SP.SOD.executeFunc('/_layouts/15/MicrosoftAjax.js', 'Sys', function () {
                SP.SOD.executeFunc('/_layouts/15/SP.Runtime.js', 'SP', function () {
                    SP.SOD.executeFunc('/_layouts/15/SP.js', 'SP', function () {

                        self.setState({ SuccessMessage: "SP.js ok!" })

                        var appweburl = "http://sp.kyivstar.ua";
                        var hostweburl = "http://intranet.kyivstar.ua";
                        var context = new SP.ClientContext(appweburl);
                        var factory = new SP.ProxyWebRequestExecutorFactory(appweburl);
                        context.set_webRequestExecutorFactory(factory);
                        var appContextSite = new SP.AppContextSite(context, hostweburl);
                        var web = appContextSite.get_web();

                        const list1: SP.List = web.get_lists().getByTitle('TicketsBooking');
                        const items1: SP.ListItemCollection = list1.getItems(SP.CamlQuery.createAllItemsQuery());

                        var arr1: any = [];
                        context.load(items1);
                        context.executeQueryAsync(function () {
                            var listEnumerator = items1.getEnumerator();
                            while (listEnumerator.moveNext()) {
                                arr1.push(listEnumerator.get_current());
                            }

                            self.setState({ SuccessMessage: "arr1= " + JSON.stringify(arr1) })
                        })
                    });
                });
            });
        });
        */


    
        
        /*
        var intranetWeb = (new Web("http://intranet.kyivstar.ua")); //.configure({ credentials: 'include' });   //mode: 'cors', 
        var promise1: Promise<any[]> = intranetWeb.lists
            .configure({ mode: "no-cors", credentials: 'include'})
            .getByTitle("TicketsBooking").items
            .select("ID", "Title", "DateTime", "Link", "Seats", "Comments")
            .expand("Author")
            .orderBy("ID", false)
            .get();

        var promise2: Promise<any[]> = intranetWeb.lists
            .configure({ mode: "no-cors", credentials: 'include' })
            .getByTitle("BookingTickets").items
            .select("ID", "Title", "BookingId", "status", "Created", "Author/ID", "Author/Title", "Author/Name", "Author/EMail")
            .expand("Author")
            .orderBy("BookingId", false)
            .get();

        Promise.all([promise1, promise2])
            .then(([intranetPlayItems, intranetBookingItems]: [any[], any[]]) => {
                this.setState({ SuccessMessage: "intranetPlayItems= " + JSON.stringify(intranetPlayItems) + "\n" + "intranetBookingItems= " + JSON.stringify(intranetBookingItems)})
                intranetPlayItems.forEach(intranetPlayItem => {

                    //sp.web.lists.getByTitle("plays").items.add({
                    //    "ID": -parseInt(intranetPlayItem["ID"])
                    //})
                    intranetBookingItems
                        .filter(intranetBookingItem => {
                            return intranetBookingItem["BookingId"] == intranetPlayItems["ID"];
                        })
                        .forEach(intranetBookingItem => {
                        })
                });
            });
            */

    }

    private _setMode = (mode: TicketTableMode) => {
        this.setState({ mode: mode, Message: "", SuccessMessage: "", ErrorMessage: "" });

        if (mode == TicketTableMode.plays) {
            this.populatePlays();
        }
        else if (mode == TicketTableMode.allbookings) {
            this.populateAllBookings();
        }
    }

    private _toggleIsManagerForm = () => {
        this.setState({ isManagerForm: !this.state.isManagerForm, Message: "", ErrorMessage: "", SuccessMessage: "" });
    }

    private _manager2persona(manager) {
        return {
            Id: manager["ID"],
            Title: manager["Title"],
            Email: manager["EMail"],
            LoginName: manager["Name"],
            text: manager["Title"],
            imageInitials: (new String(manager["Title"])).split(" ").splice(0, 2).map((n) => n[0]).join(""),
            presence: PersonaPresence.none
        };
    }

    private _persona2manager(persona) {
        return { "ID": persona["Id"], "Title": persona["text"], "EMail": persona["Email"], "Name": persona["LoginName"] };
    }

    private _parseDateFromString = (str: string): Date => {
        var date = new Date()
        var parts = str.split(" ", 2);
        if (parts.length > 0) {
            var d: any[] = parts[0].split(".", 3);
            date = new Date(parseInt(d[2]), parseInt(d[1]) - 1, parseInt(d[0]));
        }
        if (parts.length > 1) {
            var t: any[] = parts[1].split(":", 2);
            date = new Date(parseInt(d[2]), parseInt(d[1]) - 1, parseInt(d[0]), parseInt(t[0]), parseInt(t[1]));
        }
        return date;
    }


    private _saveParams = () => {

        const publishDateNew: Date = this._parseDateFromString($("#" + this._publishDateDivId + " input").attr("value"));
        const ticketCountNew: any = $("#" + this._ticketSpinnerDivId + " input").attr("aria-valuenow");
        const monthCountNew: any = $("#" + this._monthSpinnerDivId + " input").attr("aria-valuenow");
        const managersIds: number[] = this._peoplePicker.current.items.map(persona => { return parseInt(persona["Id"]); });
        const managers: any[] = this._peoplePicker.current.items.map(persona => { return this._persona2manager(persona); });

        const { id, publishDate, monthCount, ticketCount,
                TicketsOrderedEmailText, TicketsOrderedManagerEmailText, StatusChangedEmailText, StatusChangedRejectedEmailText,
                TempEmailText1, TempEmailText2, TempEmailText3, TempEmailText4,
                TempEmailSubject1, TempEmailSubject2, TempEmailSubject3, TempEmailSubject4,
                disableNotifications
            } = this.state.params;

        const subparams = {
            publishDate: publishDateNew,
            monthCount: monthCountNew,
            ticketCount: ticketCountNew,
            TicketsOrderedEmailSubject: this.state.params.TempEmailSubject1,
            TicketsOrderedManagerEmailSubj: this.state.params.TempEmailSubject2,
            StatusChangedEmailSubject: this.state.params.TempEmailSubject3,
            StatusChangedRejectedEmailSubj: this.state.params.TempEmailSubject4,
            TicketsOrderedEmailText: this.state.params.TempEmailText1,
            TicketsOrderedManagerEmailText: this.state.params.TempEmailText2,
            StatusChangedEmailText: this.state.params.TempEmailText3,
            StatusChangedRejectedEmailText: this.state.params.TempEmailText4,
            managersId: { results: managersIds }, // allows multiple users Id array [1,4]
            disableNotifications
        };

        var params: TicketTableParams = {
            id,
            publishDate: publishDateNew,
            monthCount: monthCountNew,
            ticketCount: ticketCountNew,
            TicketsOrderedEmailSubject: this.state.params.TempEmailSubject1,
            TicketsOrderedManagerEmailSubject: this.state.params.TempEmailSubject2,
            StatusChangedEmailSubject: this.state.params.TempEmailSubject3,
            StatusChangedRejectedEmailSubject: this.state.params.TempEmailSubject4,
            TicketsOrderedEmailText: this.state.params.TempEmailText1,
            TicketsOrderedManagerEmailText: this.state.params.TempEmailText2,
            StatusChangedEmailText: this.state.params.TempEmailText3,
            StatusChangedRejectedEmailText: this.state.params.TempEmailText4,
            TempEmailText1,
            TempEmailText2,
            TempEmailText3,
            TempEmailText4,
            TempEmailSubject1,
            TempEmailSubject2,
            TempEmailSubject3,
            TempEmailSubject4,
            managers,
            disableNotifications
        };

        if (id >= 0) {
            sp.web.lists.getByTitle("parameters").items.getById(id)
                .update(subparams).then(() => {
                    this.setState({ showPanel: false, params: params });
                    this.populatePlays();
                });
        }
        else {
            sp.web.lists.getByTitle("parameters").items
                .add(subparams).then((res: ItemAddResult) => {
                    params.id = res.data["ID"];
                    this.setState({ showPanel: false, params: params });
                    this.populatePlays();
                });
        }
    }

    private _showPanel = () => {
        this.setState({ showPanel: true, Message: "", ErrorMessage: "", SuccessMessage: "" });
    }

    private _hidePanel = () => {
        this.setState({ showPanel: false, Message: "", ErrorMessage: "", SuccessMessage: "" });
    }

    componentDidMount() {
        this.populatePlays();
    }

    componentDidUpdate(prevProps, prevState) {
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

    private populatePlays(playId?: string): Promise<any>{

        var message = "";

        let p = new Promise((resolve, reject) => {

            var promiseU: Promise<any> = sp.web.currentUser.get();
            var promiseP: Promise<BasePermissions> = sp.web.lists.getByTitle("Plays").getCurrentUserEffectivePermissions();
            var promiseB: Promise<BasePermissions> = sp.web.lists.getByTitle("Booking").getCurrentUserEffectivePermissions();
            var promiseJ: Promise<any> = sp.web.lists.getByTitle("parameters").items
                .select("ID", "Title", "publishDate", "ticketCount", "monthCount", "managers/ID", "managers/Title", "managers/Name", "managers/EMail", "StatusChangedEmailText", "TicketsOrderedEmailText", "StatusChangedRejectedEmailText", "TicketsOrderedManagerEmailText", "StatusChangedEmailSubject", "TicketsOrderedEmailSubject", "StatusChangedRejectedEmailSubj", "TicketsOrderedManagerEmailSubj", "disableNotifications")
                .expand("managers")
                .orderBy("ID", false)
                .get();

            // 1) user, permissions, params
            Promise.all([promiseU, promiseP, promiseB, promiseJ])
                .then(([user, playsPerm, bookingPerm, paramsArr]: [CurrentUser, BasePermissions, BasePermissions, any[]]) => {
                    // default params on Mount!
                    var loaded_params: TicketTableParams = this.state.params;

                    if (paramsArr && paramsArr.length > 0) {
                        const { ID, publishDate, monthCount, ticketCount, managers,
                            StatusChangedEmailText,
                            TicketsOrderedEmailText,
                            StatusChangedRejectedEmailText,
                            TicketsOrderedManagerEmailText,
                            TicketsOrderedEmailSubject,
                            TicketsOrderedManagerEmailSubj,
                            StatusChangedEmailSubject,
                            StatusChangedRejectedEmailSubj,
                            disableNotifications } = paramsArr[0];

                        loaded_params = {
                            id: ID,
                            publishDate: new Date(publishDate),
                            monthCount,
                            ticketCount,
                            managers: managers ? managers : [],
                            TicketsOrderedEmailSubject,
                            TicketsOrderedManagerEmailSubject: TicketsOrderedManagerEmailSubj,
                            StatusChangedEmailSubject,
                            StatusChangedRejectedEmailSubject: StatusChangedRejectedEmailSubj,
                            TicketsOrderedEmailText,
                            TicketsOrderedManagerEmailText,
                            StatusChangedEmailText,
                            StatusChangedRejectedEmailText,
                            TempEmailText1: TicketsOrderedEmailText,
                            TempEmailText2: TicketsOrderedManagerEmailText,
                            TempEmailText3: StatusChangedEmailText,
                            TempEmailText4: StatusChangedRejectedEmailText,
                            TempEmailSubject1: TicketsOrderedEmailSubject,
                            TempEmailSubject2: TicketsOrderedManagerEmailSubj,
                            TempEmailSubject3: StatusChangedEmailSubject,
                            TempEmailSubject4: StatusChangedRejectedEmailSubj,
                            disableNotifications
                        };
                    }


                    const today = new Date();
                    const archiveDate: string = this.monthAdd(new Date(), -loaded_params.monthCount).toISOString();

                    //var promiseA: Promise<any> = sp.web.lists.getByTitle("Plays").items
                    //    .select("ID", "Title", "Seats", "DateTime", "Link", "Comments")
                    //    .filter("DateTime lt datetime'" + today.toISOString() + "' and DateTime ge datetime'" + archiveDate + "'")
                    //    .orderBy("DateTime", true)
                    //    .get();

                    // Plays
                    var filterI: string = "DateTime ge datetime'" + this.state.playsFromDate.toISOString() + "'";
                    if (playId) {
                        filterI += " and ID eq " + playId;
                    }
                    var promiseI: Promise<any[]> = sp.web.lists.getByTitle("Plays").items
                        .select("ID", "Title", "Seats", "DateTime", "Link", "Comments")
                        .filter(filterI)
                        .orderBy("DateTime", true)
                        .get();

                    // Booking
                    var filterX: string = "Play/DateTime ge datetime'" + this.state.playsFromDate.toISOString() + "'";
                    if (playId) {
                        filterX += " and Play/ID eq " + playId;
                    }
                    var promiseX: Promise<any[]> = sp.web.lists.getByTitle("Booking").items
                        .select("ID", "Title", "Play/ID", "Play/DateTime", "WhoBooked/ID", "WhoBooked/Title", "WhoBooked/Name", "WhoBooked/EMail", "Seats", "Status", "Notes", "Author/ID", "Author/Title")
                        .expand("WhoBooked", "Play", "Author")
                        .filter(filterX)
                        .get();

                    var hasManagePermission: boolean = sp.web.hasPermissions(playsPerm, PermissionKind.ManagePermissions);

                    if (loaded_params.publishDate > today && !hasManagePermission) {
                        promiseI = new Promise<any[]>((resolve) => { resolve([]); });
                        promiseX = new Promise<any[]>((resolve) => { resolve([]); });
                    }

                    // 2) plays
                    Promise.all([promiseI, promiseX])
                        .then(([plays, bookingArr]) => {
                            var propsCollection: TicketTableRowProps[] = plays.map((play: any, i: number) => {
                                return {
                                    key: play.ID,
                                    user: user,
                                    bookingPerm: bookingPerm,
                                    playsPerm: playsPerm,
                                    params: loaded_params,
                                    fields: {
                                        id: play.ID,
                                        DateTime: play.DateTime,
                                        Title: play.Title,
                                        Link: play.Link,
                                        Seats: play.Seats,
                                        Comments: play.Comments,
                                        etag: play["odata.etag"]
                                    },
                                    rowRef: React.createRef<TicketTableRow>(),
                                    isManagerForm: hasManagePermission,
                                    mode: this.state.mode,
                                    bookingArr: bookingArr
                                        .filter(booking => {
                                            return booking.Play["ID"] == play.ID;
                                        })
                                        .map(booking => {
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
                                        })
                                };
                            });

                            var propsCollection1 = [];
                            var propsCollection2 = propsCollection;
                            if (playId) {
                                propsCollection1 = this.state.RowPropsCollection.filter(row => {
                                    return row.fields.id.toString() != playId;
                                });

                                propsCollection2 = this.state.RowPropsCollection.map(row => {
                                    return row.fields.id.toString() == playId ? propsCollection[0] : row;
                                });
                            }

                            this.setState({
                                RowPropsCollection: propsCollection1
                            });

                            this.setState({
                                RowPropsCollection: propsCollection2,
                                playsPerm: playsPerm,
                                bookingPerm: bookingPerm,
                                user: user,
                                Message: message,
                                ErrorMessage: "",
                                SuccessMessage: "",
                                params: loaded_params,
                                isManagerForm: hasManagePermission
                            });

                            //this.forceUpdate();

                            Promise.all([this.populateMyBookings(), this.populateAllBookings()])
                                .then((result) => {
                                    resolve();
                                })
                                .catch((err) => {
                                    this.DisplayErrorMessage(err);
                                    reject(err);
                                });
                        })
                        .catch((err) => {
                            this.DisplayErrorMessage(err);
                            reject(err);
                        });
                })
                .catch((err) => {
                    this.DisplayErrorMessage(err);
                    reject(err);
                });
        });

        return p;
    }

    private populateMyBookings = (): Promise<any> => {

        const today = new Date();
        const archiveDate: string = this.monthAdd(this.state.playsFromDate, -this.state.params.monthCount).toISOString();

        var promiseJ: Promise<any> = sp.web.lists.getByTitle("Booking").items
            .select("Id", "Title", "Play/ID", "Play/Title", "Play/Link", "Play/DateTime", "WhoBooked/ID", "WhoBooked/Title", "WhoBooked/Name", "WhoBooked/EMail", "Seats", "Status", "Notes", "Author/ID", "Author/Title")
            .expand("WhoBooked", "Play", "Author")
            .filter("WhoBooked/ID eq " + this.state.user["Id"]) // + " and (Play/DateTime ge datetime'" + archiveDate + encodeURI("' or Status eq 'В очікуванні')"))
            .get();

        promiseJ.then((items: any[]) => {
                this.setState({
                    myBookings: items
                });
        });

        return promiseJ;
    }

    private populateAllBookings = (date?: Date): Promise<any> => {

        const filterDate = date ? date : this.state.allbookingsFromDate;

        var promiseK: Promise<any> = sp.web.lists.getByTitle("Booking").items
            .select("Id", "Title", "Play/ID", "Play/Title", "Play/Link", "Play/Seats", "Play/DateTime", "WhoBooked/ID", "WhoBooked/Title", "WhoBooked/Name", "WhoBooked/EMail", "Seats", "Status", "Notes", "Author/ID", "Author/Title")
            .top(5000)
            .expand("WhoBooked", "Play", "Author")
            .filter("Play/DateTime ge datetime'" + filterDate.toISOString() + "'")  // + encodeURI("' or Status eq 'В очікуванні'")
            .orderBy("Play/DateTime", false)
            .get();

        promiseK.then((items: any[]) => {
            var reversedItems = items.reverse();

            this.setState({
                allBookings: reversedItems
            });

            var promiseD: Promise<any> = sp.web.lists.getByTitle("Booking").items
                .select("Id", "Title", "Play/ID", "Play/Title", "Play/Link", "Play/Seats", "Play/DateTime", "WhoBooked/ID", "WhoBooked/Title", "WhoBooked/Name", "WhoBooked/EMail", "Seats", "Status", "Notes", "Author/ID", "Author/Title")
                .top(5000)
                .expand("WhoBooked", "Play", "Author")
                .orderBy("Play/DateTime", false)
                .get();

            promiseD.then((all_items: any[]) => {

                reversedItems.forEach(item => {
                    var subitems = all_items.filter(altitem => { return altitem.WhoBooked["ID"] == item.WhoBooked["ID"] && altitem["Status"] == "Затверджено"; })

                    var seats: number = 0;
                    var details: string = "";

                    subitems.map(subitem => {
                        seats += parseInt(subitem["Seats"]);
                        details += (new Date(subitem.Play["DateTime"])).format('dd.MM.yyyy HH:mm') + ", " + subitem["Seats"] + ", " + subitem["Status"] + ", " + subitem.Play["Title"] + "\n";
                    });

                    item["Approved"] = seats;
                    item["ApprovedDetails"] = details;

                });

                this.setState({
                    allBookings: reversedItems
                });

            });

        });

        /*

        promiseK.then((items: any[]) => {
            this.setState({
                allBookings: items
            });

            items.reverse().map(item => {
                sp.web.lists.getByTitle("Booking").items
                    .select("Id", "Title", "Play/ID", "Play/Title", "Play/Link", "Play/Seats", "Play/DateTime", "WhoBooked/ID", "WhoBooked/Title", "WhoBooked/Name", "WhoBooked/EMail", "Seats", "Status", "Notes", "Author/ID", "Author/Title")
                    .expand("WhoBooked", "Play", "Author")
                    .filter("WhoBooked/ID eq " + item.WhoBooked["ID"] + encodeURI(" and Status eq 'Затверджено'"))
                    .orderBy("Play/DateTime", true)
                    .get()
                    .then(approveditems => {
                        var seats: number = 0;
                        var details: string = "";
                        approveditems.map(approveditem => {
                            seats += parseInt(approveditem["Seats"]);
                            details += (new Date(approveditem.Play["DateTime"])).format('dd.MM.yyyy HH:mm') + ", " + approveditem.Play["Title"] + ", " + approveditem["Seats"] + "\n";
                        });

                        item["Approved"] = seats;
                        item["ApprovedDetails"] = details;

                        this.setState({
                            allBookings: items
                        });

                    });

            });
        });
        */

        return promiseK;
    }

    private getTotalsdByPlayID(playID: string, Seats: number): { Free: number, Ordered: number, Rejected: number, Seats: number } {
        var Ordered: number = 0;
        var Rejected: number = 0;
        var Free: number = Seats;
        var Seats: number = Seats;
        this.state.allBookings.map(booking => {
            if (playID == booking.Play["ID"]) {
                if (booking.Status !== "Відхилено") {
                    Free -= parseInt(booking.Seats, 10);
                    Ordered += parseInt(booking.Seats, 10);
                }
                else {
                    Rejected += parseInt(booking.Seats, 10);
                }
            }
        });
        return { Free: Free, Ordered: Ordered, Rejected: Rejected, Seats: Seats };
    }

    private DisplayErrorMessage(err) {
        this.setState({ Message: "", SuccessMessage: "", ErrorMessage: "Помилка:" + JSON.stringify(err) });
    }

    private _onStatusChange(booking: TicketTableMyBookings, new_status: string, new_notes: string) {
        this.setState({ ErrorMessage: "" });

        var months: number = this.state.params.monthCount;

        var p3: Promise<any[]> = new Promise<any[]>((resolve) => { resolve([]); });
        var p4: Promise<any[]> = new Promise<any[]>((resolve) => { resolve([]); });

        if (booking.Status == "Відхилено" && new_status != "Відхилено") {
            p3 = sp.web.lists.getByTitle("Booking").items
                .select("ID", "Title", "Play/ID", "Play/DateTime", "WhoBooked/ID", "WhoBooked/Title", "WhoBooked/Name", "WhoBooked/EMail", "Seats", "Status", "Notes", "Author/ID", "Author/Title")
                .expand("WhoBooked", "Play", "Author")
                .filter("WhoBooked/ID eq " + booking.WhoBooked["ID"] + " and Play/DateTime ge datetime'" + this.monthAdd(new Date(), -months).toISOString() + encodeURI("' and Status ne 'Відхилено'"))
                .get();
        }
        else if (new_status == "Затверджено") {
            p3 = sp.web.lists.getByTitle("Booking").items
                .select("ID", "Title", "Play/ID", "Play/Seats", "Seats", "Status", "Notes")
                .expand("Play")
                .filter("Play/ID eq " + booking.Play.ID + encodeURI(" and Status eq 'Затверджено'"))
                .get();

            p4 = sp.web.lists.getByTitle("Plays").items
                .select("ID", "Seats")
                .filter("ID eq " + booking.Play.ID)
                .get();
        }

        Promise.all([p3, p4]).then(([items3, items4]: [any[], any[]]) => {
            if (booking.Status == "Відхилено" && new_status != "Відхилено" && items3.length > 0) {
                this.setState({ ErrorMessage: "Не можна змінювати статус замовлення. За заданний період (" + months + " місяців) вже є замовлення..." });
            }
            else if (new_status == "Затверджено" &&
                items4.reduce((sum, item) => { sum + parseInt(item["Seats"]) }, 0) -
                items3.reduce((sum, item) => { sum + parseInt(item["Seats"]) }, 0) < parseInt(booking.Seats)) {
                
                this.setState({ ErrorMessage: "Не вдалося затвердити! Помилка в кількості квитків." });
            }
            else {
                var p1: Promise<ItemUpdateResult> = sp.web.lists.getByTitle("Plays").items.getById(booking.Play.ID)
                    .update({ Title: booking.Play.Title }, booking.Play.etag);

                var p2: Promise<ItemUpdateResult> = sp.web.lists.getByTitle("Booking").items.getById(parseInt(booking.Id, 10))
                    .update({ Status: new_status, Notes: new_notes }, booking["odata.etag"]);

                Promise.all([p1, p2])
                    .then(([res1, res2]: [ItemUpdateResult, ItemUpdateResult]) => {

                        if (new_status == "Відхилено") {
                            // Параметри: %Замовлено, %Вистава, %Статус, %Дата, %Лінк", %Коментар, %Замовник
                            var body: string = this.state.params.StatusChangedRejectedEmailText
                                .replace("%Замовлено", booking.Seats.toString())
                                .replace("%Вистава", booking.Play.Title)
                                .replace("%Дата", (new Date(booking.Play.DateTime)).format('dd.MM.yyyy HH:mm'))
                                .replace("%Статус", new_status)
                                .replace("%Лінк", booking.Play.Link)
                                .replace("%Коментар", new_notes)
                                .replace("%Замовник", booking.WhoBooked["Title"]);

                            var subject: string = this.state.params.StatusChangedRejectedEmailSubject
                                .replace("%Замовлено", booking.Seats.toString())
                                .replace("%Вистава", booking.Play.Title)
                                .replace("%Дата", (new Date(booking.Play.DateTime)).format('dd.MM.yyyy HH:mm'))
                                .replace("%Статус", new_status)
                                .replace("%Лінк", booking.Play.Link)
                                .replace("%Коментар", new_notes)
                                .replace("%Замовник", booking.WhoBooked["Title"]);

                            if (!this.state.params.disableNotifications) {
                                sp.utility.sendEmail({
                                    To: [booking.WhoBooked.EMail],
                                    Subject: subject,
                                    Body: body
                                })
                            }
                        }
                        else {
                            // Параметри: %Замовлено, %Вистава, %Статус, %Дата, %Лінк", %Коментар, %Замовник
                            var body: string = this.state.params.StatusChangedEmailText
                                .replace("%Замовлено", booking.Seats.toString())
                                .replace("%Вистава", booking.Play.Title)
                                .replace("%Дата", (new Date(booking.Play.DateTime)).format('dd.MM.yyyy HH:mm'))
                                .replace("%Статус", new_status)
                                .replace("%Лінк", booking.Play.Link)
                                .replace("%Коментар", new_notes)
                                .replace("%Замовник", booking.WhoBooked["Title"]);

                            var subject: string = this.state.params.StatusChangedEmailSubject
                                .replace("%Замовлено", booking.Seats.toString())
                                .replace("%Вистава", booking.Play.Title)
                                .replace("%Дата", (new Date(booking.Play.DateTime)).format('dd.MM.yyyy HH:mm'))
                                .replace("%Статус", new_status)
                                .replace("%Лінк", booking.Play.Link)
                                .replace("%Коментар", new_notes)
                                .replace("%Замовник", booking.WhoBooked["Title"]);

                            if (!this.state.params.disableNotifications) {
                                sp.utility.sendEmail({
                                    To: [booking.WhoBooked.EMail],
                                    Subject: subject,
                                    Body: body
                                })
                            }
                        }

                        this.populateAllBookings();

                    })
                    .catch((err1: any) => {
                        // etag invalid!
                        this.populateAllBookings().then(bookings => {
                            this.setState({ ErrorMessage: "Дані, можливо, змінилися. Спробуйте перезавантажити таблицю!" });
                        });
                    });
            }
        });
    }

    private _onDeleteBookingClicked(bookingId: string) {
        var row: TicketTableRow = this.findTicketTableRowByBookingId(bookingId);

        if (row) {
            row._onDeleteBookingClicked(bookingId);
        }
        else {
            sp.web.lists.getByTitle("Booking").items.getById(parseInt(bookingId, 10))
                .delete()
                .then(res => {
                    this.populateMyBookings().then((bookings: any[]) => {});
                })
                .catch((err1: any) => {
                    this.setState({ ErrorMessage: "Дані, можливо, змінилися. Спробуйте перезавантажити таблицю!" });
                });
        }
    }

    private findTicketTableRowByBookingId(bookingId: string) : TicketTableRow {
        var row: TicketTableRow = null;
        this.state.RowPropsCollection.map(rowProps => {
            var rowByRef: TicketTableRow = rowProps.rowRef.current;
            rowByRef.state.bookingArr.map(booking => {
                if (booking.ID == bookingId) {
                    row = rowByRef;
                }
            });
        });
        return row;
    }

    private findTicketTableRowByPlayId(playId: string): TicketTableRow {
        var row: TicketTableRow = null;
        this.state.RowPropsCollection.map(rowProps => {
            if (rowProps.fields.id.toString() == playId) {
                row = rowProps.rowRef.current;
            }
        });
        return row;
    }

    private BookingChangedCallback(playId: string) {
        this.populateMyBookings();
        this.populateAllBookings();
    }

    private StatusChangedCallback(playId: string) {
        this.populateMyBookings();
        this.populateAllBookings();
    }

    private OrderAttemptFailed(playId: string, attempt: number) {
        this.populatePlays(playId)
            .then(() => {
                var row: TicketTableRow = this.findTicketTableRowByPlayId(playId);
                if (row) {
                    row._onOrderTicketsClicked(attempt + 1);
                }
            });
    }

    private MessageCallback(msg: string, err: string, scs: string) {
        this.setState({ Message: msg, ErrorMessage: err, SuccessMessage: scs });
    }
}