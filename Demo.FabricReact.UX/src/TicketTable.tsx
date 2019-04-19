import * as React from "react";
import * as ReactDOM from "react-dom";

import { sp, Web, ItemAddResult } from '@pnp/sp';
import { CurrentUser } from "@pnp/sp/src/siteusers";
import { PermissionKind, BasePermissions } from "@pnp/sp";


import { TicketTableRow, TicketTableRowProps, TicketTableRowMouseEventCode } from "./TicketTableRow";
import { MessageBar, MessageBarType, initializeIcons, Stack, DefaultButton, PrimaryButton, getId } from "office-ui-fabric-react";
import { IDatePicker, DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { SpinButton } from 'office-ui-fabric-react/lib/SpinButton';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';


import './TicketTable.scss';

// Register icons and pull the fonts from the default SharePoint cdn.
initializeIcons();

const DayPickerStrings: IDatePickerStrings = {
    months: ['Січень', 'Лютий', 'Березень', 'Квітень', 'Травень', 'Червень', 'Липень', 'Серпень', 'Вересень', 'Жовтень', 'Листопад', 'Грудень'],
    shortMonths: ['Січ', 'Лют', 'Бер', 'Кві', 'Тра', 'Чер', 'Лип', 'Сер', 'Вер', 'Жов', 'Лис', 'Dec'],
    days: ['Неділя', 'Понеділок', 'Вівторок', 'Середа', 'Четверг', 'П\'ятниця', 'Субота'],
    shortDays: ['Нд', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'],
    goToToday: 'Сьогодні',
    prevMonthAriaLabel: 'Попередній місяць',
    nextMonthAriaLabel: 'Наступний місяць',
    prevYearAriaLabel: 'Попередній рік',
    nextYearAriaLabel: 'Наступний рік',
    closeButtonAriaLabel: 'Закрити'
};


export interface TicketTableParams {
    id: number,
    publishDate: Date,
    monthCount: number,
    ticketCount: number
}

export interface TicketTableProps {
}

export interface TicketTableState {
    Message?: string;
    SuccessMessage?: string;
    ErrorMessage?: string;
    user?: CurrentUser,
    bookingPerm?: BasePermissions,
    playsPerm?: BasePermissions,
    RowPropsCollection: TicketTableRowProps[];
    showPanel: boolean,
    params: TicketTableParams
}

export class TicketTable extends React.Component<TicketTableProps, TicketTableState> {

    private _dafaultParams: TicketTableParams;
    private _spinButtonMonthRef = React.createRef<SpinButton>();
    private _spinButtonTicketRef = React.createRef<SpinButton>();
    private _ticketSpinnerDivId = getId("spinnerDiv");
    private _monthSpinnerDivId = getId("spinnerDiv");

    //private _panelPublishDate;
    private _panelTicketCount;
    private _panelMonthCount;
    private tdOrderdRef = React.createRef<HTMLTableDataCellElement>();


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
            ticketCount: 2
        };
        this.state = { RowPropsCollection: [], showPanel: false, params: this._dafaultParams };
    }

    render(): JSX.Element {
        return (
            <div className="TicketTableDiv">
                {this.state.playsPerm && sp.web.hasPermissions(this.state.playsPerm, PermissionKind.ManagePermissions) && 
                    <div>
                    <DefaultButton style={{ marginBottom: '10px' }} secondaryText="Відкрити панель параметрів" onClick={this._showPanel} text="Параметри..." />
                    <Panel
                        isOpen={this.state.showPanel}
                        type={PanelType.smallFixedFar}
                        onDismiss={this._hidePanel}
                        headerText="Параметри"
                        closeButtonAriaLabel="Close"
                        //onRenderFooterContent={this._onRenderFooterContent}
                    >
                        <Stack padding={10} gap={10}>
                            <div style={{ width: '150px' }}>
                                <DatePicker
                                    label="Дата публікації"
                                    firstDayOfWeek={DayOfWeek.Monday}
                                    isRequired={false}
                                    allowTextInput={false}
                                    strings={DayPickerStrings}
                                    value={this.state.params.publishDate}
                                    formatDate={(date?: Date) => { return date ? date.format('dd.MM.yyyy') : ""; }}
                                    onSelectDate={(date: Date) => {
                                        const { id, publishDate, monthCount, ticketCount } = this.state.params;
                                        this.setState({ params: { id, ticketCount, monthCount, publishDate: date } });
                                    }}
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
                            <div style={{ marginTop: '300px' }}>
                                <PrimaryButton onClick={this._saveParams} style={{ marginRight: '8px' }}> Save </PrimaryButton>
                                <DefaultButton onClick={this._hidePanel}> Cancel </DefaultButton>
                            </div>
                        </Stack>
                    </Panel>
                    </div>
                }
                {this.state.Message &&
                    <MessageBar className="MessageBar" isMultiline={true}>
                        {this.state.Message}
                    </MessageBar>
                }
                {this.state.SuccessMessage &&
                    <MessageBar className="MessageBar" messageBarType={MessageBarType.success} isMultiline={true} dismissButtonAriaLabel="Close">
                        {this.state.SuccessMessage}
                    </MessageBar>
                }
                {this.state.ErrorMessage &&
                    <MessageBar className="MessageBar" messageBarType={MessageBarType.error} isMultiline={true} dismissButtonAriaLabel="Close">
                        {this.state.ErrorMessage}
                    </MessageBar>
                }
                <table className="TicketTable">
                    <tbody>
                        <TicketTableRow key="0" isHeader={true} user={this.state.user} bookingPerm={null} playsPerm={null} params={this.state.params} />
                        {this.state.RowPropsCollection.map(rowProps => {
                            return <TicketTableRow key={rowProps.key} fields={rowProps.fields} user={this.state.user} bookingPerm={rowProps.bookingPerm} playsPerm={rowProps.playsPerm} params={this.state.params} />;
                        })}
                    </tbody>
                </table>
                <div id="TicketTableCalloutDiv">
                </div>
            </div>
        );
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

    private _saveParams = () => {

        const ticketCountNow: any = $("#" + this._ticketSpinnerDivId + " input").attr("aria-valuenow");
        const monthCountNow: any = $("#" + this._monthSpinnerDivId + " input").attr("aria-valuenow");

        const { id, publishDate, monthCount, ticketCount } = this.state.params;

        const subparams = {
            publishDate,
            monthCount: monthCountNow,
            ticketCount: ticketCountNow
        };

        var params: TicketTableParams = {
            id,
            publishDate,
            monthCount: monthCountNow,
            ticketCount: ticketCountNow
        };

        if (id >= 0) {
            sp.web.lists.getByTitle("Playsparam").items.getById(id)
                .update(subparams).then(() => {
                    this.setState({ showPanel: false, params: params });
                });
        }
        else {
            sp.web.lists.getByTitle("Playsparam").items
                .add(subparams).then((res: ItemAddResult) => {
                    params.id = res.data["ID"];
                    this.setState({ showPanel: false, params: params });
                });
        }
    }

    private _showPanel = () => {
        this.setState({ showPanel: true });
    }

    private _hidePanel = () => {
        this.setState({ showPanel: false });
    }

    componentDidMount() {

        const today = new Date();

        var promiseU: Promise<any> = sp.web.currentUser.get();
        var promiseP: Promise<BasePermissions> = sp.web.lists.getByTitle("Plays").getCurrentUserEffectivePermissions();
        var promiseB: Promise<BasePermissions> = sp.web.lists.getByTitle("Booking").getCurrentUserEffectivePermissions();
        var promiseI: Promise<any> = sp.web.lists.getByTitle("Plays").items
            .select("ID", "Title", "Seats", "DateTime", "Link", "Comments")
            .filter("DateTime ge datetime'" + today.toISOString() + "'")
            .orderBy("DateTime", false).get();

        if (this.state.params.publishDate > today && !sp.web.hasPermissions(this.state.playsPerm, PermissionKind.ManagePermissions)) {
            promiseI = new Promise<any>((resolve) => { resolve([])});
        }

        // and DateTime le datetime'" + this.state.params.publishDate.toISOString() + "'
        var promiseJ: Promise<any> = sp.web.lists.getByTitle("Playsparam").items
            .select("ID", "Title", "publishDate", "ticketCount", "monthCount")
            .orderBy("ID", false).get();

        this.setState({ Message: "Завантажується...", SuccessMessage: "", ErrorMessage: "" });

        Promise.all([promiseU, promiseP, promiseB, promiseI, promiseJ])
            .then(([user, playsPerm, bookingPerm, plays, paramsArr]: [CurrentUser, BasePermissions, BasePermissions, any[], any[]]) => {

                var loaded_params: TicketTableParams = this.state.params;

                if (paramsArr && paramsArr.length > 0) {
                    const { ID, publishDate, monthCount, ticketCount } = paramsArr[0];
                    loaded_params = { id: ID, publishDate: new Date(publishDate), monthCount, ticketCount };
                }

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
                        }
                    };
                });

                this.setState({
                    RowPropsCollection: propsCollection,
                    playsPerm: playsPerm,
                    bookingPerm: bookingPerm,
                    user: user,
                    Message: "",
                    ErrorMessage: "",
                    SuccessMessage: "",
                    params: loaded_params
                });
            })
            .catch(this.catchError);
    }

    componentDidUpdate(prevProps, prevState) {
    }

    private catchError(err) {
        this.setState({ Message: "", SuccessMessage: "", ErrorMessage: "Помилка:" + JSON.stringify(err) });
    }

}