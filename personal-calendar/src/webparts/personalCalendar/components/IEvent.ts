export interface IEvents {
    value: IEvent[];
}

export interface IEvent {
    isAllDay: boolean;
    start: IEventTime;
    end: IEventTime;
    location: {
        displayName: string;
    };
    showAs: string;
    subject: string;
    webLink: string;
}

export interface IEventTime {
    dateTime: string;
    timeZone: string;
}