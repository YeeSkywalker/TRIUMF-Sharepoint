import { IEvent } from '.';

export interface IPersonalCalendarState {
    error: string;
    loading: boolean;
    events: IEvent[];
    renderedDateTime: Date;
    timeZone?: string; 
}