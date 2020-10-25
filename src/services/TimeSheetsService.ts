import { IApiConfigService, ApiConfigServiceKey } from './ApiConfigService';
import HttpClient from '@microsoft/sp-http/lib/httpClient/HttpClient';
import { ITimeSheet } from '../entities/ITimeSheet';
import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';

export interface ITimeSheetsService {
    getAllTimeSheets(): Promise<ITimeSheet[]>;
    
    // The other commented out query of GET
    //getMyTimeSheets(): Promise<ITimeSheet[]>;
    
    getTimeSheets(id: number): Promise<ITimeSheet>;
    // FIX the below after retrieval..
    
    createTimeSheet(timeSheet: ITimeSheet): Promise<any>;
    updateTimeSheet(id: number, update: ITimeSheet): Promise<any>;
	removeTimeSheet(id: number): Promise<ITimeSheet>;
}

export class TimeSheetsService implements ITimeSheetsService {
    private httpClient: HttpClient;
    private apiConfig: IApiConfigService;

    constructor(private serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
        this.httpClient = serviceScope.consume(HttpClient.serviceKey);
        this.apiConfig = serviceScope.consume(ApiConfigServiceKey);
        });
    }

    public getAllTimeSheets(): Promise<ITimeSheet[]> {
        return this.httpClient.get(this.apiConfig.apiUrl, HttpClient.configurations.v1, 
            {
                mode: 'cors',
                credentials: 'include'
            }).then((resp) => resp.json());
    }

    // getMyTimeSheets(): Promise<ITimeSheet[]>;
    /*public getMytimeSheets(): Promise<ItimeSheet[]> {
        return this.httpClient.get(this.apiConfig.apiMyDocumentsUrl, HttpClient.configurations.v1, {
        mode: 'cors',
        credentials: 'include'
    }).then((resp) => resp.json());
    }*/

    public getTimeSheets(id: number): Promise<ITimeSheet> {
        return this.httpClient.get(`${this.apiConfig.apiUrl}/${id}`, HttpClient.configurations.v1,
        {
            mode: 'cors',
            credentials: 'include'
        }).then((resp) => resp.json());
    }

    public createTimeSheet(timeSheet: ITimeSheet): Promise<any> {
        return this.httpClient
            .post(`${this.apiConfig.apiUrl}`, HttpClient.configurations.v1, {
        body: JSON.stringify(timeSheet),
        headers: [
            ['Content-Type','application/json']
        ],
                mode: 'cors',
                credentials: 'include'
            })
            .then((resp) => resp.json());
    } 

    public updateTimeSheet(id: number, update: ITimeSheet): Promise<any> {
        return this.httpClient
            .fetch(`${this.apiConfig.apiUrl}/${id}`, HttpClient.configurations.v1, {
        body: JSON.stringify(update),
        headers: [
            ['Content-Type','application/json']
        ],
                mode: 'cors',
                credentials: 'include',
                method: 'PUT'
            });
    } 

    public removeTimeSheet(id: number): Promise<any> {
        return this.httpClient
            .fetch(`${this.apiConfig.apiUrl}/${id}`, HttpClient.configurations.v1, {
                mode: 'cors',
        credentials: 'include',
        method:'DELETE'
            });
    } 
}

export const TimeSheetsServiceKey = ServiceKey.create<ITimeSheetsService>(
	'ypcode:bizdocs-service',
	TimeSheetsService
);