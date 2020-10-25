import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';

export interface IApiConfigService {
	apiUrl: string;
	apiMyTimeSheetsUrl: string;
	appRedirectUri: string;
	configure(currentWebUrl: string, remoteHostUrl: string, appInstanceId: string);
}

export class ApiConfigService implements IApiConfigService {
	public apiUrl: string;
	public apiMyTimeSheetsUrl: string;
	public appRedirectUri: string;  

	constructor(private serviceScope: ServiceScope) {}

	public configure(currentWebUrl: string, apiHostUrl: string, appInstanceId: string) {
		this.apiUrl = apiHostUrl + '/api/TimeSheets';
		this.apiMyTimeSheetsUrl = apiHostUrl + '/api/MyTimeSheets';
		this.appRedirectUri = `${currentWebUrl}/_layouts/15/appredirect.aspx?instance_id=${appInstanceId}`;
	}
}

export const ApiConfigServiceKey = ServiceKey.create<IApiConfigService>('ypcode:bizdocs-apiconfig', ApiConfigService);