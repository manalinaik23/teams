import * as config from '../AppConfig.json';
export class Environment {
    private static _appInsightsKey: string;
    private static _spoDashboardAPI: string;
    private static _functionAPI: string;
    private static _rootSiteUrl:string;
    private static _maxFileDownloadSize:number;
    private static _helpManualPath:string;

    public static initialize() {
        this.SetEnvironmentVariable(window.location.href);
    }
    private static SetEnvironmentVariable(CurrentPageUrl: string): void {
        CurrentPageUrl = CurrentPageUrl.toLowerCase();
        if (CurrentPageUrl.indexOf('/securitasnareportingportaldev/') > -1) {
            this._appInsightsKey = config.Environment.Dev.appInsightsKey;
            this._spoDashboardAPI = config.Environment.Dev.SPODashboardAPI;
            this._functionAPI = config.Environment.Dev.FunctionAPI;
            this._rootSiteUrl=config.Environment.Dev.RootSiteURL;
            this._maxFileDownloadSize=config.Environment.Dev.MaxFileDownloadSize;
            this._helpManualPath=config.Environment.Dev.HelpManualPath;
        }
        else if (CurrentPageUrl.indexOf('/securitasnareportingportaltest/') > -1) {
            this._appInsightsKey = config.Environment.Test.appInsightsKey;
            this._spoDashboardAPI = config.Environment.Test.SPODashboardAPI;
            this._functionAPI = config.Environment.Test.FunctionAPI;
            this._rootSiteUrl=config.Environment.Test.RootSiteURL;
            this._maxFileDownloadSize=config.Environment.Test.MaxFileDownloadSize;
            this._helpManualPath=config.Environment.Test.HelpManualPath;
        }
        else {
            this._appInsightsKey = config.Environment.Prod.appInsightsKey;
            this._spoDashboardAPI = config.Environment.Prod.SPODashboardAPI;
            this._functionAPI = config.Environment.Prod.FunctionAPI;
            this._rootSiteUrl=config.Environment.Prod.RootSiteURL;
            this._maxFileDownloadSize=config.Environment.Prod.MaxFileDownloadSize;
            this._helpManualPath=config.Environment.Prod.HelpManualPath;

        }
    }
    public static get ClientID() {
        return config.ClientId;
    }
    public static get TenantUrl(){
        return config.TenantUrl;
    }
    public static get AppInsightsKey() {
        return this._appInsightsKey;
    }
    public static get SPODashboardAPI() {
        return this._spoDashboardAPI;
    }
    public static get FunctionAPI() {
        return this._functionAPI;
    }
    public static get RootSiteURL() {
        return this._rootSiteUrl;
    }
    public static get MaxFileDownloadSize() {
        return this._maxFileDownloadSize;
    }
    public static get HelpManualPath() {
        return this._helpManualPath;
    }
}
