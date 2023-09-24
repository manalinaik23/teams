
import { HttpClient, IHttpClientOptions } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { AadHttpClient, MSGraphClient, HttpClientResponse } from '@microsoft/sp-http';
import { GraphError } from '@microsoft/microsoft-graph-client';
import { Environment } from './Environment';

export enum APISource {
  WebAPI = "https://us-spodashboardwebapi",
  FunctionAPI = "https://nareportingportalfunctions",
  MSGraphAPI = "https://graph.microsoft.com",
}



export class WebAPIHelper {
  public userName: string;
  private _aadClient: AadHttpClient;
  private _graphClient: MSGraphClient;
  private _httpClient: HttpClient;
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    Environment.initialize();
    this.context = context;
    this._httpClient = context.httpClient;
    this.userName = context.pageContext.user.email;
  }

  private intializeAADHttpClient(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      if (this._aadClient) {
        resolve();
      }
      else {
        this.context.aadHttpClientFactory
          .getClient(Environment.ClientID)
          .then((client: AadHttpClient): void => {
            this._aadClient = client;
            resolve();
          }, (err: any) => reject(err));
      }
    });
  }

  private intializeGraphHttpClient(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      if (this._graphClient) {
        resolve();
      }
      else {
        this.context.msGraphClientFactory
          .getClient()
          .then((client: MSGraphClient): void => {
            this._graphClient = client;
            resolve();
          }, (err: any) => reject(err));
      }
    });
  }

  public getAPIBaseURL(apiSource: APISource): string {
    if (apiSource == APISource.WebAPI)
      return Environment.SPODashboardAPI + "/api/";
    else if (apiSource == APISource.FunctionAPI)
      return Environment.FunctionAPI + "/api/";
  }

  private GetHttpHeader(contentType: string = "") {
    const requestHeaders: Headers = new Headers();
    //  requestHeaders.append('Accept', 'application/json;odata=nometadata');
    if (contentType) {
      requestHeaders.append('Content-Type', contentType);
    }
    requestHeaders.append('User-Agent', 'NONISV|Securitas|NAReportingPortalWebPart/1.0');
    return requestHeaders;
  }

  public GetWebAPI(apiSource: APISource, apiRelativePath: string, contentType: string = "application/json"): Promise<any> {
    return new Promise<any>((resolve: (result: any) => void, reject: (error: any) => void): void => {

      let fullAPIUrl: string;
      const requestHeaders: Headers = this.GetHttpHeader(contentType);
      const httpClientOptions: IHttpClientOptions = {
        headers: requestHeaders
      };
      if (apiSource == APISource.MSGraphAPI)
        fullAPIUrl = apiRelativePath;
      else
        fullAPIUrl = this.getAPIBaseURL(apiSource) + apiRelativePath;


      if (apiSource == APISource.MSGraphAPI) {
        this.intializeGraphHttpClient()
          .then(async (): Promise<void> => {
            try {
              this._graphClient
                .api(apiRelativePath.replace("v1.0/", ""))
                .get((error: GraphError, response: any, rawResponse?: any) => {
                  if (error) {
                    reject(error);
                  }
                  else {
                    resolve(response);
                  }
                });
            }
            catch (error) {
              reject('Graph Client Error: ' + error);
            }
          });
      }
      else {
        this.intializeAADHttpClient()
          .then(async (): Promise<void> => {
            try {
              this._aadClient
                .get(fullAPIUrl, AadHttpClient.configurations.v1, httpClientOptions)
                .then((res: HttpClientResponse): Promise<any> => {
                  if (contentType.indexOf("json") > -1)
                    return res.json();
                  else if (contentType.indexOf("text") > -1)
                    return res.text();
                  else
                    return res.blob();
                })
                .then((responseOutput: any): void => {
                  resolve(responseOutput);
                })
                .catch((error: any) => {
                  reject(error);
                });
            }
            catch (error) {
              reject('AAD Http Client Error: ' + error);
            }
          });
      }
    });
  }

  public PostWebAPI(apiSource: APISource, apiRelativePath: string, jsonbody: string = null, contentType: string = "application/json", returnType: string = "application/json", isResponseCodeRequired: boolean = false): Promise<any> {
    return new Promise<any>((resolve: (result: any) => void, reject: (error: any) => void): void => {
      let fullAPIUrl: string;

      const requestHeaders: Headers = this.GetHttpHeader(contentType);

      const httpClientOptions: IHttpClientOptions = {
        body: jsonbody,
        headers: requestHeaders
      };

      if (apiSource == APISource.MSGraphAPI) {
        this.intializeGraphHttpClient()
          .then(async (): Promise<void> => {
            try {
              this._graphClient
                .api("/" + apiRelativePath)
                .post((error: GraphError, response: any, rawResponse?: any) => {
                  if (error) {
                    reject(error);
                  }
                  else {
                    resolve(response.value);
                  }
                });
            }
            catch (error) {
              reject('Graph Http Client Error: ' + error);
            }
          });
      }
      else {
        this.intializeAADHttpClient()
          .then(async (): Promise<void> => {
            try {
              this._aadClient
                .post(this.getAPIBaseURL(apiSource) + apiRelativePath, AadHttpClient.configurations.v1, httpClientOptions)
                .then((res: HttpClientResponse): any => {
                  let output: any;
                  if (returnType.indexOf("json") > -1)
                    output = res.json();
                  else if (returnType.indexOf("text") > -1)
                    output = res.text();
                  else
                    output = res.blob();

                  if (isResponseCodeRequired) {
                    resolve(output + "####" + res.status);
                  }
                  else {
                    resolve(output);
                  }
                })
                .catch((error: any) => {
                  reject(error);
                });
            }
            catch (error) {
              reject('AAD Http Client Error: ' + error);
            }
          });
      }
    });
  }

}
