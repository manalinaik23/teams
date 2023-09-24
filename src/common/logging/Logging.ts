import { appInsights } from './AppInsights';
import { SeverityLevel } from '@microsoft/applicationinsights-web';

export enum EventType {
  Information = "Information",
  Error = "Error",
  Exception = "Exception"
}
export class Logging {
  //static appInsights: any;
  private static userAlias: string;

  public static init(alias: string) {
    Logging.userAlias = alias;
  }

  public static AppInsightsTrackEvent(eventName: string, eventType: EventType, Method: string, message: string): void {
    var eventData = {
      "EventType": eventType,
      "EventSubType": Method,
      "Message": message,
      "Url": window.location.href,
      "Alias": Logging.userAlias
    };
    if (!Logging.userAlias)
      throw new Error("Logging not initiated. Please initiate at base component");

    try {
      appInsights.trackEvent(
        {
          name: eventName,
          properties: eventData
        });
    }
    catch (ex) {
      console.log("Error: " + ex.message.toString());
    }
  }
  public static AppInsightsTrackException(ComponentName: string, Method: string, Exception: Error,
    CustomMessage?: string): void {

    if (!Logging.userAlias)
      throw new Error("Logging not initiated. Please initiate at base component");

    appInsights.trackException(
      {
        exception: Exception,
        severityLevel: SeverityLevel.Error,
        properties: {
          Message: Exception.message,
          component: ComponentName,
          EventSubType: Method,
          Alias: Logging.userAlias
        }
      }
    );
  }
  public catch(ex: { message: { toString: () => string; }; }) {
    console.log("Error: " + ex.message.toString());
  }
}

//}




