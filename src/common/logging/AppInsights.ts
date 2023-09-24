import { ApplicationInsights } from '@microsoft/applicationinsights-web';
import { ReactPlugin, withAITracking } from '@microsoft/applicationinsights-react-js';
import { globalHistory } from "@reach/router";
import { Environment } from '../Environment';
const reactPlugin = new ReactPlugin();
Environment.initialize();
const ai = new ApplicationInsights({
    config: {
        instrumentationKey: Environment.AppInsightsKey,
        extensions: [reactPlugin],
        extensionConfig: {
            [reactPlugin.identifier]: { history: globalHistory }
        }
    }
});
ai.loadAppInsights();
 
export default (Component) => withAITracking(reactPlugin, Component);
export const appInsights = ai.appInsights;