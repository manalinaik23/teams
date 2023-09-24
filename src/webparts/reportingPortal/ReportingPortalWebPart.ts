
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { Logging } from '@common/logging/Logging';
import * as strings from 'ReportingPortalWebPartStrings';
import ReportingPortal from './components/ReportingPortal';
import { IReportingPortalProps } from './components/IReportingPortalProps';
import { initializeIcons } from '@uifabric/icons';
//Added Custom class : to increase height of webpart
import './cssFile/ReportingPortalcustome.css';
initializeIcons();

export interface IReportingPortalWebPartProps {
  description: string;
  NoOfIndexValueChar: string;
}

export default class ReportingPortalWebPart extends BaseClientSideWebPart<IReportingPortalWebPartProps> {

  public onInit(): Promise<void> {
    /*Start Commented to visible Edit */
    if (document.querySelectorAll('div[class="commandBarWrapper"]').length > 0) {
      document.querySelectorAll('div[class="commandBarWrapper"]')[0].setAttribute("style", "display:none"); //  style.display = "none";
    }
    /*End Commented to visible Edit */

    Logging.init(this.context.pageContext.user.email);

    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<IReportingPortalProps> = React.createElement(
      ReportingPortal,
      {
        description: this.properties.description,
        context: this.context,
        NoOfIndexValueChar: this.properties.NoOfIndexValueChar

      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('NoOfIndexValueChar', {
                  
                  label:"NoOfIndexValueChar"
                 
                }),
               
              ]
            }
          ]
        }
      ]
    };
  }
}
