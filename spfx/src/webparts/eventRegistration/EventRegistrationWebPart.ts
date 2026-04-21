import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'EventRegistrationWebPartStrings';

// Reference the solution
import "main-lib";
declare const EventRegistration: {
  render: (el: HTMLElement, context: WebPartContext, configuration: string) => void;
};

export interface IEventRegistrationWebPartProps {
  configuration: string;
}

export default class EventRegistrationWebPart extends BaseClientSideWebPart<IEventRegistrationWebPartProps> {

  public render(): void {
    // Clear the element
    while (this.domElement.firstChild) { this.domElement.removeChild(this.domElement.firstChild); }

    // Render the application
    EventRegistration.render(this.domElement, this.context, this.properties.configuration);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
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
                PropertyPaneTextField('configuration', {
                  label: "Configuration:",
                  description: "The json configuration for the solution.",
                  multiline: true,
                  rows: 10
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
