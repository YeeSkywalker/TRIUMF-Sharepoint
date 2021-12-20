import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PersonalEmailWebPartStrings';
import PersonalEmail from './components/PersonalEmail';
import { IPersonalEmailProps } from './components/IPersonalEmailProps';
import { loadTheme } from 'office-ui-fabric-react';
import {
  IReadonlyTheme,
  ThemeChangedEventArgs,
  ThemeProvider
} from '@microsoft/sp-component-base';

import { MSGraphClient } from '@microsoft/sp-http';

export interface IPersonalEmailWebPartProps {
  title: string;
  numOfMessages: number;
}

export default class PersonalEmailWebPart extends BaseClientSideWebPart<IPersonalEmailWebPartProps> {
  private graphClient: MSGraphClient;
  private propertyFeildNum;
  private themeProvider: ThemeProvider;
  private themeVariant: IReadonlyTheme | undefined;

  public onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this.context.msGraphClientFactory.getClient()
      .then((client: MSGraphClient): void => {
        this.graphClient = client
        resolve()
      }, err => reject(err))

      this.themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

      this.themeVariant = this.themeProvider.tryGetTheme();

      this.themeProvider.themeChangedEvent.add(
        this,
        this.themeChangeEventHandler
      );
    });
  }

  // unfinished
  public render(): void {
    const element: React.ReactElement<IPersonalEmailProps> = React.createElement(
      PersonalEmail,
      {
        title: this.properties.title,
        numOfMessages: this.properties.numOfMessages,
        displayMode: this.displayMode,
        graphClient: this.graphClient,
        themeVariant: this.themeVariant,
        updateProperty: (value: string): void => {
          this.properties.title = value;
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected async loadPropertyPaneResources(): Promise<void> {
    const { PropertyFieldNumber } = await import("@pnp/spfx-property-controls/lib/propertyFields/number");
    this.propertyFeildNum = PropertyFieldNumber;
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
              groupFields: [
                this.propertyFeildNum('numOfMessages', {
                  key: 'numOfMessages',
                  label: strings.NumOfMessagesToShow,
                  value: this.properties.numOfMessages,
                  minValue: 1,
                  maxValue: 10,
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private themeChangeEventHandler(args: ThemeChangedEventArgs): void {
    this.themeVariant = args.theme;

    this.render();
  }
}
