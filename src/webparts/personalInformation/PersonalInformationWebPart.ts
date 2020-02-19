import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { Logger, LogLevel } from '@pnp/logging';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PersonalInformationWebPartStrings';
import PersonalInformation from './components/PersonalInformation';
import { IPersonalInformationProps } from './components/IPersonalInformationProps';
import { ServiceETools } from '../../utils/ServiceETools';

export interface IPersonalInformationWebPartProps {
  description: string;
}

export default class PersonalInformationWebPart extends BaseClientSideWebPart<IPersonalInformationWebPartProps> {

  /**
   * Metodo inicial del webpart
   * @protected
   * @returns {Promise<void>}
   * @memberof VisoWebPart
   */
  protected async onInit(): Promise<void> {
    try {
      //Configuracion inicial de los servicios de SP
      ServiceETools.spSetup(this.context);

      //servicio
      // this._entityNameService = new EntityNameService(IDLists.MYLIST/*listname*/, this.context);

    } catch (err) {
      Logger.log({ data: err, level: LogLevel.Error, message: "VisorWebPart > onInit > " });
    }
    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<IPersonalInformationProps> = React.createElement(
      PersonalInformation,
      {
        description: this.properties.description
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

  protected get disableReactivePropertyChanges(): boolean {
    return true;
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
