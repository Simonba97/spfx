import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldToggleWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout';

import { Logger, LogLevel } from '@pnp/logging';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PersonalInformationWebPartStrings';
import PersonalInformation from './components/PersonalInformation';
import { IPersonalInformationProps } from './components/IPersonalInformationProps';
import { ServiceETools } from '../../utils/ServiceETools';
import { IDLists } from '../../utils/IDLists';
import { PersonalInformationServices } from '../../services/PersonalInformationServices';

export interface IPersonalInformationWebPartProps {
  urlList: string;
  toggleWithPhoto: boolean;
  fieldsShow: string[];
}

export default class PersonalInformationWebPart extends BaseClientSideWebPart<IPersonalInformationWebPartProps> {

  private _personalInformationServices: PersonalInformationServices;


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
    if (this.properties.urlList) {
      this._personalInformationServices = new PersonalInformationServices(this.properties.urlList.toString(), this.context);
    }
    const element: React.ReactElement<IPersonalInformationProps> = React.createElement(
      PersonalInformation,
      {
        _personalInformationServices: this._personalInformationServices,
        context: this.context,

        urlList: this.properties.urlList,
        toggleWithPhoto: this.properties.toggleWithPhoto,
        fieldsShow: this.properties.fieldsShow,
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
                PropertyPaneTextField('urlList', {
                  label: "Url del sitio",
                  placeholder: "https://edeasco.sharepoint.com/sites/ed-pers-simonb/",
                }),
                PropertyFieldMultiSelect('fieldsShow', {
                  key: 'fieldsShow',
                  label: "Qué desea ver",
                  options: [
                    // {
                    //   key: "Title",
                    //   text: "Nombre completo"
                    // },
                    {
                      key: "departamento",
                      text: "Departamento"
                    },
                    {
                      key: "correo",
                      text: "Correo electrónico"
                    },
                    // {
                    //   key: "descripcion",
                    //   text: "Descripción"
                    // },
                    // {
                    //   key: "telefono",
                    //   text: "Teléfono"
                    // },
                    // {
                    //   key: "direccion",
                    //   text: "Dirección"
                    // },
                    // {
                    //   key: "edad",
                    //   text: "Edad"
                    // }
                  ],
                  selectedKeys: this.properties.fieldsShow
                }),
                PropertyFieldToggleWithCallout('toggleWithPhoto', {
                  calloutTrigger: CalloutTriggers.Click,
                  key: 'toggleWithPhotoId',
                  label: 'Con foto',
                  calloutContent: React.createElement('p', {}, 'Puede elegir si desea ver la información del usuario con o sin foto'),
                  onText: 'Sí',
                  offText: 'No',
                  checked: this.properties.toggleWithPhoto
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
