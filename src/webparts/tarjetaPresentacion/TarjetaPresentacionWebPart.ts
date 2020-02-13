import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'TarjetaPresentacionWebPartStrings';
import TarjetaPresentacion from './components/TarjetaPresentacion';
import { ITarjetaPresentacionProps } from './components/ITarjetaPresentacionProps';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

export interface ITarjetaPresentacionWebPartProps {
  nombre: string;
  imagen: string;
  descripcion: string;
  redes: any[];
}

export default class TarjetaPresentacionWebPart extends BaseClientSideWebPart<ITarjetaPresentacionWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ITarjetaPresentacionProps > = React.createElement(
      TarjetaPresentacion,
      {
        nombre: this.properties.nombre,
        imagen: this.properties.imagen,
        descripcion: this.properties.descripcion,
        redes: this.properties.redes
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
                PropertyPaneTextField('nombre', {
                  label: "Nombre del empleado"
                }),
                PropertyPaneTextField('descripcion', {
                  label: "¿Quién eres?"
                }),
                PropertyPaneTextField('imagen', {
                  label: "Imagen de perfil",
                  resizable: true
                }),
                PropertyFieldCollectionData("redes", {
                  key: "redes",
                  label: "Redes sociales",
                  panelHeader: "Agrega, quita o modifica tus redes sociales",
                  manageBtnLabel: "Gestionar redes sociales",
                  value: this.properties.redes,
                  fields: [
                    {
                      id: "Icono",
                      title: "Nombre del Icono",
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: "URL",
                      title: "Url de la red",
                      type: CustomCollectionFieldType.string
                    }                   
                  ],
                  disabled: false
                })             
              ]
            }
          ]
        }
      ]
    };
  }
}
