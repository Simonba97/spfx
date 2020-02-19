import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
} from '@microsoft/sp-webpart-base';

import VisorTabs from './components/VisorTabs';
import { IVisorTabsProps } from './components/IVisorTabsProps';

import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldDropdownWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldDropdownWithCallout';
import { PropertyFieldToggleWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { PropertyFieldTextWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldTextWithCallout';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { IPropertyPaneField } from '@microsoft/sp-property-pane';


import { Logger, LogLevel } from '@pnp/logging';
import { TabsInvormativosServices } from '../../services/TabsInvormativosServices';
import { ServiceETools } from '../../utils/ServiceETools';
import { IDLists } from "../../utils/IDLists";

import { sp } from "@pnp/sp";

export interface IVisorTabsWebPartProps {
  dropdownListSelected: string | string[]; // Stores the list ID(s)
  dropdownTypeVisor: string;
  toggleSearchInfo: boolean;
  collectionData: any[];
  collectionDataDinamic: any[];
  textUrlSite: string;
  textNameTitleFld: string;
  textNameContentFld: string;
  numberCantElements: number;  
  textFilterBy: string;
  operatorFilterBy: string;
  textValueFilter: string;
  orderBy: string;
  toggleAsc: boolean;
}

export default class VisorTabsWebPart extends BaseClientSideWebPart<IVisorTabsWebPartProps> {

  //servicio
  private _tabsInformativosServices: TabsInvormativosServices;
  private _dinamycPropertyPaneField: IPropertyPaneField<any>[];

  /**
    * Metodo inicial del webpart
    * @protected
    * @returns {Promise<void>}
    * @memberof VisorTabsWebPart
    */
  protected async onInit(): Promise<void> {
    try {
      //Configuracion inicial de los servicios de SP
      ServiceETools.spSetup(this.context);

      //servicio
     // this._tabsInformativosServices = new TabsInvormativosServices(IDLists.tabsinformativos, this.context);//inicializamos servicios
    } catch (err) {
      Logger.log({ data: err, level: LogLevel.Error, message: "VisorWebPart > onInit > " });
    }
    return Promise.resolve();
  }

  public render(): void {
    if (this.properties.dropdownListSelected) {
      this._tabsInformativosServices = new TabsInvormativosServices(this.properties.dropdownListSelected.toString(), this.context);
    }
    const element: React.ReactElement<IVisorTabsProps> = React.createElement(
      VisorTabs,
      {
        tabsInformativosServices: this._tabsInformativosServices,
        context: this.context,

        dropdownTypeVisor: this.properties.dropdownTypeVisor,
        toggleSearchInfo: this.properties.toggleSearchInfo,
        collectionData: this.properties.collectionData,
        collectionDataDinamic: this.properties.collectionDataDinamic,
        textUrlSite: this.properties.textUrlSite,
        dropdownListSelected: this.properties.dropdownListSelected,
        textNameTitleFld: this.properties.textNameTitleFld,
        textNameContentFld: this.properties.textNameContentFld,
        numberCantElements: this.properties.numberCantElements,
        textFilterBy: this.properties.textFilterBy,
        operatorFilterBy: this.properties.operatorFilterBy,
        textValueFilter: this.properties.textValueFilter,
        orderBy: this.properties.orderBy,
        toggleAsc: this.properties.toggleAsc,
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

  protected async getListsSite() {
    return await sp.web.lists.select('Title, Id')
      .filter('IsCatalog eq false and Hidden eq false and BaseTemplate eq 100').get();
  }

  protected async getListSiteById(idList, strSelect, intTop) {
    return await sp.web.lists.getById(idList).items
      .select(strSelect)
      .orderBy('orden')
      .top(intTop)
      .get();
  }  

  protected setDinamicPropertyPaneField() {
    if (this.properties.toggleSearchInfo) {
      this._dinamycPropertyPaneField = [
        PropertyFieldTextWithCallout('textUrlSite', {
          calloutTrigger: CalloutTriggers.Hover,
          key: 'textUrlSiteId',
          label: 'URL del sitio',
          value: this.properties.textUrlSite
        }),
        PropertyFieldListPicker('dropdownListSelected', {
          label: 'Lista ',
          selectedList: this.properties.dropdownListSelected,
          includeHidden: false,
          orderBy: PropertyFieldListPickerOrderBy.Title,
          disabled: false,
          onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
          properties: this.properties,
          context: this.context,
          onGetErrorMessage: null,
          deferredValidationTime: 0,
          key: 'listPickerFieldId',
          webAbsoluteUrl: this.properties.textUrlSite ? this.properties.textUrlSite : '',
          baseTemplate: 100,
  
        }),   
        PropertyFieldTextWithCallout('textNameTitleFld', {
          calloutTrigger: CalloutTriggers.Hover,
          key: 'textNameTitleFldId',
          label: 'Título',
          placeholder: 'Título',
          description: 'Nombre interno del campo',
          value: this.properties.textNameTitleFld
        }),
        PropertyFieldTextWithCallout('textNameContentFld', {
          calloutTrigger: CalloutTriggers.Hover,
          key: 'textNameContentFldId',
          label: 'Contenido',
          placeholder: 'Contenido',
          description: 'Nombre interno del campo',
          value: this.properties.textNameContentFld
        }),
        PropertyFieldNumber("numberCantElements", {
          key: "numberCantElements",
          label: "Cantidad de elementos",
          placeholder: 'Cantidad de elementos',
          value: this.properties.numberCantElements,
          minValue: 1,
          disabled: false
        }),
        PropertyFieldTextWithCallout('textFilterBy', {
          calloutTrigger: CalloutTriggers.Hover,
          key: 'textFilterBy',
          label: 'Filtrado por',
          description: 'Nombre interno del campo',
          value: this.properties.textFilterBy
        }),
        PropertyFieldDropdownWithCallout('operatorFilterBy', {
          calloutTrigger: CalloutTriggers.Hover,
          key: 'operatorFilterBy',
          label: "Tipo de comparación",
          options: [{
            key: 'eq',
            text: "Igual a"
          }, {
            key: 'ne',
            text: "Diferente a"
          }, {
            key: 'lt',
            text: "Menor que"
          }, {
            key: 'le',
            text: "Menor igual que"
          }, {
            key: 'gt',
            text: "Mayor que"
          }, {
            key: 'ge',
            text: "Mayor igual que"
          }],
          selectedKey: this.properties.operatorFilterBy,
        }),
        PropertyFieldTextWithCallout('textValueFilter', {
          calloutTrigger: CalloutTriggers.Hover,
          key: 'textValueFilter',
          label: 'Valor a filtrar',
          value: this.properties.textValueFilter
        }),
        PropertyFieldTextWithCallout('orderBy', {
          calloutTrigger: CalloutTriggers.Hover,
          key: 'orderBy',
          label: 'Ordenado por',
          placeholder: 'Title',
          description: 'Nombre interno del campo',
          value: this.properties.orderBy
        }),
        PropertyFieldToggleWithCallout('toggleAsc', {
          calloutTrigger: CalloutTriggers.Click,
          key: 'toggleAscId',
          label: "Descendente/Ascendente",
          calloutContent: React.createElement('p', {}, 'With this control you can enable or disable the PnP features in your web part'),
          onText: 'Asc',
          offText: 'Desc',
          checked: this.properties.toggleAsc
        })
      ];     
    } else {
      this._dinamycPropertyPaneField = [
        PropertyFieldCollectionData("collectionData", {
          key: "collectionData",
          label: "Collection data",
          panelHeader: "Collection data panel header",
          manageBtnLabel: "Manage collection data",
          value: this.properties.collectionData,
          fields: [
            {
              id: "title",
              title: "Título",
              type: CustomCollectionFieldType.string,
              required: true
            },
            {
              id: "contenido",
              title: "Contenido",
              type: CustomCollectionFieldType.string
            },
            {
              id: "orden",
              title: "Orden",
              type: CustomCollectionFieldType.number,
              required: true
            },
          ],
          disabled: false
        })
      ];
    }
  } // end getProperties()

  // protected get disableReactivePropertyChanges(): boolean {
  //   return true;
  // }

  public getPropertyPaneConfiguration(): IPropertyPaneConfiguration { // onchange panel de propiedades
    let groupNameMoreOptions = this.properties.toggleSearchInfo ? "Basados en datos dinámicos" : "Basados en datos estáticos";
    this.setDinamicPropertyPaneField(); // set fields to properties
    return {
      pages: [
        {
          header: {
            description: "Permite agregar un visor en forma de pestañas",
          },
          groups: [
            {
              groupName: "Configuración",
              groupFields: [
                PropertyFieldDropdownWithCallout('dropdownTypeVisor', {
                  calloutTrigger: CalloutTriggers.Hover,
                  key: 'dropdownTypeVisorId',
                  label: "Estilo de las pestañas",
                  options: [{
                    key: 'primaryTheme',
                    text: "Primario"
                  }, {
                    key: 'secundaryTheme',
                    text: "Secundario"
                  }],
                  selectedKey: this.properties.dropdownTypeVisor,
                }),
                PropertyFieldToggleWithCallout('toggleSearchInfo', {
                  calloutTrigger: CalloutTriggers.Click,
                  key: 'toggleSearchInfoId',
                  label: 'Mostrar datos a partir de una',
                  calloutContent: React.createElement('p', {}, 'With this control you can enable or disable the PnP features in your web part'),
                  onText: 'Lista',
                  offText: 'Colección de datos',
                  checked: this.properties.toggleSearchInfo
                })
              ]
            },
            {
              groupName: groupNameMoreOptions,
              groupFields: [...this._dinamycPropertyPaneField]
            }
          ]
        },
      ]
    };
  }
}