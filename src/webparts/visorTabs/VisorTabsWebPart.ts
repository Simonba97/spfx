import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'VisorTabsWebPartStrings';
import VisorTabs from './components/VisorTabs';
import { IVisorTabsProps } from './components/IVisorTabsProps';

import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldDropdownWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldDropdownWithCallout';
import { PropertyFieldToggleWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { PropertyFieldTextWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldTextWithCallout';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';

import { Logger, LogLevel } from '@pnp/logging';
import { TabsInvormativosServices } from '../../services/TabsInvormativosServices';
import { ServiceETools } from '../../utils/ServiceETools';
import { IDLists } from "../../utils/IDLists";

import { sp } from "@pnp/sp";
import { LocalizedFontFamilies } from '@uifabric/styling/lib/styles/fonts';
import { func } from 'prop-types';

export interface IVisorTabsWebPartProps {
  dropdownTypeVisor: string;
  toggleSearchInfo: boolean;
  collectionData: any[];
  collectionDataDinamic: any[];
  textUrlSite: string;
  dropdownListSelected: string;
  textNameTitleFld: string;
  textNameContentFld: string;
  numberCantElements: number;
}

export default class VisorTabsWebPart extends BaseClientSideWebPart<IVisorTabsWebPartProps> {

  //servicio
  private _tabsInformativosServices: TabsInvormativosServices;
  private _listsSite = [];

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
          this._tabsInformativosServices = new TabsInvormativosServices(IDLists.tabsinformativos, this.context);
          this.congigureInitial();
          const items = await this._tabsInformativosServices.getItems(); // Render

        } catch (err) {
          Logger.log({ data: err, level: LogLevel.Error, message: "VisorWebPart > onInit > " });
      }
      return Promise.resolve();
    }

  public render(): void {
    const element: React.ReactElement<IVisorTabsProps > = React.createElement(
      VisorTabs,
      {
        dropdownTypeVisor: this.properties.dropdownTypeVisor,
        toggleSearchInfo: this.properties.toggleSearchInfo,
        collectionData: this.properties.collectionData,
        collectionDataDinamic: this.properties.collectionDataDinamic,
        textUrlSite: this.properties.textUrlSite,
        dropdownListSelected: this.properties.dropdownListSelected,
        textNameTitleFld: this.properties.textNameTitleFld,
        textNameContentFld: this.properties.textNameContentFld,
        numberCantElements: this.properties.numberCantElements,
        context: this.context,
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

  protected async congigureInitial() {
    if(this.properties.toggleSearchInfo && this.properties.collectionDataDinamic){
      this.getListsSite()
      .then((lists)=>{
          console.warn(lists);

          this._listsSite = []; //Clear array information
          
          lists.forEach((l)=>{
            this._listsSite.push({
              key: l.Id,
              text: l.Title,
            });//end push

          }) //end forEach
          const idListSelected = this.properties.dropdownListSelected;
          const strSelect = `${this.properties.textNameTitleFld}, ${this.properties.textNameContentFld}`;
          const intTop = this.properties.numberCantElements;
          this.getListSiteById(idListSelected, strSelect, intTop)
            .then((items: any[])=>{
              this.properties.collectionDataDinamic = items;
              console.log(items);
            })
            .catch((e)=>{
              console.error(`Hubo un error: ${e.message}`);
            })
        }); // end then 
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration { // onchange panel de propiedades
    this.congigureInitial();
    let groupNameMoreOptions = "";
    let arrayGroupFields = [];
    if(this.properties.toggleSearchInfo) {
      arrayGroupFields = [];
      groupNameMoreOptions = "Basados en una lista";
      arrayGroupFields.push(PropertyFieldTextWithCallout('textUrlSite', {
                              calloutTrigger: CalloutTriggers.Hover,
                              key: 'textUrlSiteId',
                              label: 'URL del sitio',
                              value: this.properties.textUrlSite
                            }),
                            PropertyFieldDropdownWithCallout('dropdownListSelected', {
                              calloutTrigger: CalloutTriggers.Hover,
                              key: 'dropdownListSelectedId',
                              label: "Lista",
                              options: this._listsSite,
                              selectedKey: this.properties.dropdownListSelected,
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
                            }));
    } else {
      arrayGroupFields = [];
      groupNameMoreOptions = "Basados en datos estáticos";
      arrayGroupFields.push(PropertyFieldCollectionData("collectionData", {
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
                            }));
    };  

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
              groupFields: arrayGroupFields,
            }
          ]
        },        
      ]
    };
  }
}
