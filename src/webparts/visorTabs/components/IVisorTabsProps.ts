import { WebPartContext } from "@microsoft/sp-webpart-base";
import { TabsInvormativosServices } from '../../../services/TabsInvormativosServices';

export interface IVisorTabsProps {
  tabsInformativosServices: TabsInvormativosServices;
  context: WebPartContext;

  dropdownTypeVisor: string;
  toggleSearchInfo: boolean;
  collectionData: any[];
  collectionDataDinamic: any[];
  textUrlSite: string;
  dropdownListSelected: string | string[]; 
  textNameTitleFld: string;
  textNameContentFld: string;
  numberCantElements: number;
  textFilterBy: string;
  operatorFilterBy: string;
  textValueFilter: string;
  orderBy: string;
  toggleAsc: boolean;
}
