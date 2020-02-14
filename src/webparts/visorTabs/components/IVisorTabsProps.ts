import { WebPartContext } from "@microsoft/sp-webpart-base";
import { TabsInvormativosServices } from '../../../services/TabsInvormativosServices';

export interface IVisorTabsProps {
  dropdownTypeVisor: string;
  toggleSearchInfo: boolean;
  collectionData: any[];
  collectionDataDinamic: any[];
  textUrlSite: string;
  dropdownListSelected: string | string[]; 
  textNameTitleFld: string;
  textNameContentFld: string;
  numberCantElements: number;
  context: WebPartContext;
  tabsInformativosServices: TabsInvormativosServices;
}
