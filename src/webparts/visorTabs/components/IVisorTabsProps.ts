import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IVisorTabsProps {
  dropdownTypeVisor: string;
  toggleSearchInfo: boolean;
  collectionData: any[];
  collectionDataDinamic: any[];
  textUrlSite: string;
  dropdownListSelected: string;
  textNameTitleFld: string;
  textNameContentFld: string;
  numberCantElements: number;
  context: WebPartContext;
}
