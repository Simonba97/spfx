import { IPersonalInformation } from "../../../models/IPersonalInformation";

export interface IPersonalInformationState {
    //elementos a renderizar
    items: IPersonalInformation[];
    showPanelDetail:boolean;
    itemDetail: IPersonalInformation;
}
