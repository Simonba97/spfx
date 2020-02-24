import { WebPartContext } from "@microsoft/sp-webpart-base";
import { PersonalInformationServices } from "../../../services/PersonalInformationServices";

export interface IPersonalInformationProps {
  _personalInformationServices: PersonalInformationServices;
  context: WebPartContext;

  urlList: string;
  toggleWithPhoto: boolean;
  fieldsShow: string[];
}
