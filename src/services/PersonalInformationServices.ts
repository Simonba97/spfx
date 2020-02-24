import "@pnp/polyfill-ie11";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
//import { sp } from "@pnp/sp";
//locals
import { ListItemService, IQueryable } from "./core/ListItemService";
import { IPersonalInformation } from "./../models/IPersonalInformation";

/**
 * Servicio que contiene los metodos necesarios de conexión con los servicios de sharePoint para una lista "listName"
 * @see {@link https://www.npmjs.com} @pnp/sp
 * @copyright 2018 e-deas (http://www.e-deas.com.co)  El uso de esta libreria esta reservador para este sitio y cualquier cambio o reutilización debe ser autorizado por e-deas.
 * @author Diego Campo <diegoc@e-deas.com.co> / Fecha: 14.01.2019 - Creado
 * @author Diego Campo <diegoc@e-deas.com.co> / Fecha: 23.01.2019 - Modificado
 *
 * @export
 * @class PersonalInformationServices
 * @extends {ListItemService}
 */
export class PersonalInformationServices extends ListItemService {
    private _mycontext: WebPartContext | ApplicationCustomizerContext;

    /**
     * Crear una instancia de PersonalInformationServices
     * @param {string} listName
     * @param {(WebPartContext | ApplicationCustomizerContext)} context
     * @memberof PersonalInformationServices
     */
    public constructor(listName: string, context: WebPartContext | ApplicationCustomizerContext) {
        super(listName, context);

        this._mycontext = context;
    }

    /**
   * Funcion para obtener los elementos con la estructura base de los momentos
   */
    public async getItemsPrincipalPanel(fieldsShow: string[], withPhoto) {
        //Query para traer los datos necesarios para la estructura base de los momentos
        const query: IQueryable = {
            select: `ID, Title ${fieldsShow.length > 0 ? ',' + fieldsShow.join(',') : ''} ${withPhoto ? ', foto' : ''}`,
            order: {
                asc: true,
                by: 'ID'
            }
        };

        //Obtener elementos con la consulta definida
        const itemsPersons: IPersonalInformation[] = await this.getItems(query);

        return itemsPersons;
    } //  end getItemsPrincipalPanel

    /**
     * Obtener el último id de la lista
     * @returns {Promise<number>}
     * @memberof PersonalInformationServices
    */
    public async getLatestItemId(): Promise<number> {
        let query: IQueryable = {
            select: "Id",
            order: { by: "Id", asc: false },
            top: 1
        };

        //Opcion 1
        const items: IPersonalInformation[] = await this.getItems(query);
        //Opcion 2
        //const items: IPersonalInformation[] = await sp.web.lists.getByTitle(this.getListName()).items.orderBy('Id', false).top(1).select('Id').get();

        if (items.length === 0) {
            return -1;
        } else {
            return items[0].Id;
        }
    }



    /**
     * Obtener el último id de la lista
     * @returns {Promise<number>}
     * @memberof PersonalInformationServices
    */
    public async getDetailById(query?: IQueryable): Promise<IPersonalInformation> {

        //Opcion 1
        const items: IPersonalInformation[] = await this.getItems(query);
        //Opcion 2
        //const items: IPersonalInformation[] = await sp.web.lists.getByTitle(this.getListName()).items.orderBy('Id', false).top(1).select('Id').get();

        return items[0];
    }

    //TOFIX: Restructurar funcion
    /*public async multiple() {
  
      var promise1 = sp.web.lists.getByTitle(this.getListName()).items.filter("Title eq 'Uno'");
      var promise2 = sp.web.lists.getByTitle(this.getListName()).items.filter("Title eq 'Dos'");;
      var promise3 = sp.web.lists.getByTitle(this.getListName()).items.filter("Title eq 'Tres'");
  
      const calls = await Promise.all([promise1, promise2, promise3]);
      return calls;
    }*/

    //TODO: getByField
    //TODO: getByLookupId
    //TODO: getByManyLookups

}
