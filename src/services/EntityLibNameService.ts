import "@pnp/polyfill-ie11";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
//import { sp } from "@pnp/sp";
//locals
import { LibDocService, IQueryable } from "./core/LibDocService";
import { IListItem } from "./../models/IListItem";

/**
 * Servicio que contiene los metodos necesarios de conexión con los servicios de sharePoint para una librería "libName"
 * @see {@link https://www.npmjs.com} @pnp/sp
 * @copyright 2018 e-deas (http://www.e-deas.com.co)  El uso de esta libreria esta reservador para este sitio y cualquier cambio o reutilización debe ser autorizado por e-deas.
 * @author Diego Campo <diegoc@e-deas.com.co> / Fecha: 14.01.2019 - Creado
 * @author Diego Campo <diegoc@e-deas.com.co> / Fecha: 23.01.2019 - Modificado
 *
 * @export
 * @class EntityLibNameService
 * @extends {LibDocService}
 */
export class EntityLibNameService extends LibDocService {
    private _mycontext: WebPartContext | ApplicationCustomizerContext;

    /**
     * Crear una instancia de EntityLibNameService
     * @param {string} listName
     * @param {(WebPartContext | ApplicationCustomizerContext)} context
     * @memberof EntityLibNameService
     */
    public constructor(listName: string, context: WebPartContext | ApplicationCustomizerContext) {
        super(listName, context);

        this._mycontext = context; 
    }



}
