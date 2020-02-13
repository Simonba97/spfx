import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { sp, SearchQuery, SearchResults } from "@pnp/sp";
//Locals
import { IResult } from "./../../models/IResult";
/**
 * Clase que ofrece diferentes servicios de busqueda
 *
 * @see {@link https://www.npmjs.com} @pnp/sp
 * @copyright 2018 e-deas (http://www.e-deas.com.co)  El uso de esta libreria esta reservador para este sitio y cualquier cambio o reutilización debe ser autorizado por e-deas.
 * @author Diego Campo <diegoc@e-deas.com.co> / Fecha: 19.02.2019 - Creado
 * @author Diego Campo <diegoc@e-deas.com.co> / Fecha: 19.02.2019 - Modificado
 *
 * @export
 * @class SiteService
 */
export class SearchService {

    private _context: WebPartContext | ApplicationCustomizerContext;


    /**
      * Crear una instancia de SearchService
      * @param {(WebPartContext | ApplicationCustomizerContext)} context
      * @memberof SearchService
      */
    public constructor(context: WebPartContext | ApplicationCustomizerContext) {
        this._context = context;
    }


    public async searchAny(querytext: string): Promise<IResult[]> {
        const result: SearchResults = await sp.search(querytext);
        return result.PrimarySearchResults as IResult[];
    }
}
