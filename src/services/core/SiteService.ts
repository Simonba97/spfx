import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
/**
 * Clase que ofrece diferentes servicios de sitio
 *
 * @see {@link https://www.npmjs.com} @pnp/sp
 * @copyright 2018 e-deas (http://www.e-deas.com.co)  El uso de esta libreria esta reservador para este sitio y cualquier cambio o reutilización debe ser autorizado por e-deas.
 * @author Diego Campo <diegoc@e-deas.com.co> / Fecha: 19.02.2019 - Creado
 * @author Diego Campo <diegoc@e-deas.com.co> / Fecha: 19.02.2019 - Modificado
 *
 * @export
 * @class SiteService
 */
export class SiteService {

    private _context: WebPartContext | ApplicationCustomizerContext;


    /**
      * Crear una instancia de SiteService
      * @param {(WebPartContext | ApplicationCustomizerContext)} context
      * @memberof SiteService
      */
    public constructor(context: WebPartContext | ApplicationCustomizerContext) {
        this._context = context;
    }

}