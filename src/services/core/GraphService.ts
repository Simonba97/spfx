import "@pnp/polyfill-ie11";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { graph } from "@pnp/graph";
import { dateAdd } from "@pnp/common";
import { Logger, LogLevel } from '@pnp/logging';
import { User } from '@microsoft/microsoft-graph-types';
//locals


/**
 * Clase que ofrece diferentes servicios de Graph
 *
 * @see {@link https://www.npmjs.com} @pnp/graph
 * @copyright 2018 e-deas (http://www.e-deas.com.co)  El uso de esta libreria esta reservador para este sitio y cualquier cambio o reutilización debe ser autorizado por e-deas.
 * @author John Freddy Torres <jhont@e-deas.com.co> / Fecha: 19.02.2019 - Creado
 * @author Diego Campo <diegoc@e-deas.com.co> / Fecha: 19.02.2019 - Modificado
 *
 * @export
 * @class GraphService
 */
export class GraphService {

    //contexto del webpart
    private _context: WebPartContext | ApplicationCustomizerContext;


    /**
    * Crear una instancia de GraphService
    * @memberof GraphService
    */
    public constructor(context: WebPartContext | ApplicationCustomizerContext) {
        this._context = context;

    }

    /**
     * Obtener datos de "me" sobre graph
     * Id,Title,userPrincipalName,jobTitle,onPremisesExtensionAttributes,department
     *
     * @returns {Promise<IUserGraph>}
     * @memberof GraphService
     */
    public async getMe(): Promise<User> {

        let me = graph.me;

        //filtro
        me.select("Id,displayName,userPrincipalName,jobTitle,department,onPremisesExtensionAttributes");
        me.usingCaching({
            expiration: dateAdd(new Date(), "minute", 15),//second
            key: me.toUrlAndQuery()
        });

        let user: User = await me.get();

        Logger.log({ data: user, level: LogLevel.Warning, message: "GraphService > getMe > " });

        return user;

    }





}
