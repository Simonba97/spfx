import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { sp } from "@pnp/sp";
import { Logger, LogLevel, ConsoleListener } from '@pnp/logging';
import { ICachingOptions } from "@pnp/odata";
//import { graph } from "@pnp/graph";

/**
 * Clase que ofrece diferentes utilidades a nivel de servicios
 *
 * @see {@link https://www.npmjs.com} @pnp/sp
 * @copyright 2018 e-deas (http://www.e-deas.com.co)  El uso de esta libreria esta reservador para este sitio y cualquier cambio o reutilización debe ser autorizado por e-deas.
 * @author Diego Campo <diegoc@e-deas.com.co> / Fecha: 14.01.2019 - Creado
 * @author Diego Campo <diegoc@e-deas.com.co> / Fecha: 23.01.2019 - Modificado
 *
 * @export
 * @class ServiceETools
 */
export abstract class ServiceETools {

    /**
     * Se configura el headers y otros elementos de pnpjs
     *
     * @static
     * @memberof ServiceETools
     */
    public static spSetup(context: WebPartContext | ApplicationCustomizerContext): void {
        sp.setup({
            sp: {
                headers: {
                    "Accept": "application/json; odata=nometadata"//verbose, nometadata
                },
                baseUrl: context.pageContext.web.absoluteUrl
            }
        });

        let isDebug = window.location.href.indexOf("debug") > 0 ? true : false;

        if (isDebug) {
            //Only for Debug
            Logger.activeLogLevel = LogLevel.Info;
        } else {
            //Only for Production
            Logger.activeLogLevel = LogLevel.Error;
        }

        const consoleListener = new ConsoleListener();
        Logger.subscribe(consoleListener);

    }

    /**
     * Se configura el headers y otros elementos de pnpjs
     *
     * @static
     * @memberof ServiceETools
     */
    /*public static graphSetup(context: WebPartContext | ApplicationCustomizerContext): void {
        graph.setup({
            spfxContext: context
        });

        let isDebug = window.location.href.indexOf("debug") > 0 ? true : false;

        if (isDebug) {
            //Only for Debug
            Logger.activeLogLevel = LogLevel.Info;
        } else {
            //Only for Production
            Logger.activeLogLevel = LogLevel.Error;
        }

        const consoleListener = new ConsoleListener();
        Logger.subscribe(consoleListener);

    }*/

    /**
     * Permite realizar un delay
     *
     * @static
     * @param {number} ms
     * @memberof ServiceETools
     */
    public static async delay(ms: number) {
        await new Promise(resolve => setTimeout(() => resolve(), ms)).then(() => console.log(ms + " end delay"));
    }

    /**
   * Obtener elemento de cache basado en ICachingOptions
   *
   * @static
   * @param {ICachingOptions} options
   * @returns {*}
   * @memberof ServiceETools
   */
    public static getCaching(options: ICachingOptions): any {
        let storage = null;
        if (options.storeName === "local") {
            storage = localStorage;
        } else {
            storage = sessionStorage;
        }

        let data = null;
        let dataStorage = null;
        try {
            dataStorage = storage.getItem(options.key);

            //if not empty get storage
            if (dataStorage != null && dataStorage != "") {
                dataStorage = JSON.parse(dataStorage);
                var expiration = dataStorage.expiration;
                data = dataStorage.value;//get data

                if (new Date(expiration) < new Date()) {
                    data = null;
                }
            }


        } catch (e) {
            data = null;
        }

        return data;
    }


    /**
     * ASignar elemento de cache basado en ICachingOptions
     *
     * @static
     * @param {ICachingOptions} options
     * @param {*} data
     * @memberof ServiceETools
     */
    public static setCaching(options: ICachingOptions, data: any): void {
        let storage = null;
        if (options.storeName === "local") {
            storage = localStorage;
        } else {
            storage = sessionStorage;
        }
        try {
            //storage
            let fullData = {
                value: data,
                expiration: options.expiration
            };
            storage.setItem(options.key, JSON.stringify(fullData));
        } catch (e) { }
    }

}
