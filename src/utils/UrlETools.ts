

/**
 * Clase que ofrece diferentes utilidades a nivel de URL
 *
 * @copyright 2018 e-deas (http://www.e-deas.com.co)  El uso de esta libreria esta reservador para este sitio y cualquier cambio o reutilización debe ser autorizado por e-deas.
 * @author Diego Campo <diegoc@e-deas.com.co> / Fecha: 19.02.2019 - Creado
 * @author Diego Campo <diegoc@e-deas.com.co> / Fecha: 19.02.2019 - Modificado
 *
 * @export
 * @class UrlETools
 */
export abstract class UrlETools {


	/**
   * Validar si el tenant tiene licencia para visualizar la app
   *
   * @static
   * @returns {boolean}
   * @memberof UrlETools
   */
    public static isAvailTenant(): boolean {
        let isAvail = false;

        let tenants: string[] = [
            "https://bancolombia.sharepoint",
            "https://ambientesbc.sharepoint",
            "https://edeasco.sharepoint",
            "https://localhost:4321"
        ];

        let url = window.location.href.toLowerCase();
        for (let i = 0; i < tenants.length; i++) {
            const tenant = tenants[i].toLocaleLowerCase();
            if (url.indexOf(tenant) == 0) {
                isAvail = true;
                break;
            }
        }

        return isAvail;
    }

    /**
     * Get the value of a querystring
     * @param  {String} field The field to get the value of
     * @param  {String} url   The URL to get the value from (optional)
     * @return {String}       The field value
     */
    public static getQueryStringParam(field: string, url: string) {
        const href = url ? url : window.location.href;
        const reg = new RegExp("[?&#]" + field + "=([^&#]*)", "i");
        const qs = reg.exec(href);
        return qs ? qs[1] : null;
    }

    /**
     * @param {String} field The field name of the query string to remove
     * @param {String} sourceURL The source URL
     * @return {String}       The updated URL
     */
    public static removeQueryStringParam(field: string, sourceURL: string) {
        let rtn = sourceURL.split("?")[0];
        let param = null;
        let paramsArr = [];
        const queryString = (sourceURL.indexOf("?") !== -1) ? sourceURL.split("?")[1] : "";

        if (queryString !== "") {
            paramsArr = queryString.split("&");
            for (let i = paramsArr.length - 1; i >= 0; i -= 1) {
                param = paramsArr[i].split("=")[0];
                if (param === field) {
                    paramsArr.splice(i, 1);
                }
            }

            if (paramsArr.length > 0) {
                rtn = rtn + "?" + paramsArr.join("&");
            }
        }
        return rtn;
    }

    /**
     * Add or replace a query string parameter
     * @param url The current URL
     * @param param The query string parameter to add or replace
     * @param value The new value
     */
    public static addOrReplaceQueryStringParam(url: string, param: string, value: string) {
        const re = new RegExp("[\\?&]" + param + "=([^&#]*)");
        const match = re.exec(url);
        let delimiter;
        let newString;

        if (match === null) {
            // Append new param
            const hasQuestionMark = /\?/.test(url);
            delimiter = hasQuestionMark ? "&" : "?";
            newString = url + delimiter + param + "=" + value;
        } else {
            delimiter = match[0].charAt(0);
            newString = url.replace(re, delimiter + param + "=" + value);
        }

        return newString;
    }

	/**
	   * Obtiene la url relativa
	   *
	   * @static
	   * @returns
	   * @memberof UrlETools
	   */
    public static getRelativeUrl() {
        let server = window.location.protocol + "//" + window.location.hostname;
        return window.location.pathname.replace(server, "");
    }

}