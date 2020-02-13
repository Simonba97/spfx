//import * as strings from 'VisorDocsRecientesWebPartStrings';
import { Text } from '@microsoft/sp-core-library';
/**
 * Clase que ofrece diferentes utilidades a nivel de formato
 *
 * @copyright 2018 e-deas (http://www.e-deas.com.co)  El uso de esta libreria esta reservador para este sitio y cualquier cambio o reutilización debe ser autorizado por e-deas.
 * @author Diego Campo <diegoc@e-deas.com.co> / Fecha: 19.02.2019 - Creado
 * @author Diego Campo <diegoc@e-deas.com.co> / Fecha: 19.02.2019 - Modificado
 *
 * @export
 * @class FormatETools
 */
export abstract class FormatETools {

    private static _months: string[] = [/*strings.LblMonthEne, strings.LblMonthFeb, strings.LblMonthMar, strings.LblMonthAbr, strings.LblMonthMay, strings.LblMonthJun, strings.LblMonthJul, strings.LblMonthAgo, strings.LblMonthSep, strings.LblMonthOct, strings.LblMonthNov, strings.LblMonthDic*/];

    /**
     * Returns the relative date for the document activity
     *
     * @static
     * @param {string} crntDate
     * @returns {string}
     * @memberof FormatETools
     */
    /*public static getRelativeDate(crntDate: string): string {
      //const date = new Date((crntDate || "").replace(/-/g, "/").replace(/[TZ]/g, " "));
      const date = new Date(crntDate);
      const diff = (((new Date()).getTime() - date.getTime()) / 1000);
      const day_diff = Math.floor(diff / 86400);
  
      if (isNaN(day_diff) || day_diff < 0) {
        return;
      }
  
      return day_diff === 0 && (
        diff < 60 && strings.LblDateJustNow ||
        diff < 120 && strings.LblDateMinute ||
        diff < 3600 && Text.format(strings.LblDateMinutesAgo, `${Math.floor(diff / 60)}`) ||
        diff < 7200 && strings.LblDateHour ||
        diff < 86400 && Text.format(strings.LblDateHoursAgo, `${Math.floor(diff / 3600)}`)) ||
        day_diff == 1 && strings.LblDateDay ||
        day_diff <= 30 && Text.format(strings.LblDateDaysAgo, `${day_diff}`) ||
        day_diff > 30 && this.getDateFormat(date);
    }*/

    /**
     * Fecha con formato de texto
     *
     * @static
     * @param {Date} thisDate
     * @returns {string}
     * @memberof FormatETools
     */
    public static getDateFormat(thisDate: Date): string {
        let dd = thisDate.getDate();

        let mmFull = this._months[thisDate.getMonth()];

        let yyyyFull = "";
        if (thisDate.getFullYear() != (new Date()).getFullYear()) {
            yyyyFull = " de " + thisDate.getFullYear().toString();
        }

        return dd.toString() + ' de ' + mmFull + yyyyFull;
    }

    /**
     * Fecha en formato dd/mm/aaaa
     *
     * @static
     * @param {Date} thisDate
     * @returns {string}
     * @memberof FormatETools
     */
    public static getDateFormatDDMMYYYY(thisDate: Date): string {

        let dd = thisDate.getDate();
        let ddFull = dd.toString();
        let mm = thisDate.getMonth() + 1; //January is 0!
        let mmFull = mm.toString();
        let yyyy = thisDate.getFullYear();
        let yyyyFull = yyyy.toString();

        if (dd < 10) {
            ddFull = '0' + ddFull;
        }
        if (mm < 10) {
            mmFull = '0' + mmFull;
        }
        return ddFull + '/' + mmFull + '/' + yyyyFull;
    }

}
