import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
/*import { ITermStore, ITerms, ITermData, Session, ITerm } from "@pnp/sp-taxonomy";*/
/**
 * Clase que ofrece diferentes servicios de sitio
 *
 * @see {@link https://www.npmjs.com} @pnp/sp-taxonomy
 * @copyright 2018 e-deas (http://www.e-deas.com.co)  El uso de esta libreria esta reservador para este sitio y cualquier cambio o reutilización debe ser autorizado por e-deas.
 * @author Diego Campo <diegoc@e-deas.com.co> / Fecha: 19.02.2019 - Creado
 * @author Diego Campo <diegoc@e-deas.com.co> / Fecha: 19.02.2019 - Modificado
 *
 * @export
 * @class SiteService
 */
export class TaxonomyService {
    private _context: WebPartContext | ApplicationCustomizerContext;
    private _siteUrl: string;


    /**
      * Crear una instancia de TaxonomyService
      * @param {(WebPartContext | ApplicationCustomizerContext)} context
      * @memberof TaxonomyService
      */
    public constructor(siteUrl: string, context: WebPartContext | ApplicationCustomizerContext) {
        this._context = context;
        this._siteUrl = siteUrl;
    }


    /**
     * Gets multiple terms by their ids using the current taxonomy context
     * @param termIds An array of term ids to search for
     */
    /*public async getTermsById(termIds: string[]): Promise<(ITerm & ITermData)[]> {

        if (termIds.length > 0) {

            const taxonomySession = new Session(this._siteUrl);
            taxonomySession.setup({
                sp: {
                    headers: {
                        Accept: "application/json;odata=nometadata",
                    },
                },
            });

            // Get the default termstore
            const store: ITermStore = await taxonomySession.getDefaultSiteCollectionTermStore();    
            const terms: ITerms = await store.getTermsById(...termIds);

            return await terms.select('Id','Labels').get();
        } else {
            return [];
        }
    }*/
}