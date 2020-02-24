import { sp, Items, ItemAddResult, ItemUpdateResult, List } from "@pnp/sp";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
//locals
import { IListItemService } from "./IListItemService";
import { IListItem } from "./../../models/IListItem";
    
//Representación de una consulta SP
export interface IQueryable {
    select?: string;
    expand?: string;
    filter?: string;
    order?: {
        by: string,
        asc: boolean
    };
    top?: number;
}


/**
 * Servicio "Wrap" que contiene los metodos necesarios de conexión con los servicios de sharePoint para una lista
 * @see {@link https://www.npmjs.com} @pnp/sp
 * @copyright 2018 e-deas (http://www.e-deas.com.co)  El uso de esta libreria esta reservador para este sitio y cualquier cambio o reutilización debe ser autorizado por e-deas.
 * @author Diego Campo <diegoc@e-deas.com.co> / Fecha: 14.01.2019 - Creado
 * @author Diego Campo <diegoc@e-deas.com.co> / Fecha: 23.01.2019 - Modificado
 *
 * @export
 * @class ListItemService
 * @implements {IListItemService}
 */
export abstract class ListItemService implements IListItemService {

    //nombre de la lista que controla la instancia
    private _listName: string;
    private _context: WebPartContext | ApplicationCustomizerContext;


    /**
      * Crear una instancia de ListItemService
      * @param {string} listName
      * @param {(WebPartContext | ApplicationCustomizerContext)} context
      * @memberof ListItemService
      */
    public constructor(listName: string, context: WebPartContext | ApplicationCustomizerContext) {
        this._listName = listName;
        this._context = context;
    }

    /**
     * Obtener el nombre de la lista de la instancia
     * @returns {string}
     * @memberof ListItemService
     */
    public getListName(): string {
        return this._listName;
    }

    /**
   * Define si se quiere consultar por GUID o por TITLE la lista
   *
   * @returns {List}
   * @memberof ListItemService
   */
    public getList(): List {
        if (this._listName.length == 36 && this._listName.indexOf("-") > 0) {
            return sp.web.lists.getById(this._listName);
        } else {
            return sp.web.lists.getByTitle(this._listName);
        }

    }

    /**
     * Leer varios items en una lista
     * @param {IQueryable} [query]
     * @param {boolean} [usingCaching]
     * @returns {Promise<any[]>}
     * @memberof ListItemService
     */
    public async getItems(query?: IQueryable, usingCaching?: boolean): Promise<any[]> {
        //const items: any[] = await sp.web.lists.getByTitle(this._listName).items.get();

        let items: Items = this.getList().items;

        if (query != null) {
            if (query.select != null && query.select.trim() != "") {
                items.select(query.select);
            }
            if (query.expand != null && query.expand.trim() != "") {
                items.expand(query.expand);
            }
            if (query.filter != null && query.filter.trim() != "") {
                items.filter(query.filter);
            }
            if (query.order != null) {
                items.orderBy(query.order.by, query.order.asc);
            }
            if (query.top != null && query.top > 0) {
                items.top(query.top);
            } else {
                items.top(5000);
            }
            if (usingCaching != null && usingCaching) {
                items.usingCaching();
            }
        }

        const item: any = await items.get();
        return item;
    }

    /**
     * Leer un item en una lista
     * @param {number} itemId
     * @param {IQueryable} [query]
     * @param {boolean} [usingCaching]
     * @returns {Promise<any>}
     * @memberof ListItemService
     */
    public async getById(itemId: number, query?: IQueryable, usingCaching?: boolean): Promise<any> {
        if (itemId === -1) {
            throw new Error('No items found in the list');
        }

        let items: Items = this.getList().items;

        if (query != null) {
            if (query.select != null && query.select.trim() != "") {
                items.select(query.select);
            }
            if (query.expand != null && query.expand.trim() != "") {
                items.expand(query.expand);
            }
        }
        if (usingCaching != null && usingCaching) {
            items.usingCaching();
        }

        const item: any = await items.getById(itemId).get();
        return item;
    }

    /**
     * Leer el último item modificado por cualquier persona en una lista
     * @param {IQueryable} [query]
     * @returns {Promise<any>}
     * @memberof ListItemService
     */
    public async getLastModified(query?: IQueryable): Promise<any> {

        let items: Items = this.getList().items;

        if (query != null) {
            if (query.select != null && query.select.trim() != "") {
                items.select(query.select);
            }
            if (query.expand != null && query.expand.trim() != "") {
                items.expand(query.expand);
            }
            if (query.filter != null && query.filter.trim() != "") {
                items.filter(query.filter);
            }
        }

        const item: any = await items.orderBy("Modified", false).top(1).get();
        return item;
    }

    /**
     * Crear un item en una lista
     * @param {*} newItem
     * @returns {Promise<any>}
     * @memberof ListItemService
     */
    public async addItem(newItem: any): Promise<any> {
        const result: ItemAddResult = await this.getList().items.add(newItem);
        return result.data as any;
    }


    /**
     * Actualizar un item en una lista
     *
     * @param {number} itemId
     * @param {*} updateItem
     * @param {boolean} [hasTag]
     * @returns {Promise<ItemUpdateResult>}
     * @memberof ListItemService
     */
    public async updateItem(itemId: number, updateItem: any, hasTag?: boolean): Promise<ItemUpdateResult> {
        if (itemId === -1) {
            throw new Error('No items found in the list');
        }

        let itemUpdated = null;

        if (hasTag) {
            let etag: string = undefined;
            let headers = {
                headers: {
                    'Accept': 'application/json;odata=minimalmetadata'
                }
            };
            const item = await this.getList().items.getById(itemId).get(undefined, headers);
            etag = item["odata.etag"];
            itemUpdated = await this.getList().items.getById(itemId).update(updateItem, etag);
        } else {
            itemUpdated = await this.getList().items.getById(itemId).update(updateItem);
        }
        return itemUpdated;
    }

    /**
     * Eliminar un item en una lista
     * @param {number} itemId
     * @returns {Promise<void>}
     * @memberof ListItemService
     */
    public async deleteItem(itemId: number): Promise<void> {
        if (itemId === -1) {
            throw new Error('No items found in the list');
        }
        let etag: string = undefined;
        let headers = {
            headers: {
                'Accept': 'application/json;odata=minimalmetadata'
            }
        };
        const item = await this.getList().items.getById(itemId).get(undefined, headers);
        etag = item["odata.etag"];
        const item2Update = (item as any) as IListItem;
        return this.getList().items.getById(item2Update.Id).delete(etag);
    }

}

