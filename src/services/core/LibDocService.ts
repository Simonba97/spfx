import { sp, Items, FileAddResult, ItemUpdateResult, Item, List } from "@pnp/sp";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { Logger, LogLevel } from '@pnp/logging';
//Locals
import { ILibDocService } from "./ILibDocService";
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
 * Servicio "Wrap" que contiene los metodos necesarios de conexión con los servicios de sharePoint para una librería documental
 * @see {@link https://www.npmjs.com} @pnp/sp
 * @copyright 2018 e-deas (http://www.e-deas.com.co)  El uso de esta libreria esta reservador para este sitio y cualquier cambio o reutilización debe ser autorizado por e-deas.
 * @author nombre <usuario@e-deas.com.co> / Fecha: dd.mm.aaaa - Creado
 * @author nombre <usuario@e-deas.com.co> / Fecha: dd.mm.aaaa - Modificado
 *
 * @export
 * @class LibDocService
 * @implements {ILibDocService}
 */
export abstract class LibDocService implements ILibDocService {


    //nombre de la lista que controla la instancia
    private _listName: string;
    private _context: WebPartContext | ApplicationCustomizerContext;


    /**
     * Crear una instancia de LibDocService
     * @param {string} listName
     * @param {(WebPartContext | ApplicationCustomizerContext)} context
     * @memberof LibDocService
     */
    public constructor(listName: string, context: WebPartContext | ApplicationCustomizerContext) {
        this._listName = listName;
        this._context = context;
    }


    /**
     * Obtener el nombre de la lista de la instancia
     * @returns {string}
     * @memberof LibDocService
     */
    public getListName(): string {
        return this._listName;
    }

    /**
   * Define si se quiere consultar por GUID o por TITLE la lista
   *
   * @private
   * @returns {List}
   * @memberof ListItemService
   */
    public getList(): List {
        if (this._listName.length == 36 && this._listName.indexOf("-") > 0) {
            return sp.web.lists.getById(this._listName);
        } else {
            return this.getList();
        }

    }

    /**
     * Leer varios items en una lista
     * @param {IQueryable} [query]
     * @param {boolean} [usingCaching]
     * @returns {Promise<any[]>}
     * @memberof LibDocService
     */
    public async getItems(query?: IQueryable, usingCaching?: boolean): Promise<any[]> {
        //const items: any[] = await this.getList().items.get();

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
     * @memberof LibDocService
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
     * @memberof LibDocService
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
     * @param {*} file
     * @param {*} newItem
     * @returns {Promise<any>}
     * @memberof LibDocService
     */
    public async addItem(file: any, folderRelativeUrl: string, newItem: any): Promise<any> {

        // you can adjust this number to control what size files are uploaded in chunks
        let result: FileAddResult;
        if (file.size <= 10485760) {
            // small upload
            result = await sp.web.getFolderByServerRelativeUrl(folderRelativeUrl).files.add(file.name, file, true);
        } else {
            // large upload
            result = await sp.web.getFolderByServerRelativeUrl(folderRelativeUrl).files.addChunked(file.name, file,
                data => {
                    Logger.log({ data: data, level: LogLevel.Verbose, message: "progress" });
                }, true);
        }

        return result.file as any;
    }


    /**
     * Actualizar un item en una lista
     *
     * @param {number} itemId
     * @param {*} updateItem
     * @param {boolean} [hasTag]
     * @returns {Promise<ItemUpdateResult>}
     * @memberof LibDocService
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
     * @memberof LibDocService
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
        const itemDeleted = await this.getList().items.getById(item2Update.Id).delete(etag);
        return itemDeleted;
    }

    /**
     * listItemAllFields
     */
    public async getItemFileFromUrl(urlFile: string, query?: IQueryable, usingCaching?: boolean): Promise<any> {
        let item = sp.web.getFileByServerRelativeUrl(urlFile).listItemAllFields;

        if (query != null) {
            if (query.select != null && query.select.trim() != "") {
                item.select(query.select);
            }
            if (query.expand != null && query.expand.trim() != "") {
                item.expand(query.expand);
            }
            if (usingCaching != null && usingCaching) {
                item.usingCaching();
            }
        }

        return await item.get();
    }


    /**
     * File
     */
    public async getFile(urlFile: string, usingCaching?: boolean): Promise<any> {
        let file = sp.web.getFileByServerRelativeUrl(urlFile);

        if (usingCaching != null && usingCaching) {
            file.usingCaching();
        }

        let text = await file.getText();
        return text;
    }

}
