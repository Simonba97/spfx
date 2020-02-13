import { List } from "@pnp/sp";
import { IQueryable } from "./LibDocService";
/**
 * @summary Interfaz que representa los metodos que implementa la clase LibDocServices
 * @copyright 2018 e-deas (http://www.e-deas.com.co)  El uso de esta libreria esta reservador para este sitio y cualquier cambio o reutilización debe ser autorizado por e-deas.
 *
 * @export
 * @interface ILibDocService
 */
export interface ILibDocService {

    getListName(): string;
    getList(): List;

    getItems(query?: IQueryable, usingCaching?: boolean): Promise<any[]>;
    getById(itemId: number, query?: IQueryable, usingCaching?: boolean): Promise<any>;
    getLastModified(query?: IQueryable): Promise<any>;

    updateItem(itemId: number, updateItem: any, hasTag?: boolean): Promise<any>;
    deleteItem(itemId: number): Promise<void>;

    getFile(urlFile: string, usingCaching?: boolean): Promise<any>;
    getItemFileFromUrl(urlFile: string, query?: IQueryable, usingCaching?: boolean): Promise<any>;

}
