import { List } from "@pnp/sp";
/**
 * Interfaz que representa los metodos que implementa la clase ListItemService
 * @copyright 2018 e-deas (http://www.e-deas.com.co)  El uso de esta libreria esta reservador para este sitio y cualquier cambio o reutilización debe ser autorizado por e-deas. 
 * 
 * @export
 * @interface IListItemService
 */
export interface IListItemService {

    getListName(): string;
    getList(): List;

    getItems(): Promise<any[]>;
    getById(itemId: number): Promise<any>;
    getLastModified(): Promise<any>;

    addItem(newItem: any): Promise<any>;
    updateItem(itemId: number, updateItem: any): Promise<any>;
    deleteItem(itemId: number): Promise<void>;
}

