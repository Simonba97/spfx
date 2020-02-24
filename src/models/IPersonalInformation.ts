/**
 * Representa un registro de sharepoint en una lista o libreria documental
 *
 * @export
 * @interface IPersonalInformation
 */
export interface IPersonalInformation {
	Id: number;
	ID: number;
	Title: string;
	departamento: string;
	[foto: string]: any;
	correo: string;
	descripcion: string;
	telefono: string;
	direccion: string;
	edad: string;
}