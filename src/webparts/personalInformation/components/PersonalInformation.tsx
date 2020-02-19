import * as React from 'react';
import styles from './PersonalInformation.module.scss';
import { Logger, LogLevel } from '@pnp/logging';
import { IPersonalInformationProps } from './IPersonalInformationProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IPersonalInformationState } from './IPersonalInformationState';

/**
 * Componente REACT para visualizar la información de personas  
 *
 * @copyright 2020 e-deas (http://www.e-deas.com.co)  El uso de esta libreria esta reservador para este sitio y cualquier cambio o reutilización debe ser autorizado por e-deas.
 * @author Simón Bustamante Alzate <simon.bustamante@e-deas.com.co> / Fecha: 18.02.2020 - Creado
 * @author Simón Bustamante Alzate <simon.bustamante@e-deas.com.co> / Fecha: 18.02.2020 - Modificado
 *
 * @export
 * @class PersonasInformación
 * @extends {React.Component<IPersonalInformationProps, IPersonalInformationState>}
 */
export default class PersonalInformation extends React.Component<IPersonalInformationProps, IPersonalInformationState> {

  /**
  *  Crea una instancia de  Visor.
  * @param {IVisorProps} props
  * @param {IVisorState} state
  * @memberof Visor
  */
  constructor(props: IPersonalInformationProps, state: IPersonalInformationState) {
    super(props);

    //inicializar el estado
    this.state = {
      items: []
    };
  }


  public render(): React.ReactElement<IPersonalInformationProps> {
    return (
      <div className={ styles.personalInformation }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
  
  //al terminar la carga inicial del componente
  public componentDidMount() {
      this._fetchData(this.props);
  }

  //Al cambiar las propiedades
  public componentWillReceiveProps(props: IPersonalInformationProps) {
      this._fetchData(props);
  }

  //Consulta de datos por medio de servicios rest
  private _fetchData(props: IPersonalInformationProps) {
    //se consulta el servicio rest y se asigna el state
    const asyncCall = async () => {
        try {

            // const items = await props.entityNameService.getItems();  //el servicio viene en los parametros de entrada
            const items = [];  //el servicio viene en los parametros de entrada

            this.setState({
                items: items
            }); //se inicializa el render()


        } catch (err) {
            this.setState({
                items: []
            }); //se inicializa el render()

            Logger.log({ data: err, level: LogLevel.Error, message: "Visor (TSX) > _fetchData > " });
        }
    };

    //llamada al servicio anteriormente definido
    asyncCall();
  }

}
