import * as React from 'react';
import styles from './PersonalInformation.module.scss';
import { Logger, LogLevel } from '@pnp/logging';
import { IPersonalInformationProps } from './IPersonalInformationProps';
import { IPersonalInformationState } from './IPersonalInformationState';
import { autobind } from '@uifabric/utilities';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";

import { IPersonaSharedProps, Persona, PersonaSize, PersonaPresence } from 'office-ui-fabric-react/lib/Persona';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Stack } from 'office-ui-fabric-react/lib/Stack';

import { Panel, FontSizes } from "office-ui-fabric-react";

import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Separator } from 'office-ui-fabric-react/lib/Separator';
import { Text } from 'office-ui-fabric-react/lib/Text';
import { IPersonalInformation } from '../../../models/IPersonalInformation';
import { IQueryable } from '../../../services/core/ListItemService';

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
      items: [],
      showPanelDetail: false,
      itemDetail: null
    };
  }

  public render(): React.ReactElement<IPersonalInformationProps> {
    if (this.state.items.length > 0) {
      return this._renderPersons(this.props);
    } else {
      return this._renderConfigInitial();
    }
  }

  //Render de configuración del WP cuando no hay datos para mostrar
  public _renderConfigInitial(): React.ReactElement<IPersonalInformationProps> {
    return (
      <Placeholder iconName='Edit'
        iconText='Configure your web part'
        description='Please configure the web part.'
        buttonLabel='Configure'
        onConfigure={this._onConfigure} />
    );
  } // end _renderConfigInitial

  //Render de personas cuando hay datos para pintar
  public _renderPersons(props: IPersonalInformationProps): React.ReactElement<IPersonalInformationProps> {
    return (
      <div>
        <Stack tokens={{ childrenGap: 10 }}>
          {
            this.state.items.map(element => {

              const personInformation: IPersonaSharedProps = this._buildPropsPerson(element, props);

              return (
                <Persona
                  {...personInformation}
                  size={PersonaSize.size72}
                  presence={PersonaPresence.none}
                  // onRenderSecondaryText={_onRenderSecondaryText}
                  styles={{ root: { margin: '0 0 10px 0' } }}
                  imageAlt={element.Title}
                  imageInitials={element.Title.split(' ').map(v => { return v.substring(0, 1) }).join('').substr(0, 2)}
                />
              );
            })
          }
        </Stack>
        {this.state.showPanelDetail &&
          <Panel isOpen={this.state.showPanelDetail}>
            {this._renderDetailPerson()}
          </Panel>
        }
      </div>
    );
  } // end _renderPersons

  // Construimos el detalle de la persona (panel)
  private _renderDetailPerson() {
    return (
      <Stack>
        {!this.state.itemDetail &&
          <Spinner label="Por favor espere...estámos cargando la información" size={SpinnerSize.large} />
        }
        {this.state.itemDetail &&
          <Stack>
            <Stack.Item align="center">
              <Persona
                imageUrl={this.state.itemDetail.foto ? this.state.itemDetail.foto.Url : ''}
                imageInitials={this.state.itemDetail.Title.split(' ').map(v => { return v.substring(0, 1) }).join('').substr(0, 2)}
                size={PersonaSize.extraLarge}
                presence={PersonaPresence.none}
                styles={{ root: { margin: '0 auto', textAlign: 'center' } }}
                imageAlt={this.state.itemDetail.Title}
              />
            </Stack.Item>
            <Separator></Separator>
            <Stack.Item>
              <Icon iconName="ReminderPerson" style={{ fontSize: FontSizes.large }} /> <Text key="name" variant="xLarge"> {this.state.itemDetail.Title}</Text>  <br />
              <Icon iconName="Work" style={{ fontSize: FontSizes.large }} /> <Text key="departamento" variant="mediumPlus"> {this.state.itemDetail.departamento}</Text> <br />
              <Icon iconName="TextDocumentShared" style={{ fontSize: FontSizes.large }} /> <Text key="descripcion" variant="mediumPlus"> {this.state.itemDetail.descripcion}</Text> <br />
              <Icon iconName="Phone" style={{ fontSize: FontSizes.large }} /> <Text key="telefono" variant="mediumPlus"> {this.state.itemDetail.telefono}</Text> <br />
              <Icon iconName="Mail" style={{ fontSize: FontSizes.large }} /> <Text key="correo" variant="mediumPlus"> {this.state.itemDetail.correo}</Text> <br />
              <Icon iconName="MapPin" style={{ fontSize: FontSizes.large }} /> <Text key="direccion" variant="mediumPlus"> {this.state.itemDetail.direccion} </Text> <br />
              <Icon iconName="Medical" style={{ fontSize: FontSizes.large }} /> <Text key="edad" variant="mediumPlus"> {this.state.itemDetail.edad} Años</Text> <br />
            </Stack.Item>

          </Stack>
        }
      </Stack>
    );
  } // end _renderDetailPerson

  //Construir los datos informativos de la persona
  public _buildPropsPerson(element: IPersonalInformation, props: IPersonalInformationProps): IPersonaSharedProps {
    const _personInformation: IPersonaSharedProps = {
      text: element.Title,
      onClick: () => this._buildDetailInformation(element.ID)
    };

    try {
      if (element.foto && props.toggleWithPhoto) {
        _personInformation.imageUrl = element.foto.Url;
      }

      if (props.fieldsShow.indexOf('departamento') != -1) {
        _personInformation.secondaryText = element.departamento;
      }
      if (props.fieldsShow.indexOf('correo') != -1) {
        _personInformation.tertiaryText = element.correo;
      }
    } catch (e) { }
    return _personInformation;
  } //  end _buildPropsPerson

  //al terminar la carga inicial del componente
  public componentDidMount() {
    if (this.props.urlList) {
      this._fetchData(this.props);
    }
  }

  //Al cambiar las propiedades
  public componentWillReceiveProps(props: IPersonalInformationProps) {
    if (props.urlList) {
      this._fetchData(props);
    }
  }

  //Consulta de datos por medio de servicios rest
  private _fetchData(props: IPersonalInformationProps) {
    //se consulta el servicio rest y se asigna el state
    const asyncCall = async () => {
      try {
        const items = await props._personalInformationServices.getItemsPrincipalPanel(props.fieldsShow, props.toggleWithPhoto);//el servicio viene en los parametros de entrada
        Logger.log({ data: items, level: LogLevel.Info, message: "PersonalInformation (TSX) > _fetchData > " });

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


  private _buildDetailInformation(idPerson: number) {
    this.setState({ showPanelDetail: true, itemDetail: null });// Lo hacemos separados por si es mejor traer la información de una consulta  mostrar spinner

    const asyncCall = async () => {
      try {
        const infoPersonal = await this.props._personalInformationServices.getDetailById({
          select: 'Title,departamento,descripcion,telefono,correo,direccion,edad,foto',
          filter: `ID eq ${idPerson}`,
          top: 1
        });

        let item = infoPersonal;

        this.setState({
          itemDetail: item
        }) // Set info person in state

      } catch (err) {
        this.setState({
          itemDetail: null
        }); //se inicializa el render()

        Logger.log({ data: err, level: LogLevel.Error, message: "Visor (TSX) > _buildDetailInformation > " });
      }

    }

    //llamada al servicio anteriormente definido
    asyncCall();

  }// end _buildDetailInformation

  @autobind
  private _onConfigure() {
    // Context of the web part
    this.props.context.propertyPane.open();
  }

}
