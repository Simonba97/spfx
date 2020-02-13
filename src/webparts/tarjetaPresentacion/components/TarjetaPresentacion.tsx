import * as React from 'react';
import styles from './TarjetaPresentacion.module.scss';
import { ITarjetaPresentacionProps } from './ITarjetaPresentacionProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as ReactDOM from 'react-dom';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { nullRender } from '@uifabric/utilities';

export default class TarjetaPresentacion extends React.Component<ITarjetaPresentacionProps, {}> {
  public render(): React.ReactElement<ITarjetaPresentacionProps> {

    return (
      <div id="contPrincipal" style={{display: "flex", padding: "10px"}}>
        <div id="rowLeft" style={{width: "40%", textAlign: "center", padding:"0 20px 0 0"}}>
          <div id="contImgPerfil">
            <img style={{width: "250px", borderRadius: "50%"}} src={this.props.imagen} alt="Imagen de perfil" id="imgProfile"/>
          </div>
          <div id="contRedes" style={{fontSize: "40px", paddingTop: "10px"}}>
          {
            this.props.redes != undefined &&(
              this.props.redes.map(red => {
                return <div><a href={red.URL}><Icon iconName={red.Icono} /></a></div>;
              })
            )
          }
          </div>
        </div>
        <div id="rowRigth" style={{width: "50%"}}>
          <div id="nombre" style={{margin: "10px 0"}}>
            <h1>{this.props.nombre}</h1>
          </div>
          <div id="descripcion">
            <span>{this.props.descripcion}</span>
          </div>
        </div>
      </div>
    );
  }
}
