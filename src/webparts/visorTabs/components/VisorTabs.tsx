import * as React from 'react';
import styles from './VisorTabs.module.scss';
import { IVisorTabsProps } from './IVisorTabsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Pivot, PivotItem, PivotLinkFormat } from 'office-ui-fabric-react/lib/Pivot';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { autobind, imgProperties } from '@uifabric/utilities';
import { IVisorTabsState } from './IVisorTabsState';
import { sp } from '@pnp/sp';

export default class VisorTabs extends React.Component<IVisorTabsProps, IVisorTabsState, {}> {

  @autobind
  private _onConfigure() {
    // Context of the web part
    this.props.context.propertyPane.open();
  }

  constructor(props: IVisorTabsProps, state: IVisorTabsState){
    super(props);
    
    //inicializar el estado
    this.state = {
      items: []
    };
  }

  public render(): React.ReactElement<IVisorTabsProps> {

    const themeTab = this.props.dropdownTypeVisor == "primaryTheme" ? PivotLinkFormat.tabs : PivotLinkFormat.links;
    if (this.props.collectionData != undefined && this.props.collectionData.length > 0 && !this.props.toggleSearchInfo){     
      this.props.collectionData.sort((a,b)=>a.orden-b.orden); 
      return <Pivot linkFormat={themeTab}>
                this.props.collectionData &&{
                  this.props.collectionData.map(pivot => {
                    return <PivotItem headerText={pivot.title}>
                            <Label>{pivot.contenido}</Label>
                          </PivotItem> 
                  })
                }         
            </Pivot>
    } else if(this.props.toggleSearchInfo && this.state.items.length > 0){
      return <Pivot linkFormat={themeTab}>
                this.state.items &&{
                  this.state.items.map(pivot => {
                    return <PivotItem headerText={pivot.Title}>
                            <Label>{pivot.contenido}</Label>
                          </PivotItem> 
                  })
                }         
            </Pivot>
    } else {
      return <Placeholder iconName='Edit'
                          iconText='Configure your web part'
                          description='Please configure the web part.'
                          buttonLabel='Configure'
                          onConfigure={this._onConfigure} />;
    }
  } // end Render
    
  //al terminar la carga inicial del componente
  public componentDidMount() {
    // alert('componentDidMount');
  } // end componentDidMount
  
  //Al cambiar las propiedades
  public componentWillReceiveProps(props: IVisorTabsProps) {
    if(props.toggleSearchInfo) this._fetchData(props);
  } // end componentWillReceiveProps
  
//   //Consulta de datos por medio de servicios rest
  private _fetchData(props: IVisorTabsProps) {
    //se consulta el servicio rest y se asigna el state
    const asyncCall = async () => {
      try{

        const canFilter = props.textFilterBy && props.operatorFilterBy && props.textValueFilter;
        let items = await props.tabsInformativosServices.getItems({
          select:`${props.textNameTitleFld}, ${props.textNameContentFld}`,
          order: {by: props.orderBy, asc: props.toggleAsc},
          top: props.numberCantElements ? props.numberCantElements : 10,
          filter: canFilter ? `${props.textFilterBy} ${props.operatorFilterBy} ${props.textValueFilter}` : ''
        }); //el servicio viene en los parametros de entrada
        
    
        this.setState({
          items: items
        }); //se inicializa el render()
        

      } catch (err){
        console.log(err);
        this.setState({
          items: []
        }); //se inicializa el render()
      }
    }

    //llamada al servicio anteriormente definido
    asyncCall();
  }
}
