import * as React from 'react';
import styles from './VisorTabs.module.scss';
import { IVisorTabsProps } from './IVisorTabsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Pivot, PivotItem, PivotLinkFormat } from 'office-ui-fabric-react/lib/Pivot';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { autobind, imgProperties } from '@uifabric/utilities';


export default class VisorTabs extends React.Component<IVisorTabsProps, {}> {

  @autobind
  private _onConfigure() {
    // Context of the web part
    this.props.context.propertyPane.open();
  }

  public render(): React.ReactElement<IVisorTabsProps> {

    if (this.props.collectionData != undefined && this.props.collectionData.length > 0 && !this.props.toggleSearchInfo){     
      this.props.collectionData.sort((a,b)=>a.orden-b.orden); 
      return <Pivot linkFormat={PivotLinkFormat.tabs}>
                this.props.collectionData &&{
                  this.props.collectionData.map(pivot => {
                    return <PivotItem headerText={pivot.title}>
                            <Label>{pivot.contenido}</Label>
                          </PivotItem> 
                  })
                }         
            </Pivot>
    } else if(this.props.toggleSearchInfo){
      return <Pivot linkFormat={PivotLinkFormat.tabs}>
                this.props.collectionDataDinamic &&{
                  this.props.collectionDataDinamic.map(pivot => {
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
  }
}
