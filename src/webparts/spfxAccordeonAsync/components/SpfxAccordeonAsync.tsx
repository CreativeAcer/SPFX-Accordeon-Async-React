import * as React from 'react';
import styles from './SpfxAccordeonAsync.module.scss';
import { ISpfxAccordeonAsyncProps } from './ISpfxAccordeonAsyncProps';

import { sp } from "@pnp/sp/presets/all";
import './reactaccordeon.css';

import {
  Accordion,
  AccordionItem,
  AccordionItemHeading,
  AccordionItemButton,
  AccordionItemPanel,
} from 'react-accessible-accordion';

export interface ISpfxAccordeonAsyncState {
  items: Array<any>;
}

export default class SpfxAccordeonAsync extends React.Component<ISpfxAccordeonAsyncProps, ISpfxAccordeonAsyncState> {
  
  constructor(props: ISpfxAccordeonAsyncProps) {
    super(props);
    
    this.state = {
      items: new Array<any>()
    };
    this.getListItems();
  }

  private getListItems(): void {
    if(typeof this.props.listId !== "undefined" && this.props.listId.length > 0) {
      sp.web.lists.getById(this.props.listId).items.select(this.props.titleField,this.props.valueField).get()
        .then((results: Array<any>) => {
          this.setState({
            items: results
          });
        })
        .catch((error:any) => {
          console.log("Failed to get list items!");
          console.log(error);
        });
    }
  }

  public componentDidUpdate(prevProps:ISpfxAccordeonAsyncProps): void {
    if(prevProps.listId !== this.props.listId || prevProps.titleField !== this.props.titleField || prevProps.valueField !== this.props.valueField) {
      this.getListItems();
    }
  }
  
  public render(): React.ReactElement<ISpfxAccordeonAsyncProps> {
    return (
      <div className={ styles.spfxAccordeonAsync }>
        <div>
          <h2>{this.props.accordionTitle}</h2>
          <Accordion allowZeroExpanded> 
            {this.state.items.map((item:any) => {
              return (
                <AccordionItem>
                  <AccordionItemHeading>
                    <AccordionItemButton>
                      {item[this.props.titleField]}
                    </AccordionItemButton>
                  </AccordionItemHeading>
                    <AccordionItemPanel>
                      <p  dangerouslySetInnerHTML={{__html: item[this.props.valueField]}} />
                    </AccordionItemPanel>
                </AccordionItem>
                );
              })
            }
          </Accordion>
        </div> 
        
      </div>
    );
  }
}
