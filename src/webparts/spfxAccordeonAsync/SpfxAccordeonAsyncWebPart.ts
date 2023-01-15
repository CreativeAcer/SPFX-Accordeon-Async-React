import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SpfxAccordeonAsyncWebPartStrings';
import SpfxAccordeonAsync from './components/SpfxAccordeonAsync';
import { ISpfxAccordeonAsyncProps } from './components/ISpfxAccordeonAsyncProps';

import { sp } from "@pnp/sp";
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { isEmpty } from '@microsoft/sp-lodash-subset';

export interface ISpfxAccordeonAsyncWebPartProps {
  listId: string;
  accordionTitle: string;
  titleField: string;
  valueField: string;
}

export default class SpfxAccordeonAsyncWebPart extends BaseClientSideWebPart<ISpfxAccordeonAsyncWebPartProps> {
  private _initComplete = false;
  private _placeholder = null;

  public async onInit(): Promise<void> {

    this._initializeRequiredProperties();
    sp.setup({
      spfxContext: this.context
    });
    this._initComplete = true;

    return super.onInit();

  }

  private _initializeRequiredProperties() {
    
  }

  public async render(): Promise<void> {
    if (!this._initComplete) {
      return;
    }

    if (this.displayMode === DisplayMode.Edit) {
      const { Placeholder } = await import(
          /* webpackChunkName: 'search-property-pane' */
          '@pnp/spfx-controls-react/lib/Placeholder'
      );
      this._placeholder = Placeholder;
    }

    this.renderCompleted();
  }

  private _isWebPartConfigured(): boolean {
    return (!isEmpty(this.properties.listId) && !isEmpty(this.properties.titleField) && !isEmpty(this.properties.valueField));
  }

  protected renderCompleted(): void {
    super.renderCompleted();

    let renderElement: React.ReactElement<ISpfxAccordeonAsyncProps> = null;

    if (this._isWebPartConfigured()) {
      
      const element: React.ReactElement<ISpfxAccordeonAsyncProps> = React.createElement(
        SpfxAccordeonAsync,
        {
          listId: this.properties.listId,
          accordionTitle: this.properties.accordionTitle,
          titleField: this.properties.titleField,
          valueField: this.properties.valueField,
          onConfigure: () => {
            this.context.propertyPane.open();
          }
        }
      );
      renderElement = element;
    } else {
      if (this.displayMode === DisplayMode.Edit) {
          const placeholder: React.ReactElement<any> = React.createElement(
              this._placeholder,
              {
                  iconName: strings.placeholderIconName,
                  iconText: strings.placeholderName,
                  description: strings.placeholderDescription,
                  buttonLabel: strings.placeholderbtnLbl,
                  onConfigure: this._setupWebPart.bind(this)
              }
          );
          renderElement = placeholder;
      } else {
          renderElement = React.createElement('div', null);
      }
    }

    ReactDom.render(renderElement, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  private _setupWebPart() {
    this.context.propertyPane.open();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('accordionTitle', {
                  label: 'Accordion Title (optional)'
                }),
                PropertyFieldListPicker('listId', {
                  label: 'Select a list',
                  selectedList: this.properties.listId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: <any>this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                }),
                PropertyPaneTextField('titleField', {
                  label: 'Column name for Title',
                }),,
                PropertyPaneTextField('valueField', {
                  label: 'Column name for Value',
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
