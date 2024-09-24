import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import PeopleDirectory from './components/PeopleDirectory';
import { IPeopleDirectoryProps } from './components/IPeopleDirectoryProps';
import { GraphService } from './services/GraphService';

export interface IPeopleDirectoryWebPartProps {}

export default class PeopleDirectoryWebPart extends BaseClientSideWebPart<IPeopleDirectoryWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IPeopleDirectoryProps> = React.createElement(PeopleDirectory, {
      graphService: new GraphService(this.context)
    });
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, "People Directory.");
    setTimeout(() => {
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      ReactDom.render(element, this.domElement);
    },5000);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: []
    };
  }
}
