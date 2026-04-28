import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '../../tailwind-generated.css';
import "./Hide.scss"
import Routing from './components/Routing';

export default class NfaWebPart extends BaseClientSideWebPart<{}> {
  public render(): void {
    const sp = spfi().using(SPFx(this.context));

    const element = React.createElement(Routing, {
      context: this.context,
      sp: sp,
    });

    ReactDom.render(element, this.domElement);
  }
}