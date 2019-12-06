import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import {
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';

import * as strings from 'LicensesFieldCustomizerStrings';
import styles from './LicensesFieldCustomizer.module.scss';

import {
  sp, CamlQuery
}
from '@pnp/sp';
import * as pnp from 'sp-pnp-js';
import * as moment from 'moment';
import * as $ from 'jquery';

import { SPHttpClient, SPHttpClientResponse,SPHttpClientConfiguration } from '@microsoft/sp-http';  

/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ILicensesFieldCustomizerProperties {
  // This is an example; replace with your own property
  sampleText?: string;
}

const LOG_SOURCE: string = 'LicensesFieldCustomizer';

export default class LicensesFieldCustomizer
  extends BaseFieldCustomizer<ILicensesFieldCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    // Add your custom initialization to this method.  The framework will wait
    // for the returned promise to resolve before firing any BaseFieldCustomizer events.
    Log.info(LOG_SOURCE, 'Activated LicensesFieldCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    Log.info(LOG_SOURCE, `The following string should be equal: "LicensesFieldCustomizer" and "${strings.Title}"`);
    return Promise.resolve();
  }

  @override
  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    // Use this method to perform your custom cell rendering.
  //  const text: string = `${this.properties.sampleText}: ${event.fieldValue}`;
    event.domElement.classList.add(styles.cell);
    const items = [];
    var Name = event.listItem["_values"].get("Name");
    // this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Employee and Company Licenses')/Items?$select=Employee/Title,License/Title&$expand=Employee/Title, License/Title&$filter=Employee/Title eq '${Name}'`,  
    //   SPHttpClient.configurations.v1)  
    //   .then((response: SPHttpClientResponse) => {  
    //     response.json().then((responseJSON: any) => {  
    //       console.log(responseJSON);  

    //       for (var i = 0; i < responseJSON.value.length; i++) {
    //         items.push(responseJSON.value[i].License.Title);
    //       }

    //       let ul = document.createElement('ul');
    //       let li;
    //       items.forEach(function (item) {
    //           li = document.createElement('li');
    //           ul.appendChild(li);
          
    //           li.innerHTML += item;
    //       });
    //       event.domElement.innerHTML = ul.innerHTML;
    //     });  
    //   });
    event.domElement.innerText =  Name;
  }

  @override
  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    // This method should be used to free any resources that were allocated during rendering.
    // For example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    super.onDisposeCell(event);
  }
}
