import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import {
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';

import * as strings from 'ExpiryDateFieldCustomizerStrings';
import styles from './ExpiryDateFieldCustomizer.module.scss';
import {
  sp, CamlQuery
}
from '@pnp/sp';
import * as pnp from 'sp-pnp-js';
import * as moment from 'moment';
import * as $ from 'jquery';
/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IExpiryDateFieldCustomizerProperties {
  // This is an example; replace with your own property
  sampleText?: string;
}

const LOG_SOURCE: string = 'ExpiryDateFieldCustomizer';

export default class ExpiryDateFieldCustomizer
  extends BaseFieldCustomizer<IExpiryDateFieldCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    // Add your custom initialization to this method.  The framework will wait
    // for the returned promise to resolve before firing any BaseFieldCustomizer events.
    Log.info(LOG_SOURCE, 'Activated ExpiryDateFieldCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    Log.info(LOG_SOURCE, `The following string should be equal: "ExpiryDateFieldCustomizer" and "${strings.Title}"`);
    return Promise.resolve();
  }

  @override
  public async onRenderCell(event: IFieldCustomizerCellEventParameters) {
    // Use this method to perform your custom cell rendering.
  // Use this method to perform your custom cell rendering.
      //const text: string = `${this.properties.sampleText}: ${event.listItem.fields.}`;

      // event.domElement.innerText = text;
      event.domElement.classList.add(styles.cell);

      var EffectiveDate = event.listItem["_values"].get("Effective_x0020_Date");
      var LicenseID = event.listItem["_values"].get("License")[0].lookupId;

      let RenewalMonth;
      let RenewalDay;
      var ExpiryDate;
      await pnp.sp.web.lists.getByTitle("LicensesAndPermits").items.getById(LicenseID)
      .select("RenewalMonth", "RenewalDay").get().then((item: any) => {
        debugger;
        RenewalMonth = parseInt(item["RenewalMonth"]);
        RenewalDay = parseInt(item["RenewalDay"]);
        var Fullyear = new Date().getFullYear();
        var NextFullYear = new Date().getFullYear() + 1;
        var thisYearDate = moment(RenewalMonth + "/" + RenewalDay  + "/" + Fullyear,"MM/DD/YYYY");
        var EffectiveDate2 = moment(EffectiveDate,"MM/DD/YYYY");
       
        if(EffectiveDate2 > thisYearDate)
        {
          ExpiryDate = RenewalMonth + "/" + RenewalDay  + "/" + NextFullYear;
        }
        else{
          ExpiryDate = RenewalMonth + "/" + RenewalDay  + "/" + Fullyear;
        }
      });

      event.domElement.innerText = ExpiryDate;
  }

  @override
  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    // This method should be used to free any resources that were allocated during rendering.
    // For example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    super.onDisposeCell(event);
  }
}
