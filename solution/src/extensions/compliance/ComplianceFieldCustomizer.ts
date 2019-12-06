import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import {
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';

import * as strings from 'ComplianceFieldCustomizerStrings';
import styles from './ComplianceFieldCustomizer.module.scss';
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
export interface IComplianceFieldCustomizerProperties {
  // This is an example; replace with your own property
  TechnicalExpected: Number;
  PersonalSkillsExpected: Number;
  ManagementExpected: Number;
  EthicsExpected: Number;

  TechnicalSpent: Number;
  PersonalSkillsSpent: Number;
  ManagementSpent: Number;
  EthicsSpent: Number;

  TechnicalValue: Number;
  PersonalSkillsValue: Number;
  ManagementValue: Number;
  EthicsValue: Number;
}

const LOG_SOURCE: string = 'ComplianceFieldCustomizer';

export default class ComplianceFieldCustomizer
  extends BaseFieldCustomizer<IComplianceFieldCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    // Add your custom initialization to this method.  The framework will wait
    // for the returned promise to resolve before firing any BaseFieldCustomizer events.
    Log.info(LOG_SOURCE, 'Activated ComplianceFieldCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    Log.info(LOG_SOURCE, `The following string should be equal: "ComplianceFieldCustomizer" and "${strings.Title}"`);
    return Promise.resolve();
  }

  @
  override
  public async onRenderCell(event: IFieldCustomizerCellEventParameters) {
      // Use this method to perform your custom cell rendering.
      //const text: string = `${this.properties.sampleText}: ${event.listItem.fields.}`;

      // event.domElement.innerText = text;
      event.domElement.classList.add(styles.cell);

      var EffectiveDate = event.listItem["_values"].get("Effective_x0020_Date");
      var ExpireDate = event.listItem["_values"].get("Expiry_x0020_Date");
      var LicenseID = event.listItem["_values"].get("License")[0].lookupId;
      var EmployeeID = event.listItem["_values"].get("Employee")[0].lookupId;

      var EffectiveDay = parseInt(moment(EffectiveDate,"MM/DD/YYYY").format("DD"));
      var EffectiveMonth = parseInt(moment(EffectiveDate,"MM/DD/YYYY").format("MM"));
      let RenewalMonth;
      let RenewalDay;
      var Technical = 0;
      var PersonalSkills = 0;
      var Management = 0;
      var Ethics = 0;
      var Technical1 = 0;
      var PersonalSkills1 = 0;
      var Management1 = 0;
      var Ethics1 = 0;
      var TechnicalString;
      let PersonalSkillsString;
      let ManagementString;
      let EthicsString;
      var ExpiryDate;
      await pnp.sp.web.lists.getByTitle("LicensesAndPermits").items.getById(LicenseID)
      .select("RenewalMonth", "RenewalDay").get().then((item: any) => {
        debugger;
        RenewalMonth = parseInt(item["RenewalMonth"]);
        RenewalDay = parseInt(item["RenewalDay"]);
        var Fullyear = new Date().getFullYear();
        var NextFullYear = new Date().getFullYear() + 1;
        var thisYearDate = moment(RenewalMonth + "/" + RenewalDay  + "/" + Fullyear,"MM/DD/YYYY");
        var NextYearDate = moment(RenewalMonth + "/" + RenewalDay  + "/" + NextFullYear,"MM/DD/YYYY");
        var EffectiveDate2 = moment(EffectiveDate,"MM/DD/YYYY");
       
        if(EffectiveDate2 > thisYearDate)
        {
          ExpiryDate = RenewalMonth + "/" + RenewalDay  + "/" + NextFullYear;
        }
        else{
          ExpiryDate = RenewalMonth + "/" + RenewalDay  + "/" + Fullyear;
        }
      });

      const _camlQuery: CamlQuery = {};
      _camlQuery.ViewXml =
          `<View>  
              <Query> 
                <Where><And><And><Geq><FieldRef Name='Date_x0020_Completed' /><Value Type='DateTime'>${moment(EffectiveDate,"MM/DD/YYYY").toISOString()}</Value></Geq><Leq><FieldRef Name='Date_x0020_Completed' /><Value Type='DateTime'>${moment(ExpiryDate,"MM/DD/YYYY").toISOString()}</Value></Leq></And><Eq><FieldRef Name='Employee' LookupId='True' /><Value Type='Lookup'>${EmployeeID}</Value></Eq></And></Where> 
              </Query> 
              <ViewFields><FieldRef Name='AccreditationCategory' /><FieldRef Name='Course' /><FieldRef Name='Date_x0020_Completed' /><FieldRef Name='Employee' /><FieldRef Name='ID' /><FieldRef Name='Number_x0020_of_x0020_Hours' /><FieldRef Name='Title' /></ViewFields> 
        </View>`;
       await  pnp.sp.web.lists.getByTitle("Continuing Education Tracking").getItemsByCAMLQuery(_camlQuery).then((items: any[]) => {
          if (items != null && items.length > 0) {
              Technical = 0;
              PersonalSkills = 0;
              Management = 0;
              Ethics = 0;
              for (let item of items) {
                  if (item["AccreditationCategory"] == "Technical") {
                      Technical += parseInt(item["Number_x0020_of_x0020_Hours"]);
                  } else if (item["AccreditationCategory"] == "Personal Skills") {
                      PersonalSkills += parseInt(item["Number_x0020_of_x0020_Hours"]);
                  } else if (item["AccreditationCategory"] == "Management") {
                      Management += parseInt(item["Number_x0020_of_x0020_Hours"]);
                  } else if (item["AccreditationCategory"] == "Ethics") {
                      Ethics += parseInt(item["Number_x0020_of_x0020_Hours"]);
                  }
              }
          }
          else{
            Technical = 0;
            PersonalSkills = 0;
            Management = 0;
            Ethics = 0;
          }
      });
      await pnp.sp.web.lists.getByTitle("LicensesAndPermits").items.getById(LicenseID)
          .select("Technical", "Personal_x0020_Skills", "Management", "Ethics").get().then((item: any) => {
            Technical1 = parseInt(item["Technical"]);
            PersonalSkills1 = parseInt(item["Personal_x0020_Skills"]);
            Management1 = parseInt(item["Management"]);
            Ethics1 = parseInt(item["Ethics"]);
      });
      
      if(Technical >= Technical1){
        TechnicalString = "<div style='color:green'><b>Technical</b> : <b>" + Technical + "/" + Technical1 + "</b></div><br />" ;
      }
      else{
        TechnicalString = "<div style='color:red'><b>Technical</b> : <b>" + Technical + "/" + Technical1 + "</b></div><br />" ;
      }

      if(PersonalSkills >= PersonalSkills1){
        PersonalSkillsString = "<div style='color:green'><b>Personal Skills</b> : <b>" + PersonalSkills + "/" + PersonalSkills1 + "</b></div><br />" ;
      }
      else{
        PersonalSkillsString = "<div style='color:red'><b>Personal Skills</b> : <b>" + PersonalSkills + "/" + PersonalSkills1 + "</b></div><br />" ;
      }

      if(Management >= Management1){
        ManagementString = "<div style='color:green'><b>Management</b> : <b>" + Management + "/" + Management1 + "</b></div><br />" ;
      }
      else{
        ManagementString = "<div style='color:red'><b>Management</b> : <b>" + Management + "/" + Management1 + "</b></div><br />" ;
      }

      if(Ethics >= Ethics1){
        EthicsString = "<div style='color:green'><b>Ethics</b> : <b>" + Ethics + "/" + Ethics1 + "</b></div><br />" ;
      }
      else{
        EthicsString = "<div style='color:red'><b>Ethics</b> : <b>" + Ethics + "/" + Ethics1 + "</b></div><br />" ;
      }
      event.domElement.innerHTML = TechnicalString + PersonalSkillsString + ManagementString + EthicsString ;
  }

  @override
  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    // This method should be used to free any resources that were allocated during rendering.
    // For example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    super.onDisposeCell(event);
  }
}
