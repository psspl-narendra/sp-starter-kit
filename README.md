#SharePoint Online Site Design

##Objective
To provide guidance on how to Install and use the Pre-configured SharePoint Online Portal along with Subsites and required Site Designs. 
Understanding Site Design
Site Designs are re-usable templates deployed in Office 365 to apply a consistent theme, layout and set of actions for SharePoint Sites created in an organization.
A Site Design can be accessed from “Create Site” menu under SharePoint tab in your Office 365 App launcher.
 

By selecting the appropriate “Site Design” you can implement the design – layout, Lists and Libraries to a blank Site Collection template. https://github.com/psspl-narendra/sp-starter-kit/blob/master/README.md




Tips and Tricks 

 



Prerequisites
Permissions Required
Permissions needed (In your target tenant):
1.	Ensure you are connecting to your tenant site using a tenant admin account.
Tenant Configuration
1.	Account you are using to connect to your tenant site must be added as a term store administrator. Use Term Store option in SharePoint Online admin Center to assign this.
 

2.	Set Release preferences for your tenant to be set as "Targeted release for everyone".  Navigate to Office 365 admin Center -> Organization profile.
 
3. Create a tenant 'App Catalog' site. This must be Created with the 'Apps' option of the SharePoint Admin Center.
 











Installing the Solution
Objective
To provide guidance on how to setup Portal solution in a new Office 365 Tenant.
Install PowerShell Module
Start Windows PowerShell as Admin. 
 
Install SharePoint PNP Online Module by running the following command in Windows PowerShell.
Install-Module SharePointPnPPowerShellOnline
 
Select “y”
 

Connect to Client’s Office 365 Site
Once the Pnp is installed. Run the following to connect to your Client’s SharePoint site. {domain} is the Office 365 tenant domain you are using.
Connect-PnPOnline –Url https://{domain}-admin.sharepoint.com/ –Credentials (Get-Credential)
i.e. Connect-PnPOnline –Url https://nsitedev-admin.sharepoint.com/ –Credentials (Get-Credential)
Enter the Credentials.
 

Install the starter-kit
Copy the Development Package to your local computer and then navigate to the path of the ‘provisioning’ folder which is one of the sub-folders in your development folder. Development Package is upload in Foundation Core Development site under Development library. 
For example, copy the following path of your starterkit.pnp file and then run the ‘cd’ command to change to that directory. 
Cd ‘I:\SPFXSTARTERKIT\sp-starter-kit13\provisioning’
 
Finally, run the following to Install the package: 
Apply-PnPTenantTemplate -Path starterkit.pnp
 
Wait until the following sites are provisioned.
 
Verify the package and the sites. 









Finally Apply the Site Design as below. 

 
