# Masamune Industries SharePoint Site Provisioning

## Overview

This project contains a PowerShell script designed to automate the provisioning of SharePoint sites using PnP (Patterns and Practices) templates. The script leverages the SharePoint Online Management Shell and PnP PowerShell modules to create multiple sites efficiently while applying a predefined template to each site. The script is designed to be flexible and customizable, allowing users to specify the number of sites to create, the site names, URLs, and descriptions, and the template to apply to each site. The script also includes retry logic to handle any errors that could occur during site creation.

## Features

- Creates multiple SharePoint sites with customized names and URLs.
- Applies PnP template/s to each site.
- Implements retry logic to handle any errors.
- Customizable site titles, descriptions, and URLs based on predefined department and committee names.

## Prerequisites

- [PowerShell 5.1 or later](https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell)
- [PnP.PowerShell module](https://pnp.github.io/powershell/articles/installation.html)
- [NameIT module](https://www.powershellgallery.com/packages/NameIT)
- [SharePoint Online tenant with Administrative Access](https://www.microsoft.com/en-au/microsoft-365/enterprise/microsoft365-plans-and-pricing) </br>
[*Offers 1 month free trials.. but remember to cancel the subscription before the trial ends*]

## Setup

1. **Clone the repository**:
   ```bash
   git clone https://github.com/TimBroderick44/Masamune_Industries
   cd Masaume_Industries
    ```
2. **Set up SharePoint Online Tenant**:
   - Create a SharePoint Online tenant if you don't have one.
   - Create a SharePoint site to store the PnP templates (Can do through the SharePoint Admin Center or via [Powershell](https://pnp.github.io/powershell/cmdlets/New-PnPSite.html) - *make sure that the module is first installed!*). 
   - Once created, go to the site contents of the new site (top of the page), create a new 'document library' and upload the 'contosoworks.pnp' file here.
   - **Alternatively**, work with the 'contosoworks.pnp' locally.
3. **Install required modules (Optional as included in the script)**:
   ```bash
   Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
   Install-Module -Name NameIT -Scope CurrentUser -Force -AllowClobber
    ```
4. **Implement Variables and Run the Script**:
   - Open the `CreateSites.ps1` script in a text editor.
   - Customize the variables at the top of the script as needed (*See below for usage*).
   - Through the PowerShell terminal, go to the directory with the 'CreateSites.ps1' file and run the script.
   ```bash
   .\CreateSites.ps1
    ```

### Variables

The `CreateSites.ps1` script has the following variables that can be customized:

- **`$numSites`**: *(int)* Number of sites to create. Default is 5.  
Example: `-numSites 10`

- **`$tenantUrl`**: *(string)* URL of your SharePoint tenant.  
Example: `-tenantUrl "https://masamuneindustries.sharepoint.com"`

- **`$templateSite`**: *(string)* URL of the site where the template is stored.  
Example: `-templateSite "https://masamuneindustries.sharepoint.com/sites/PNPTemplates"`
  
- **`$templateUrl`**: *(string)* **COMPLETE** URL of the PnP template **file**.  
  Example: `-templateUrl "https://masamuneindustries.sharepoint.com/sites/PNPTemplates/PNP_Templates/contosoworks.pnp"`
  
- **`$templateFilePath`**: *(string)* **RELATIVE** Path to the PnP template file within the SharePoint site.  
  Example: `-templateFilePath "/sites/PNPTemplates/PNP_Templates/contosoworks.pnp"`
  
- **`$localTemplatePath`**: *(string)* Local path to store the downloaded template file.  
  Example: `-localTemplatePath "C:\temp\contosoworks.pnp"`
  
- **`$tempFolderPath`**: *(string)* Temporary folder path to store the pnp file temporarily.  
  Example: `-tempFolderPath "C:\temp"`
  
- **`$siteOwner`**: *(string)* Owner of the newly created sites.
  Example: `-siteOwner "BobBobson@mYourTenant.onmicrosoft.com"`

## Challenges and Lessons:

- **PowerShell Scripting**: At first, I was unfamiliar with PowerShell scripting, but after working on this project, I have gained a much stronger understanding of the language and its capabilities.
- **PnP Templates**: Gained experience in working with PnP templates for SharePoint. Specifically, the structure of XML files and its syntax.
- **Azure Integration**: Explored methods for automating processes with Azure services. For example, Azure Functions or Power Automate (Not implemented in this project... *yet*...)
- **Error Handling**: Trying to get more verbose and expletive error handling in the script. This was a challenge as when errors arose, it was often difficult to pinpoint the exact cause and/or get detailed error messages.

## Future Improvements

- **Scalability**: Optimize the script for better performance with large-scale site creation. For example, create a wider variety of templates for different departments and committees.
- **Customization**: Add more customization options for site creation and template application.
- **Automation**: Research and implement automation for the script using Azure Functions or Power Automate.

## Contact

Feel free to reach out to me with any questions, suggestions, or feedback!

Happy provisioning! 🚀
