# Masamune Industries SharePoint Site Provisioning

## Overview

This project contains a PowerShell script designed to automate the provisioning of SharePoint sites using PnP (Patterns and Practices) templates. The script leverages the SharePoint Online Management Shell and PnP PowerShell modules to create multiple sites efficiently while applying a [predefined template](https://github.com/SharePoint/sp-dev-provisioning-templates/blob/master/tenant/contosoworks/contosoworks.pnp) to each site. The script is designed to be flexible and customizable, allowing users to specify the number of sites to create, the site names, URLs, and descriptions, and the template to apply to each site. The script also includes retry logic to handle any errors that could occur during site creation.

## Features

- Creates multiple SharePoint sites with customized names and URLs.
- Applies PnP template/s to each site.
- Implements retry logic to handle any errors.
- Customizable site titles, descriptions, and URLs based on predefined department and committee names.

## Screenshot

![Script Execution](./Assets/screenshot.png)
 _Screenshot of the script execution in PowerShell terminal_

![SharePoint Site Creation](./Assets/sharepoint.png)
_Screenshot of the hub site applied by the script_

![SharePoint Site Creation](./Assets/sharepoint2.png)
_Screenshot of the subsite applied by the script_

![Script Retry](./Assets/retry.png)
_Screenshot of the script retrying after an error_

#### NOTE: The template used above was altered so that it was more applicable for an organisation based in Australia (e.g. dates were altered, terminology was changed, etc.)

## The Script

```Powershell
# Variables
$numSites = XXXX  ## Change this value to the number of sites you want to create ##

$tenantUrl = "https://XXXXX.sharepoint.com"
$templateUrl = "https://XXXXXXX.sharepoint.com/sites/PNPTemplates/PNP_Templates/contosoworks.pnp"   ## Can be changed for any other uploaded template too.
$templateSite = "https://XXXXXXX.sharepoint.com/sites/PNPTemplates"
$templateFilePath = "/sites/PNPTemplates/PNP_Templates/contosoworks.pnp"
$localTemplatePath = "C:\temp\contosoworks.pnp"
$tempFolderPath = "C:\temp"
$siteOwner = "XXXXXXX@XXXXXX.onmicrosoft.com"

# Custom department and committee names
$CustomData = @{
    department = @("HR", "Finance", "Marketing", "Sales", "Engineering", "IT", "Legal", "Operations", "Customer Service", "Research and Development", "Quality Assurance", 
    "Product Management", "Purchasing", "Logistics", "Supply Chain", "Manufacturing", "Facilities", "Safety", "Security", "Training", "Human Resources", "Quidditch Team", 
    "Wand Maintenance", "Time Travel Research", "Muggle Relations", "Potion Development", "Broomstick Testing", "Spell Innovation", "Magical Creature Care", 
    "Hogwarts Alumni Relations", "Dunder Mifflin Paper Supply", "Galactic Empire Liaison", "Time Machine Maintenance", "Alien Relations", "Meme Curation", 
    "Prank Department", "Dragon Taming", "Unicorn Care", "Troll Management", "Wizarding World Marketing", "Intergalactic Trade", "Superhero Coordination", 
    "Villain Rehabilitation", "Paranormal Investigation", "Secret Agent Training", "Fictional Character Recruitment", "Fairy Tale Compliance", "Mythical Artifact Procurement",
    "Comic Relief", "Mystery Department", "Paper Sales", "Dundie Awards Planning", "Pranks and Shenanigans", "World's Best Boss Training", "Beet Farming Operations", 
    "Pretzel Day Coordination", "Warehouse Management", "Scranton Branch Relations", "Regional Manager Support", "Assistant to the Regional Manager", 
    "Party Planning Committee", "Vance Refrigeration Collaboration", "Michael Scott Paper Company", "Schrute Farms Agritourism", "Sabre Compliance", 
    "Athlead Development", "Wuphf.com Operations", "Janitorial Humor", "Conference Room Scheduling", "Muckduck Analysis", "Workplace Safety Drills", 
    "Threat Level Midnight Production", "Procurement of Office Supplies", "Dunder Mifflin Infinity Support", "Finer Things Club", "Crossword Puzzles",
    "Pretending to Care", "Dunder Mifflin Olympics", "Office Romance Management", "Dunder Mifflin Infinity Training", "Dunder Mifflin Infinity Sales")
    
    committee = @("Council", "League", "Cabal", "Order of", "Squad", "Guild", "Society", "Fellowship", "Assembly", "Horde", "Band", "Team", "Group", 
    "Task Force", "Force", "Consortium", "Congregation", "Tribe", "Coterie", "Crew", "Clan", "Brigade", "Alliance", "Circle", "Union", "Coalition",
    "Federation", "Syndicate", "Network", "Confederation", "Corporation", "Corps", "Division", "Sect", "Chapter", "Department",
    "Office", "Bureau", "Agency", "Institute", "Center ", "Foundation", "Incorporated", "Corporation", "Company", "Association")
}

# Ensure the temp folder exists
if (-not (Test-Path -Path $tempFolderPath)) {
    New-Item -ItemType Directory -Path $tempFolderPath -Force
}

# Check if NameIT module is installed, if not install it
if (-not (Get-Module -ListAvailable -Name NameIT)) {
    Write-Host "NameIT module not found. Installing..."
    Install-Module -Name NameIT -Scope CurrentUser -Force -AllowClobber
}
Import-Module NameIT

# Check if PnP.PowerShell module is installed, if not install it
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "PnP.PowerShell module not found. Installing..."
    Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
}
Import-Module PnP.PowerShell

Write-Host "Connecting to main site $siteUrl to apply the template (Attempt $try)"

# Download template locally
Write-Host "Downloading PnP template from $templateUrl"
Connect-PnPOnline -Url $templateSite -Interactive
Get-PnPFile -Url $templateFilePath -Path $tempFolderPath -FileName "contosoworks.pnp" -AsFile -force

# Generate random site details
function Get-RandomSiteDetails {
    $department = Invoke-Generate "[department]" -CustomData $CustomData
    $committee = Invoke-Generate "[committee]" -CustomData $CustomData 
    $title = "The $department $committee"
    $description = "Internal communications and project collaboration for the $department $committee"
    $url = Invoke-Generate "$department$committee-[state abbr]##" -CustomData $CustomData
    $url = $url -replace '\s', '' 
    return [PSCustomObject]@{
        Title       = $title
        Description = $description
        URL         = "$tenantUrl/sites/$url"
        BenefitsTitle = "The $department $committee Benefits"
        BenefitsUrl = "$tenantUrl/sites/$url-benefits"
    }
}

# No prompts
$ConfirmPreference = 'None'

# Apply template with retry logic
function InitialiseTemplate {
    param (
        [string]$siteTitle,
        [string]$siteUrl,
        [string]$benefitsTitle,
        [string]$benefitsUrl,
        [int]$retryCount = 3
    )
    for ($try = 1; $try -le $retryCount; $try++) {
        try {
            Write-Host "================================================================================="
            Write-Host "Attempt $try to apply template to site: $siteTitle"
            Write-Host "================================================================================="

            Connect-PnPOnline -Url $siteUrl -Interactive
            Write-Host "Applying template to main site $siteTitle at $siteUrl (Attempt $try)"
            Invoke-PnPTenantTemplate -Path $localTemplatePath -Parameters @{
                "SiteTitle" = $siteTitle
                "SiteUrl" = $siteUrl
                "BenefitsSiteTitle" = $benefitsTitle
                "BenefitsSiteUrl" = $benefitsUrl
            }
            Write-Host "================================================================================="
            Write-Host "Successfully applied template!"
            Write-Host "Take a look: $siteUrl"
            Write-Host "================================================================================="

            # Cleanup the temporary file
            if (Test-Path $localTemplatePath) {
                Remove-Item $localTemplatePath
            }
            return
        } catch {
            Write-Warning "================================================================================="
            Write-Warning "Attempt $try failed: $_. Retrying..."
            Write-Warning "================================================================================="
            Write-Verbose "Exception message: $($_.Exception.Message)"
            Write-Verbose "Stack trace: $($_.Exception.StackTrace)"
            Write-Host "Detailed Exception Data: $($_ | Format-List -Force)"
            Start-Sleep -Seconds 60
        }
    }
    throw "=================================================================================`nFailed to apply template to $siteTitle at $siteUrl after $retryCount attempts"
}

# Loop to create sites
for ($i = 1; $i -le $numSites; $i++) {
    $siteDetails = Get-RandomSiteDetails
    Write-Host "================================================================================="
    Write-Host "Processing set $i of $numSites"
    Write-Host "================================================================================="
    Write-Host "Site Details:"
    Write-Host "Title             ---> $($siteDetails.Title)"
    Write-Host "Description       ---> $($siteDetails.Description)"
    Write-Host "URL               ---> $($siteDetails.URL)"
    Write-Host "Benefits Title    ---> $($siteDetails.BenefitsTitle)"
    Write-Host "Benefits URL      ---> $($siteDetails.BenefitsUrl)"
    Write-Host "================================================================================="

    try {
        # Create the main site
        Write-Host "Creating main site:     $($siteDetails.Title)"
        Connect-PnPOnline -Url $tenantUrl -Interactive
        New-PnPSite -Type CommunicationSite -Title $siteDetails.Title -Url $siteDetails.URL -Description $siteDetails.Description -Owner $siteOwner -Verbose -Debug
        Write-Host "---------------------------------------------------------------------------------"

        # Create the benefits site
        Write-Host "Creating benefits site: $($siteDetails.BenefitsTitle)"
        New-PnpSite -Type CommunicationSite -Title $siteDetails.BenefitsTitle -Url $siteDetails.BenefitsUrl -Description "Benefits site for $($siteDetails.Title)" -Owner $siteOwner -Verbose -Debug
        Write-Host "---------------------------------------------------------------------------------"

        Write-Host "================================================================================="
        Write-Host "Successfully created main site:     $($siteDetails.Title)"
        Write-Host "URL:                                $($siteDetails.URL)"
        Write-Host "Successfully created benefits site: $($siteDetails.BenefitsTitle)" 
        Write-Host "URL:                                $($siteDetails.BenefitsUrl)"
        Write-Host "================================================================================="
        Write-Host "Waiting 10 seconds before applying template to main site... It's a bit sensitive..."
        Start-Sleep -Seconds 10

        # Apply the main template to the main site with retry logic
        InitialiseTemplate -siteTitle $siteDetails.Title -siteUrl $siteDetails.URL -benefitsTitle $siteDetails.BenefitsTitle -benefitsUrl $siteDetails.BenefitsUrl
    } catch {
        Write-Host "================================================================================="
        Write-Host "Failed to create or apply template to site $($siteDetails.Title) at $($siteDetails.URL). Error: $_"
        Write-Host "================================================================================="
        Write-Debug "Exception details: $_"
        Write-Verbose "Exception message: $($_.Exception.Message)"
        Write-Verbose "Stack trace: $($_.Exception.StackTrace)"
        Write-Host "Detailed Exception Data: $($_ | Format-List -Force)"
    }
}

Write-Host "================================================================================="
Write-Host "Cleaning up temp folder."
Write-Host "================================================================================="

# Remove temp folder
if (Test-Path $tempFolderPath) {
    Remove-Item $tempFolderPath -Recurse -Force
}

Write-Host "================================================================================="
Write-Host "Script execution completed!"
Write-Host "================================================================================="
```

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
   - Create a SharePoint site to store the PnP templates (Can do through the SharePoint Admin Center or via [Powershell](https://pnp.github.io/powershell/cmdlets/New-PnPSite.html) - _make sure that the module is first installed!_).
   - Once created, go to the site contents of the new site (top of the page), create a new 'document library' and upload the 'contosoworks.pnp' file here.
   - **Alternatively**, work with the 'contosoworks.pnp' locally.
3. **Install required modules (Optional as included in the script)**:
   ```bash
   Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
   Install-Module -Name NameIT -Scope CurrentUser -Force -AllowClobber
   ```
4. **Implement Variables and Run the Script**:
   - Open the `CreateSites.ps1` script.
   - Customize and/or confirm the variables at the top of the script as needed (_See below for usage_).
   - Through the PowerShell terminal, go to the directory with the 'CreateSites.ps1' file and run the script (_i.e. the cloned repo from step 1_).
   ```bash
   .\CreateSites.ps1
   ```

### Variables

The `CreateSites.ps1` script has the following variables that can be customized:

- **`$numSites`**: _(int)_ Number of sites to create. Default is 5.  
  Example: `-numSites 10`

- **`$tenantUrl`**: _(string)_ URL of your SharePoint tenant.  
  Example: `-tenantUrl "https://masamuneindustries.sharepoint.com"`

- **`$templateSite`**: _(string)_ URL of the site where the template is stored.  
  Example: `-templateSite "https://masamuneindustries.sharepoint.com/sites/PNPTemplates"`
- **`$templateUrl`**: _(string)_ **COMPLETE** URL of the PnP template **file**.  
  Example: `-templateUrl "https://masamuneindustries.sharepoint.com/sites/PNPTemplates/PNP_Templates/contosoworks.pnp"`
- **`$templateFilePath`**: _(string)_ **RELATIVE** Path to the PnP template file within the SharePoint site.  
  Example: `-templateFilePath "/sites/PNPTemplates/PNP_Templates/contosoworks.pnp"`
- **`$localTemplatePath`**: _(string)_ Local path to store the downloaded template file.  
  Example: `-localTemplatePath "C:\temp\contosoworks.pnp"`
- **`$tempFolderPath`**: _(string)_ Temporary folder path to store the pnp file temporarily.  
  Example: `-tempFolderPath "C:\temp"`
- **`$siteOwner`**: _(string)_ Owner of the newly created sites.
  Example: `-siteOwner "BobBobson@mYourTenant.onmicrosoft.com"`

## Challenges and Lessons:

- **PowerShell Scripting**: At first, I was unfamiliar with PowerShell scripting, but after working on this project, I have gained a much stronger understanding of the language and its capabilities. Even beyond the scope of this project, it'll be applicable to other tasks and projects in the future.
- **PnP Templates**: Gained experience in working with PnP templates for SharePoint. Specifically, the structure of XML files and its syntax. I still have some difficulties understanding how it goes through the XML and applies it to the screen. With practice, I'll get there; however, for now, I much prefer traditional front-end development (i.e. React with TypeScript).
- **Azure Integration**: Explored methods for automating processes with Azure services. For example, Azure Functions or Power Automate (Not implemented in this project... _yet_...)
- **Error Handling**: Trying to get more verbose and expletive error handling in the script. This was a challenge as when errors arose, it was often difficult to pinpoint the exact cause and/or get detailed error messages. To date, I still have CSOM errors and am not sure how to get more verbose or accurate information on how to resolve these kinds of errors.

## Future Improvements

- **Batching and Asynchronous Processing**: Implement batching and asynchronous processing to create sites more efficiently. I did try to implement this, but there were issues with throttling. It's definitely something that I'll be revisiting.
- **Scalability**: Optimize the script for better performance with large-scale site creation. For example, create a wider variety of templates for different departments and committees.
- **Customization**: Add more customization options for site creation and template application.
- **Automation**: Research and implement automation for the script using Azure Functions or Power Automate.
- **Error Handling**: Research improving error handling and logging to provide more detailed information on errors and how to resolve them.

## Contact

Feel free to reach out to me with any questions, suggestions, or feedback!

Happy provisioning! 🚀
