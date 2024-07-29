# Variables
$numSites = 5  ## Change this value to the number of sites you want to create ##

$tenantUrl = "https://masamuneindustries.sharepoint.com"
$templateUrl = "https://masamuneindustries.sharepoint.com/sites/PNPTemplates/PNP_Templates/contosoworks.pnp"
$templateSite = "https://masamuneindustries.sharepoint.com/sites/PNPTemplates"
$templateFilePath = "/sites/PNPTemplates/PNP_Templates/contosoworks.pnp"
$localTemplatePath = "C:\temp\contosoworks.pnp"
$tempFolderPath = "C:\temp"
$siteOwner = "timbroderick@masamuneindustries.onmicrosoft.com"

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
    committee = @("Council of", "League of", "Cabal of", "Order of", "Squad of", "Guild of", "Society of", "Fellowship of", "Assembly of", "Horde of", "Band of", 
    "Consortium of", "Congregation of", "Tribe of", "Coterie of", "Crew of", "Clan of", "Brigade of", "Alliance of", "Circle of", "Union of", "Coalition of",
    "Federation of", "Syndicate of", "Network of", "Confederation of", "Corporation of", "Corps of", "Division of", "Sect of", "Chapter of", "Department of",
    "Office of", "Bureau of", "Agency of", "Institute of", "Center of", "Foundation of", "Incorporated of", "Corporation of", "Company of", "Association of")
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

# Generate random site details
function Get-RandomSiteDetails {
    $department = Invoke-Generate "[department]" -CustomData $CustomData
    $committee = Invoke-Generate "[committee]" -CustomData $CustomData 
    $title = "$committee $department"
    $description = "The $committee of $department site"
    $url = Invoke-Generate "$committee$department-[state abbr]##" -CustomData $CustomData
    $url = $url -replace '\s', '' 
    return [PSCustomObject]@{
        Title       = $title
        Description = $description
        URL         = "$tenantUrl/sites/$url"
        BenefitsTitle = "$department Benefits"
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
            Write-Host "Connecting to main site $siteUrl to apply the template (Attempt $try)"
            Connect-PnPOnline -Url $siteUrl -Interactive
            
            # Download template locally
            Write-Host "Downloading PnP template from $templateUrl"
            Connect-PnPOnline -Url $templateSite -Interactive
            Get-PnPFile -Url $templateFilePath -Path $tempFolderPath -FileName "contosoworks.pnp" -AsFile
            
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
        Write-Host "Waiting 60 seconds before applying template to main site... It's a bit sensitive..."
        Start-Sleep -Seconds 60

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
