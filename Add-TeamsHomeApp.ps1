<#
.SYNOPSIS
    This is a script to deploy the Team Home app and supporting content

.DESCRIPTION
    This script will attempt to create or connect to a supplied SharePoint site and
    populate it wil sepcific Teams related content to enable end users to better understand
    the intent and useage of Teams.

    A Teams app will be produced that can then be sideloaded by a user to test or deployed
    throughout a tenant via policy in Teams.

.EXAMPLE
    .\Add-TeamsHomeApp.ps1 -url https://contoso.sharepoint.com/sites/TeamsHome -Language en -TeamsHomeManagerEmail TeamManager@contoso.onmicrosoft.com -LogoFilePath "C:\TeamsApp\CompanyLogo.PNG" -UpdateTeamsAppDetails
    
.INPUTS
    -Url                   - The desired url of the Team Home app content
    -TeamsHomeManagerEmail - This must be the tenant email address of the contact for the application
    -Language              - The desired language for the site.
    -LogoFilePath          - The full Windows path to a .PNG file to be used for the logos (optional)
    -UpdateTeamsAppDetails - Change the name and description of the app before uploading the package (optional)

.OUTPUTS
    Configured SPO site
    Teams Home app in the application root, MCSTeamsHomeApp.zip
 
-----------------------------------------------------------------------------------------------------------------------------------
Script name : Add-TeamsHomeApp.ps1
Authors : Microsoft
Version : V2.0
Dependencies : PnP PowerShell modules
-----------------------------------------------------------------------------------------------------------------------------------
-----------------------------------------------------------------------------------------------------------------------------------
Version Changes:
Date:       Version: Changed By:     Info:
01/12/2020  V1.0     Microsoft       Initial script creation
04/02/2021  V2.0     Microsoft       Updates for PnP.PowerShell module

-----------------------------------------------------------------------------------------------------------------------------------
DISCLAIMER
   THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
   MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES
   OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR
   PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR
   ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS, BUSINESS
   INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR
   INABILITY TO USE THE SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES.
   BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION OF LIABILITY FOR CONSEQUENTIAL OR
   INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.
#>
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true, HelpMessage = "What is the site or Page URL you want to work with?")]
    [string]$Url,

    [Parameter(Mandatory = $true, HelpMessage = "What is the email address of the Teams Home Manager?")]
    [string]$TeamsHomeManagerEmail,

    [Parameter(Mandatory = $true, HelpMessage = "Which language code do you want to install?")]
    [ValidateSet("en", "da", "sv", "fi", "nb")]
    [string]$Language,

    [Parameter(Mandatory = $false, HelpMessage = "Where is the PNG logo file you want to work with? It must be a PNG formatted image")]
    [string]$LogoFilePath,

    [Parameter(Mandatory = $false, HelpMessage = "Use this switch to control the App Name and Descriptions that are used for the deployment")]
    [switch]$UpdateTeamsAppDetails
)

#If the host is an appropriate PowerShell version, set reusethreading.
if ($Host.Version.Major -gt 1) {$Host.Runspace.ThreadOptions = "ReuseThread"}

#Inform the user how long the operation took. This is the start date/time.
$dtStartTime = Get-Date
$Version = "V2.0"
 
Write-Host "#################################################################################" -ForegroundColor Cyan
Write-Host "[$($Version)] Starting the Team Home App install process at $($dtStartTime.ToString('HH:mm:ss'))" -ForegroundColor Cyan
Write-Host "#################################################################################" -ForegroundColor Cyan


<#
.SYNOPSIS
    This function is used to output screen information

.DESCRIPTION
    This function will will output to current console formatted information.

.EXAMPLE
    LogStatus -Routine "The routine running" -LogType Information -Detail "This is information to log and [TEXT] between brackets will be coloured"

.INPUTS
    -Routine - Used to output information on where the code is currently
    -logType - Either Information, Success or Fail. Each type will render the output differently
    -Detail  - String information to be displayed to screen

.OUTPUTS
    Output to screen
#>
function LogStatus($Routine, $logType, $Detail) {
    switch ($logType) {
        "FAIL" {
            Write-Host "$($Routine) " -ForegroundColor Cyan -NoNewline; Write-Host "$($Detail)" -ForegroundColor Red
        }
        "HIGHLIGHT" {
            Write-Host "$($Routine) " -ForegroundColor Cyan -NoNewline; Write-Host "$($Detail)" -ForegroundColor Green
        }
        "INFORMATION" {
            #We place characters between "[" and "]" in yellow to ease viewing
            Write-Host "$($Routine) " -ForegroundColor Cyan -NoNewline

            foreach ($sDetail in ($Detail.Replace("[", "{[~").Replace("]", "]}").Split("[").Split("]"))) {
                $sDetail = $sDetail.Replace("{", "[").Replace("}", "]")

                if ($sDetail.IndexOf("~") -ne -1) {
                    Write-Host $sDetail.Replace("~", "") -ForegroundColor Yellow -NoNewline
                }
                else {
                    Write-Host $sDetail -ForegroundColor White -NoNewline 
                }
            }
            Write-Host ""
        }
        "SUCCESS" {
            #We place characters between "[" and "]" in green to ease viewing
            Write-Host "$($Routine) " -ForegroundColor Cyan -NoNewline

            foreach ($sDetail in ($Detail.Replace("[", "{[~").Replace("]", "]}").Split("[").Split("]"))) {
                $sDetail = $sDetail.Replace("{", "[").Replace("}", "]")

                if ($sDetail.IndexOf("~") -ne -1) { Write-Host $sDetail.Replace("~", "") -ForegroundColor Green -NoNewline }
                else { Write-Host $sDetail -ForegroundColor White -NoNewline }
            }
            Write-Host ""
        }
    }
}


<#
.SYNOPSIS
    This function is used to create or connect to the Team Home app site

.DESCRIPTION
    This function will check to see if the supplied site url already exists.

    If it does, then further checks are made to ensure that the Lcid is as expected
    and that the site template used is set to a communication site. If either of these
    checks fails, then the user is informed and the process stops.

    If the site does not exist, then we attempt to create it. This does assume that the
    user running the tool has permission to connect to the root site for the tenant, i.e.
    https://contoso.sharepoint.com, and that they have permission to create new sites in
    the tenant.

.EXAMPLE
    Get-TASuppliedUrl -Url https://contoso.sharepoint.com/sites/TeamsHome -Lcid 1033

.INPUTS
    -Url  - The Full rul of the desired site to create/connect to
    -Lcid - The language Id needed for the site

.OUTPUTS
    $siteUrl - The url of the site. If $null, then there was an issue and the process will stop
        
#>
function Get-TASuppliedUrl($Url, $Lcid) {
    try {
        #Does the site already exist?
        # Connect to the tenant rootweb and check. This assumes the user has appropriate permissions to https://<tenant>.sharepoint.com
        LogStatus -Routine Get-TASuppliedUrl -logType Highlight -Detail "If prompted, please supply the credentials of a user account permissioned to create new site collections or has 'Owner' permission on the existing site."
        Connect-PnPOnline -Url $Url.Substring(0, $Url.IndexOf('.com/') + 4) -UseWebLogin

        #This is used to pass back the url of the site to use. It will be empty if there is a problem
        $siteUrl = $null

        #This will fail if the supplied url does not exist. We will try to create the site instead
        Get-PnPTenantSite -Url $Url -ErrorAction Stop | Out-Null

        #The site exists. Do we have permissions to it and is it the correct type/Lcid
        try { 
            #Set context to the correct web
            Connect-PnPOnline -Url $Url -UseWebLogin

            $web = Get-PnPWeb -Includes Language, WebTemplate -ErrorAction Stop

            #Managed to connect to the requested site
            LogStatus -Routine Get-TASuppliedUrl -logType Information -Detail "Connected to [$($Url)]. Performing checks..."

            #If we get to this stage, then this means the site collection exists and we have privs to access it.
            # Now check to see if the site type is correct
            if ($web.Language -eq $Lcid) {
                #It is in the correct language. Check the site type
                if ($web.WebTemplate -eq 'SITEPAGEPUBLISHING') {
                    #It is the correct site type
                    $siteUrl = $Url
                }
                else {
                    #The site exists as requested but the template type is different.
                    LogStatus -Routine Get-TASuppliedUrl -logType Fail -Detail "The supplied url [$($Url)] exists but it's template type is [$($web.WebTemplate)]. The site should be a Communication site [SITEPAGEPUBLISHING]. Please try again."
                }
            }
            else {
                #The site exists as requested but the Lcid value is different.
                LogStatus -Routine Get-TASuppliedUrl -logType Fail -Detail "The supplied url [$($Url)] exists but it's Lcid value is [$($web.Language)]. You requested [$($Lcid)]. Please try again."
            }
        }
        catch {
            LogStatus -Routine Get-SuppliedUrl -logType Fail -Detail "You do not have permissions to access the existing site [$($Url)]. Please retry with another url."
        }
    }
    catch {
        LogStatus -Routine Get-TASuppliedUrl -logType Information -Detail "The supplied Url [$($Url)] does not exist. Attempting to create it..."
        $siteOwner = Read-Host "Please enter your cloud identity that you are running this process as to create the new site collection. i.e. <user>@$($url.Substring($url.IndexOf('://') + 3, ($url.IndexOf('.sharepoint.com/')) - ($url.IndexOf('://') + 3))).onmicrosoft.com"
        $siteUrl = New-PnPSite -Type CommunicationSite -Title "Microsoft Teams Home" -Url $Url -SiteDesign Showcase -Lcid $Lcid -Description "This site hosts the content for the Teams Home app" -Owner $siteOwner
        Connect-PnPOnline -Url $siteUrl -UseWebLogin
    }
    return $siteUrl
}


<#
.SYNOPSIS
    This function is used to update the manifest file

.DESCRIPTION
    This function will use the manifest.template file that is in the application
    deployment pack in the app_package directory, and create a relevant 
    manifest.json for this deployment

.EXAMPLE
    Update-TAManifest -manifestPath ".\en\TeamsApp\app-package\manifest.template" -siteUrl https://contoso.sharepoint.com/sites/TeamsHomeApp

.INPUTS
    $manifestPath - the directory location to the deployment manifest.template file
    $siteUrl - The full SPO site url to the Team Home app site

.OUTPUTS
    Creation of a new manifest.json file
    
#>
function Update-TAManifest($manifestPath, $siteUrl) {

    $manifestJson = Get-Content $manifestPath | ConvertFrom-Json

    #Update name and description
    $manifestJson.packageName = $appName
    $manifestJson.name.short = $appName
    $manifestJson.description.short = $appName
    $manifestJson.description.full = $appDescription

    
    #Update the Teams app staticTabs information
    ForEach ($staticTab in $manifestJson.staticTabs) {
        $staticTab.contentUrl = $staticTab.contentUrl -replace "<TeamsHomeSiteURL>", "$($siteUrl.Substring($siteUrl.IndexOf('://') + 3))" -replace "<TeamsHomeSiteRelative>", "$($siteUrl.Substring($siteUrl.IndexOf('.com/') + 4))"
        $staticTab.websiteUrl = $staticTab.websiteUrl -replace "<TeamsHomeSiteURL>", "$($siteUrl.Substring($siteUrl.IndexOf('://') + 3))" -replace "<TeamsHomeSiteRelative>", "$($siteUrl.Substring($siteUrl.IndexOf('.com/') + 4))"
    }

    #Update the Teams app validDomains information
    $manifestJson.validDomains = $manifestJson.validDomains -replace "<TeamsHomeSiteDomain>", "$($siteUrl.Substring($siteUrl.IndexOf('://') + 3, $siteUrl.IndexOf('.com/') - 4))"

    #Update the Teams app webApplicationInfo resource information
    $manifestJson.webApplicationInfo.resource = $siteUrl

    #Save the template as the new manifest.json file
    $manifestJson | ConvertTo-Json -Depth 10 | ForEach-Object { [System.Text.RegularExpressions.Regex]::Unescape($_) } | Set-Content -Encoding UTF8 ($manifestPath -replace "-template.", ".")
}


<#
.SYNOPSIS
    This function is used to create and copy logo renditions

.DESCRIPTION
    For a Teams app, you can use 2 logo files of set sizes. Color.png must
    by 192*192px and the outline.png 32*32px. We take the supplied logo
    file and as long as it is a PNG, we will create the necessary versions.

.EXAMPLE
    Get-TALogos -filePath "C:\temp\companylogo.png"

.INPUTS
    $filePath - the file location to the required logo file

.OUTPUTS
    2 renditions of the logo file, Color.png and Outline.png
    Returns status of the operation
    
#>
function Get-TALogos($filePath) {

    try {
        if ((Get-ChildItem -Path $filePath).Extension.ToLower() -eq ".png") {
            #Get the supplied image
            $imgLogo = [System.Drawing.Image]::FromFile((Get-Item $filePath))

            #Create the canvas with the required dimensions
            $imgColor = New-Object System.Drawing.Bitmap(192, 192)
            $imgOutline = New-Object System.Drawing.Bitmap(32, 32)

            #Create a new drawing surface from the image
            $graphColor = [System.Drawing.Graphics]::FromImage($imgColor)
            $graphOutline = [System.Drawing.Graphics]::FromImage($imgOutline)

            #Draws the images
            $graphColor.DrawImage($imgLogo, 0, 0, 192, 192)
            $graphOutline.DrawImage($imgLogo, 0, 0, 32, 32)

            #Saves the resized images
            $imgColor.Save(".\$($Language)\TeamsApp\app-package\color.png")
            $imgOutline.Save(".\$($Language)\TeamsApp\app-package\outline.png")

            LogStatus -Routine Get-TAlogos -logType Success -Detail "The [color.png] and [outline.png] logo files have been created."

            return $true
        }
        else {
            LogStatus -Routine Get-TAlogos -logType Failure -Detail "The supplied logo [$($filePath)] is not a .PNG file. The default ones will be used."
            return $false
        }
    }
    catch {
        return $false
    }
}


###
#Main routine
###

#Check for SharePoint PnP module
if (!(((Get-Module -ListAvailable SharePointPnPPowerShellOnline).Version -ge "3.26.2010.0") -or ((Get-Module -ListAvailable PnP.PowerShell).Version -ge "1.2.0"))) {
    LogStatus -Routine Main -logType Fail -Detail "Please install the latest appropriate PnP PowerShell module. Visit https://github.com/pnp/powershell for further details."
}
else {
    #Get the root to copy the files to or else use the current location
    if (Test-Path -LiteralPath ".\$($Language)\SharePointSite\template.xml" -PathType Leaf) {

        #Put the supplied url into a known state
        $Url = $Url.TrimEnd('/').ToLower()
        $Language = $Language.ToLower()

        if ($Language -eq 'da') {
            $Lcid = "1030"
        }
        elseif ($Language -eq 'sv') {
            $Lcid = "1053"
        }
        elseif ($Language -eq 'fi') {
            $Lcid = "1035"
        }
        elseif ($Language -eq 'nb') {
            $Lcid = "1044"
        }
        else {
            $Lcid = "1033"
        }

        #Get any further information needed before progressing
        If ($UpdateTeamsAppDetails.IsPresent) {
            #The user has requested to update the Teams app detail
            $appName = Read-Host "Please enter your name for the Teams Home App (Examples: 'Home App', 'Contoso', etc.)"
            $appDescription = Read-Host "Please enter your description for the Teams Home App (Examples: 'This is an application to demonstrate a home application for Microsoft Teams in English.')"
        }
        else {
            $appName = "Teams Home"
            $appDescription = "This is an application to demonstrate a home application for Microsoft Teams"
        }

        ###
        #Logo management
        $companyLogoUsed = $false #This is for the logo management
        if ($logoFilePath){
            if (-not (Test-Path $logoFilePath))
            {
                do
                {
                    $logoFilePath = Read-Host "Cannot find the supplied logo file [$($logoFilePath)]. Please RE-ENTER "
                } while (-not (Test-Path $logoFilePath))
            }

            #Copy renditions of the file to the Teams package
            $companyLogoUsed = Get-TALogos -filePath $logoFilePath
        }

        #There was an issue with the logo file supplied or no logo file was supplied
        if (!$companyLogoUsed) {
            #We want to make a copy of the template logo files 
            Copy-Item ".\$($Language)\TeamsApp\app-package\color-template.png" -Destination ".\$($Language)\TeamsApp\app-package\color.png"
            Copy-Item ".\$($Language)\TeamsApp\app-package\outline-template.png" -Destination ".\$($Language)\TeamsApp\app-package\outline.png"
        }
        
        ###
        #SPO section
        #
        #Get context into the site and apply the PnP template
        if (Get-TASuppliedUrl -Url $Url -Lcid $Lcid -PSModule $PnPModule) {

            #We have a valid site and context. Apply the PnP site template to it.
            try {
                $TAManager = New-PnPUser -LoginName $TeamsHomeManagerEmail -ErrorAction Stop

                #Run the correct cmdlet for the apporpiate PowerShell module used.
                # The PnP.PowerShell modules uses a differenet provisioning cmdlet
                if ($null -ne (Get-Module -ListAvailable PnP.PowerShell)) {
                    Invoke-PnPSiteTemplate -Path ".\$($Language)\SharePointSite\template.xml" -Parameter @{"TAManagerTitle" = $TAManager.Title; "TAManagerEmail" = $TAManager.Email; "TASiteOwner" = $TAManager.LoginName; "TAAllUsers" = "c:0-.f|rolemanager|spo-grid-all-users/$(Get-PnPAuthenticationRealm)";}
                }
                else {
                    Apply-PnPProvisioningTemplate -Path ".\$($Language)\SharePointSite\template.xml" -Parameter @{"TAManagerTitle" = $TAManager.Title; "TAManagerEmail" = $TAManager.Email; "TASiteOwner" = $TAManager.LoginName; "TAAllUsers" = "c:0-.f|rolemanager|spo-grid-all-users/$(Get-PnPAuthenticationRealm)";}
                }

                ###
                #Teams section
                #
                #We need to modify the approriate manifest file and package up
                Update-TAManifest -manifest ".\$($Language)\TeamsApp\app-package\manifest-template.json" -siteUrl (Get-PnPContext).Url

                #Create the Teams app package/zip file
                Get-ChildItem -Path ".\$($Language)\TeamsApp\app-package\*" -Exclude "*-template.*" | Compress-Archive -DestinationPath .\MCSTeamsHomeApp.zip -Force

                #Tell the user that the app is ready to sideload/deploy
                LogStatus -Routine Main -logType Highlight -Detail "The Team app is ready to deploy. Load it from [$($(Get-Location).Path)\MCSTeamsHomeApp.zip]"
            }
            catch{
                #Some issue, probably incorrect email address though
                LogStatus -Routine Main -logType Fail -Detail "Problem encountered. Is the Teams Application manager email [$($TeamsHomeManagerEmail)] correct?. Please try again."
            }
        }
        else {
            #Can't access the supplied url or create it
            LogStatus -Routine Main -logType Fail -Detail "Cannot connect or create the requested site [$($Url)]. Please try again."
        }
    }
    else {
        #Package not downloaded correctly or the user hasn't set the current directory to the root of the package
        LogStatus -Routine Main -logType Fail -Detail "Cannot find [$(".\$($Language)\SharePointSite\template.xml")]. Please make sure you set your default location (cd) to the root of the package where this script file is located, then try again."
    }
}

Write-Host "#################################################################################" -ForegroundColor Cyan
Write-Host "Completed..." -ForegroundColor Cyan
Write-Host "The process took - $((New-TimeSpan -Start $dtStartTime -End (Get-Date)).Minutes) minutes $((New-TimeSpan -Start $dtStartTime -End (Get-Date)).Seconds) seconds to complete." -ForegroundColor Cyan
Write-Host "#################################################################################" -ForegroundColor Cyan
