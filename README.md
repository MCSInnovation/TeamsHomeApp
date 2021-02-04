# Microsoft Teams Home app #

### Summary ###
This is a repository that hosts the Teams Home app which intends to help users new to Microsoft Teams.

More people are working remotely than ever before, and they may be new to Microsoft Teams. We wanted to provide a solution for individual users to answer the most frequently asked questions straight from Microsoft Teams client - in their native language and in their own way. 

### Our Solution ###
- A simple “Home app” for customized guides for Teams usage
- Solution components:
  - A SharePoint site template
  - Localized Teams guide
  - App manifest
   
### Prerequisites ###
In order to be able to successfully use the Teams Home app:
- The installer should have the PnP.PowerShell modules installed and available. The module can be deployed from https://pnp.github.io/powershell/
- A person nominated within the tenant who will be a point of contact for further questions on the app.
- A SharePoint site that will host the necessary pages. The site can either already exist or the installation script will attempt to create it for you. In the latter case, the installer must have permissions to create new site collections in the SPO environment.
- Optionally a logo file. This is a .PNG file that you want to use within Teams to point to this app. Note, the .PNG will be resized as necessary by the installation script.

### Installation Process ###
To install the Teams Home app:
- Copy the repository files to your local environment and via PowerShell, navigate to the root folder of the app.
- Run ".\Add-TeamsHomeApp.ps1 -url https://contoso.sharepoint.com/sites/TeamsHome -Language en -TeamsHomeManagerEmail TeamManager@contoso.onmicrosoft.com -LogoFilePath "C:\TeamsApp\CompanyLogo.PNG""
  - "url" - Mandatory, the SPO site collection to host the help files.
  - "TeamHomeManagerEmail" - Mandatory, the Teams Home app contact person.
  - "Language" - Mandatory, currently choose between EN/DA/SV/FI/NB.
  - "LogoFilePath" - Optional, the full file path to the logo .PNG file to use.
  - "UpdateTeamsAppDetails" - Optional, provides the ability for theinstaller to control the application presentation name and description.

### Disclaimer ###
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
