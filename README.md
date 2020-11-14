# PowerShell-Scripts

This repository contains some PowerShell scripts and functions that I have used to carry out tasks on cloud resources through the years. 

## Platforms

- Office 365
- Security & Compliance Center
- Licensing
- SharePoint PnP
- Azure
- Power Platform

## Scripts

Below is a breakdown of some of what you will find in this repository.

|Script|Function|Description|
|---|---|---|
|PnP-FileRetention.ps1|Add-FolderWithLabel|Adds a folder to a user's OneDrive or SharePoint Library and applies a retention label|

## More to come

Please visit [AlanPs1.io](AlanPs1.io) to learn more about what I do.

## Using PnP-FileRetention.ps1

Firstly you will need to dot source the script

```powershell
. .\PnP-FileRetention.ps1
```

Then you will run the following command

```powershell
Add-FolderWithLabel -CertPass <Certificate Password> -UserPrincipalName "john.doe@contoso.com"
```

Other examples available using:

```powershell
Get-Help Add-FolderWithLabel -Examples
```