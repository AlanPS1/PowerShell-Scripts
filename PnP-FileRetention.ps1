New-Module {

    Function Invoke-Prerequisites {

        [OutputType()]
        [CmdletBinding()]
        Param (
        [Parameter(Mandatory = $true, Position = 1)]
        [string]$CertPass,
        [Parameter(Mandatory = $false, HelpMessage = "Enter your O365 tenant name, like 'contoso'")]
        [ValidateNotNullorEmpty()]
        [string] $Tenant = "wonderful12345",
        [Parameter(Mandatory = $false, HelpMessage = "Enter your Az App Client ID")]
        [ValidateNotNullorEmpty()]
        [string] $ClientID = "1f7e91f2-0d86-45bc-9c66-404e4958498f",
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, HelpMessage = "Enter certificate .pfx path")]
        [String]$CertPath = "$Home\AppData\Local\LabSPOAccess.pfx"
        )

        $Script:MySPUrl     = "https://$($Tenant)-my.sharepoint.com/personal"
        $Script:AadDomain   = "$($Tenant).onmicrosoft.com"
        $Script:ClientID    = $ClientID
        $Script:CertPass    = $CertPass
        $Script:CertPath    = $CertPath

        If ($CertPath -eq $null) {

            Add-Type -AssemblyName System.Windows.Forms

            $Dialog = New-Object System.Windows.Forms.OpenFileDialog
            $Dialog.InitialDirectory = "$InitialDirectory"
            $Dialog.Title = "Select certificate file"
            $Dialog.Filter = "Certificate file|*.pfx"        
            $Dialog.Multiselect = $false
            $Result = $Dialog.ShowDialog()

            if ($Result -eq 'OK') {

                Try {
    
                    $Script:CertPath = $Dialog.FileNames
                }

                Catch {

                    $Script:CertPath = $null
                    Break
                }
            }

            else {

                #Shows upon cancellation of Save Menu
                Write-Host -ForegroundColor Yellow "Notice: No file selected."
                Break
            }
        }

    }

    Function Add-FolderWithLabel {

        <#
        .SYNOPSIS
        Add-FolderWithLabel firstly checks if the user has provisioned there OneDrive. If the
        user has not provisioned their OneDrive the script will do it for them. After that, the 
        script goes on to create either 3 or 4 folders within their OneDrive and adds the
        required retention label to each folder. There is a an additional label that gets added
        to the root "Documents" folder. There is no need to create that folder, just to apply
        the label. 
        
        This function will handle both authentication to PnP Online using a service pricipal and 
        then disconnect the session at the end. 
        
        This function can process multiple users at one time.

        .DESCRIPTION
        This Function will deploy 3 folders within a user's OneDrive Folder.
        When using -Designs switch, a 4th folder is deployed cad "My Designs".
        Each folder's label has a 6 month retention applied except for "My Designs"
        which has a 7 year retention applied.

        .PARAMETER Password
        Mandatory parameter for the service principal certificate password.

        .PARAMETER UserPrincipalName
        Mandatory parameter for the service principal certificate password.
        
        .PARAMETER Designs
        Creates a 4th folder called "My Des Wosk" and applies a label called "DesinsWork"

        .EXAMPLE
        PS C:\> Add-FolderWithLabel -CertPass <Cert Pass> -UserPrincipalName "john.doe@contoso.com" -Verbose

        This will create the following 3 folders and apply the required label to each folder.

        My Recipes      : Recipes
        My WorkOuts     : WorkOuts
        My Certificates : Certificates

        A label called "Default" is applied to the root Documents folder.

        .EXAMPLE
        PS C:\> Add-FolderWithLabel -CertPass <Cert Pass> -UserPrincipalName "john.doe@contoso.com", "jane.doe@contoso.com"

        This will create the following 3 folders and apply the required label to each folder.

        My Recipes      : Recipes
        My WorkOuts     : WorkOuts
        My Certificates : Certificates

        A label called "Default" is applied to the root Documents folder.

        .EXAMPLE
        PS C:\> Add-FolderWithLabel -CertPass <Cert Pass>  -UserPrincipalName "john.doe@contoso.com" -Designsork

        This will create the following 4 folders and apply the required label to each folder.

        My Recipes      : Recipes
        My WorkOuts     : WorkOuts
        My Certificates : Certificates
        My Designs      : Designs

        A label called "Default" is applied to the root Documents folder.

        .NOTES

        Author:  Alan Wightman
        Website: http://AlanPs1.io
        Twitter: @AlanO365

        #>

        # Todo: Implement a prompt and ask for a mandatory pair of Folder & Label
        # Once user enters all folder & label combinations type "Done" or similar to move on

        [OutputType()]
        [CmdletBinding()]
        Param (
        [Parameter(Mandatory = $true, Position = 1)]
        [string]$CertPass,
        [Parameter(Mandatory = $true, Position = 2)]
        [string[]] $UserPrincipalName,
        [Parameter(Mandatory = $false, Position = 3)]
        [switch]$Designs
        )

        Invoke-Prerequisites $CertPass

        Foreach ($User in $UserPrincipalName) {

        $SiteUrl = "$($MySPUrl)/$(($User).Replace("@", "_").Replace(".", "_"))"

        $Params = @{
            ClientId            = $ClientID
            CertificatePath     = $CertPath
            CertificatePassword = (ConvertTo-SecureString -AsPlainText $CertPass -Force)
            Url                 = $SiteUrl
            Tenant              = $AadDomain
        }

        Connect-PnPOnline @Params

        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl) 
        $web = $Context.Web 
        $Context.Load($Web)

        $TheUser = ($User).Split("@")[0]

        Write-Host "Verifying User:" $TheUser

            Try {
                $Response = Invoke-WebRequest -Uri $SiteUrl
                $StatusCode = $Response.StatusCode
            }
            Catch {
                $StatusCode = $_.Exception.Response.StatusCode.value__
            }

            If ($StatusCode -eq 200 -or $StatusCode -eq 400) {
                Write-Host -ForegroundColor Green "Site for user $TheUser has already been provisioned"
            }
            Else {
                Write-Host -ForegroundColor Yellow "Provisioning OneDrive for User: $TheUser"
                New-PnPPersonalSite -Email $User
            }

        $Response = $null
        $StatusCode = $null

        $Result = Get-PNPLabel -List "Documents"

            If ($Result[0] -notlike "*Default*") {

                Try {
                    Set-PnPLabel -List "Documents" -Label "Default" -SyncToItems $true -ErrorAction Stop
                    Write-Verbose "Label called 'Default' applied to 'Documents - Root' Folder"
                } 
                Catch [Microsoft.SharePoint.Client.ServerException] {
                    Write-Warning -Message $($_.Exception.Message)
                } 
                Catch {
                    Write-Warning -Message $($_.Exception.Message)
                }

            }

        $Folder = "My Recipes"

            Try {
                Get-PnPFolder -Url "Documents/$($Folder)" -ErrorAction Stop
            } 
            Catch [Microsoft.SharePoint.Client.ServerException] {
                Add-PnPFolder -Name $Folder  -Folder "Documents" -ErrorAction Stop
            } 
            Catch {
                Write-Warning -Message $($_.Exception.Message)
            }

        $Folder = Get-PnPFolder -Url "Documents/$($Folder)"
        $Folder.ListItemAllFields.SetComplianceTagWithNoHold("Recipes") 
        Invoke-PnPQuery
        Write-Verbose "Label called 'Recipes' applied to '$($Folder)' Folder"

        $Folder = "My Workouts"

            Try {
                Get-PnPFolder -Url "Documents/$($Folder)" -ErrorAction Stop
            } 
            Catch [Microsoft.SharePoint.Client.ServerException] {
                Add-PnPFolder -Name $Folder  -Folder "Documents" -ErrorAction Stop
            } 
            Catch {
                Write-Warning -Message $($_.Exception.Message)
            }

        $Folder = Get-PnPFolder -Url "Documents/$($Folder)"
        $Folder.ListItemAllFields.SetComplianceTagWithNoHold("WorkOuts") 
        Invoke-PnPQuery
        Write-Verbose "Label called 'WorkOuts' applied to '$($Folder)' Folder"

        $Folder = "My Certificates"

            Try {
                Get-PnPFolder -Url "Documents/$($Folder)" -ErrorAction Stop
            } 
            Catch [Microsoft.SharePoint.Client.ServerException] {
                Add-PnPFolder -Name $Folder  -Folder "Documents" -ErrorAction Stop
            } 
            Catch {
                Write-Warning -Message $($_.Exception.Message)
            }

        $Folder = Get-PnPFolder -Url "Documents/$($Folder)"
        $Folder.ListItemAllFields.SetComplianceTagWithNoHold("Certificates")
        Invoke-PnPQuery
        Write-Verbose "Label called 'Certificates' applied to '$($Folder)' Folder"

            If ($PSBoundParameters.Keys -contains "Designs") {
                
                $Folder = "My Designs"
                
                Try {
                    Get-PnPFolder -Url "Documents/$($Folder)" -ErrorAction Stop
                } 
                Catch [Microsoft.SharePoint.Client.ServerException] {
                    Add-PnPFolder -Name $Folder  -Folder "Documents" -ErrorAction Stop
                } 
                Catch {
                    Write-Warning -Message $($_.Exception.Message)
                }
                
                $Folder = Get-PnPFolder -Url "Documents/$($Folder)"
                $Folder.ListItemAllFields.SetComplianceTagWithNoHold("Designs")
                Invoke-PnPQuery
                Write-Verbose "Label called 'Designs' applied to '$($Folder)' Folder"
            }

        Disconnect-PnPOnline

        }

    }

    # ToDo: Find-MissingFolder - This should be performed on 1 to many userprincipalnames (Pipeline compatible)
    # ToDo: Find-MissingLabel - This should be performed on 1 to many userprincipalnames (Pipeline compatible)
    # ToDo: Get-FolderStatus - This should be performed on 1 to many userprincipalnames (Pipeline compatible) & check for a single folder's existence. 
    # Record output of above, Folder exists $boolean, Label, Folder Created Date, Creator of folder

Export-ModuleMember Add-FolderWithLabel

} | Out-Null

<#

$Result = Get-PnPFolder -Url "Documents/My Recipes"
Remove-PnPFolder -Folder "Documents" -Name "My Recipres"

Get-PnPFolder -Url "Documents"
Get-PnPFolder -Url "Documents/My Recipes"
Get-PnPFolder -Url "Documents/My WorkOuts"
Get-PnPFolder -Url "Documents/My Certificates"
Get-PnPFolder -Url "Documents/My Designs"

Get-PnPFolderItem -FolderSiteRelativeUrl "Documents"
Get-PnPFolderItem -FolderSiteRelativeUrl "Documents/My Recipes"
Get-PnPFolderItem -FolderSiteRelativeUrl "Documents/My WorkOuts"
Get-PnPFolderItem -FolderSiteRelativeUrl "Documents/My Certificates"
Get-PnPFolderItem -FolderSiteRelativeUrl "Documents/My Designs"

Get-PnPFolderItem -FolderSiteRelativeUrl "Documents/My Designs" | Select -First 1

Get-PnPLabel -List "Documents"
Get-PnPLabel -List "Documents/My Recipes"
Get-PnPLabel -List "Documents/My WorkOuts"
Get-PnPLabel -List "Documents/My Certificates"
Get-PnPLabel -List "Documents/My Designs"

Reset-PnPLabel -List "Documents"
Reset-PnPLabel -List "Documents/My Recipes"
Reset-PnPLabel -List "Documents/My WorkOuts"
Reset-PnPLabel -List "Documents/My Certificates"
Reset-PnPLabel -List "Documents/My Designs"

Set-PnPLabel -List "Documents" -Label "Default"
Set-PnPLabel -List "Documents/My Recipes" -Label "Recipes"
Set-PnPLabel -List "Documents/My WorkOuts" -Label "WorkOuts"
Set-PnPLabel -List "Documents/My Certificates" -Label "Certificates"
Set-PnPLabel -List "Documents/My Designs" -Label "Designs"

Set-PnPLabel -List "Documents" -Label "Default"

#>