New-Module {

    Function Invoke-Prerequisites {

            [OutputType()]
            [CmdletBinding()]
            Param (
                [Parameter(Position = 1)]
                [string] $Tenant,
                [Parameter(Position = 2)]
                [string] $ClientID,
                [Parameter(Position = 3)]
                [string] $CertPath,
                [Parameter(Position = 4)]
                [string] $CertPass
            )

            If ((Get-Culture).LCID -eq "1033") {
                $Script:Date = (Get-Date).tostring("MM-dd-yy")
            }
            Else {
                $Script:Date = (Get-Date).tostring("dd-MM-yy")
            }

            $Script:Tenant = $Tenant

            $Script:TenantUrl   = "https://$($Tenant).sharepoint.com"
            $Script:AadDomain   = "$($Tenant).onmicrosoft.com"
            $Script:ClientID    = $ClientID
            $Script:CertPass    = $CertPass
            $Script:CertPath    = $CertPath

    }

    Function Get-Administrators {
        $Admins = Get-PnPSiteCollectionAdmin

        <# Below gets users who have full control - set as administrator via admin portal #>

        ForEach ($Admin in $Admins | Where-Object { $_.Title -notlike "*(Admin)*" -and $_ -ne "System Account" }) {

            $Datum = New-Object -TypeName PSObject

            $Datum | Add-Member -MemberType NoteProperty -Name Tenant -Value $Tenant
            $Datum | Add-Member -MemberType NoteProperty -Name Site -Value $SiteUrl
            $Datum | Add-Member -MemberType NoteProperty -Name Group -Value "Aministrators"
            $Datum | Add-Member -MemberType NoteProperty -Name Member -Value $Admin.Title
            $Datum | Add-Member -MemberType NoteProperty -Name Subsite -Value "No"

            $Script:Data += $Datum

        }
    }

    Function Get-OwnerNoGroup {

        Param(
            [Parameter(Mandatory = $true, Position = 0)]
            [string]$Subsite
        )

        $Web = Get-PnPWeb -Includes RoleAssignments
        ForEach ($RA in $Web.RoleAssignments) {
            $LoginName = Get-PnPProperty -ClientObject $($RA.Member) -Property LoginName
            $RoleBindings = Get-PnPProperty -ClientObject $RA -Property RoleDefinitionBindings
            If ($RoleBindings.Name -like "*Full Control*" -and $LoginName -notlike "*Owners*") {

                If ($LoginName -like "i:0#.f|membership|*") {
                    $LoginName = $LoginName.Split('|')[2]
                    $DisplayName = $LoginName.Split('@')[0].Replace('.', ' ')
                    $DisplayName = (Get-Culture).TextInfo.ToTitleCase($DisplayName)
                }

                $Datum = New-Object -TypeName PSObject

                $Datum | Add-Member -MemberType NoteProperty -Name Tenant -Value $Tenant

                If ($Subsite -eq "No") {
                    $Datum | Add-Member -MemberType NoteProperty -Name Site -Value $SiteUrl
                } 
                Else {
                    $Datum | Add-Member -MemberType NoteProperty -Name Site -Value $SubsiteUrl
                }

                $Datum | Add-Member -MemberType NoteProperty -Name Group -Value "N/A"
                $Datum | Add-Member -MemberType NoteProperty -Name Member -Value $DisplayName
                $Datum | Add-Member -MemberType NoteProperty -Name Subsite -Value $Subsite

                $Script:Data += $Datum
            }
        }
    }

    Function Get-OwnerFromGroup {

        Param(
            [Parameter(Mandatory = $true, Position = 0)]
            [string]$Subsite,
            [Parameter(Position = 1)]
            [string]$SiteUrl,
            [Parameter(Position = 1)]
            [string]$SubsiteUrl
        )

        $Groups = Get-PnPGroup | Where-Object { $_.Title -like "*Owner*" -or $_.Title -like "*Admin*" } | Select-Object Title, Users

        If ($Subsite -eq "No") { 
            Write-Host "Processing data for $($SiteUrl)"
        } 
        Else { 
            Write-Host "Processing data for $($SubsiteUrl)"
        }

        ForEach ($Group in $Groups) {

            $GroupPermission = Get-PnPGroupPermissions -Identity $Group.Title -ErrorAction SilentlyContinue | Where-Object { $_.Hidden -like "False" } 

            If ($GroupPermission.RoleTypeKind -eq "Administrator") {

                ForEach ($G in $Group.Users.Title | Where-Object { $_ -notlike "*(Admin)*" -and $_ -ne "System Account" }) {

                    $Datum = New-Object -TypeName PSObject

                    $Datum | Add-Member -MemberType NoteProperty -Name Tenant -Value $Tenant

                    If ($Subsite -eq "No") {
                        $Datum | Add-Member -MemberType NoteProperty -Name Site -Value $SiteUrl
                    } 
                    Else {
                        $Datum | Add-Member -MemberType NoteProperty -Name Site -Value $SubsiteUrl
                    }

                    $Datum | Add-Member -MemberType NoteProperty -Name Group -Value $Group.Title
                    $Datum | Add-Member -MemberType NoteProperty -Name Member -Value $G
                    $Datum | Add-Member -MemberType NoteProperty -Name Subsite -Value $Subsite

                    $Script:Data += $Datum

                }

            }

        }

    }

    Function Invoke-FilePicker {

        Write-Host "Select certificate .pfx file"

        Add-Type -AssemblyName System.Windows.Forms

        $Dialog = New-Object System.Windows.Forms.OpenFileDialog
        $Dialog.InitialDirectory = "$InitialDirectory"
        $Dialog.Title = "Select certificate .pfx file"
        $Dialog.Filter = "Certificate file|*.pfx"  
        $Dialog.Multiselect = $false
        $Result = $Dialog.ShowDialog()

        If ($Result -eq 'OK') {

            Try {

                $Script:CertPath = $Dialog.FileNames
            }

            Catch {

                $Script:CertPath = $null
                Break
            }
        }

        Else {

            #Shows upon cancellation of Save Menu
            Write-Host -ForegroundColor Yellow "Notice: No file selected."
            Break
        }

        Write-Host "CertPath - Invoke: $CertPath" # Testing only
        
    }

    Function Export-SPOAdmin {

        [OutputType()]
        [CmdletBinding()]
        Param (
            [Parameter(Mandatory = $true, Position = 1, HelpMessage = "Enter your O365 tenant name, like 'contoso'")]
            [ValidateNotNullorEmpty()]
            [string] $Tenant,
            [Parameter(Mandatory = $true, Position = 2, HelpMessage = "Enter your Az App Client ID")]
            [ValidateNotNullorEmpty()]
            [string] $ClientID
        )

        #ToDo: Exclude SubSites Switch

        BEGIN {

            Invoke-FilePicker

            $Script:CertPass = Read-Host "Enter your certificate password"

            Invoke-Prerequisites -Tenant $Tenant -ClientID $ClientID -CertPath "$CertPath" -CertPass $CertPass

            $Params = @{

                ClientId            = $ClientID
                CertificatePath     = $CertPath
                CertificatePassword = (ConvertTo-SecureString -AsPlainText $CertPass -Force)
                Url                 = $TenantUrl
                Tenant              = $AadDomain

            }

        }
        PROCESS {
            Connect-PnPOnline @Params -WarningAction SilentlyContinue

            $Script:Sites = Get-PnPTenantSite | Where-Object -Property Template -NotIn ("SRCHCEN#0", "SPSMSITEHOST#0", "APPCATALOG#0", "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1")

            $Sites = $Sites.Url

            # $Sites = $Sites | Select -First 5 -Skip 17 # Test only

            Disconnect-PnPOnline

            $Script:Data = @()

            ForEach ($SiteUrl in $Sites) {

                $Subsite = "No"

                $Params = @{

                    ClientId            = $ClientID
                    CertificatePath     = $CertPath
                    CertificatePassword = (ConvertTo-SecureString -AsPlainText $CertPass -Force)
                    Url                 = $SiteUrl
                    Tenant              = $AadDomain

                }

                Connect-PnPOnline @Params -WarningAction SilentlyContinue

                <# Below gets users who have full control - directly applied at the root #>
                Get-OwnerNoGroup -Subsite $Subsite

                <# Below gets users who have full control - inherited from a group #>
                Get-OwnerFromGroup -Subsite $Subsite -SiteUrl $SiteUrl

                <# Below gets users who have full control - set as administrator via admin portal #>
                Get-Administrators

                $SubSites = Get-PnPSubWebs -Recurse

                Disconnect-PnPOnline

                If ($SubSites) {

                    ForEach ($Site in $SubSites) {

                        $Subsite = "Yes"
                        $SubsiteUrl = $Site.Url

                        $Params = @{

                            ClientId            = $ClientID
                            CertificatePath     = $CertPath
                            CertificatePassword = (ConvertTo-SecureString -AsPlainText $CertPass -Force)
                            Url                 = $SubsiteUrl
                            Tenant              = $AadDomain

                        }

                        Connect-PnPOnline @Params -WarningAction SilentlyContinue

                        <# Below gets users who have full control - directly applied at the root #>
                        Get-OwnerNoGroup -Subsite $Subsite

                        <# Below gets users who have full control - inherited from a group #>
                        Get-OwnerFromGroup -Subsite $Subsite -SubsiteUrl $SubsiteUrl

                        Write-Host "Data processing complete for $($SubsiteUrl)" -ForegroundColor Yellow

                        Disconnect-PnPOnline

                    }

                }

                Write-Host "Data processing complete for $($SiteUrl)" -ForegroundColor Cyan

            }
        }
        END {
            # To do - Import-Excel module to be used to group the data in excel by Url

            # Export results to CSV
            If ($Data) {
                $Path = ".\"
                $FileName = "$Tenant-SPOAdmins-$Date.csv"
                $Data | Export-Csv -Path "$Path$FileName" -NoTypeInformation
                $Location = Get-Location
                Write-Host
                Write-Host "File called '$FileName' exported to $Location"
            }
            Else {
                Write-Host "No Data to Export"
            }
        }

    }

    Export-ModuleMember Export-SPOAdmin

} | Out-Null

