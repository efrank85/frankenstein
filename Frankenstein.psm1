<#
.SYNOPSIS
    Creation of Eric Frank. Discovers Exchange On-Premises and Online Information.

.DESCRIPTION
    This module contains functions used to collect and export data in preparation
    for an Exchange to Exchange Online migration.

.EXAMPLE
    Get-FrankensteinExchangeDiscovery -Online -CSV -UseCurrentSession -PublicFolders
    Get-FrankensteinGSuiteDiscovery -CSV

.OUTPUTS
    CSV and .txt transcript files

.NOTES
    Author:  Eric D. Frank
    11/07/23 - Updated to use GitHub as repository
#>

#region Helpers

function Get-FrankensteinHelp {
    [CmdletBinding()]
    Param()

    Write-Host @"

    Frankenstein offers several functions to assist in the Exchange, Azure, and GSuite discovery processes.

    1) Get-FrankensteinExchangeDiscovery
       Provides Exchange on-premises or Exchange Online discovery. Outputs a transcript and optional CSVs.
       Switches: [-VirtualDirectories] [-CSV] [-UseCurrentSession] [-Online] [-PublicFolders]

    2) Get-FrankensteinPublicFolderDiscovery
       Outputs CSVs for Exchange Public Folder information.

    3) Get-FrankensteinGSuiteDiscovery
       Outputs G Suite discovery CSV files.
       Prerequisites: PSGsuite - https://psgsuite.io/
       Switches: [-CSV] [-IncludeGroupSettings] [-IncludeGroupMembership] [-IncludeDelegates] [-IncludeSendAsSettings] [-IncludeAutoForwardSettings]

    4) Install-M365Modules
       Installs M365 PowerShell modules. Use -All to install everything, or pick workloads individually.
       Switches: [-All] [-ExchangeOnline] [-Graph] [-SharePoint] [-PnP] [-Teams] [-Compliance] [-PowerPlatform]

    5) Connect-M365
       Connects to one or more M365 services using modern authentication.
       Switches: [-All] [-ExchangeOnline] [-Graph] [-SharePoint] [-PnP] [-Teams] [-Compliance] [-PowerPlatform]
       Parameters: [-SharePointAdminUrl <url>] [-GraphScopes <string[]>]

    6) Connect-ExchangeOnPremServer
       Connects to an on-premises Exchange server using FQDN.

    7) Get-FrankensteinRecipientCounts
       Displays a summary of all recipient types. Auto-detects Exchange Online vs On-Premises.

    8) Get-FrankensteinMailboxPermissions
       Retrieves FullAccess, SendAs, and SendOnBehalf permissions.
       Switches: [-FullAccess] [-SendAs] [-SendOnBehalf] [-UseCurrentSession] [-CSV] [-Help]

    9) Get-FrankensteinVirtualDirectories
       Reports on Exchange virtual directory URLs and authentication methods.
       Switches: [-CSV]

    10) Get-FrankensteinEntraDiscovery
        Comprehensive Entra ID (Azure AD) discovery via Microsoft Graph. Covers org info, users,
        MFA registration, admin roles, groups, devices, Conditional Access, apps, and security posture.
        Switches: [-CSV] [-UseCurrentSession]

"@
}

function Get-Linebreak {
    Write-Host "`n################################################################################################`n"
}

#endregion

#region Connection

function Connect-ExchangeOnPremServer {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [String]$ExchangeServerFQDN
    )
    $UserCredential = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange `
        -ConnectionUri "http://$ExchangeServerFQDN/PowerShell/" `
        -Authentication Kerberos `
        -Credential $UserCredential
    Import-PSSession $Session -DisableNameChecking
}

function Connect-M365 {
    [CmdletBinding()]
    Param (
        [Switch]$All,
        [Switch]$ExchangeOnline,
        [Switch]$Graph,
        [Switch]$SharePoint,
        [Switch]$PnP,
        [Switch]$Teams,
        [Switch]$Compliance,
        [Switch]$PowerPlatform,

        # Required for SharePoint and PnP connections
        [string]$SharePointAdminUrl,

        # Optional — defaults to a broad read/write admin scope set
        [string[]]$GraphScopes = @(
            "Directory.ReadWrite.All",
            "User.ReadWrite.All",
            "Group.ReadWrite.All",
            "Organization.Read.All",
            "Reports.Read.All",
            "RoleManagement.Read.Directory",
            "Policy.Read.All",
            "AuditLog.Read.All"
        )
    )

    if ($All) {
        $ExchangeOnline = $Graph = $SharePoint = $PnP = $Teams = $Compliance = $PowerPlatform = $true
    }

    if ((-not $ExchangeOnline) -and (-not $Graph) -and (-not $SharePoint) -and
        (-not $PnP) -and (-not $Teams) -and (-not $Compliance) -and (-not $PowerPlatform)) {
        Write-Warning "No workload specified. Use -All or one of: -ExchangeOnline, -Graph, -SharePoint, -PnP, -Teams, -Compliance, -PowerPlatform"
        return
    }

    if (($SharePoint -or $PnP) -and -not $SharePointAdminUrl) {
        $SharePointAdminUrl = Read-Host "SharePoint Admin URL (e.g. https://contoso-admin.sharepoint.com)"
    }

    if ($ExchangeOnline) {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
        Connect-ExchangeOnline
    }

    if ($Graph) {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes $GraphScopes
        # Import Graph sub-modules so cmdlets are available immediately
        @(
            "Microsoft.Graph.Identity.DirectoryManagement",
            "Microsoft.Graph.Users",
            "Microsoft.Graph.Groups",
            "Microsoft.Graph.Identity.SignIns",
            "Microsoft.Graph.Applications",
            "Microsoft.Graph.Reports"
        ) | ForEach-Object {
            if (-not (Get-Module -Name $_ -ErrorAction SilentlyContinue)) {
                Import-Module $_ -ErrorAction SilentlyContinue
            }
        }
    }

    if ($SharePoint) {
        Write-Host "Connecting to SharePoint Online..." -ForegroundColor Cyan
        Connect-SPOService -Url $SharePointAdminUrl
    }

    if ($PnP) {
        Write-Host "Connecting to PnP PowerShell..." -ForegroundColor Cyan
        Connect-PnPOnline -Url $SharePointAdminUrl -Interactive
    }

    if ($Teams) {
        Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Cyan
        Connect-MicrosoftTeams
    }

    if ($Compliance) {
        Write-Host "Connecting to Security & Compliance / Purview..." -ForegroundColor Cyan
        Connect-IPPSSession
    }

    if ($PowerPlatform) {
        Write-Host "Connecting to Power Platform..." -ForegroundColor Cyan
        Add-PowerAppsAccount
    }

    Write-Host "`nConnections complete." -ForegroundColor Green
}

#endregion

#region Installation

function Install-M365Modules {
    [CmdletBinding()]
    Param (
        [Switch]$All,
        [Switch]$ExchangeOnline,
        [Switch]$Graph,
        [Switch]$SharePoint,
        [Switch]$PnP,
        [Switch]$Teams,
        [Switch]$Compliance,     # Included in ExchangeOnlineManagement; listed for clarity
        [Switch]$PowerPlatform
    )

    if ($All) {
        $ExchangeOnline = $Graph = $SharePoint = $PnP = $Teams = $PowerPlatform = $true
    }

    if ((-not $ExchangeOnline) -and (-not $Graph) -and (-not $SharePoint) -and
        (-not $PnP) -and (-not $Teams) -and (-not $Compliance) -and (-not $PowerPlatform)) {
        Write-Warning "No workload specified. Use -All or one of: -ExchangeOnline, -Graph, -SharePoint, -PnP, -Teams, -Compliance, -PowerPlatform"
        return
    }

    # Ensure NuGet and PowerShellGet are up to date
    Write-Host "Bootstrapping NuGet and PowerShellGet..." -ForegroundColor Cyan
    Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
    Install-Module -Name PowerShellGet -Force -Scope CurrentUser -AllowClobber | Out-Null

    $modules = [ordered]@{}

    if ($ExchangeOnline -or $Compliance) {
        # ExchangeOnlineManagement covers both EXO (Connect-ExchangeOnline)
        # and Security & Compliance / Purview (Connect-IPPSSession)
        $modules["ExchangeOnlineManagement"] = "Exchange Online + Security & Compliance / Purview"
    }

    if ($Graph) {
        # Microsoft.Graph is the modern replacement for both AzureAD and MSOnline
        $modules["Microsoft.Graph"] = "Microsoft Graph (replaces AzureAD + MSOnline)"
    }

    if ($SharePoint) {
        $modules["Microsoft.Online.SharePoint.PowerShell"] = "SharePoint Online Administration"
    }

    if ($PnP) {
        # PnP.PowerShell is the recommended module for SharePoint and Teams site-level management
        $modules["PnP.PowerShell"] = "PnP PowerShell (SharePoint + Teams site management)"
    }

    if ($Teams) {
        $modules["MicrosoftTeams"] = "Microsoft Teams Administration"
    }

    if ($PowerPlatform) {
        $modules["Microsoft.PowerApps.Administration.PowerShell"] = "Power Platform Administration"
        $modules["Microsoft.PowerApps.PowerShell"]                = "Power Apps PowerShell"
    }

    foreach ($moduleName in $modules.Keys) {
        Write-Host "Installing $moduleName  ($($modules[$moduleName]))..." -ForegroundColor Cyan
        $existing = Get-Module -Name $moduleName -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
        $gallery  = Find-Module -Name $moduleName -ErrorAction SilentlyContinue

        if ($existing -and $gallery -and ($existing.Version -ge $gallery.Version)) {
            Write-Host "  $moduleName is already up to date (v$($existing.Version))." -ForegroundColor Gray
        }
        else {
            Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber -Confirm:$false
            Write-Host "  Installed $moduleName." -ForegroundColor Green
        }
    }

    Write-Host "`nInstallation complete. Run Connect-M365 to authenticate." -ForegroundColor Green
}

#endregion

#region Discovery

function Get-FrankensteinRecipientCounts {
    [CmdletBinding()]
    Param()

    Write-Host "Detecting Exchange environment..." -ForegroundColor Cyan

    if (Get-Command Get-EXOMailbox -ErrorAction SilentlyContinue) {
        $Environment  = "Exchange Online"
        $AllMailboxes = Get-EXOMailbox -ResultSize Unlimited -PropertySets All
        $AllDistGroups = Get-DistributionGroup -ResultSize Unlimited
        $CASMailbox   = Get-EXOCASMailbox -ResultSize Unlimited
    }
    elseif (Get-Command Get-Mailbox -ErrorAction SilentlyContinue) {
        $Environment  = "Exchange On-Premises"
        $AllMailboxes = Get-Mailbox -ResultSize Unlimited
        $AllDistGroups = Get-DistributionGroup -ResultSize Unlimited
        $CASMailbox   = Get-CASMailbox -ResultSize Unlimited
    }
    else {
        Write-Error "No Exchange environment detected. Load the Exchange module first."
        return
    }

    Write-Host "Building CAS lookup table..." -ForegroundColor Cyan
    $CASLookup = @{}
    foreach ($cas in $CASMailbox) {
        $CASLookup[$cas.Identity.ToString()] = $cas
    }

    $UserMBXCount = $SharedMBXCount = $RoomMBXCount = $EquipmentMBXCount = $PublicFolderMailboxCount = 0
    $LitHoldCount = $RetentionHoldCount = $ADPDisabledCount = 0
    $POPCount = $IMAPCount = $MAPICount = $ActiveSyncCount = $OWACount = 0

    Write-Host "Processing $($AllMailboxes.Count) mailboxes..." -ForegroundColor Cyan
    $total = $AllMailboxes.Count
    $count = 0

    foreach ($mbx in $AllMailboxes) {
        $count++
        Write-Progress -Activity "Processing Mailboxes" `
            -Status "Mailbox $count of $total ($($mbx.DisplayName))" `
            -PercentComplete ([math]::Round(($count / $total) * 100))

        switch ($mbx.RecipientTypeDetails) {
            "UserMailbox"         { $UserMBXCount++ }
            "SharedMailbox"       { $SharedMBXCount++ }
            "RoomMailbox"         { $RoomMBXCount++ }
            "EquipmentMailbox"    { $EquipmentMBXCount++ }
            "PublicFolderMailbox" { $PublicFolderMailboxCount++ }
        }

        if ($mbx.RecipientTypeDetails -in @("UserMailbox", "SharedMailbox")) {
            if ($mbx.LitigationHoldEnabled)          { $LitHoldCount++ }
            if ($mbx.RetentionHoldEnabled)            { $RetentionHoldCount++ }
            if (-not $mbx.EmailAddressPolicyEnabled)  { $ADPDisabledCount++ }
        }

        $key = $mbx.Identity.ToString()
        if ($CASLookup.ContainsKey($key)) {
            $cas = $CASLookup[$key]
            if ($cas.PopEnabled)        { $POPCount++ }
            if ($cas.ImapEnabled)       { $IMAPCount++ }
            if ($cas.MAPIEnabled)       { $MAPICount++ }
            if ($cas.ActiveSyncEnabled) { $ActiveSyncCount++ }
            if ($cas.OWAEnabled)        { $OWACount++ }
        }
    }
    Write-Progress -Activity "Processing Mailboxes" -Completed

    $Stats = [ordered]@{
        Environment                = $Environment
        TotalMailboxes             = $AllMailboxes.Count
        UserMailboxes              = $UserMBXCount
        SharedMailboxes            = $SharedMBXCount
        RoomMailboxes              = $RoomMBXCount
        EquipmentMailboxes         = $EquipmentMBXCount
        MailUsers                  = (Get-MailUser -ResultSize Unlimited -ErrorAction SilentlyContinue).Count
        MailContacts               = (Get-MailContact -ResultSize Unlimited -ErrorAction SilentlyContinue).Count
        DistributionGroups         = $AllDistGroups.Count
        DynamicDistributionGroups  = (Get-DynamicDistributionGroup -ResultSize Unlimited -ErrorAction SilentlyContinue).Count
        UnifiedGroups              = (Get-UnifiedGroup -ResultSize Unlimited -ErrorAction SilentlyContinue).Count
        LitigationHoldMailboxes    = $LitHoldCount
        RetentionHoldMailboxes     = $RetentionHoldCount
        PublicFolders              = (Get-PublicFolder -Recurse -ErrorAction SilentlyContinue | Measure-Object).Count
        MailPublicFolders          = (Get-MailPublicFolder -ResultSize Unlimited -ErrorAction SilentlyContinue | Measure-Object).Count
        PublicFolderMailboxes      = $PublicFolderMailboxCount
        POPEnabled                 = $POPCount
        IMAPEnabled                = $IMAPCount
        MAPIEnabled                = $MAPICount
        ActiveSyncEnabled          = $ActiveSyncCount
        OWAEnabled                 = $OWACount
        EmailAddressPolicyDisabled = $ADPDisabledCount
    }

    Write-Host "`nExchange Recipient Counts:" -ForegroundColor Cyan
    foreach ($key in $Stats.Keys) {
        $value = $Stats[$key]
        if ($value -gt 0) {
            Write-Host ("{0,-30} : {1}" -f $key, $value) -ForegroundColor White -BackgroundColor DarkGreen
        }
        else {
            Write-Host ("{0,-30} : {1}" -f $key, $value)
        }
    }
}

function Get-FrankensteinVirtualDirectories {
    [CmdletBinding()]
    Param (
        [Switch]$CSV
    )

    Get-Linebreak
    Write-Host "Get-VirtualDirectories" -ForegroundColor Cyan

    $VDirProps = @(
        "server", "name", "internalurl", "externalurl",
        @{Name = "InternalAuthenticationMethods"; Expression = { $_.InternalAuthenticationMethods -join ";" } },
        @{Name = "ExternalAuthenticationMethods"; Expression = { $_.ExternalAuthenticationMethods -join ";" } },
        "IISAuthenticationMethods", "internalhostname", "externalhostname",
        "InternalClientAuthenticationMethod", "ExternalClientAuthenticationMethod"
    )

    $VDirCmdlets = @(
        "Get-AutoDiscoverVirtualDirectory",
        "Get-OwaVirtualDirectory",
        "Get-ECPVirtualDirectory",
        "Get-MAPIVirtualDirectory",
        "Get-ActiveSyncVirtualDirectory",
        "Get-WebServicesVirtualDirectory",
        "Get-OABVirtualDirectory",
        "Get-OutlookAnywhere"
    )

    $DateStamp = (Get-Date).ToString('MMddyy')
    $CsvPath   = ".\VirtualDirectories_$DateStamp.csv"
    $first     = $true

    foreach ($cmdlet in $VDirCmdlets) {
        Write-Host "  Running $cmdlet..." -ForegroundColor Gray
        $results = & $cmdlet -ADPropertiesOnly | Select-Object $VDirProps

        if ($CSV) {
            if ($first) {
                $results | Export-Csv $CsvPath -NoTypeInformation
                $first = $false
            }
            else {
                $results | Export-Csv $CsvPath -NoTypeInformation -Append
            }
        }
        else {
            $results
        }
    }
}

function Get-FrankensteinPublicFolderDiscovery {
    [CmdletBinding()]
    Param()

    $DateStamp = (Get-Date).ToString('MMddyy')

    Get-Linebreak
    Write-Host "Getting Public Folders..." -ForegroundColor Cyan
    Get-PublicFolder -Recurse -ErrorAction SilentlyContinue |
        Select-Object RunspaceId, Identity, Name, MailEnabled, MailRecipientGuid, ParentPath,
            LostAndFoundFolderOriginalPath, ContentMailboxName, ContentMailboxGuid,
            PerUserReadStateEnabled, EntryId, DumpsterEntryId, ParentFolder, OrganizationId,
            AgeLimit, RetainDeletedItemsFor, ProhibitPostQuota, IssueWarningQuota, MaxItemSize,
            LastMovedTime, AdminFolderFlags, FolderSize, HasSubfolders, FolderClass, FolderPath,
            AssociatedDumpsterFolders, DefaultFolderType, ExtendedFolderFlags, MailboxOwnerId,
            IsValid, ObjectState |
        Export-Csv ".\Get_PublicFolder_$DateStamp.csv" -NoTypeInformation

    Get-Linebreak
    Write-Host "Getting Mail Public Folders..." -ForegroundColor Cyan
    Get-MailPublicFolder -ResultSize Unlimited -ErrorAction SilentlyContinue |
        Select-Object RunspaceId, DisplayName, PrimarySmtpAddress,
            @{Name = "EmailAddresses"; Expression = { $_.EmailAddresses -join ";" } },
            Contacts, ContentMailbox, DeliverToMailboxAndForward, ExternalEmailAddress,
            OnPremisesObjectId, IgnoreMissingFolderLink, ForwardingAddress,
            AcceptMessagesOnlyFrom, AcceptMessagesOnlyFromDLMembers,
            AcceptMessagesOnlyFromSendersOrMembers, GrantSendOnBehalfTo,
            AddressListMembership, AdministrativeUnits, Alias, ArbitrationMailbox,
            BypassModerationFromSendersOrMembers, OrganizationalUnit,
            HiddenFromAddressListsEnabled, LastExchangeChangedTime, LegacyExchangeDN,
            MaxSendSize, MaxReceiveSize, ModerationEnabled, ModeratedBy,
            EmailAddressPolicyEnabled, RequireSenderAuthenticationEnabled,
            WindowsEmailAddress, WhenChanged, WhenCreated, ExchangeObjectId, Guid |
        Export-Csv ".\Get_MailPublicFolder_$DateStamp.csv" -NoTypeInformation

    Get-Linebreak
    Write-Host "Getting Public Folder Mailboxes..." -ForegroundColor Cyan
    Get-Mailbox -PublicFolder -ResultSize Unlimited -ErrorAction SilentlyContinue |
        Select-Object RunspaceId, DisplayName, PrimarySmtpAddress, LegacyExchangeDN, Database,
            DeliverToMailboxAndForward, IsHierarchyReady, IsHierarchySyncEnabled,
            LitigationHoldEnabled, SingleItemRecoveryEnabled, RetentionHoldEnabled,
            EndDateForRetentionHold, StartDateForRetentionHold, LitigationHoldDate,
            LitigationHoldOwner, LitigationHoldDuration, ComplianceTagHoldApplied,
            DelayHoldApplied, RetentionPolicy, AddressBookPolicy, ExchangeGuid,
            @{Name = "MailboxLocations"; Expression = { $_.MailboxLocations -join ";" } },
            ExchangeUserAccountControl, AdminDisplayVersion, ForwardingAddress,
            ForwardingSmtpAddress, RetainDeletedItemsFor, IsMailboxEnabled,
            ProhibitSendQuota, ProhibitSendReceiveQuota, RecoverableItemsQuota,
            RecoverableItemsWarningQuota, CalendarLoggingQuota, RecipientLimits,
            ImListMigrationCompleted, IsRootPublicFolderMailbox, LinkedMasterAccount,
            SamAccountName, UserPrincipalName, RoleAssignmentPolicy, SharingPolicy,
            @{Name = "EmailAddresses"; Expression = { $_.EmailAddresses -join ";" } },
            MaxSendSize, MaxReceiveSize, ModerationEnabled, ModeratedBy,
            RecipientTypeDetails, WhenChanged, WhenCreated |
        Export-Csv ".\Get_MailboxPF_$DateStamp.csv" -NoTypeInformation
}

function Get-FrankensteinEntraDiscovery {
    [CmdletBinding()]
    Param (
        [Switch]$CSV,
        [Switch]$UseCurrentSession
    )

    # Verify required Graph sub-modules are installed before doing anything else
    $GraphSubModules = @(
        "Microsoft.Graph.Identity.DirectoryManagement",  # Get-MgOrganization, Get-MgDomain, Get-MgDevice, Get-MgSubscribedSku, Get-MgDirectoryRole
        "Microsoft.Graph.Users",                         # Get-MgUser
        "Microsoft.Graph.Groups",                        # Get-MgGroup
        "Microsoft.Graph.Identity.SignIns",              # Get-MgIdentityConditionalAccessPolicy, security defaults, auth method policy
        "Microsoft.Graph.Applications",                  # Get-MgApplication, Get-MgServicePrincipal
        "Microsoft.Graph.Reports"                        # Get-MgReportAuthenticationMethodUserRegistrationDetail
    )
    $missing = $GraphSubModules | Where-Object { -not (Get-Module -Name $_ -ListAvailable -ErrorAction SilentlyContinue) }
    if ($missing) {
        Write-Error "The following required Graph sub-modules are not installed:`n  $($missing -join "`n  ")`n`nRun: Install-M365Modules -Graph"
        return
    }

    # Import any sub-modules not yet loaded in this session
    foreach ($mod in $GraphSubModules) {
        if (-not (Get-Module -Name $mod -ErrorAction SilentlyContinue)) {
            Write-Host "Importing $mod..." -ForegroundColor Gray
            Import-Module $mod -ErrorAction Stop
        }
    }

    if (-not $UseCurrentSession) {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
        try {
            Connect-MgGraph -Scopes @(
                "Organization.Read.All",
                "Domain.Read.All",
                "User.Read.All",
                "Group.Read.All",
                "Device.Read.All",
                "Policy.Read.All",
                "Application.Read.All",
                "RoleManagement.Read.Directory",
                "UserAuthenticationMethod.Read.All",
                "Reports.Read.All",
                "AuditLog.Read.All",
                "Directory.Read.All"
            ) -ErrorAction Stop
        }
        catch {
            Write-Error "Microsoft Graph authentication failed: $_"
            return
        }
        # Confirm connection was successful
        if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
            Write-Error "Graph connection could not be confirmed. Run Connect-M365 -Graph first."
            return
        }
    }

    $DateStamp = (Get-Date).ToString('MMddyy')
    $OutputDir = ".\Frankenstein_EntraDiscovery_$DateStamp"
    New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
    Push-Location $OutputDir
    Start-Transcript ".\EntraDiscovery_Transcript_$DateStamp.txt"

    # ── Organization ─────────────────────────────────────────────────────────
    Get-Linebreak
    Write-Host "Organization Info" -ForegroundColor Cyan
    $Org = Get-MgOrganization
    $TenantName    = $Org.DisplayName
    $TenantId      = $Org.Id
    $CreatedDate   = $Org.CreatedDateTime
    $TechEmail     = ($Org.TechnicalNotificationMails -join ";")
    $OnPremSync    = $Org.OnPremisesSyncEnabled
    $LastSync      = $Org.OnPremisesLastSyncDateTime

    Write-Host ("{0,-30} : {1}" -f "Tenant Name",  $TenantName)
    Write-Host ("{0,-30} : {1}" -f "Tenant ID",    $TenantId)
    Write-Host ("{0,-30} : {1}" -f "Country",       $Org.CountryLetterCode)
    Write-Host ("{0,-30} : {1}" -f "Created",       $CreatedDate)
    Write-Host ("{0,-30} : {1}" -f "Tech Email",    $TechEmail)
    Write-Host ("{0,-30} : {1}" -f "DirSync",       $(if ($OnPremSync) { "Enabled" } else { "Disabled / Cloud-Only" }))
    Write-Host ("{0,-30} : {1}" -f "Last Sync",     $LastSync)

    if ($CSV) {
        [PSCustomObject]@{
            TenantName                 = $TenantName
            TenantId                   = $TenantId
            Country                    = $Org.CountryLetterCode
            Created                    = $CreatedDate
            TechnicalNotificationEmail = $TechEmail
            OnPremisesSyncEnabled      = $OnPremSync
            LastDirSync                = $LastSync
        } | Export-Csv ".\Organization_$DateStamp.csv" -NoTypeInformation
    }

    # ── Domains ──────────────────────────────────────────────────────────────
    Get-Linebreak
    Write-Host "Domains" -ForegroundColor Cyan
    $Domains = Get-MgDomain -All
    $Domains | Format-Table Id, IsDefault, IsVerified, AuthenticationType -AutoSize
    if ($CSV) {
        $Domains | Select-Object Id, IsDefault, IsInitial, IsVerified, AuthenticationType,
            @{Name="SupportedServices"; Expression={$_.SupportedServices -join ";"}} |
            Export-Csv ".\Domains_$DateStamp.csv" -NoTypeInformation
    }

    # ── Licenses ─────────────────────────────────────────────────────────────
    Get-Linebreak
    Write-Host "License Summary" -ForegroundColor Cyan
    $Skus = Get-MgSubscribedSku -All
    $Skus | ForEach-Object {
        $avail = $_.PrepaidUnits.Enabled - $_.ConsumedUnits
        $color = if ($avail -le 5) { "Yellow" } else { "White" }
        Write-Host ("{0,-45} Assigned: {1,-6} Total: {2,-6} Available: {3}" -f `
            $_.SkuPartNumber, $_.ConsumedUnits, $_.PrepaidUnits.Enabled, $avail) -ForegroundColor $color
    }
    if ($CSV) {
        $Skus | Select-Object SkuPartNumber, SkuId,
            @{Name="TotalLicenses";     Expression={$_.PrepaidUnits.Enabled}},
            @{Name="AssignedLicenses";  Expression={$_.ConsumedUnits}},
            @{Name="AvailableLicenses"; Expression={$_.PrepaidUnits.Enabled - $_.ConsumedUnits}},
            @{Name="SuspendedLicenses"; Expression={$_.PrepaidUnits.Suspended}},
            @{Name="WarningLicenses";   Expression={$_.PrepaidUnits.Warning}},
            CapabilityStatus |
            Export-Csv ".\Licenses_$DateStamp.csv" -NoTypeInformation
    }

    # ── Users ─────────────────────────────────────────────────────────────────
    Get-Linebreak
    Write-Host "Gathering Users..." -ForegroundColor Cyan
    $AllUsers = Get-MgUser -All -Property `
        Id, DisplayName, UserPrincipalName, UserType, AccountEnabled,
        AssignedLicenses, OnPremisesSyncEnabled, CreatedDateTime,
        LastPasswordChangeDateTime, JobTitle, Department, Mail,
        UsageLocation, SignInActivity

    $MemberUsers   = $AllUsers | Where-Object { $_.UserType -eq "Member" }
    $GuestUsers    = $AllUsers | Where-Object { $_.UserType -eq "Guest" }
    $Licensed      = $AllUsers | Where-Object { $_.AssignedLicenses.Count -gt 0 }
    $Unlicensed    = $MemberUsers | Where-Object { $_.AssignedLicenses.Count -eq 0 }
    $SyncedUsers   = $AllUsers | Where-Object { $_.OnPremisesSyncEnabled -eq $true }
    $CloudOnly     = $MemberUsers | Where-Object { $_.OnPremisesSyncEnabled -ne $true }
    $DisabledUsers = $AllUsers | Where-Object { $_.AccountEnabled -eq $false }

    # Flag stale accounts (no sign-in in 90+ days)
    $StaleThreshold = (Get-Date).AddDays(-90)
    $StaleUsers = $MemberUsers | Where-Object {
        $_.SignInActivity.LastSignInDateTime -and
        [datetime]$_.SignInActivity.LastSignInDateTime -lt $StaleThreshold
    }

    Write-Host "`nUser Summary" -ForegroundColor Cyan
    $UserStats = [ordered]@{
        "Total Users"             = $AllUsers.Count
        "Member Users"            = $MemberUsers.Count
        "Guest / External Users"  = $GuestUsers.Count
        "Licensed Users"          = $Licensed.Count
        "Unlicensed Members"      = $Unlicensed.Count
        "Disabled Accounts"       = $DisabledUsers.Count
        "Synced from On-Premises" = $SyncedUsers.Count
        "Cloud-Only Members"      = $CloudOnly.Count
        "Stale (90+ days)"        = $StaleUsers.Count
    }
    foreach ($k in $UserStats.Keys) {
        $v = $UserStats[$k]
        if ($v -gt 0) { Write-Host ("{0,-30} : {1}" -f $k, $v) -ForegroundColor White -BackgroundColor DarkGreen }
        else          { Write-Host ("{0,-30} : {1}" -f $k, $v) }
    }
    if ($CSV) {
        $AllUsers | Select-Object DisplayName, UserPrincipalName, UserType, AccountEnabled,
            @{Name="Licensed";          Expression={$_.AssignedLicenses.Count -gt 0}},
            @{Name="AssignedLicenses";  Expression={$_.AssignedLicenses.SkuId -join ";"}},
            OnPremisesSyncEnabled, JobTitle, Department, Mail, UsageLocation,
            CreatedDateTime, LastPasswordChangeDateTime,
            @{Name="LastSignIn";        Expression={$_.SignInActivity.LastSignInDateTime}} |
            Export-Csv ".\Users_$DateStamp.csv" -NoTypeInformation
    }

    # ── MFA & Authentication Registration ────────────────────────────────────
    Get-Linebreak
    Write-Host "MFA & Authentication Registration..." -ForegroundColor Cyan
    try {
        $AuthReg        = Get-MgReportAuthenticationMethodUserRegistrationDetail -All
        $MfaRegistered  = ($AuthReg | Where-Object { $_.IsMfaRegistered }).Count
        $MfaNotReg      = ($AuthReg | Where-Object { -not $_.IsMfaRegistered }).Count
        $MfaCapable     = ($AuthReg | Where-Object { $_.IsMfaCapable }).Count
        $SsprRegistered = ($AuthReg | Where-Object { $_.IsSsprRegistered }).Count
        $Passwordless   = ($AuthReg | Where-Object { $_.IsPasswordlessCapable }).Count
        $AdminCount     = ($AuthReg | Where-Object { $_.IsAdmin }).Count

        $MfaStats = [ordered]@{
            "MFA Registered"       = $MfaRegistered
            "MFA Not Registered"   = $MfaNotReg
            "MFA Capable"          = $MfaCapable
            "SSPR Registered"      = $SsprRegistered
            "Passwordless Capable" = $Passwordless
            "Admin Accounts"       = $AdminCount
        }
        foreach ($k in $MfaStats.Keys) {
            $v    = $MfaStats[$k]
            $warn = ($k -eq "MFA Not Registered" -and $v -gt 0)
            if ($warn)       { Write-Host ("{0,-30} : {1}" -f $k, $v) -ForegroundColor Yellow }
            elseif ($v -gt 0){ Write-Host ("{0,-30} : {1}" -f $k, $v) -ForegroundColor White -BackgroundColor DarkGreen }
            else              { Write-Host ("{0,-30} : {1}" -f $k, $v) }
        }
        if ($CSV) {
            $AuthReg | Select-Object UserPrincipalName, DisplayName, IsAdmin,
                IsMfaRegistered, IsMfaCapable, IsSsprRegistered, IsSsprCapable,
                IsPasswordlessCapable,
                @{Name="MethodsRegistered"; Expression={$_.MethodsRegistered -join ";"}} |
                Export-Csv ".\MFARegistration_$DateStamp.csv" -NoTypeInformation
        }
    }
    catch {
        Write-Warning "Could not retrieve MFA registration data. Ensure UserAuthenticationMethod.Read.All or Reports.Read.All is consented."
    }

    # ── Admin Role Assignments ────────────────────────────────────────────────
    Get-Linebreak
    Write-Host "Admin Role Assignments..." -ForegroundColor Cyan
    $Roles = Get-MgDirectoryRole -All
    $RoleAssignments = foreach ($role in $Roles) {
        $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction SilentlyContinue
        foreach ($member in $members) {
            [PSCustomObject]@{
                RoleName   = $role.DisplayName
                RoleId     = $role.Id
                MemberName = $member.AdditionalProperties["displayName"]
                MemberUPN  = $member.AdditionalProperties["userPrincipalName"]
                MemberType = $member.OdataType
            }
        }
    }
    $RoleAssignments | Sort-Object RoleName | Format-Table RoleName, MemberName, MemberUPN -AutoSize
    $roleColor = if ($RoleAssignments.Count -gt 25) { "Yellow" } else { "White" }
    Write-Host "Total privileged role assignments: $($RoleAssignments.Count)" -ForegroundColor $roleColor
    if ($CSV) {
        $RoleAssignments | Export-Csv ".\AdminRoleAssignments_$DateStamp.csv" -NoTypeInformation
    }

    # ── Groups ────────────────────────────────────────────────────────────────
    Get-Linebreak
    Write-Host "Gathering Groups..." -ForegroundColor Cyan
    $AllGroups = Get-MgGroup -All -Property `
        Id, DisplayName, GroupTypes, SecurityEnabled, MailEnabled,
        MembershipRule, OnPremisesSyncEnabled, AssignedLicenses, Visibility

    $SecurityGroups      = $AllGroups | Where-Object { $_.SecurityEnabled -and -not $_.MailEnabled -and $_.GroupTypes -notcontains "Unified" }
    $M365Groups          = $AllGroups | Where-Object { $_.GroupTypes -contains "Unified" }
    $MailEnabledSecurity = $AllGroups | Where-Object { $_.SecurityEnabled -and $_.MailEnabled -and $_.GroupTypes -notcontains "Unified" }
    $DynamicGroups       = $AllGroups | Where-Object { $_.GroupTypes -contains "DynamicMembership" }
    $SyncedGroups        = $AllGroups | Where-Object { $_.OnPremisesSyncEnabled -eq $true }
    $LicensedGroups      = $AllGroups | Where-Object { $_.AssignedLicenses.Count -gt 0 }
    $PublicM365          = $M365Groups | Where-Object { $_.Visibility -eq "Public" }

    $GroupStats = [ordered]@{
        "Total Groups"            = $AllGroups.Count
        "Security Groups"         = $SecurityGroups.Count
        "Microsoft 365 Groups"    = $M365Groups.Count
        "  Public M365 Groups"    = $PublicM365.Count
        "Mail-Enabled Security"   = $MailEnabledSecurity.Count
        "Dynamic Groups"          = $DynamicGroups.Count
        "Synced from On-Premises" = $SyncedGroups.Count
        "License-Assigned Groups" = $LicensedGroups.Count
    }
    foreach ($k in $GroupStats.Keys) {
        $v = $GroupStats[$k]
        if ($v -gt 0) { Write-Host ("{0,-30} : {1}" -f $k, $v) -ForegroundColor White -BackgroundColor DarkGreen }
        else          { Write-Host ("{0,-30} : {1}" -f $k, $v) }
    }
    if ($CSV) {
        $AllGroups | Select-Object DisplayName, Visibility,
            @{Name="GroupTypes";        Expression={$_.GroupTypes -join ";"}},
            SecurityEnabled, MailEnabled, OnPremisesSyncEnabled,
            @{Name="IsDynamic";         Expression={$_.GroupTypes -contains "DynamicMembership"}},
            MembershipRule,
            @{Name="AssignedLicenses";  Expression={$_.AssignedLicenses.SkuId -join ";"}} |
            Export-Csv ".\Groups_$DateStamp.csv" -NoTypeInformation
    }

    # ── Devices ───────────────────────────────────────────────────────────────
    Get-Linebreak
    Write-Host "Gathering Devices..." -ForegroundColor Cyan
    $Devices = Get-MgDevice -All -Property `
        Id, DisplayName, OperatingSystem, OperatingSystemVersion,
        TrustType, IsCompliant, IsManaged, AccountEnabled,
        RegisteredDateTime, ApproximateLastSignInDateTime

    $EntraJoined  = $Devices | Where-Object { $_.TrustType -eq "AzureAd" }
    $HybridJoined = $Devices | Where-Object { $_.TrustType -eq "ServerAd" }
    $Registered   = $Devices | Where-Object { $_.TrustType -eq "Workplace" }
    $Compliant    = $Devices | Where-Object { $_.IsCompliant -eq $true }
    $NonCompliant = $Devices | Where-Object { $_.IsCompliant -eq $false }
    $Managed      = $Devices | Where-Object { $_.IsManaged -eq $true }
    $EnabledDevs  = $Devices | Where-Object { $_.AccountEnabled -eq $true }

    $DeviceStats = [ordered]@{
        "Total Devices"       = $Devices.Count
        "Entra Joined"        = $EntraJoined.Count
        "Hybrid Joined"       = $HybridJoined.Count
        "Registered (BYOD)"   = $Registered.Count
        "Compliant"           = $Compliant.Count
        "Non-Compliant"       = $NonCompliant.Count
        "Managed (Intune)"    = $Managed.Count
        "Enabled"             = $EnabledDevs.Count
    }
    foreach ($k in $DeviceStats.Keys) {
        $v    = $DeviceStats[$k]
        $warn = ($k -eq "Non-Compliant" -and $v -gt 0)
        if ($warn)        { Write-Host ("{0,-30} : {1}" -f $k, $v) -ForegroundColor Yellow }
        elseif ($v -gt 0) { Write-Host ("{0,-30} : {1}" -f $k, $v) -ForegroundColor White -BackgroundColor DarkGreen }
        else              { Write-Host ("{0,-30} : {1}" -f $k, $v) }
    }
    Write-Host "`nOS Breakdown:" -ForegroundColor Cyan
    $Devices | Group-Object OperatingSystem | Sort-Object Count -Descending |
        ForEach-Object { Write-Host ("  {0,-25} : {1}" -f $_.Name, $_.Count) }

    if ($CSV) {
        $Devices | Select-Object DisplayName, OperatingSystem, OperatingSystemVersion,
            TrustType, IsCompliant, IsManaged, AccountEnabled,
            RegisteredDateTime, ApproximateLastSignInDateTime |
            Export-Csv ".\Devices_$DateStamp.csv" -NoTypeInformation
    }

    # ── Conditional Access ────────────────────────────────────────────────────
    Get-Linebreak
    Write-Host "Conditional Access Policies..." -ForegroundColor Cyan
    $CAPolicies   = Get-MgIdentityConditionalAccessPolicy -All
    $CAEnabled    = $CAPolicies | Where-Object { $_.State -eq "enabled" }
    $CAReportOnly = $CAPolicies | Where-Object { $_.State -eq "enabledForReportingButNotEnforced" }
    $CADisabled   = $CAPolicies | Where-Object { $_.State -eq "disabled" }

    Write-Host ("{0,-30} : {1}" -f "Total CA Policies",  $CAPolicies.Count)
    Write-Host ("{0,-30} : {1}" -f "  Enabled",          $CAEnabled.Count)
    Write-Host ("{0,-30} : {1}" -f "  Report-Only",      $CAReportOnly.Count)
    Write-Host ("{0,-30} : {1}" -f "  Disabled",         $CADisabled.Count)
    $CAPolicies | Sort-Object State, DisplayName | Format-Table DisplayName, State -AutoSize

    try {
        $NamedLocations = Get-MgIdentityConditionalAccessNamedLocation -All
        Write-Host ("{0,-30} : {1}" -f "Named Locations", $NamedLocations.Count)
    } catch {}

    if ($CSV) {
        $CAPolicies | Select-Object DisplayName, State, Id,
            @{Name="IncludeUsers";   Expression={$_.Conditions.Users.IncludeUsers -join ";"}},
            @{Name="ExcludeUsers";   Expression={$_.Conditions.Users.ExcludeUsers -join ";"}},
            @{Name="IncludeGroups";  Expression={$_.Conditions.Users.IncludeGroups -join ";"}},
            @{Name="IncludeApps";    Expression={$_.Conditions.Applications.IncludeApplications -join ";"}},
            @{Name="ExcludeApps";    Expression={$_.Conditions.Applications.ExcludeApplications -join ";"}},
            @{Name="Platforms";      Expression={$_.Conditions.Platforms.IncludePlatforms -join ";"}},
            @{Name="GrantControls";  Expression={$_.GrantControls.BuiltInControls -join ";"}} |
            Export-Csv ".\ConditionalAccessPolicies_$DateStamp.csv" -NoTypeInformation
    }

    # ── Applications ─────────────────────────────────────────────────────────
    Get-Linebreak
    Write-Host "Applications..." -ForegroundColor Cyan
    $AppRegs = Get-MgApplication -All -Property Id, DisplayName, CreatedDateTime, SignInAudience, PublisherDomain
    $EntApps = Get-MgServicePrincipal -All -Property Id, DisplayName, ServicePrincipalType, AccountEnabled, AppId, Tags |
        Where-Object { $_.ServicePrincipalType -eq "Application" }
    $EntAppsEnabled  = $EntApps | Where-Object { $_.AccountEnabled }
    $EntAppsDisabled = $EntApps | Where-Object { -not $_.AccountEnabled }

    Write-Host ("{0,-30} : {1}" -f "App Registrations",    $AppRegs.Count)
    Write-Host ("{0,-30} : {1}" -f "Enterprise Apps",      $EntApps.Count)
    Write-Host ("{0,-30} : {1}" -f "  Enabled",            $EntAppsEnabled.Count)
    Write-Host ("{0,-30} : {1}" -f "  Disabled",           $EntAppsDisabled.Count)

    if ($CSV) {
        $AppRegs | Select-Object DisplayName, CreatedDateTime, SignInAudience, PublisherDomain, Id |
            Export-Csv ".\AppRegistrations_$DateStamp.csv" -NoTypeInformation
        $EntApps | Select-Object DisplayName, ServicePrincipalType, AccountEnabled, AppId, Id |
            Export-Csv ".\EnterpriseApps_$DateStamp.csv" -NoTypeInformation
    }

    # ── Security Posture ─────────────────────────────────────────────────────
    Get-Linebreak
    Write-Host "Security Posture..." -ForegroundColor Cyan

    try {
        $SecDefaults = Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy
        $secColor    = if ($SecDefaults.IsEnabled) { "Green" } else { "Yellow" }
        Write-Host ("{0,-30} : {1}" -f "Security Defaults", $(if ($SecDefaults.IsEnabled) { "ENABLED" } else { "Disabled" })) -ForegroundColor $secColor
    }
    catch { Write-Warning "Could not retrieve Security Defaults policy." }

    try {
        $AuthMethodPolicy = Get-MgPolicyAuthenticationMethodPolicy
        Write-Host ("{0,-30} : {1}" -f "Auth Method Policy", $AuthMethodPolicy.DisplayName)
    }
    catch { Write-Warning "Could not retrieve Authentication Method Policy." }

    try {
        $PasswordPolicy = Get-MgDomain | Where-Object { $_.IsDefault }
        Write-Host ("{0,-30} : {1}" -f "Default Domain", $PasswordPolicy.Id)
    }
    catch {}

    # ── Discovery Summary ─────────────────────────────────────────────────────
    Get-Linebreak
    Write-Host "ENTRA ID DISCOVERY SUMMARY" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Tenant  : $TenantName  ($TenantId)"
    Write-Host "  Users   : $($AllUsers.Count) total  |  $($Licensed.Count) licensed  |  $($GuestUsers.Count) guests  |  $($DisabledUsers.Count) disabled"
    Write-Host "  Groups  : $($AllGroups.Count) total  |  $($SecurityGroups.Count) security  |  $($M365Groups.Count) M365  |  $($DynamicGroups.Count) dynamic"
    Write-Host "  Devices : $($Devices.Count) total  |  $($EntraJoined.Count) Entra joined  |  $($HybridJoined.Count) hybrid  |  $($Compliant.Count) compliant"
    Write-Host "  CA      : $($CAPolicies.Count) policies  ($($CAEnabled.Count) enabled)"
    Write-Host "  Apps    : $($AppRegs.Count) registrations  |  $($EntApps.Count) enterprise apps"
    Write-Host "  DirSync : $(if($OnPremSync){'Enabled — Last sync: ' + $LastSync}else{'Cloud-Only'})"
    Write-Host ""
    Write-Host "Output saved to: $OutputDir" -ForegroundColor Gray

    Stop-Transcript
    Pop-Location
}

function Get-FrankensteinExchangeDiscovery {
    [CmdletBinding()]
    Param (
        [Switch]$VirtualDirectories,
        [Switch]$CSV,
        [Switch]$UseCurrentSession,
        [Switch]$Online,
        [Switch]$PublicFolders
    )

    if ($UseCurrentSession) {
        # Use whatever session is already active
    }
    elseif ($Online) {
        Connect-ExchangeOnline
    }
    else {
        Connect-ExchangeOnPremServer
    }

    $DateStamp = (Get-Date).ToString('MMddyy')
    if ($Online) {
        $OutputDir      = ".\Frankenstein_ExchangeOnline_Discovery_$DateStamp"
        $TranscriptName = "ExchangeOnline_DiscoveryTranscript_$DateStamp.txt"
    }
    else {
        $OutputDir      = ".\Frankenstein_ExchangeOnPrem_Discovery_$DateStamp"
        $TranscriptName = "ExchangeOnPrem_DiscoveryTranscript_$DateStamp.txt"
    }

    New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
    Push-Location $OutputDir
    Start-Transcript -Path ".\$TranscriptName"

    Get-Linebreak
    Get-FrankensteinRecipientCounts

    if (-not $Online) {
        Get-Linebreak
        Write-Host "Get-ExchangeServer" -ForegroundColor Cyan
        $ExchangeServers = Get-ExchangeServer
        $ExchangeServers | Format-List
        if ($CSV) {
            $ExchangeServers | Select-Object Name, Domain, Edition, FQDN,
                IsHubTransportServer, IsClientAccessServer, IsEdgeServer, IsMailboxServer,
                IsUnifiedMessagingServer, IsFrontendTransportServer,
                OrganizationalUnit, AdminDisplayVersion, Site, ServerRole |
                Export-Csv ".\ExchangeServers_$DateStamp.csv" -NoTypeInformation
        }

        Get-Linebreak
        Write-Host "Get-MailboxDatabase" -ForegroundColor Cyan
        Get-MailboxDatabase | Format-List
        if ($CSV) {
            Get-MailboxDatabase | Select-Object Name, Server, MailboxRetention,
                ProhibitSendReceiveQuota, ProhibitSendQuota, RecoverableItemsQuota,
                RecoverableItemsWarningQuota, IsExcludedFromProvisioning, ReplicationType,
                DeletedItemRetention, CircularLoggingEnabled, AdminDisplayVersion |
                Export-Csv ".\Databases_$DateStamp.csv" -NoTypeInformation
        }

        Get-Linebreak
        Write-Host "Get-DatabaseAvailabilityGroup" -ForegroundColor Cyan
        Get-DatabaseAvailabilityGroup | Format-List
        if ($CSV) {
            Get-DatabaseAvailabilityGroup | Export-Csv ".\DAG_$DateStamp.csv" -NoTypeInformation
        }
    }

    Get-Linebreak
    Write-Host "Get-RetentionPolicy" -ForegroundColor Cyan
    Get-RetentionPolicy | Format-List
    if ($CSV) {
        Get-RetentionPolicy | Select-Object Name,
            @{Name="RetentionPolicyTagLinks"; Expression={$_.RetentionPolicyTagLinks -join ";"}} |
            Export-Csv ".\RetentionPolicies_$DateStamp.csv" -NoTypeInformation
    }

    Get-Linebreak
    Write-Host "Get-RetentionPolicyTag" -ForegroundColor Cyan
    Get-RetentionPolicyTag | Format-List
    if ($CSV) {
        Get-RetentionPolicyTag | Select-Object Name, Type, AgeLimitForRetention, RetentionAction |
            Export-Csv ".\RetentionPoliciesTag_$DateStamp.csv" -NoTypeInformation
    }

    Get-Linebreak
    Write-Host "Get-JournalRule" -ForegroundColor Cyan
    Get-JournalRule | Format-List
    if ($CSV) {
        Get-JournalRule | Select-Object Name, Recipient, JournalEmailAddress, Scope, Enabled |
            Export-Csv ".\JournalRules_$DateStamp.csv" -NoTypeInformation
    }

    Get-Linebreak
    Write-Host "Get-AcceptedDomain" -ForegroundColor Cyan
    $AcceptedDomain = Get-AcceptedDomain
    $AcceptedDomain | Format-List
    if ($CSV) {
        $AcceptedDomain | Select-Object Name, DomainName, DomainType, Default |
            Export-Csv ".\AcceptedDomains_$DateStamp.csv" -NoTypeInformation
    }
    Write-Host "Domain MX Records" -ForegroundColor Cyan
    foreach ($domain in $AcceptedDomain) { Resolve-DnsName -Name $domain.DomainName -Type MX -ErrorAction SilentlyContinue }
    Write-Host "Domain TXT Records" -ForegroundColor Cyan
    foreach ($domain in $AcceptedDomain) { Resolve-DnsName -Name $domain.DomainName -Type TXT -ErrorAction SilentlyContinue }
    Write-Host "Domain CNAME Records" -ForegroundColor Cyan
    foreach ($domain in $AcceptedDomain) { Resolve-DnsName -Name $domain.DomainName -Type CNAME -ErrorAction SilentlyContinue }

    Get-Linebreak
    Write-Host "Get-RemoteDomain" -ForegroundColor Cyan
    Get-RemoteDomain | Format-List
    if ($CSV) {
        Get-RemoteDomain | Select-Object Name, DomainName, AllowedOOFType |
            Export-Csv ".\RemoteDomains_$DateStamp.csv" -NoTypeInformation
    }

    Get-Linebreak
    Write-Host "Get-EmailAddressPolicy" -ForegroundColor Cyan
    Get-EmailAddressPolicy | Format-List
    if ($CSV) {
        Get-EmailAddressPolicy | Select-Object Name, Priority, IncludedRecipients,
            @{Name="EnabledEmailAddressTemplates"; Expression={$_.EnabledEmailAddressTemplates -join ";"}},
            RecipientFilterApplied |
            Export-Csv ".\EmailAddressPolicies_$DateStamp.csv" -NoTypeInformation
    }

    Get-Linebreak
    Write-Host "Get-TransportRule" -ForegroundColor Cyan
    Get-TransportRule | Format-List
    if ($CSV) {
        Get-TransportRule | Select-Object Name, Description, State, Priority |
            Export-Csv ".\TransportRules_$DateStamp.csv" -NoTypeInformation
        $file = Export-TransportRuleCollection
        Set-Content -Path ".\Rules.xml" -Value $file.FileData -Encoding Byte
    }

    Get-Linebreak
    if ($Online) {
        Write-Host "Get-OutboundConnector" -ForegroundColor Cyan
        Get-OutboundConnector | Format-List
        if ($CSV) {
            Get-OutboundConnector | Select-Object Name,
                @{Name="SmartHosts";            Expression={$_.SmartHosts -join ";"}},
                Enabled,
                @{Name="AddressSpaces";         Expression={$_.AddressSpaces -join ";"}},
                @{Name="SourceTransportServers";Expression={$_.SourceTransportServers -join ";"}},
                FQDN, MaxMessageSize, ProtocolLoggingLevel, RequireTLS |
                Export-Csv ".\OutboundConnectors_$DateStamp.csv" -NoTypeInformation
        }

        Get-Linebreak
        Write-Host "Get-InboundConnector" -ForegroundColor Cyan
        Get-InboundConnector | Format-List
        if ($CSV) {
            Get-InboundConnector | Select-Object Name, AuthMechanism,
                @{Name="Bindings";        Expression={$_.Bindings -join ";"}},
                Enabled,
                @{Name="RemoteIPRanges"; Expression={$_.RemoteIPRanges -join ";"}},
                RequireTLS, OriginatingServer |
                Export-Csv ".\InboundConnectors_$DateStamp.csv" -NoTypeInformation
        }
    }
    else {
        Write-Host "Get-SendConnector" -ForegroundColor Cyan
        Get-SendConnector | Format-List
        if ($CSV) {
            Get-SendConnector | Select-Object Name,
                @{Name="SmartHosts";            Expression={$_.SmartHosts -join ";"}},
                Enabled,
                @{Name="AddressSpaces";         Expression={$_.AddressSpaces -join ";"}},
                @{Name="SourceTransportServers";Expression={$_.SourceTransportServers -join ";"}},
                FQDN, MaxMessageSize, ProtocolLoggingLevel, RequireTLS |
                Export-Csv ".\SendConnectors_$DateStamp.csv" -NoTypeInformation
        }

        Get-Linebreak
        Write-Host "Get-ReceiveConnector" -ForegroundColor Cyan
        Get-ReceiveConnector | Format-List
        if ($CSV) {
            Get-ReceiveConnector | Select-Object Name, AuthMechanism,
                @{Name="Bindings";       Expression={$_.Bindings -join ";"}},
                Enabled,
                @{Name="RemoteIPRanges";Expression={$_.RemoteIPRanges -join ";"}},
                RequireTLS, OriginatingServer |
                Export-Csv ".\ReceiveConnectors_$DateStamp.csv" -NoTypeInformation
        }

        Get-Linebreak
        Write-Host "Get-TransportAgent" -ForegroundColor Cyan
        Get-TransportAgent | Format-List

        Get-Linebreak
        Write-Host "Get-AddressList / Get-AddressBookPolicy" -ForegroundColor Cyan
        Get-AddressList
        Get-AddressBookPolicy
    }

    Get-Linebreak
    Write-Host "Get-OrganizationConfig" -ForegroundColor Cyan
    Get-OrganizationConfig | Format-List

    Get-Linebreak
    Write-Host "Get-FederationTrust" -ForegroundColor Cyan
    Get-FederationTrust | Format-List

    Get-Linebreak
    Write-Host "Get-OrganizationRelationship" -ForegroundColor Cyan
    Get-OrganizationRelationship | Format-List
    if ($CSV) {
        Get-OrganizationRelationship | Select-Object Name,
            @{Name="DomainNames"; Expression={$_.DomainNames -join ";"}},
            TargetAutoDiscoverEpr, TargetOWAUrl, TargetSharingEpr,
            TargetApplicationUri, Enabled |
            Export-Csv ".\OrganizationRelationships_$DateStamp.csv" -NoTypeInformation
    }

    Get-Linebreak
    Write-Host "Get-IntraOrganizationConnector / Get-IntraOrganizationConfiguration" -ForegroundColor Cyan
    Get-IntraOrganizationConnector | Format-List
    Get-IntraOrganizationConfiguration

    if (-not $Online) {
        Get-Linebreak
        Write-Host "Get-ExchangeCertificate" -ForegroundColor Cyan
        Get-ExchangeCertificate | Format-List
        if ($CSV) {
            Get-ExchangeCertificate | Select-Object Subject, Issuer, Thumbprint, FriendlyName, NotAfter |
                Export-Csv ".\ExchangeCertificates_$DateStamp.csv" -NoTypeInformation
        }

        Get-Linebreak
        Write-Host "Get-HybridConfiguration" -ForegroundColor Cyan
        $Hybrid = Get-HybridConfiguration -ErrorAction SilentlyContinue
        if ($null -ne $Hybrid) {
            Write-Host "Hybrid configuration detected:" -ForegroundColor Yellow
            $Hybrid | Format-List
        }
        else {
            Write-Host "No hybrid configuration detected." -ForegroundColor Gray
        }
    }

    Get-Linebreak

    if ($VirtualDirectories) {
        Get-FrankensteinVirtualDirectories -CSV:$CSV
    }

    if ($PublicFolders) {
        Get-FrankensteinPublicFolderDiscovery
    }

    Stop-Transcript
    Pop-Location
}

function Get-FrankensteinGSuiteDiscovery {
    [CmdletBinding()]
    Param (
        [Switch]$CSV,
        [Switch]$IncludeGroupSettings,
        [Switch]$IncludeGroupMembership,
        [Switch]$IncludeDelegates,
        [Switch]$IncludeSendAsSettings,
        [Switch]$IncludeAutoForwardSettings
    )

    $DateStamp = (Get-Date).ToString('MMddyy')
    $OutputDir = ".\GSuiteDiscovery_$DateStamp"
    New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
    Push-Location $OutputDir
    Start-Transcript ".\GSuiteDiscoveryTranscript_$DateStamp.txt"

    Get-Linebreak
    Write-Host "Building Variables..." -ForegroundColor Cyan
    $GSUser                   = Get-GSUser -Filter *
    $GSGroup                  = Get-GSGroup
    $GSDomain                 = Get-GSDomain
    $GSResource               = Get-GSResource -Filter *
    $GSOrganizationalUnitList = Get-GSOrganizationalUnitList
    $GSUserLicenseInfo        = Get-GSUserLicenseInfo

    Write-Host "$($GSUser.Count) Total Users"
    Write-Host "$($GSGroup.Count) Total Groups"
    Write-Host "$($GSDomain.Count) Total Domains"
    Write-Host "$($GSResource.Count) Total Resources"
    Write-Host "$($GSOrganizationalUnitList.Count) Total Org Units"
    Write-Host "$($GSUserLicenseInfo.Count) Licenses applied across $($GSUser.Count) users"

    Get-Linebreak
    if ($CSV) {
        Write-Host "Creating GSUser Report..." -ForegroundColor Cyan
        $GSUser | Select-Object User, PrimaryEmail, AgreedToTerms,
            @{Name="Aliases";             Expression={$_.Aliases -join ";"}},
            Archived, ChangePasswordAtNextLogin, CreationTime, DeletionTime, Id,
            IncludeInGlobalAddressList, IpWhitelisted, IsAdmin, IsDelegate, IsEnforced,
            IsEnrolledIn2Sv, IsMailboxSetup, LastLoginTime,
            @{Name="NonEditableAliases";  Expression={$_.NonEditableAliases -join ";"}},
            OrgUnitPath,
            @{Name="Organizations";       Expression={$_.Organizations -join ";"}},
            @{Name="Phones";              Expression={$_.Phones -join ";"}},
            RecoveryEmail, Suspended, SuspensionReason |
            Export-Csv ".\GSUsers_$DateStamp.csv" -NoTypeInformation

        $GSUser | Get-GSUserAlias |
            Select-Object AliasValue, PrimaryEmail |
            Export-Csv ".\GSUserAlias_$DateStamp.csv" -NoTypeInformation
    }

    if ($IncludeDelegates) {
        Get-Linebreak
        Write-Host "Processing GSUser Delegates..." -ForegroundColor Cyan
        $WarningPreference = "SilentlyContinue"
        $DelegationList = foreach ($User in $GSUser) {
            $Delegates = Get-GSGmailDelegate -User $User.PrimaryEmail -ErrorAction SilentlyContinue
            if ($Delegates) {
                $Delegates | ForEach-Object {
                    [PSCustomObject]@{
                        User               = $User.PrimaryEmail
                        DelegateEmail      = $_.DelegateEmail
                        VerificationStatus = $_.VerificationStatus
                    }
                }
            }
        }
        $DelegationList | Export-Csv ".\GSDelegates_$DateStamp.csv" -NoTypeInformation
        $WarningPreference = "Continue"
    }

    if ($IncludeSendAsSettings) {
        Get-Linebreak
        Write-Host "Processing GSUser Send As Settings..." -ForegroundColor Cyan
        $SendAsSettings = foreach ($User in $GSUser) {
            $SendAs = Get-GSGmailSendAsSettings -User $User.PrimaryEmail
            if ($SendAs) {
                $SendAs | ForEach-Object {
                    [PSCustomObject]@{
                        User        = $User.PrimaryEmail
                        SendAsEmail = $_.SendAsEmail
                        IsDefault   = $_.IsDefault
                        IsPrimary   = $_.IsPrimary
                    }
                }
            }
        }
        $SendAsSettings | Export-Csv ".\GSSendAsSettings_$DateStamp.csv" -NoTypeInformation
    }

    if ($IncludeAutoForwardSettings) {
        Get-Linebreak
        Write-Host "Collecting Auto Forward Settings..." -ForegroundColor Cyan
        $GSUser | Get-GSGmailAutoForwardingSettings |
            Where-Object { $_.Enabled -eq $true } |
            Select-Object User, Disposition, EmailAddress, Enabled |
            Export-Csv ".\GSAutoForwardSettings_$DateStamp.csv" -NoTypeInformation
    }

    if ($IncludeGroupSettings) {
        Get-Linebreak
        Write-Host "Collecting Group Settings..." -ForegroundColor Cyan
        $GSGroup | Get-GSGroupSettings |
            Export-Csv ".\GSGroupSettings_$DateStamp.csv" -NoTypeInformation
    }

    if ($IncludeGroupMembership) {
        Get-Linebreak
        Write-Host "Collecting Group Membership..." -ForegroundColor Cyan
        $GSGroup | Get-GSGroupMember |
            Export-Csv ".\GSGroupMembers_$DateStamp.csv" -NoTypeInformation
    }

    if ($CSV) {
        Get-Linebreak
        Write-Host "Collecting Org Units..." -ForegroundColor Cyan
        $GSOrganizationalUnitList | Select-Object BlockInheritance, Description, Name,
            OrgUnitId, OrgUnitPath, ParentOrgUnitId, ParentOrgUnitPath |
            Export-Csv ".\GSOrganizationalUnitList_$DateStamp.csv" -NoTypeInformation

        Get-Linebreak
        Write-Host "Collecting User License Information..." -ForegroundColor Cyan
        $GSUserLicenseInfo | Select-Object UserId, ProductId, ProductName, SkuId, SkuName |
            Export-Csv ".\GSUserLicenseInfo_$DateStamp.csv" -NoTypeInformation
    }

    Stop-Transcript
    Pop-Location
}

function Get-FrankensteinMailboxPermissions {
    [CmdletBinding()]
    Param (
        [Switch]$FullAccess,
        [Switch]$SendAs,
        [Switch]$SendOnBehalf,
        [Switch]$UseCurrentSession,
        [Switch]$CSV,
        [Switch]$Help
    )

    if ($Help) {
        Write-Host @"
SYNOPSIS
    Retrieves Full Access, SendAs, and SendOnBehalf permissions.

DESCRIPTION
    Requires minimum of Exchange Reader. Global Reader will not work.

PARAMETERS
    -UseCurrentSession  Use the current Exchange session instead of prompting to connect.
    -FullAccess         Scope to FullAccess permissions only.
    -SendAs             Scope to SendAs permissions only.
    -SendOnBehalf       Scope to SendOnBehalf permissions only.
    -CSV                Export results to CSV.

EXAMPLE
    Get-FrankensteinMailboxPermissions -UseCurrentSession -FullAccess -SendAs -SendOnBehalf

NOTES
    Author: Eric D. Frank
    09/26/25 - Added UserWithAccess/Mailbox ExchangeGUID and caching for recipients.
"@
        return
    }

    if (-not $UseCurrentSession) {
        Connect-ExchangeOnline
    }

    $Mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox
    $total     = $Mailboxes.Count
    $count     = 0
    $Results   = [System.Collections.Generic.List[PSCustomObject]]::new()
    $RecipientCache = @{}

    Write-Host "Gathering permissions for $total mailboxes..." -ForegroundColor Cyan

    function Resolve-Recipient ([string]$identity) {
        if (-not $RecipientCache.ContainsKey($identity)) {
            $RecipientCache[$identity] = Get-Recipient -Identity $identity -ErrorAction SilentlyContinue
        }
        return $RecipientCache[$identity]
    }

    foreach ($mbx in $Mailboxes) {
        $count++
        Write-Progress -Activity "Gathering Permissions" `
            -Status "Processing $($mbx.DisplayName) ($count of $total)" `
            -PercentComplete ([math]::Round(($count / $total) * 100))

        $MailboxExchangeGuid = $mbx.ExchangeGuid
        $MailboxSMTP         = $mbx.PrimarySmtpAddress
        $MailboxType         = $mbx.RecipientTypeDetails

        if ($FullAccess) {
            Get-MailboxPermission -Identity $mbx.Identity -ErrorAction SilentlyContinue |
                Where-Object { -not $_.IsInherited -and $_.User -notlike "NT AUTHORITY\SELF" } |
                ForEach-Object {
                    $r = Resolve-Recipient $_.User
                    $Results.Add([PSCustomObject]@{
                        DisplayName                = $mbx.DisplayName
                        UserPrincipalName          = $MailboxSMTP
                        MailboxType                = $MailboxType
                        MailboxExchangeGuid        = $MailboxExchangeGuid
                        AccessType                 = "FullAccess"
                        UserWithAccess             = if ($r) { $r.PrimarySmtpAddress }    else { $_.User }
                        UserWithAccessType         = if ($r) { $r.RecipientTypeDetails }  else { "Unknown/External" }
                        UserWithAccessExchangeGuid = if ($r) { $r.ExchangeGuid }          else { $null }
                    })
                }
        }

        if ($SendAs) {
            Get-RecipientPermission -Identity $mbx.Identity -ErrorAction SilentlyContinue |
                Where-Object { $_.Trustee -ne "NT AUTHORITY\SELF" } |
                ForEach-Object {
                    $r = Resolve-Recipient $_.Trustee
                    $Results.Add([PSCustomObject]@{
                        DisplayName                = $mbx.DisplayName
                        UserPrincipalName          = $MailboxSMTP
                        MailboxType                = $MailboxType
                        MailboxExchangeGuid        = $MailboxExchangeGuid
                        AccessType                 = "SendAs"
                        UserWithAccess             = if ($r) { $r.PrimarySmtpAddress }    else { $_.Trustee }
                        UserWithAccessType         = if ($r) { $r.RecipientTypeDetails }  else { "Unknown/External" }
                        UserWithAccessExchangeGuid = if ($r) { $r.ExchangeGuid }          else { $null }
                    })
                }
        }

        if ($SendOnBehalf) {
            foreach ($delegate in $mbx.GrantSendOnBehalfTo) {
                $r = Resolve-Recipient $delegate
                $Results.Add([PSCustomObject]@{
                    DisplayName                = $mbx.DisplayName
                    UserPrincipalName          = $MailboxSMTP
                    MailboxType                = $MailboxType
                    MailboxExchangeGuid        = $MailboxExchangeGuid
                    AccessType                 = "SendOnBehalf"
                    UserWithAccess             = if ($r) { $r.PrimarySmtpAddress }    else { $delegate }
                    UserWithAccessType         = if ($r) { $r.RecipientTypeDetails }  else { "Unknown/External" }
                    UserWithAccessExchangeGuid = if ($r) { $r.ExchangeGuid }          else { $null }
                })
            }
        }
    }
    Write-Progress -Activity "Gathering Permissions" -Completed

    if ($CSV) {
        $FileName = ".\MailboxPermissions_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
        $Results | Export-Csv $FileName -NoTypeInformation -Encoding UTF8
        Write-Host "Export complete: $FileName" -ForegroundColor Green
    }
    else {
        $Results
    }
}

function Set-FrankensteinPSWindowTitle {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [String]$Title
    )
    $host.UI.RawUI.WindowTitle = $Title
}

#endregion
