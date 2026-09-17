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
       Retrieves FullAccess, SendAs, and SendOnBehalf permissions. DL delegates are automatically
       expanded to individual members (including nested DLs). Expanded rows include an
       ExpandedFromGroup column showing which DL the member came from.
       Switches: [-FullAccess] [-SendAs] [-SendOnBehalf] [-UseCurrentSession] [-CSV] [-Help]

    9) Get-FrankensteinVirtualDirectories
       Reports on Exchange virtual directory URLs and authentication methods.
       Switches: [-CSV]

    10) Get-FrankensteinEntraDiscovery
        Comprehensive Entra ID (Azure AD) discovery via Microsoft Graph. Covers org info, users,
        MFA registration, admin roles, groups, devices, Conditional Access, apps, and security posture.
        Switches: [-CSV] [-UseCurrentSession]

    11) Import-FrankensteinMailboxPermissions
        Imports FullAccess, SendAs, and SendOnBehalf permissions using a permissions export CSV and a
        migration mapping CSV. Both the mailbox and delegate must exist in the mapping to apply a permission.
        Outputs a timestamped log CSV showing successes, skips, and failures.
        Switches: [-FullAccess] [-SendAs] [-SendOnBehalf] [-UseCurrentSession] [-Help]

    12) Get-FrankensteinGroups
        Exports properties for all mail-enabled group types: Distribution Groups, Mail-Enabled Security Groups,
        Dynamic Distribution Groups, and M365 (Unified) Groups. A GroupType column identifies each record.
        Switches: [-DistributionGroups] [-MailEnabledSecurityGroups] [-DynamicDistributionGroups] [-M365Groups]
                  [-UseCurrentSession] [-CSV] [-Help]

    13) Get-FrankensteinGroupMember
        Exports group membership for all mail-enabled group types. DDG members are resolved live from the
        recipient filter and tagged as Dynamic. M365 groups include both Members and Owners with a Role column.
        Switches: [-DistributionGroups] [-MailEnabledSecurityGroups] [-DynamicDistributionGroups] [-M365Groups]
                  [-UseCurrentSession] [-CSV] [-Help]

    14) Import-FrankensteinGroupMembers
        Imports group membership using a membership export CSV and a migration mapping CSV (Source/Target).
        DDG rows are skipped automatically. M365 Owner rows are applied via the Owners link type.
        Outputs a timestamped log CSV showing successes, skips, and failures.
        Switches: [-DistributionGroups] [-MailEnabledSecurityGroups] [-DynamicDistributionGroups] [-M365Groups]
                  [-UseCurrentSession] [-Help]

    15) Get-FrankensteinMailboxReport
        Exports a comprehensive mailbox report for UserMailbox, SharedMailbox, RoomMailbox, and EquipmentMailbox.
        Defaults to enabled mailboxes only. Supports scoping to a CSV list of addresses via -ImportCSV.
        Always outputs a timestamped CSV. Use -IncludeStatistics to append size, item count, and last logon.
        Switches:    [-IncludeDisabled] [-IncludeStatistics] [-ImportCSV] [-UseCurrentSession] [-Help]
        Parameters:  [-OutputPath <path>]

    18) Invoke-FrankensteinDLMigrator
        GUI-driven distribution group membership migrator. Source is always M365; target is M365
        or on-prem Exchange. One mapping CSV covers both group rows and member rows (Source + Target
        columns). Supports adding members to existing groups or creating groups with a prefix and
        new SMTP domain. Nested groups are added directly if mapped, or expanded if not. Exports
        a timestamped log CSV with Created, Success, AlreadyMember, Conflict, Skipped, and Failed.

    17) Invoke-FrankensteinPermissionMigrator
        GUI-driven permission migration tool for tenant-to-tenant cutovers.
        Load a Source-to-Target mapping CSV, connect to source (reads permissions live with optional
        DL expansion), connect to target, then run the migration. Exports a timestamped log CSV
        with Success, Skipped, AlreadyExists, and Failed results per permission entry.
        Switches: [-Help]

    16) Get-FrankensteinAliasReport
        Exports a flat address map with one row per address per recipient. Columns: DisplayName,
        PrimarySmtpAddress, Address, AddressType (SMTP/smtp/X500/SIP), RecipientType.
        Covers mailboxes, groups (DG/MESG/M365), mail users, and mail contacts.
        Use -LegacyDN to append an X500 row from each object's LegacyExchangeDN.
        Switches:    [-Mailboxes] [-Groups] [-MailUsers] [-MailContacts] [-LegacyDN] [-SIP]
                     [-ImportCSV] [-UseCurrentSession] [-CSV] [-Help]
        Parameters:  [-OutputPath <path>]

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

        # Optional -- defaults to a broad read/write admin scope set
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

    # -- Organization ---------------------------------------------------------
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

    # -- Domains --------------------------------------------------------------
    Get-Linebreak
    Write-Host "Domains" -ForegroundColor Cyan
    $Domains = Get-MgDomain -All
    $Domains | Format-Table Id, IsDefault, IsVerified, AuthenticationType -AutoSize
    if ($CSV) {
        $Domains | Select-Object Id, IsDefault, IsInitial, IsVerified, AuthenticationType,
            @{Name="SupportedServices"; Expression={$_.SupportedServices -join ";"}} |
            Export-Csv ".\Domains_$DateStamp.csv" -NoTypeInformation
    }

    # -- Licenses -------------------------------------------------------------
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

    # -- Users -----------------------------------------------------------------
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

    # -- MFA & Authentication Registration ------------------------------------
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

    # -- Admin Role Assignments ------------------------------------------------
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

    # -- Groups ----------------------------------------------------------------
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

    # -- Devices ---------------------------------------------------------------
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

    # -- Conditional Access ----------------------------------------------------
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

    # -- Applications ---------------------------------------------------------
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

    # -- Security Posture -----------------------------------------------------
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

    # -- Discovery Summary -----------------------------------------------------
    Get-Linebreak
    Write-Host "ENTRA ID DISCOVERY SUMMARY" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Tenant  : $TenantName  ($TenantId)"
    Write-Host "  Users   : $($AllUsers.Count) total  |  $($Licensed.Count) licensed  |  $($GuestUsers.Count) guests  |  $($DisabledUsers.Count) disabled"
    Write-Host "  Groups  : $($AllGroups.Count) total  |  $($SecurityGroups.Count) security  |  $($M365Groups.Count) M365  |  $($DynamicGroups.Count) dynamic"
    Write-Host "  Devices : $($Devices.Count) total  |  $($EntraJoined.Count) Entra joined  |  $($HybridJoined.Count) hybrid  |  $($Compliant.Count) compliant"
    Write-Host "  CA      : $($CAPolicies.Count) policies  ($($CAEnabled.Count) enabled)"
    Write-Host "  Apps    : $($AppRegs.Count) registrations  |  $($EntApps.Count) enterprise apps"
    Write-Host "  DirSync : $(if($OnPremSync){'Enabled -- Last sync: ' + $LastSync}else{'Cloud-Only'})"
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

function Expand-FrankensteinDLMembers {
    [CmdletBinding()]
    Param (
        [string]$Identity,
        [Switch]$Recurse,
        [int]$Depth    = 0,
        [int]$MaxDepth = 5,
        [System.Collections.Generic.HashSet[string]]$Visited = $null
    )
    if ($null -eq $Visited) {
        $Visited = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }
    if ($Depth -gt $MaxDepth -or -not $Visited.Add($Identity)) { return }

    $members = @(Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction SilentlyContinue)
    foreach ($m in $members) {
        if ($Recurse -and $m.RecipientTypeDetails -in @('MailUniversalDistributionGroup','MailUniversalSecurityGroup')) {
            Expand-FrankensteinDLMembers -Identity $m.PrimarySmtpAddress -Recurse:$Recurse -Depth ($Depth + 1) -MaxDepth $MaxDepth -Visited $Visited
        } else {
            $m
        }
    }
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
    Any delegate that is a Distribution Group or Mail-Enabled Security Group is automatically
    expanded to its individual members, including nested DLs recursively. Expanded rows include
    an ExpandedFromGroup column showing the originating DL address.

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

    $Mailboxes      = Get-Mailbox -RecipientTypeDetails UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox
    $total          = $Mailboxes.Count
    $count          = 0
    $Results        = [System.Collections.Generic.List[PSCustomObject]]::new()
    $RecipientCache = @{}
    $SeenKeys       = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

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

        $dlTypes = @('MailUniversalDistributionGroup','MailUniversalSecurityGroup')

        function Add-PermissionRow ([string]$AccessType, $r, [string]$rawId, [string]$expandedFrom) {
            $smtp = if ($r) { $r.PrimarySmtpAddress } else { $rawId }
            $key  = "$MailboxSMTP|$smtp|$AccessType"
            if ($SeenKeys.Add($key)) {
                $Results.Add([PSCustomObject]@{
                    DisplayName                = $mbx.DisplayName
                    UserPrincipalName          = $MailboxSMTP
                    MailboxType                = $MailboxType
                    MailboxExchangeGuid        = $MailboxExchangeGuid
                    AccessType                 = $AccessType
                    UserWithAccess             = $smtp
                    UserWithAccessType         = if ($r) { $r.RecipientTypeDetails } else { 'Unknown/External' }
                    UserWithAccessExchangeGuid = if ($r) { $r.ExchangeGuid }         else { $null }
                    ExpandedFromGroup          = $expandedFrom
                })
            }
        }

        function Resolve-AndAdd ([string]$rawId, [string]$AccessType) {
            $r = Resolve-Recipient $rawId
            $smtp = if ($r) { $r.PrimarySmtpAddress } else { $rawId }
            if ($r -and $r.RecipientTypeDetails -in $dlTypes) {
                $members = @(Expand-FrankensteinDLMembers -Identity $smtp -Recurse)
                foreach ($m in $members) { Add-PermissionRow $AccessType $m $m.PrimarySmtpAddress $smtp }
            } else {
                Add-PermissionRow $AccessType $r $rawId ''
            }
        }

        if ($FullAccess) {
            Get-MailboxPermission -Identity $mbx.Identity -ErrorAction SilentlyContinue |
                Where-Object { -not $_.IsInherited -and $_.User -notlike "NT AUTHORITY\SELF" } |
                ForEach-Object { Resolve-AndAdd $_.User 'FullAccess' }
        }

        if ($SendAs) {
            Get-RecipientPermission -Identity $mbx.Identity -ErrorAction SilentlyContinue |
                Where-Object { $_.Trustee -ne "NT AUTHORITY\SELF" } |
                ForEach-Object { Resolve-AndAdd $_.Trustee 'SendAs' }
        }

        if ($SendOnBehalf) {
            foreach ($delegate in $mbx.GrantSendOnBehalfTo) {
                Resolve-AndAdd $delegate 'SendOnBehalf'
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

function Import-FrankensteinMailboxPermissions {
    [CmdletBinding()]
    Param (
        [Switch]$FullAccess,
        [Switch]$SendAs,
        [Switch]$SendOnBehalf,
        [Switch]$UseCurrentSession,
        [Switch]$Help
    )

    if ($Help) {
        Write-Host @"
SYNOPSIS
    Imports FullAccess, SendAs, and SendOnBehalf permissions using a permissions export and two mapping files.

DESCRIPTION
    Reads a permissions CSV exported by Get-FrankensteinMailboxPermissions and applies those permissions
    in the target environment. Two separate mapping files (Source/Target CSVs) are required:
      - Mailbox mapping  : translates source mailbox SMTP addresses to destination addresses
      - Delegate mapping : translates source delegate SMTP addresses to destination addresses
    Both the mailbox AND the delegate must resolve for a permission to be applied.
    Supports mailboxes and distribution groups as delegates.
    Results are written to a timestamped log CSV.

PARAMETERS
    -UseCurrentSession  Use the current Exchange Online session instead of prompting to connect.
    -FullAccess         Process only FullAccess rows from the permissions export.
    -SendAs             Process only SendAs rows from the permissions export.
    -SendOnBehalf       Process only SendOnBehalf rows from the permissions export.
    -Help               Display this help text.

    Note: If no permission type switch is specified, all three types are processed.

EXAMPLE
    Import-FrankensteinMailboxPermissions -UseCurrentSession -FullAccess -SendAs -SendOnBehalf

NOTES
    Author: Eric D. Frank
"@
        return
    }

    Add-Type -AssemblyName System.Windows.Forms

    # --- File picker 1: Permissions export ---
    [System.Windows.Forms.MessageBox]::Show(
        "Step 1 of 3: Select the permissions export CSV produced by Get-FrankensteinMailboxPermissions.`n`nExpected columns:`n  - UserPrincipalName  : SMTP address of the mailbox`n  - UserWithAccess     : SMTP address of the delegate`n  - AccessType         : FullAccess, SendAs, or SendOnBehalf",
        "Select Permissions Export File",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null

    $permDialog        = New-Object System.Windows.Forms.OpenFileDialog
    $permDialog.Title  = "Select Permissions Export CSV"
    $permDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    if ($permDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "No permissions file selected. Exiting." -ForegroundColor Yellow
        return
    }
    $PermissionsFile = $permDialog.FileName

    # --- File picker 2: Mailbox mapping ---
    [System.Windows.Forms.MessageBox]::Show(
        "Step 2 of 3: Select your MAILBOX mapping CSV.`n`nRequired columns:`n  - Source : original mailbox SMTP address (matches UserPrincipalName in the permissions export)`n  - Target : destination mailbox SMTP address in the new environment",
        "Select Mailbox Mapping File",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null

    $mbxMapDialog        = New-Object System.Windows.Forms.OpenFileDialog
    $mbxMapDialog.Title  = "Select Mailbox Mapping CSV"
    $mbxMapDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    if ($mbxMapDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "No mailbox mapping file selected. Exiting." -ForegroundColor Yellow
        return
    }
    $MailboxMappingFile = $mbxMapDialog.FileName

    # --- File picker 3: Delegate mapping ---
    [System.Windows.Forms.MessageBox]::Show(
        "Step 3 of 3: Select your DELEGATE mapping CSV.`n`nRequired columns:`n  - Source : original delegate SMTP address (matches UserWithAccess in the permissions export)`n  - Target : destination delegate SMTP address in the new environment`n`nDelegates may be users or distribution groups.",
        "Select Delegate Mapping File",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null

    $delMapDialog        = New-Object System.Windows.Forms.OpenFileDialog
    $delMapDialog.Title  = "Select Delegate Mapping CSV"
    $delMapDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    if ($delMapDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "No delegate mapping file selected. Exiting." -ForegroundColor Yellow
        return
    }
    $DelegateMappingFile = $delMapDialog.FileName

    # --- Load and validate files ---
    $Permissions = Import-Csv $PermissionsFile

    $MailboxMapping = @{}
    foreach ($entry in (Import-Csv $MailboxMappingFile)) {
        if ($entry.Source -and $entry.Target) {
            $MailboxMapping[$entry.Source.Trim().ToLower()] = $entry.Target.Trim()
        }
    }

    $DelegateMapping = @{}
    foreach ($entry in (Import-Csv $DelegateMappingFile)) {
        if ($entry.Source -and $entry.Target) {
            $DelegateMapping[$entry.Source.Trim().ToLower()] = $entry.Target.Trim()
        }
    }

    if ($MailboxMapping.Count -eq 0) {
        Write-Host "Mailbox mapping file is empty or missing Source/Target columns. Exiting." -ForegroundColor Red
        return
    }
    if ($DelegateMapping.Count -eq 0) {
        Write-Host "Delegate mapping file is empty or missing Source/Target columns. Exiting." -ForegroundColor Red
        return
    }

    $processAll = (-not $FullAccess -and -not $SendAs -and -not $SendOnBehalf)

    if (-not $UseCurrentSession) {
        Connect-ExchangeOnline
    }

    $Log            = [System.Collections.Generic.List[PSCustomObject]]::new()
    $total          = $Permissions.Count
    $count          = 0
    $RecipientCache = @{}
    $SeenKeys       = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    function Resolve-TargetRecipient ([string]$identity) {
        $key = $identity.ToLower()
        if (-not $RecipientCache.ContainsKey($key)) {
            $RecipientCache[$key] = Get-Recipient -Identity $identity -ErrorAction SilentlyContinue |
                Select-Object -First 1
        }
        return $RecipientCache[$key]
    }

    Write-Host "Processing $total permission rows..." -ForegroundColor Cyan

    foreach ($row in $Permissions) {
        $count++
        Write-Progress -Activity "Importing Permissions" `
            -Status "Row $count of $total - $($row.AccessType): $($row.UserPrincipalName)" `
            -PercentComplete ([math]::Round(($count / $total) * 100))

        $accessType = $row.AccessType

        if (-not $processAll) {
            if ($accessType -eq 'FullAccess'  -and -not $FullAccess)  { continue }
            if ($accessType -eq 'SendAs'       -and -not $SendAs)       { continue }
            if ($accessType -eq 'SendOnBehalf' -and -not $SendOnBehalf) { continue }
        }

        $seenKey = "$($row.UserPrincipalName.Trim())|$($row.UserWithAccess.Trim())|$accessType"
        if (-not $SeenKeys.Add($seenKey)) {
            $Log.Add([PSCustomObject][ordered]@{
                SourceMailbox  = $row.UserPrincipalName
                TargetMailbox  = $null
                SourceDelegate = $row.UserWithAccess
                TargetDelegate = $null
                AccessType     = $accessType
                Status         = 'Skipped'
                Details        = 'Duplicate entry in permissions file'
            })
            continue
        }

        $mbxKey      = $row.UserPrincipalName.Trim().ToLower()
        $delegateKey = $row.UserWithAccess.Trim().ToLower()
        $mbxTarget      = $MailboxMapping[$mbxKey]
        $delegateTarget = $DelegateMapping[$delegateKey]

        $logEntry = [ordered]@{
            SourceMailbox  = $row.UserPrincipalName
            TargetMailbox  = $mbxTarget
            SourceDelegate = $row.UserWithAccess
            TargetDelegate = $delegateTarget
            AccessType     = $accessType
            Status         = $null
            Details        = $null
        }

        if (-not $mbxTarget -and -not $delegateTarget) {
            $logEntry.Status  = 'Skipped'
            $logEntry.Details = 'Mailbox not found in mailbox mapping; Delegate not found in delegate mapping'
            $Log.Add([PSCustomObject]$logEntry)
            continue
        }
        if (-not $mbxTarget) {
            $logEntry.Status  = 'Skipped'
            $logEntry.Details = 'Mailbox not found in mailbox mapping'
            $Log.Add([PSCustomObject]$logEntry)
            continue
        }
        if (-not $delegateTarget) {
            $logEntry.Status  = 'Skipped'
            $logEntry.Details = 'Delegate not found in delegate mapping'
            $Log.Add([PSCustomObject]$logEntry)
            continue
        }

        # Resolve mapping values to PrimarySmtpAddress (handles proxy addresses)
        $resolvedMbx      = Resolve-TargetRecipient $mbxTarget
        $resolvedDelegate = Resolve-TargetRecipient $delegateTarget

        if (-not $resolvedMbx -and -not $resolvedDelegate) {
            $logEntry.Status  = 'Failed'
            $logEntry.Details = "Neither target mailbox ('$mbxTarget') nor target delegate ('$delegateTarget') found in environment"
            $Log.Add([PSCustomObject]$logEntry)
            continue
        }
        if (-not $resolvedMbx) {
            $logEntry.Status  = 'Failed'
            $logEntry.Details = "Target mailbox '$mbxTarget' not found in environment"
            $Log.Add([PSCustomObject]$logEntry)
            continue
        }
        if (-not $resolvedDelegate) {
            $logEntry.Status  = 'Failed'
            $logEntry.Details = "Target delegate '$delegateTarget' not found in environment"
            $Log.Add([PSCustomObject]$logEntry)
            continue
        }

        $mbxPrimary      = $resolvedMbx.PrimarySmtpAddress
        $delegatePrimary = $resolvedDelegate.PrimarySmtpAddress

        $logEntry.TargetMailbox  = $mbxPrimary
        $logEntry.TargetDelegate = $delegatePrimary

        try {
            switch ($accessType) {
                'FullAccess' {
                    Add-MailboxPermission -Identity $mbxPrimary -User $delegatePrimary `
                        -AccessRights FullAccess -InheritanceType All `
                        -AutoMapping $false -ErrorAction Stop | Out-Null
                }
                'SendAs' {
                    Add-RecipientPermission -Identity $mbxPrimary -Trustee $delegatePrimary `
                        -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                }
                'SendOnBehalf' {
                    Set-Mailbox -Identity $mbxPrimary -GrantSendOnBehalfTo @{Add = $delegatePrimary} `
                        -ErrorAction Stop
                }
            }
            $logEntry.Status  = 'Success'
            $logEntry.Details = 'Permission applied'
        } catch {
            $logEntry.Status  = 'Failed'
            $logEntry.Details = $_.Exception.Message
        }

        $Log.Add([PSCustomObject]$logEntry)
    }

    Write-Progress -Activity "Importing Permissions" -Completed

    $LogFile = ".\ImportPermissionsLog_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
    $Log | Export-Csv $LogFile -NoTypeInformation -Encoding UTF8

    $success = ($Log | Where-Object Status -eq 'Success').Count
    $skipped = ($Log | Where-Object Status -eq 'Skipped').Count
    $failed  = ($Log | Where-Object Status -eq 'Failed').Count

    Write-Host "`nImport complete." -ForegroundColor Green
    Write-Host "  Success : $success" -ForegroundColor Green
    Write-Host "  Skipped : $skipped" -ForegroundColor Yellow
    Write-Host "  Failed  : $failed"  -ForegroundColor $(if ($failed -gt 0) { 'Red' } else { 'Green' })
    Write-Host "  Log     : $LogFile" -ForegroundColor Cyan
}

function Get-FrankensteinGroups {
    [CmdletBinding()]
    Param (
        [Switch]$DistributionGroups,
        [Switch]$MailEnabledSecurityGroups,
        [Switch]$DynamicDistributionGroups,
        [Switch]$M365Groups,
        [Switch]$UseCurrentSession,
        [Switch]$CSV,
        [Switch]$Help
    )

    if ($Help) {
        Write-Host @"
SYNOPSIS
    Exports properties for Distribution Groups, Mail-Enabled Security Groups,
    Dynamic Distribution Groups, and M365 (Unified) Groups.

DESCRIPTION
    By default all four group types are exported. Use type switches to scope the export.
    A GroupType column identifies each record. Type-specific properties (RecipientFilter,
    Visibility, SharePointSiteUrl, etc.) are included for all records and left blank where
    not applicable so the CSV schema remains consistent across types.

PARAMETERS
    -DistributionGroups         Include Distribution Groups.
    -MailEnabledSecurityGroups  Include Mail-Enabled Security Groups.
    -DynamicDistributionGroups  Include Dynamic Distribution Groups.
    -M365Groups                 Include M365 (Unified) Groups.
    -UseCurrentSession          Use the current Exchange Online session.
    -CSV                        Export results to a timestamped CSV file.
    -Help                       Display this help text.

    Note: If no group type switch is specified, all types are exported.

EXAMPLE
    Get-FrankensteinGroups -UseCurrentSession -CSV
    Get-FrankensteinGroups -UseCurrentSession -M365Groups -DistributionGroups -CSV

NOTES
    Author: Eric D. Frank
"@
        return
    }

    if (-not $UseCurrentSession) { Connect-ExchangeOnline }

    $processAll = (-not $DistributionGroups -and -not $MailEnabledSecurityGroups -and -not $DynamicDistributionGroups -and -not $M365Groups)
    $Results    = [System.Collections.Generic.List[PSCustomObject]]::new()

    function Join-Prop ($col) {
        if ($col) { ($col | ForEach-Object { "$_" }) -join '; ' } else { '' }
    }

    function New-GroupRecord ($g, [string]$GroupTypeName) {
        [PSCustomObject]@{
            GroupType                            = $GroupTypeName
            DisplayName                          = $g.DisplayName
            PrimarySmtpAddress                   = $g.PrimarySmtpAddress
            Alias                                = $g.Alias
            ExchangeGuid                         = $g.ExchangeGuid
            ExternalDirectoryObjectId            = if ($g.PSObject.Properties['ExternalDirectoryObjectId'] -and $g.ExternalDirectoryObjectId) { "$($g.ExternalDirectoryObjectId)" } else { '' }
            Description                          = $g.Description
            ManagedBy                            = Join-Prop $g.ManagedBy
            HiddenFromAddressListsEnabled        = $g.HiddenFromAddressListsEnabled
            RequireSenderAuthenticationEnabled   = $g.RequireSenderAuthenticationEnabled
            MemberJoinRestriction                = if ($g.PSObject.Properties['MemberJoinRestriction'])    { "$($g.MemberJoinRestriction)" }    else { '' }
            MemberDepartRestriction              = if ($g.PSObject.Properties['MemberDepartRestriction'])  { "$($g.MemberDepartRestriction)" }  else { '' }
            ModerationEnabled                    = $g.ModerationEnabled
            ModeratedBy                          = Join-Prop $g.ModeratedBy
            SendModerationNotifications          = if ($g.PSObject.Properties['SendModerationNotifications'])          { "$($g.SendModerationNotifications)" }          else { '' }
            AcceptMessagesOnlyFrom               = Join-Prop $g.AcceptMessagesOnlyFrom
            AcceptMessagesOnlyFromDLMembers      = Join-Prop $g.AcceptMessagesOnlyFromDLMembers
            BypassModerationFromSendersOrMembers = if ($g.PSObject.Properties['BypassModerationFromSendersOrMembers']) { Join-Prop $g.BypassModerationFromSendersOrMembers } else { '' }
            ReportToManagerEnabled               = if ($g.PSObject.Properties['ReportToManagerEnabled'])               { "$($g.ReportToManagerEnabled)" }               else { '' }
            ReportToOriginatorEnabled            = if ($g.PSObject.Properties['ReportToOriginatorEnabled'])            { "$($g.ReportToOriginatorEnabled)" }            else { '' }
            RecipientFilter                      = if ($g.PSObject.Properties['RecipientFilter'])                      { "$($g.RecipientFilter)" }                      else { '' }
            RecipientContainer                   = if ($g.PSObject.Properties['RecipientContainer'])                   { "$($g.RecipientContainer)" }                   else { '' }
            Visibility                           = if ($g.PSObject.Properties['Visibility'])                           { "$($g.Visibility)" }                           else { '' }
            AccessType                           = if ($g.PSObject.Properties['AccessType'])                           { "$($g.AccessType)" }                           else { '' }
            SharePointSiteUrl                    = if ($g.PSObject.Properties['SharePointSiteUrl'])                    { "$($g.SharePointSiteUrl)" }                    else { '' }
            AutoSubscribeNewMembers              = if ($g.PSObject.Properties['AutoSubscribeNewMembers'])              { "$($g.AutoSubscribeNewMembers)" }              else { '' }
            WelcomeMessageEnabled                = if ($g.PSObject.Properties['WelcomeMessageEnabled'])                { "$($g.WelcomeMessageEnabled)" }                else { '' }
            EmailAddresses                       = Join-Prop $g.EmailAddresses
            WhenCreated                          = $g.WhenCreated
            WhenChanged                          = $g.WhenChanged
        }
    }

    if ($processAll -or $DistributionGroups) {
        Write-Host "Gathering Distribution Groups..." -ForegroundColor Cyan
        $groups = @(Get-DistributionGroup -RecipientTypeDetails MailUniversalDistributionGroup -ResultSize Unlimited)
        $i = 0; $n = $groups.Count
        foreach ($g in $groups) {
            $i++
            Write-Progress -Activity "Distribution Groups" -Status "$($g.DisplayName) ($i of $n)" -PercentComplete ([math]::Round(($i / $n) * 100))
            $Results.Add((New-GroupRecord $g "DistributionGroup"))
        }
        Write-Progress -Activity "Distribution Groups" -Completed
        Write-Host "  Found $n Distribution Groups." -ForegroundColor Gray
    }

    if ($processAll -or $MailEnabledSecurityGroups) {
        Write-Host "Gathering Mail-Enabled Security Groups..." -ForegroundColor Cyan
        $groups = @(Get-DistributionGroup -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize Unlimited)
        $i = 0; $n = $groups.Count
        foreach ($g in $groups) {
            $i++
            Write-Progress -Activity "Mail-Enabled Security Groups" -Status "$($g.DisplayName) ($i of $n)" -PercentComplete ([math]::Round(($i / $n) * 100))
            $Results.Add((New-GroupRecord $g "MailEnabledSecurityGroup"))
        }
        Write-Progress -Activity "Mail-Enabled Security Groups" -Completed
        Write-Host "  Found $n Mail-Enabled Security Groups." -ForegroundColor Gray
    }

    if ($processAll -or $DynamicDistributionGroups) {
        Write-Host "Gathering Dynamic Distribution Groups..." -ForegroundColor Cyan
        $groups = @(Get-DynamicDistributionGroup -ResultSize Unlimited)
        $i = 0; $n = $groups.Count
        foreach ($g in $groups) {
            $i++
            Write-Progress -Activity "Dynamic Distribution Groups" -Status "$($g.DisplayName) ($i of $n)" -PercentComplete ([math]::Round(($i / $n) * 100))
            $Results.Add((New-GroupRecord $g "DynamicDistributionGroup"))
        }
        Write-Progress -Activity "Dynamic Distribution Groups" -Completed
        Write-Host "  Found $n Dynamic Distribution Groups." -ForegroundColor Gray
    }

    if ($processAll -or $M365Groups) {
        Write-Host "Gathering M365 Groups..." -ForegroundColor Cyan
        $groups = @(Get-UnifiedGroup -ResultSize Unlimited)
        $i = 0; $n = $groups.Count
        foreach ($g in $groups) {
            $i++
            Write-Progress -Activity "M365 Groups" -Status "$($g.DisplayName) ($i of $n)" -PercentComplete ([math]::Round(($i / $n) * 100))
            $Results.Add((New-GroupRecord $g "M365Group"))
        }
        Write-Progress -Activity "M365 Groups" -Completed
        Write-Host "  Found $n M365 Groups." -ForegroundColor Gray
    }

    Write-Host "`nTotal groups collected: $($Results.Count)" -ForegroundColor Cyan

    if ($CSV) {
        $FileName = ".\Groups_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
        $Results | Export-Csv $FileName -NoTypeInformation -Encoding UTF8
        Write-Host "Export complete: $FileName" -ForegroundColor Green
    } else {
        $Results
    }
}

function Get-FrankensteinGroupMember {
    [CmdletBinding()]
    Param (
        [Switch]$DistributionGroups,
        [Switch]$MailEnabledSecurityGroups,
        [Switch]$DynamicDistributionGroups,
        [Switch]$M365Groups,
        [Switch]$ImportCSV,
        [Switch]$UseCurrentSession,
        [Switch]$CSV,
        [Switch]$Help
    )

    if ($Help) {
        Write-Host @"
SYNOPSIS
    Exports group membership for all mail-enabled group types.

DESCRIPTION
    For Distribution Groups and Mail-Enabled Security Groups, direct members are enumerated.
    For Dynamic Distribution Groups, membership is resolved live from the recipient filter
    and tagged with Role = Dynamic to indicate it cannot be statically imported.
    For M365 Groups, both Members and Owners are exported with a Role column (Member/Owner).

    Use -ImportCSV to scope the pull to a specific list of groups rather than querying all
    groups from Exchange Online. A file picker will prompt for a CSV with two columns:
      - PrimarySmtpAddress : the group's primary SMTP address
      - GroupType          : DistributionGroup, MailEnabledSecurityGroup,
                             DynamicDistributionGroup, or M365Group

    Type switches (-M365Groups, etc.) can be combined with -ImportCSV to further filter
    which rows from the imported list are processed.

PARAMETERS
    -DistributionGroups         Include Distribution Groups.
    -MailEnabledSecurityGroups  Include Mail-Enabled Security Groups.
    -DynamicDistributionGroups  Include Dynamic Distribution Groups (live filter resolution).
    -M365Groups                 Include M365 (Unified) Groups (Members and Owners).
    -ImportCSV                  Scope membership pull to a list of groups from a CSV file.
    -UseCurrentSession          Use the current Exchange Online session.
    -CSV                        Export results to a timestamped CSV file.
    -Help                       Display this help text.

    Note: If no group type switch is specified, all types are processed.

EXAMPLE
    Get-FrankensteinGroupMember -UseCurrentSession -CSV
    Get-FrankensteinGroupMember -UseCurrentSession -ImportCSV -CSV
    Get-FrankensteinGroupMember -UseCurrentSession -ImportCSV -M365Groups -CSV

NOTES
    Author: Eric D. Frank
"@
        return
    }

    if (-not $UseCurrentSession) { Connect-ExchangeOnline }

    $ImportedGroups = $null
    if ($ImportCSV) {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show(
            "Select a CSV containing the groups to scope this membership pull to.`n`nRequired columns:`n  - PrimarySmtpAddress : the group's primary SMTP address`n  - GroupType          : DistributionGroup, MailEnabledSecurityGroup,`n                         DynamicDistributionGroup, or M365Group`n`nOnly groups present in this file will be queried for membership.",
            "Select Group Scope CSV",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null

        $scopeDialog        = New-Object System.Windows.Forms.OpenFileDialog
        $scopeDialog.Title  = "Select Group Scope CSV"
        $scopeDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        if ($scopeDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            Write-Host "No scope file selected. Exiting." -ForegroundColor Yellow
            return
        }
        $ImportedGroups = Import-Csv $scopeDialog.FileName
        if (-not ($ImportedGroups | Get-Member -Name 'PrimarySmtpAddress') -or -not ($ImportedGroups | Get-Member -Name 'GroupType')) {
            Write-Host "Scope CSV is missing required columns (PrimarySmtpAddress, GroupType). Exiting." -ForegroundColor Red
            return
        }
        Write-Host "Scope file loaded: $($ImportedGroups.Count) group(s) to process." -ForegroundColor Cyan
    }

    $processAll = (-not $DistributionGroups -and -not $MailEnabledSecurityGroups -and -not $DynamicDistributionGroups -and -not $M365Groups)
    $Results    = [System.Collections.Generic.List[PSCustomObject]]::new()
    $SeenKeys   = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    function Add-MemberRecord ($group, [string]$groupType, $member, [string]$role) {
        $memberSmtp = if ($member.PrimarySmtpAddress) { $member.PrimarySmtpAddress } else { $member.WindowsLiveID }
        $key = "$($group.PrimarySmtpAddress)|$memberSmtp|$role"
        if ($SeenKeys.Add($key)) {
            $Results.Add([PSCustomObject]@{
                GroupDisplayName           = $group.DisplayName
                GroupPrimarySmtp           = $group.PrimarySmtpAddress
                GroupType                  = $groupType
                MemberDisplayName          = $member.DisplayName
                MemberPrimarySmtp          = $memberSmtp
                MemberRecipientTypeDetails = $member.RecipientTypeDetails
                Role                       = $role
            })
        }
    }

    if ($processAll -or $DistributionGroups) {
        Write-Host "Gathering Distribution Group membership..." -ForegroundColor Cyan
        if ($ImportCSV) {
            $scopedSmtps = @($ImportedGroups | Where-Object { $_.GroupType -eq 'DistributionGroup' } | ForEach-Object { $_.PrimarySmtpAddress })
            $groups = foreach ($smtp in $scopedSmtps) {
                $g = Get-DistributionGroup -Identity $smtp -ErrorAction SilentlyContinue
                if (-not $g) { Write-Warning "Distribution Group not found: $smtp" }
                $g
            }
            $groups = @($groups | Where-Object { $_ })
        } else {
            $groups = @(Get-DistributionGroup -RecipientTypeDetails MailUniversalDistributionGroup -ResultSize Unlimited)
        }
        $i = 0; $n = $groups.Count
        foreach ($g in $groups) {
            $i++
            Write-Progress -Activity "DG Membership" -Status "$($g.DisplayName) ($i of $n)" -PercentComplete ([math]::Round(($i / $n) * 100))
            $members = @(Get-DistributionGroupMember -Identity $g.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue)
            foreach ($m in $members) { Add-MemberRecord $g "DistributionGroup" $m "Member" }
        }
        Write-Progress -Activity "DG Membership" -Completed
        Write-Host "  Processed $n Distribution Groups." -ForegroundColor Gray
    }

    if ($processAll -or $MailEnabledSecurityGroups) {
        Write-Host "Gathering Mail-Enabled Security Group membership..." -ForegroundColor Cyan
        if ($ImportCSV) {
            $scopedSmtps = @($ImportedGroups | Where-Object { $_.GroupType -eq 'MailEnabledSecurityGroup' } | ForEach-Object { $_.PrimarySmtpAddress })
            $groups = foreach ($smtp in $scopedSmtps) {
                $g = Get-DistributionGroup -Identity $smtp -ErrorAction SilentlyContinue
                if (-not $g) { Write-Warning "Mail-Enabled Security Group not found: $smtp" }
                $g
            }
            $groups = @($groups | Where-Object { $_ })
        } else {
            $groups = @(Get-DistributionGroup -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize Unlimited)
        }
        $i = 0; $n = $groups.Count
        foreach ($g in $groups) {
            $i++
            Write-Progress -Activity "MESG Membership" -Status "$($g.DisplayName) ($i of $n)" -PercentComplete ([math]::Round(($i / $n) * 100))
            $members = @(Get-DistributionGroupMember -Identity $g.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue)
            foreach ($m in $members) { Add-MemberRecord $g "MailEnabledSecurityGroup" $m "Member" }
        }
        Write-Progress -Activity "MESG Membership" -Completed
        Write-Host "  Processed $n Mail-Enabled Security Groups." -ForegroundColor Gray
    }

    if ($processAll -or $DynamicDistributionGroups) {
        Write-Host "Gathering Dynamic Distribution Group membership (resolving filters live)..." -ForegroundColor Cyan
        if ($ImportCSV) {
            $scopedSmtps = @($ImportedGroups | Where-Object { $_.GroupType -eq 'DynamicDistributionGroup' } | ForEach-Object { $_.PrimarySmtpAddress })
            $groups = foreach ($smtp in $scopedSmtps) {
                $g = Get-DynamicDistributionGroup -Identity $smtp -ErrorAction SilentlyContinue
                if (-not $g) { Write-Warning "Dynamic Distribution Group not found: $smtp" }
                $g
            }
            $groups = @($groups | Where-Object { $_ })
        } else {
            $groups = @(Get-DynamicDistributionGroup -ResultSize Unlimited)
        }
        $i = 0; $n = $groups.Count
        foreach ($g in $groups) {
            $i++
            Write-Progress -Activity "DDG Membership" -Status "$($g.DisplayName) ($i of $n)" -PercentComplete ([math]::Round(($i / $n) * 100))
            try {
                $members = @(Get-Recipient -RecipientPreviewFilter $g.RecipientFilter -ResultSize Unlimited -ErrorAction Stop)
                foreach ($m in $members) { Add-MemberRecord $g "DynamicDistributionGroup" $m "Dynamic" }
            } catch {
                Write-Warning "Could not resolve membership for DDG '$($g.DisplayName)': $($_.Exception.Message)"
            }
        }
        Write-Progress -Activity "DDG Membership" -Completed
        Write-Host "  Processed $n Dynamic Distribution Groups." -ForegroundColor Gray
    }

    if ($processAll -or $M365Groups) {
        Write-Host "Gathering M365 Group membership..." -ForegroundColor Cyan
        if ($ImportCSV) {
            $scopedSmtps = @($ImportedGroups | Where-Object { $_.GroupType -eq 'M365Group' } | ForEach-Object { $_.PrimarySmtpAddress })
            $groups = foreach ($smtp in $scopedSmtps) {
                $g = Get-UnifiedGroup -Identity $smtp -ErrorAction SilentlyContinue
                if (-not $g) { Write-Warning "M365 Group not found: $smtp" }
                $g
            }
            $groups = @($groups | Where-Object { $_ })
        } else {
            $groups = @(Get-UnifiedGroup -ResultSize Unlimited)
        }
        $i = 0; $n = $groups.Count
        foreach ($g in $groups) {
            $i++
            Write-Progress -Activity "M365 Membership" -Status "$($g.DisplayName) ($i of $n)" -PercentComplete ([math]::Round(($i / $n) * 100))
            $members = @(Get-UnifiedGroupLinks -Identity $g.Identity -LinkType Members -ResultSize Unlimited -ErrorAction SilentlyContinue)
            foreach ($m in $members) { Add-MemberRecord $g "M365Group" $m "Member" }
            $owners = @(Get-UnifiedGroupLinks -Identity $g.Identity -LinkType Owners -ResultSize Unlimited -ErrorAction SilentlyContinue)
            foreach ($m in $owners)  { Add-MemberRecord $g "M365Group" $m "Owner" }
        }
        Write-Progress -Activity "M365 Membership" -Completed
        Write-Host "  Processed $n M365 Groups." -ForegroundColor Gray
    }

    Write-Host "`nTotal membership records: $($Results.Count)" -ForegroundColor Cyan

    if ($CSV) {
        $FileName = ".\GroupMembership_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
        $Results | Export-Csv $FileName -NoTypeInformation -Encoding UTF8
        Write-Host "Export complete: $FileName" -ForegroundColor Green
    } else {
        $Results
    }
}

function Import-FrankensteinGroupMembers {
    [CmdletBinding()]
    Param (
        [Switch]$DistributionGroups,
        [Switch]$MailEnabledSecurityGroups,
        [Switch]$DynamicDistributionGroups,
        [Switch]$M365Groups,
        [Switch]$UseCurrentSession,
        [Switch]$Help
    )

    if ($Help) {
        Write-Host @"
SYNOPSIS
    Imports group membership using a membership export CSV and two mapping files.

DESCRIPTION
    Reads a membership CSV exported by Get-FrankensteinGroupMember and applies membership
    in the target environment. Two separate mapping files (Source/Target CSVs) are required:
      - Group mapping  : translates source group SMTP addresses to destination addresses
      - Member mapping : translates source member SMTP addresses to destination addresses
    Dynamic Distribution Group rows are always skipped (filter-based membership cannot be
    statically imported). M365 Owner rows are applied via the Owners link type.
    Results are written to a timestamped log CSV.

PARAMETERS
    -DistributionGroups         Process Distribution Group rows.
    -MailEnabledSecurityGroups  Process Mail-Enabled Security Group rows.
    -DynamicDistributionGroups  Process DDG rows (they will be logged as Skipped).
    -M365Groups                 Process M365 Group rows.
    -UseCurrentSession          Use the current Exchange Online session.
    -Help                       Display this help text.

    Note: If no group type switch is specified, all types are processed.

EXAMPLE
    Import-FrankensteinGroupMembers -UseCurrentSession
    Import-FrankensteinGroupMembers -UseCurrentSession -M365Groups -DistributionGroups

NOTES
    Author: Eric D. Frank
"@
        return
    }

    Add-Type -AssemblyName System.Windows.Forms

    # --- File picker 1: Membership export ---
    [System.Windows.Forms.MessageBox]::Show(
        "Step 1 of 3: Select the group membership export CSV produced by Get-FrankensteinGroupMember.`n`nExpected columns:`n  - GroupPrimarySmtp  : SMTP address of the source group`n  - GroupType         : DistributionGroup, MailEnabledSecurityGroup, DynamicDistributionGroup, or M365Group`n  - MemberPrimarySmtp : SMTP address of the source member`n  - Role              : Member, Owner (M365 only), or Dynamic (DDG - will be skipped)",
        "Select Group Membership Export File",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null

    $memberDialog        = New-Object System.Windows.Forms.OpenFileDialog
    $memberDialog.Title  = "Select Group Membership Export CSV"
    $memberDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    if ($memberDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "No membership file selected. Exiting." -ForegroundColor Yellow
        return
    }
    $MembershipFile = $memberDialog.FileName

    # --- File picker 2: Group mapping ---
    [System.Windows.Forms.MessageBox]::Show(
        "Step 2 of 3: Select your GROUP mapping CSV.`n`nRequired columns:`n  - Source : original group SMTP address (matches GroupPrimarySmtp in the membership export)`n  - Target : destination group SMTP address in the new environment`n`nDynamic Distribution Group rows are always skipped regardless of mapping.",
        "Select Group Mapping File",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null

    $grpMapDialog        = New-Object System.Windows.Forms.OpenFileDialog
    $grpMapDialog.Title  = "Select Group Mapping CSV"
    $grpMapDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    if ($grpMapDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "No group mapping file selected. Exiting." -ForegroundColor Yellow
        return
    }
    $GroupMappingFile = $grpMapDialog.FileName

    # --- File picker 3: Member mapping ---
    [System.Windows.Forms.MessageBox]::Show(
        "Step 3 of 3: Select your MEMBER mapping CSV.`n`nRequired columns:`n  - Source : original member SMTP address (matches MemberPrimarySmtp in the membership export)`n  - Target : destination member SMTP address in the new environment`n`nMembers may be users, contacts, or mail-enabled groups.",
        "Select Member Mapping File",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null

    $memMapDialog        = New-Object System.Windows.Forms.OpenFileDialog
    $memMapDialog.Title  = "Select Member Mapping CSV"
    $memMapDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    if ($memMapDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "No member mapping file selected. Exiting." -ForegroundColor Yellow
        return
    }
    $MemberMappingFile = $memMapDialog.FileName

    # --- Load and validate files ---
    $MembershipRows = Import-Csv $MembershipFile

    $GroupMapping = @{}
    foreach ($entry in (Import-Csv $GroupMappingFile)) {
        if ($entry.Source -and $entry.Target) {
            $GroupMapping[$entry.Source.Trim().ToLower()] = $entry.Target.Trim()
        }
    }

    $MemberMapping = @{}
    foreach ($entry in (Import-Csv $MemberMappingFile)) {
        if ($entry.Source -and $entry.Target) {
            $MemberMapping[$entry.Source.Trim().ToLower()] = $entry.Target.Trim()
        }
    }

    if ($GroupMapping.Count -eq 0) {
        Write-Host "Group mapping file is empty or missing Source/Target columns. Exiting." -ForegroundColor Red
        return
    }
    if ($MemberMapping.Count -eq 0) {
        Write-Host "Member mapping file is empty or missing Source/Target columns. Exiting." -ForegroundColor Red
        return
    }

    $processAll = (-not $DistributionGroups -and -not $MailEnabledSecurityGroups -and -not $DynamicDistributionGroups -and -not $M365Groups)

    if (-not $UseCurrentSession) { Connect-ExchangeOnline }

    $Log            = [System.Collections.Generic.List[PSCustomObject]]::new()
    $total          = $MembershipRows.Count
    $count          = 0
    $RecipientCache = @{}
    $SeenKeys       = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    function Resolve-TargetRecipient ([string]$identity) {
        $key = $identity.ToLower()
        if (-not $RecipientCache.ContainsKey($key)) {
            $RecipientCache[$key] = Get-Recipient -Identity $identity -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        return $RecipientCache[$key]
    }

    Write-Host "Processing $total membership rows..." -ForegroundColor Cyan

    foreach ($row in $MembershipRows) {
        $count++
        Write-Progress -Activity "Importing Group Membership" `
            -Status "Row $count of $total - $($row.GroupType): $($row.GroupPrimarySmtp)" `
            -PercentComplete ([math]::Round(($count / $total) * 100))

        $groupType = $row.GroupType
        $role      = $row.Role

        if (-not $processAll) {
            if ($groupType -eq 'DistributionGroup'        -and -not $DistributionGroups)       { continue }
            if ($groupType -eq 'MailEnabledSecurityGroup' -and -not $MailEnabledSecurityGroups) { continue }
            if ($groupType -eq 'DynamicDistributionGroup' -and -not $DynamicDistributionGroups) { continue }
            if ($groupType -eq 'M365Group'                -and -not $M365Groups)                { continue }
        }

        $seenKey = "$($row.GroupPrimarySmtp.Trim())|$($row.MemberPrimarySmtp.Trim())|$role"
        if (-not $SeenKeys.Add($seenKey)) {
            $Log.Add([PSCustomObject][ordered]@{
                SourceGroup  = $row.GroupPrimarySmtp
                TargetGroup  = $null
                SourceMember = $row.MemberPrimarySmtp
                TargetMember = $null
                GroupType    = $groupType
                Role         = $role
                Status       = 'Skipped'
                Details      = 'Duplicate entry in membership file'
            })
            continue
        }

        if ($groupType -eq 'DynamicDistributionGroup') {
            $Log.Add([PSCustomObject][ordered]@{
                SourceGroup  = $row.GroupPrimarySmtp
                TargetGroup  = $null
                SourceMember = $row.MemberPrimarySmtp
                TargetMember = $null
                GroupType    = $groupType
                Role         = $role
                Status       = 'Skipped'
                Details      = 'Dynamic Distribution Group membership is filter-based and cannot be statically imported'
            })
            continue
        }

        $groupKey    = $row.GroupPrimarySmtp.Trim().ToLower()
        $memberKey   = $row.MemberPrimarySmtp.Trim().ToLower()
        $groupTarget  = $GroupMapping[$groupKey]
        $memberTarget = $MemberMapping[$memberKey]

        $logEntry = [ordered]@{
            SourceGroup  = $row.GroupPrimarySmtp
            TargetGroup  = $groupTarget
            SourceMember = $row.MemberPrimarySmtp
            TargetMember = $memberTarget
            GroupType    = $groupType
            Role         = $role
            Status       = $null
            Details      = $null
        }

        if (-not $groupTarget -and -not $memberTarget) {
            $logEntry.Status  = 'Skipped'
            $logEntry.Details = 'Group not found in group mapping; Member not found in member mapping'
            $Log.Add([PSCustomObject]$logEntry)
            continue
        }
        if (-not $groupTarget) {
            $logEntry.Status  = 'Skipped'
            $logEntry.Details = 'Group not found in group mapping'
            $Log.Add([PSCustomObject]$logEntry)
            continue
        }
        if (-not $memberTarget) {
            $logEntry.Status  = 'Skipped'
            $logEntry.Details = 'Member not found in member mapping'
            $Log.Add([PSCustomObject]$logEntry)
            continue
        }

        $resolvedGroup  = Resolve-TargetRecipient $groupTarget
        $resolvedMember = Resolve-TargetRecipient $memberTarget

        if (-not $resolvedGroup -and -not $resolvedMember) {
            $logEntry.Status  = 'Failed'
            $logEntry.Details = "Neither target group ('$groupTarget') nor target member ('$memberTarget') found in environment"
            $Log.Add([PSCustomObject]$logEntry)
            continue
        }
        if (-not $resolvedGroup) {
            $logEntry.Status  = 'Failed'
            $logEntry.Details = "Target group '$groupTarget' not found in environment"
            $Log.Add([PSCustomObject]$logEntry)
            continue
        }
        if (-not $resolvedMember) {
            $logEntry.Status  = 'Failed'
            $logEntry.Details = "Target member '$memberTarget' not found in environment"
            $Log.Add([PSCustomObject]$logEntry)
            continue
        }

        $groupPrimary  = $resolvedGroup.PrimarySmtpAddress
        $memberPrimary = $resolvedMember.PrimarySmtpAddress
        $logEntry.TargetGroup  = $groupPrimary
        $logEntry.TargetMember = $memberPrimary

        try {
            switch ($groupType) {
                { $_ -eq 'DistributionGroup' -or $_ -eq 'MailEnabledSecurityGroup' } {
                    Add-DistributionGroupMember -Identity $groupPrimary -Member $memberPrimary `
                        -BypassSecurityGroupManagerCheck -ErrorAction Stop
                }
                'M365Group' {
                    $linkType = if ($role -eq 'Owner') { 'Owners' } else { 'Members' }
                    Add-UnifiedGroupLinks -Identity $groupPrimary -LinkType $linkType `
                        -Links $memberPrimary -ErrorAction Stop
                }
            }
            $logEntry.Status  = 'Success'
            $logEntry.Details = 'Membership applied'
        } catch {
            $errMsg = $_.Exception.Message
            if ($errMsg -match 'already a member|already exists|is already in the group') {
                $logEntry.Status  = 'AlreadyExists'
                $logEntry.Details = "Member Already Exists: $memberPrimary is already a member of $groupPrimary"
            } else {
                $logEntry.Status  = 'Failed'
                $logEntry.Details = $errMsg
            }
        }

        $Log.Add([PSCustomObject]$logEntry)
    }

    Write-Progress -Activity "Importing Group Membership" -Completed

    $LogFile = ".\ImportGroupMembershipLog_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
    $Log | Export-Csv $LogFile -NoTypeInformation -Encoding UTF8

    $success       = ($Log | Where-Object Status -eq 'Success').Count
    $skipped       = ($Log | Where-Object Status -eq 'Skipped').Count
    $alreadyExists = ($Log | Where-Object Status -eq 'AlreadyExists').Count
    $failed        = ($Log | Where-Object Status -eq 'Failed').Count

    Write-Host "`nImport complete." -ForegroundColor Green
    Write-Host "  Success        : $success"       -ForegroundColor Green
    Write-Host "  Skipped        : $skipped"       -ForegroundColor Yellow
    Write-Host "  Member Already Exists : $alreadyExists" -ForegroundColor Yellow
    Write-Host "  Failed         : $failed"        -ForegroundColor $(if ($failed -gt 0) { 'Red' } else { 'Green' })
    Write-Host "  Log            : $LogFile"       -ForegroundColor Cyan
}

function Get-FrankensteinMailboxReport {
    [CmdletBinding()]
    Param (
        [Switch]$UserMailbox,
        [Switch]$SharedMailbox,
        [Switch]$RoomMailbox,
        [Switch]$EquipmentMailbox,
        [Switch]$IncludeDisabled,
        [Switch]$IncludeStatistics,
        [Switch]$ImportCSV,
        [Switch]$UseCurrentSession,
        [String]$OutputPath = '.',
        [Switch]$Help
    )

    if ($Help) {
        Write-Host @"
SYNOPSIS
    Exports a comprehensive mailbox report for all standard mailbox types.

DESCRIPTION
    Collects properties for UserMailbox, SharedMailbox, RoomMailbox, and EquipmentMailbox.
    By default all four types are included. Use the type switches to scope to one or more types.

    -IncludeDisabled applies only to UserMailbox objects. SharedMailbox, RoomMailbox, and
    EquipmentMailbox are always included regardless of account enabled state, since those
    object types are disabled by design in Azure AD.

    Use -ImportCSV to scope the report to a specific list of mailboxes in a CSV file.
    Always outputs a timestamped CSV file.

PARAMETERS
    -UserMailbox        Include User Mailboxes (default: included).
    -SharedMailbox      Include Shared Mailboxes (default: included).
    -RoomMailbox        Include Room Mailboxes (default: included).
    -EquipmentMailbox   Include Equipment Mailboxes (default: included).
    -IncludeDisabled    Also include UserMailboxes whose Azure AD account is disabled.
    -IncludeStatistics  Append TotalItemSize, ItemCount, and LastLogonTime from Get-MailboxStatistics.
                        Note: adds one API call per mailbox; slower on large environments.
    -ImportCSV          Scope the report to a CSV file containing a PrimarySmtpAddress column.
    -UseCurrentSession  Use the current Exchange Online session instead of prompting to connect.
    -OutputPath         Folder path for the output CSV. Defaults to the current working directory.
    -Help               Display this help text.

    Note: If no type switch is specified, all four types are included.

EXAMPLE
    Get-FrankensteinMailboxReport -UseCurrentSession
    Get-FrankensteinMailboxReport -UseCurrentSession -UserMailbox -IncludeDisabled
    Get-FrankensteinMailboxReport -UseCurrentSession -SharedMailbox -RoomMailbox
    Get-FrankensteinMailboxReport -UseCurrentSession -ImportCSV -OutputPath C:\Reports

NOTES
    Author: Eric D. Frank
"@
        return
    }

    # --- Optional ImportCSV file picker ---
    $ImportedMailboxes = $null
    if ($ImportCSV) {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show(
            "Select a CSV containing the mailboxes to include in this report.`n`nRequired column:`n  - PrimarySmtpAddress : the mailbox's primary SMTP address`n`nAny additional columns in the file are ignored.",
            "Select Mailbox Scope CSV",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null

        $scopeDialog        = New-Object System.Windows.Forms.OpenFileDialog
        $scopeDialog.Title  = "Select Mailbox Scope CSV"
        $scopeDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        if ($scopeDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            Write-Host "No scope file selected. Exiting." -ForegroundColor Yellow
            return
        }
        $ImportedMailboxes = Import-Csv $scopeDialog.FileName
        if (-not ($ImportedMailboxes | Get-Member -Name 'PrimarySmtpAddress')) {
            Write-Host "Scope CSV is missing the required PrimarySmtpAddress column. Exiting." -ForegroundColor Red
            return
        }
        Write-Host "Scope file loaded: $($ImportedMailboxes.Count) address(es) to process." -ForegroundColor Cyan
    }

    if (-not $UseCurrentSession) { Connect-ExchangeOnline }

    $typeAll        = (-not $UserMailbox -and -not $SharedMailbox -and -not $RoomMailbox -and -not $EquipmentMailbox)
    $recipientTypes = @(
        if ($typeAll -or $UserMailbox)      { 'UserMailbox' }
        if ($typeAll -or $SharedMailbox)    { 'SharedMailbox' }
        if ($typeAll -or $RoomMailbox)      { 'RoomMailbox' }
        if ($typeAll -or $EquipmentMailbox) { 'EquipmentMailbox' }
    )

    # --- Build mailbox record ---
    function New-MailboxRecord ($mbx) {
        $stats = $null
        if ($IncludeStatistics) {
            $stats = Get-MailboxStatistics -Identity $mbx.ExchangeGuid.ToString() -ErrorAction SilentlyContinue
        }

        $onMicrosoftSmtp = $mbx.EmailAddresses |
            Where-Object { $_.ToString() -like 'smtp:*@*.onmicrosoft.com' } |
            Select-Object -First 1 |
            ForEach-Object { $_.ToString() -replace '^(?i)smtp:', '' }

        [PSCustomObject]@{
            # Identity
            DisplayName                      = $mbx.DisplayName
            PrimarySmtpAddress               = $mbx.PrimarySmtpAddress
            OnMicrosoftAddress               = if ($onMicrosoftSmtp) { $onMicrosoftSmtp } else { '' }
            UserPrincipalName                = $mbx.UserPrincipalName
            SamAccountName                   = $mbx.SamAccountName
            Alias                            = $mbx.Alias
            RecipientTypeDetails             = $mbx.RecipientTypeDetails

            # GUIDs / legacy identifiers
            ExchangeGuid                     = $mbx.ExchangeGuid
            ExternalDirectoryObjectId        = $mbx.ExternalDirectoryObjectId
            LegacyExchangeDN                 = $mbx.LegacyExchangeDN
            DistinguishedName                = $mbx.DistinguishedName

            # All proxy addresses
            EmailAddresses                   = ($mbx.EmailAddresses | ForEach-Object { "$_" }) -join ', '

            # Account status
            AccountDisabled                  = $mbx.AccountDisabled
            IsDirSynced                      = $mbx.IsDirSynced
            WhenCreated                      = $mbx.WhenCreated
            WhenChanged                      = $mbx.WhenChanged
            WhenMailboxCreated               = $mbx.WhenMailboxCreated

            # Forwarding
            ForwardingAddress                = $mbx.ForwardingAddress
            ForwardingSmtpAddress            = $mbx.ForwardingSmtpAddress
            DeliverToMailboxAndForward       = $mbx.DeliverToMailboxAndForward

            # Address list visibility
            HiddenFromAddressListsEnabled    = $mbx.HiddenFromAddressListsEnabled

            # Archive
            ArchiveStatus                    = $mbx.ArchiveStatus
            ArchiveGuid                      = $mbx.ArchiveGuid
            ArchiveName                      = ($mbx.ArchiveName -join ', ')

            # Compliance / legal hold
            LitigationHoldEnabled            = $mbx.LitigationHoldEnabled
            LitigationHoldDuration           = $mbx.LitigationHoldDuration
            LitigationHoldOwner              = $mbx.LitigationHoldOwner
            LitigationHoldDate               = $mbx.LitigationHoldDate
            InPlaceHolds                     = ($mbx.InPlaceHolds -join ', ')
            SingleItemRecoveryEnabled        = $mbx.SingleItemRecoveryEnabled
            RetainDeletedItemsFor            = $mbx.RetainDeletedItemsFor

            # Quota
            ProhibitSendQuota                = $mbx.ProhibitSendQuota
            ProhibitSendReceiveQuota         = $mbx.ProhibitSendReceiveQuota
            IssueWarningQuota                = $mbx.IssueWarningQuota
            UseDatabaseQuotaDefaults         = $mbx.UseDatabaseQuotaDefaults
            MaxSendSize                      = $mbx.MaxSendSize
            MaxReceiveSize                   = $mbx.MaxReceiveSize

            # Policy
            RetentionPolicy                  = $mbx.RetentionPolicy
            RetentionHoldEnabled             = $mbx.RetentionHoldEnabled
            AddressBookPolicy                = $mbx.AddressBookPolicy
            SharingPolicy                    = $mbx.SharingPolicy

            # Audit
            AuditEnabled                     = $mbx.AuditEnabled
            AuditLogAgeLimit                 = $mbx.AuditLogAgeLimit

            # Authentication / protocols (available on Get-Mailbox in EXO)
            ActiveSyncEnabled                = $mbx.ActiveSyncEnabled
            OWAEnabled                       = $mbx.OWAEnabled
            MAPIEnabled                      = $mbx.MAPIEnabled
            EwsEnabled                       = $mbx.EwsEnabled
            PopEnabled                       = $mbx.PopEnabled
            ImapEnabled                      = $mbx.ImapEnabled

            # Statistics (populated only when -IncludeStatistics is set)
            TotalItemSize                    = if ($stats) { $stats.TotalItemSize.Value } else { '' }
            ItemCount                        = if ($stats) { $stats.ItemCount }           else { '' }
            LastLogonTime                    = if ($stats) { $stats.LastLogonTime }       else { '' }
        }
    }

    $Results = [System.Collections.Generic.List[PSCustomObject]]::new()

    if ($ImportCSV) {
        $total = $ImportedMailboxes.Count
        $count = 0
        foreach ($row in $ImportedMailboxes) {
            $count++
            Write-Progress -Activity "Generating Mailbox Report" `
                -Status "Looking up $($row.PrimarySmtpAddress) ($count of $total)" `
                -PercentComplete ([math]::Round(($count / $total) * 100))

            $mbx = Get-Mailbox -Identity $row.PrimarySmtpAddress -ErrorAction SilentlyContinue
            if (-not $mbx) {
                Write-Warning "Mailbox not found: $($row.PrimarySmtpAddress)"
                continue
            }
            if (-not $typeAll -and $mbx.RecipientTypeDetails -notin $recipientTypes) { continue }
            if (-not $IncludeDisabled -and $mbx.AccountDisabled -and $mbx.RecipientTypeDetails -eq 'UserMailbox') { continue }
            $Results.Add((New-MailboxRecord $mbx))
        }
    } else {
        Write-Host "Retrieving mailboxes..." -ForegroundColor Cyan
        $Mailboxes = @(Get-Mailbox -RecipientTypeDetails $recipientTypes -ResultSize Unlimited)
        if (-not $IncludeDisabled) {
            $Mailboxes = @($Mailboxes | Where-Object { $_.RecipientTypeDetails -ne 'UserMailbox' -or -not $_.AccountDisabled })
        }
        $total = $Mailboxes.Count
        $count = 0
        Write-Host "Processing $total mailboxes..." -ForegroundColor Cyan
        foreach ($mbx in $Mailboxes) {
            $count++
            Write-Progress -Activity "Generating Mailbox Report" `
                -Status "$($mbx.DisplayName) ($count of $total)" `
                -PercentComplete ([math]::Round(($count / $total) * 100))
            $Results.Add((New-MailboxRecord $mbx))
        }
    }

    Write-Progress -Activity "Generating Mailbox Report" -Completed

    $FileName = "MailboxReport_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
    $FullPath = Join-Path $OutputPath $FileName
    $Results | Export-Csv $FullPath -NoTypeInformation -Encoding UTF8

    Write-Host "`nReport complete: $($Results.Count) mailbox(es) exported." -ForegroundColor Green
    Write-Host "File: $FullPath" -ForegroundColor Cyan
}

function Get-FrankensteinAliasReport {
    [CmdletBinding()]
    Param (
        [Switch]$Mailboxes,
        [Switch]$Groups,
        [Switch]$MailUsers,
        [Switch]$MailContacts,
        [Switch]$LegacyDN,
        [Switch]$SIP,
        [Switch]$IncludeOnMSAddress,
        [Switch]$ImportCSV,
        [Switch]$ImportMappingCSV,
        [Switch]$UseCurrentSession,
        [String]$OutputPath = '.',
        [Switch]$CSV,
        [Switch]$Help
    )

    if ($Help) {
        Write-Host @"
SYNOPSIS
    Exports a flat address map with one row per address per recipient.

DESCRIPTION
    For each recipient, one row is written per email address. This produces a lookup table
    where every SMTP alias, primary address, and optionally LegacyExchangeDN or SIP address
    maps back to the object's PrimarySmtpAddress and type.

    Default (no type switches): all mailboxes (User, Shared, Room, Equipment).
    Use type switches to include other recipient types or mix and match.

PARAMETERS
    -Mailboxes      Include UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox.
    -Groups         Include Distribution Groups, Mail-Enabled Security Groups, and M365 Groups.
    -MailUsers      Include Mail Users and Guest Mail Users.
    -MailContacts   Include Mail Contacts.
    -LegacyDN           Append an X500 row per recipient using the LegacyExchangeDN value.
    -SIP                Include SIP proxy addresses as additional rows.
    -IncludeOnMSAddress By default, onmicrosoft.com addresses are excluded since they are
                        never stamped on objects in a target tenant. Set this switch to include them.
    -ImportCSV          Scope to a CSV file with a PrimarySmtpAddress column.
    -ImportMappingCSV   Scope to a mapping CSV with Source and Target columns.
                        Source can be UPN, PrimarySMTP, alias, or any Exchange-resolvable identity.
                        Adds a TargetAddress column to the export populated with the Target value.
                        Use this to drive a ForEach loop in the target tenant to stamp aliases.
                        Mutually exclusive with -ImportCSV (mapping takes precedence if both set).
    -UseCurrentSession  Use the current Exchange Online session.
    -OutputPath         Folder path for the output CSV. Defaults to the current directory.
    -CSV                Export results to a timestamped CSV file.
    -Help               Display this help text.

    Note: If no type switch is specified, all mailboxes are included.

OUTPUT COLUMNS
    DisplayName, PrimarySmtpAddress, AddressValue, AddressType, RecipientType[, TargetAddress]

    AddressType values:
      SMTP   - primary SMTP address
      smtp   - secondary/alias SMTP address
      X500   - LegacyExchangeDN (requires -LegacyDN)
      SIP    - SIP address (requires -SIP)

    TargetAddress is only present when -ImportMappingCSV is used. It contains the
    Target value from the mapping file -- typically the object's identity in the target tenant.

EXAMPLE
    Get-FrankensteinAliasReport -UseCurrentSession -CSV
    Get-FrankensteinAliasReport -UseCurrentSession -Mailboxes -Groups -LegacyDN -CSV
    Get-FrankensteinAliasReport -UseCurrentSession -ImportCSV -LegacyDN -CSV
    Get-FrankensteinAliasReport -UseCurrentSession -ImportMappingCSV -LegacyDN -CSV

NOTES
    Author: Eric D. Frank
"@
        return
    }

    Add-Type -AssemblyName System.Windows.Forms

    # --- Optional ImportCSV file picker ---
    $ScopedAddresses = $null
    if ($ImportCSV -and -not $ImportMappingCSV) {
        [System.Windows.Forms.MessageBox]::Show(
            "Select a CSV containing the recipients to include in this report.`n`nRequired column:`n  - PrimarySmtpAddress : the recipient's primary SMTP address`n`nAny additional columns in the file are ignored.",
            "Select Recipient Scope CSV",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null

        $scopeDialog        = New-Object System.Windows.Forms.OpenFileDialog
        $scopeDialog.Title  = "Select Recipient Scope CSV"
        $scopeDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        if ($scopeDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            Write-Host "No scope file selected. Exiting." -ForegroundColor Yellow
            return
        }
        $imported = Import-Csv $scopeDialog.FileName
        if (-not ($imported | Get-Member -Name 'PrimarySmtpAddress')) {
            Write-Host "Scope CSV is missing the required PrimarySmtpAddress column. Exiting." -ForegroundColor Red
            return
        }
        $ScopedAddresses = @($imported | ForEach-Object { $_.PrimarySmtpAddress } | Where-Object { $_ })
        Write-Host "Scope file loaded: $($ScopedAddresses.Count) address(es) to process." -ForegroundColor Cyan
    }

    # --- Optional ImportMappingCSV file picker ---
    $MappingEntries  = $null
    $ResolvedMapping = @{}   # PrimarySmtpAddress.ToLower() -> Target value
    if ($ImportMappingCSV) {
        [System.Windows.Forms.MessageBox]::Show(
            "Select a mapping CSV to scope this report and populate the TargetAddress column.`n`nRequired columns:`n  - Source : UPN, PrimarySMTP, alias, or any Exchange-resolvable identity`n  - Target : identity of the matching object in the target tenant`n`nEach Source is resolved to its PrimarySmtpAddress. The Target value is stamped`non every address row for that recipient so you can drive a ForEach loop in the target tenant.",
            "Select Source-to-Target Mapping CSV",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null

        $mapDialog        = New-Object System.Windows.Forms.OpenFileDialog
        $mapDialog.Title  = "Select Source-to-Target Mapping CSV"
        $mapDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        if ($mapDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            Write-Host "No mapping file selected. Exiting." -ForegroundColor Yellow
            return
        }
        $mapImported = Import-Csv $mapDialog.FileName
        if (-not ($mapImported | Get-Member -Name 'Source') -or -not ($mapImported | Get-Member -Name 'Target')) {
            Write-Host "Mapping CSV is missing required Source and/or Target columns. Exiting." -ForegroundColor Red
            return
        }
        $MappingEntries = @($mapImported | Where-Object { $_.Source -and $_.Target })
        Write-Host "Mapping file loaded: $($MappingEntries.Count) entries to process." -ForegroundColor Cyan
    }

    if (-not $UseCurrentSession) { Connect-ExchangeOnline }

    # Build recipient type filter
    $typeAll    = (-not $Mailboxes -and -not $Groups -and -not $MailUsers -and -not $MailContacts)
    $typeFilter = @(
        if ($typeAll -or $Mailboxes)    { 'UserMailbox'; 'SharedMailbox'; 'RoomMailbox'; 'EquipmentMailbox' }
        if ($typeAll -or $Groups)       { 'MailUniversalDistributionGroup'; 'MailUniversalSecurityGroup'; 'GroupMailbox' }
        if ($typeAll -or $MailUsers)    { 'MailUser'; 'GuestMailUser' }
        if ($typeAll -or $MailContacts) { 'MailContact' }
    )

    $Results = [System.Collections.Generic.List[PSCustomObject]]::new()

    function Add-AddressRows ($recip) {
        $primarySmtp = $recip.PrimarySmtpAddress
        $displayName = $recip.DisplayName
        $recipType   = $recip.RecipientTypeDetails
        $targetAddr  = if ($ImportMappingCSV) { $ResolvedMapping[$primarySmtp.ToLower()] } else { $null }

        foreach ($addr in $recip.EmailAddresses) {
            $addrStr  = $addr.ToString()
            $colonIdx = $addrStr.IndexOf(':')
            if ($colonIdx -lt 0) { continue }
            $prefix = $addrStr.Substring(0, $colonIdx)
            $value  = $addrStr.Substring($colonIdx + 1)

            $include = ($prefix -ieq 'smtp') -or ($SIP -and $prefix -ieq 'sip')
            if (-not $include) { continue }
            if (-not $IncludeOnMSAddress -and $value -like '*onmicrosoft.com') { continue }

            $row = [ordered]@{
                DisplayName        = $displayName
                PrimarySmtpAddress = $primarySmtp
                AddressValue       = $value
                AddressType        = $prefix
                RecipientType      = $recipType
            }
            if ($ImportMappingCSV) { $row['TargetAddress'] = $targetAddr }
            $Results.Add([PSCustomObject]$row)
        }

        if ($LegacyDN -and $recip.LegacyExchangeDN) {
            $row = [ordered]@{
                DisplayName        = $displayName
                PrimarySmtpAddress = $primarySmtp
                AddressValue       = $recip.LegacyExchangeDN
                AddressType        = 'X500'
                RecipientType      = $recipType
            }
            if ($ImportMappingCSV) { $row['TargetAddress'] = $targetAddr }
            $Results.Add([PSCustomObject]$row)
        }
    }

    # Resolves a Get-Recipient stub to the type-specific object that carries LegacyExchangeDN
    function Get-FullRecipientObject ($recip) {
        $id = $recip.PrimarySmtpAddress
        switch ($recip.RecipientTypeDetails) {
            { $_ -in @('UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox') }       { Get-Mailbox           -Identity $id -ErrorAction SilentlyContinue; break }
            { $_ -in @('MailUniversalDistributionGroup','MailUniversalSecurityGroup') }        { Get-DistributionGroup -Identity $id -ErrorAction SilentlyContinue; break }
            'GroupMailbox'                                                                      { Get-UnifiedGroup      -Identity $id -ErrorAction SilentlyContinue; break }
            { $_ -in @('MailUser','GuestMailUser') }                                           { Get-MailUser          -Identity $id -ErrorAction SilentlyContinue; break }
            'MailContact'                                                                       { Get-MailContact       -Identity $id -ErrorAction SilentlyContinue; break }
            default                                                                             { $recip }
        }
    }

    if ($ImportMappingCSV) {
        $total = $MappingEntries.Count
        $count = 0
        foreach ($entry in $MappingEntries) {
            $count++
            $sourceId  = $entry.Source.Trim()
            $targetVal = $entry.Target.Trim()
            Write-Progress -Activity "Generating Alias Report" `
                -Status "Resolving $sourceId ($count of $total)" `
                -PercentComplete ([math]::Round(($count / $total) * 100))

            $recip = Get-Recipient -Identity $sourceId -ErrorAction SilentlyContinue
            if (-not $recip) {
                Write-Warning "Recipient not found: $sourceId"
                continue
            }
            if (-not $typeAll -and $recip.RecipientTypeDetails -notin $typeFilter) {
                Write-Warning "Recipient '$sourceId' is type '$($recip.RecipientTypeDetails)' which is excluded by the current type switches."
                continue
            }
            if ($LegacyDN) { $recip = Get-FullRecipientObject $recip }
            $ResolvedMapping[$recip.PrimarySmtpAddress.ToLower()] = $targetVal
            Add-AddressRows $recip
        }
    } elseif ($ImportCSV) {
        $total = $ScopedAddresses.Count
        $count = 0
        foreach ($smtp in $ScopedAddresses) {
            $count++
            Write-Progress -Activity "Generating Alias Report" `
                -Status "Looking up $smtp ($count of $total)" `
                -PercentComplete ([math]::Round(($count / $total) * 100))

            $recip = Get-Recipient -Identity $smtp -ErrorAction SilentlyContinue
            if (-not $recip) {
                Write-Warning "Recipient not found: $smtp"
                continue
            }
            if (-not $typeAll -and $recip.RecipientTypeDetails -notin $typeFilter) {
                Write-Warning "Recipient '$smtp' is type '$($recip.RecipientTypeDetails)' which is excluded by the current type switches."
                continue
            }
            if ($LegacyDN) { $recip = Get-FullRecipientObject $recip }
            Add-AddressRows $recip
        }
    } else {
        # Bulk path: use type-specific cmdlets directly -- they carry LegacyExchangeDN natively.
        # Get-Recipient is avoided here because EXO's REST module returns LegacyExchangeDN as null.
        $AllRecipients = [System.Collections.Generic.List[object]]::new()

        if ($typeAll -or $Mailboxes) {
            Write-Host "Fetching mailboxes..." -ForegroundColor Cyan
            @(Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails @('UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox')) |
                ForEach-Object { $AllRecipients.Add($_) }
        }
        if ($typeAll -or $Groups) {
            Write-Host "Fetching distribution groups..." -ForegroundColor Cyan
            @(Get-DistributionGroup -ResultSize Unlimited) | ForEach-Object { $AllRecipients.Add($_) }
            Write-Host "Fetching M365 groups..." -ForegroundColor Cyan
            @(Get-UnifiedGroup -ResultSize Unlimited)      | ForEach-Object { $AllRecipients.Add($_) }
        }
        if ($typeAll -or $MailUsers) {
            Write-Host "Fetching mail users..." -ForegroundColor Cyan
            @(Get-MailUser -ResultSize Unlimited) | ForEach-Object { $AllRecipients.Add($_) }
        }
        if ($typeAll -or $MailContacts) {
            Write-Host "Fetching mail contacts..." -ForegroundColor Cyan
            @(Get-MailContact -ResultSize Unlimited) | ForEach-Object { $AllRecipients.Add($_) }
        }

        $total = $AllRecipients.Count
        Write-Host "Processing $total recipient(s)..." -ForegroundColor Cyan
        $count = 0
        foreach ($recip in $AllRecipients) {
            $count++
            Write-Progress -Activity "Generating Alias Report" `
                -Status "$($recip.DisplayName) ($count of $total)" `
                -PercentComplete ([math]::Round(($count / $total) * 100))
            Add-AddressRows $recip
        }
    }

    Write-Progress -Activity "Generating Alias Report" -Completed
    Write-Host "`nTotal address rows: $($Results.Count) across $total recipient(s)." -ForegroundColor Cyan

    if ($CSV) {
        $FileName = "AliasReport_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
        $FullPath = Join-Path $OutputPath $FileName
        $Results | Export-Csv $FullPath -NoTypeInformation -Encoding UTF8
        Write-Host "Export complete: $FullPath" -ForegroundColor Green
    } else {
        $Results
    }
}

function Invoke-FrankensteinPermissionMigrator {
    [CmdletBinding()]
    Param ([Switch]$Help)

    if ($Help) {
        Write-Host @"
SYNOPSIS
    GUI-driven delegate permission migration for Exchange Online tenant-to-tenant cutovers.

DESCRIPTION
    Launches a WinForms interface that drives the full permission migration workflow:
      1. Load a Source-to-Target identity mapping CSV (mailboxes, delegates, DLs all in one file)
      2. Connect to the source tenant and pull FullAccess, SendAs, and SendOnBehalf permissions live
      3. Connect to the target tenant and apply permissions using the mapping
      4. Export a timestamped log CSV showing Success, Skipped, AlreadyExists, and Failed results

    DL delegates can be resolved to their individual members at read-time, with optional
    recursive expansion of nested DLs. If a DL maps cleanly through the mapping file it is
    applied as-is; expansion is only used when -ResolveDLs is chosen in the GUI.

    The mapping CSV must have Source and Target columns. All identities -- mailboxes, delegates,
    and groups -- should appear in the same file. Any identity not in the mapping is logged as Skipped.

USAGE
    Invoke-FrankensteinPermissionMigrator
    Invoke-FrankensteinPermissionMigrator -Help

NOTES
    Author: Eric D. Frank
    Requires ExchangeOnlineManagement module.
"@
        return
    }

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "ExchangeOnlineManagement module not found." -ForegroundColor Red
        Write-Host "Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser" -ForegroundColor Yellow
        return
    }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $script:PMSourceConnected = $false
    $script:PMTargetConnected = $false
    $script:PMSourceData      = [System.Collections.Generic.List[PSCustomObject]]::new()
    $script:PMMappingTable    = @{}   # source.ToLower() -> target
    $script:PMResultLog       = [System.Collections.Generic.List[PSCustomObject]]::new()

    #region ---- Helpers ----

    function Write-PMLog {
        Param ([string]$Message, [System.Drawing.Color]$Color = [System.Drawing.Color]::Silver)
        $rtbLog.SelectionStart  = $rtbLog.TextLength
        $rtbLog.SelectionLength = 0
        $rtbLog.SelectionColor  = $Color
        $rtbLog.AppendText("$(Get-Date -Format 'HH:mm:ss')  $Message`n")
        $rtbLog.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }

    function Set-PMStatusLabel {
        Param ($Label, [string]$Text, [System.Drawing.Color]$Color)
        $Label.Text      = $Text
        $Label.ForeColor = $Color
        [System.Windows.Forms.Application]::DoEvents()
    }

    function Collect-PMSourcePermissions {
        Param ([string[]]$MailboxList)
        $script:PMSourceData.Clear()
        $total   = $MailboxList.Count
        $idx     = 0
        $dlTypes = @('MailUniversalDistributionGroup','MailUniversalSecurityGroup')

        function Find-PMMapping ([object]$recipient) {
            # Returns the matching mapping table key (lowercase) or $null
            # Checks primary SMTP first, then all proxy addresses
            $primary = $recipient.PrimarySmtpAddress.ToLower()
            if ($script:PMMappingTable.ContainsKey($primary)) { return $primary }
            if ($recipient.EmailAddresses) {
                foreach ($addr in $recipient.EmailAddresses) {
                    if ($addr.ToString() -match '^smtp:(.+)$') {
                        $proxyKey = $Matches[1].ToLower()
                        if ($script:PMMappingTable.ContainsKey($proxyKey)) { return $proxyKey }
                    }
                }
            }
            return $null
        }

        foreach ($mbxSmtp in $MailboxList) {
            $idx++
            Write-PMLog "  [$idx/$total] $mbxSmtp" ([System.Drawing.Color]::DimGray)

            function Add-PMRecord ([string]$delegate, [string]$type, [string]$expandedFrom) {
                $script:PMSourceData.Add([PSCustomObject]@{
                    Mailbox        = $mbxSmtp
                    Delegate       = $delegate
                    PermissionType = $type
                    ExpandedFrom   = $expandedFrom
                })
            }

            function Resolve-PMDelegate ([string]$rawId, [string]$permType) {
                $r = Get-Recipient -Identity $rawId -ErrorAction SilentlyContinue
                if (-not $r) { return }
                $smtp       = $r.PrimarySmtpAddress
                $mappingKey = Find-PMMapping $r
                $isDL       = $r.RecipientTypeDetails -in $dlTypes
                $dlIsMapped = $null -ne $mappingKey
                if ($isDL -and $chkResolveDLs.Checked -and -not $dlIsMapped) {
                    # DL has no mapping entry -- expand members and grant individually
                    $members = @(Expand-FrankensteinDLMembers -Identity $smtp -Recurse:($chkNestedDLs.Checked))
                    if ($members.Count) {
                        Write-PMLog "    DL '$smtp' (unmapped, expanding) -> $($members.Count) member(s)" ([System.Drawing.Color]::DarkCyan)
                        foreach ($m in $members) {
                            $mKey = Find-PMMapping $m
                            $keyToStore = if ($mKey) { $mKey } else { $m.PrimarySmtpAddress }
                            Add-PMRecord $keyToStore $permType $smtp
                        }
                    }
                } else {
                    # Non-DL, DL expansion off, or DL is mapped -- record as direct delegate
                    $keyToStore = if ($mappingKey) { $mappingKey } else { $smtp }
                    Add-PMRecord $keyToStore $permType ''
                }
            }

            if ($chkFullAccess.Checked) {
                @(Get-MailboxPermission -Identity $mbxSmtp -ErrorAction SilentlyContinue) |
                    Where-Object { -not $_.IsInherited -and $_.User -notlike 'NT AUTHORITY\SELF' } |
                    ForEach-Object { Resolve-PMDelegate $_.User 'FullAccess' }
            }
            if ($chkSendAs.Checked) {
                @(Get-RecipientPermission -Identity $mbxSmtp -ErrorAction SilentlyContinue) |
                    Where-Object { $_.Trustee -ne 'NT AUTHORITY\SELF' } |
                    ForEach-Object { Resolve-PMDelegate $_.Trustee 'SendAs' }
            }
            if ($chkSendOnBehalf.Checked) {
                $mbxObj = Get-Mailbox -Identity $mbxSmtp -ErrorAction SilentlyContinue
                if ($mbxObj) {
                    foreach ($g in $mbxObj.GrantSendOnBehalfTo) { Resolve-PMDelegate $g 'SendOnBehalf' }
                }
            }
        }
        Write-PMLog "Source read complete: $($script:PMSourceData.Count) permission record(s) collected." ([System.Drawing.Color]::LimeGreen)
    }

    function Apply-PMTargetPermissions {
        $script:PMResultLog.Clear()
        $sobByMailbox = @{}

        foreach ($perm in $script:PMSourceData) {
            $mbxTarget = $script:PMMappingTable[$perm.Mailbox.ToLower()]
            $delTarget = $script:PMMappingTable[$perm.Delegate.ToLower()]

            $log = [ordered]@{
                SourceMailbox  = $perm.Mailbox
                TargetMailbox  = if ($mbxTarget) { $mbxTarget } else { 'NOT MAPPED' }
                SourceDelegate = $perm.Delegate
                TargetDelegate = if ($delTarget) { $delTarget } else { 'NOT MAPPED' }
                PermissionType = $perm.PermissionType
                ExpandedFrom   = $perm.ExpandedFrom
                Status         = ''
                Details        = ''
                Timestamp      = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            }

            if (-not $mbxTarget) {
                $log.Status = 'Skipped'; $log.Details = 'Mailbox not in mapping'
                Write-PMLog "  SKIP  [$($perm.PermissionType)] Mailbox not mapped: $($perm.Mailbox)" ([System.Drawing.Color]::DarkGoldenrod)
                $script:PMResultLog.Add([PSCustomObject]$log); continue
            }
            if (-not $delTarget) {
                $log.Status = 'Skipped'; $log.Details = 'Delegate not in mapping'
                Write-PMLog "  SKIP  [$($perm.PermissionType)] Delegate not mapped: $($perm.Delegate)" ([System.Drawing.Color]::DarkGoldenrod)
                $script:PMResultLog.Add([PSCustomObject]$log); continue
            }

            try {
                switch ($perm.PermissionType) {
                    'FullAccess' {
                        Add-MailboxPermission -Identity $mbxTarget -User $delTarget -AccessRights FullAccess `
                            -InheritanceType All -AutoMapping $false -ErrorAction Stop | Out-Null
                        $log.Status = 'Success'; $log.Details = 'FullAccess granted'
                        Write-PMLog "  OK    [FullAccess] $delTarget -> $mbxTarget" ([System.Drawing.Color]::LimeGreen)
                    }
                    'SendAs' {
                        Add-RecipientPermission -Identity $mbxTarget -Trustee $delTarget `
                            -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                        $log.Status = 'Success'; $log.Details = 'SendAs granted'
                        Write-PMLog "  OK    [SendAs] $delTarget -> $mbxTarget" ([System.Drawing.Color]::LimeGreen)
                    }
                    'SendOnBehalf' {
                        if (-not $sobByMailbox[$mbxTarget]) {
                            $sobByMailbox[$mbxTarget] = [System.Collections.Generic.List[string]]::new()
                        }
                        $sobByMailbox[$mbxTarget].Add($delTarget)
                        $log.Status = 'Queued'; $log.Details = 'SendOnBehalf queued for bulk apply'
                        Write-PMLog "  QUE   [SendOnBehalf] $delTarget -> $mbxTarget" ([System.Drawing.Color]::DimGray)
                    }
                }
            } catch {
                $err = $_.Exception.Message
                if ($err -match 'already|exists|present') {
                    $log.Status = 'AlreadyExists'; $log.Details = 'Permission already present'
                    Write-PMLog "  DUP   [$($perm.PermissionType)] Already exists: $delTarget -> $mbxTarget" ([System.Drawing.Color]::SteelBlue)
                } else {
                    $log.Status = 'Failed'; $log.Details = $err
                    Write-PMLog "  ERR   [$($perm.PermissionType)] $err" ([System.Drawing.Color]::Tomato)
                }
            }
            $script:PMResultLog.Add([PSCustomObject]$log)
        }

        foreach ($mbxTarget in $sobByMailbox.Keys) {
            try {
                Set-Mailbox -Identity $mbxTarget -GrantSendOnBehalfTo @{Add = $sobByMailbox[$mbxTarget].ToArray()} -ErrorAction Stop
                Write-PMLog "  OK    [SendOnBehalf] Applied to $mbxTarget ($($sobByMailbox[$mbxTarget].Count) delegate(s))" ([System.Drawing.Color]::LimeGreen)
                $script:PMResultLog | Where-Object { $_.TargetMailbox -eq $mbxTarget -and $_.PermissionType -eq 'SendOnBehalf' -and $_.Status -eq 'Queued' } |
                    ForEach-Object { $_.Status = 'Success'; $_.Details = 'SendOnBehalf granted' }
            } catch {
                $err = $_.Exception.Message
                Write-PMLog "  ERR   [SendOnBehalf] Failed on $mbxTarget`: $err" ([System.Drawing.Color]::Tomato)
                $script:PMResultLog | Where-Object { $_.TargetMailbox -eq $mbxTarget -and $_.PermissionType -eq 'SendOnBehalf' -and $_.Status -eq 'Queued' } |
                    ForEach-Object { $_.Status = 'Failed'; $_.Details = $err }
            }
        }
    }

    #endregion

    #region ---- Build Form ----

    $form                 = New-Object System.Windows.Forms.Form
    $form.Text            = "Frankenstein - Permission Migrator"
    $form.ClientSize      = New-Object System.Drawing.Size(700, 824)
    $form.StartPosition   = 'CenterScreen'
    $form.FormBorderStyle = 'FixedSingle'
    $form.MaximizeBox     = $false
    $form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.BackColor       = [System.Drawing.Color]::FromArgb(245, 245, 248)

    $lblHeader           = New-Object System.Windows.Forms.Label
    $lblHeader.Text      = "Frankenstein - Permission Migrator"
    $lblHeader.Font      = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
    $lblHeader.ForeColor = [System.Drawing.Color]::FromArgb(28, 28, 72)
    $lblHeader.Location  = New-Object System.Drawing.Point(14, 12)
    $lblHeader.Size      = New-Object System.Drawing.Size(660, 28)
    $form.Controls.Add($lblHeader)

    # --- 2. Identity Mapping ---
    $grpMap           = New-Object System.Windows.Forms.GroupBox
    $grpMap.Text      = "2. Identity Mapping"
    $grpMap.Location  = New-Object System.Drawing.Point(12, 154)
    $grpMap.Size      = New-Object System.Drawing.Size(676, 214)
    $form.Controls.Add($grpMap)

    # CSV option
    $radCsvMapping          = New-Object System.Windows.Forms.RadioButton
    $radCsvMapping.Text     = "Load CSV mapping file  (Source + Target columns -- mailboxes, delegates, groups)"
    $radCsvMapping.Checked  = $true
    $radCsvMapping.Location = New-Object System.Drawing.Point(10, 20)
    $radCsvMapping.Size     = New-Object System.Drawing.Size(570, 20)
    $grpMap.Controls.Add($radCsvMapping)

    $btnBrowseMap           = New-Object System.Windows.Forms.Button
    $btnBrowseMap.Text      = "Browse..."
    $btnBrowseMap.Location  = New-Object System.Drawing.Point(582, 16)
    $btnBrowseMap.Size      = New-Object System.Drawing.Size(82, 26)
    $grpMap.Controls.Add($btnBrowseMap)

    $lblMapPath             = New-Object System.Windows.Forms.Label
    $lblMapPath.Text        = "No file loaded"
    $lblMapPath.ForeColor   = [System.Drawing.Color]::Gray
    $lblMapPath.Location    = New-Object System.Drawing.Point(10, 42)
    $lblMapPath.Size        = New-Object System.Drawing.Size(654, 16)
    $grpMap.Controls.Add($lblMapPath)

    $ttMap = New-Object System.Windows.Forms.ToolTip
    $ttMap.AutoPopDelay = 12000
    $ttMap.InitialDelay = 350
    $ttMap.ReshowDelay  = 200
    $ttMap.ShowAlways   = $true
    $csvTip = @"
CSV Requirements
------------------------------------------------
Column headers (exact, case-insensitive):
  Source  -- the identity in the SOURCE tenant
  Target  -- the corresponding identity in the TARGET tenant

Accepted identity formats:
  * Primary SMTP address  (user@source.com)
  * User Principal Name   (user@source.com)
  * Any value Exchange PowerShell can resolve

Supported object types per row:
  * User mailbox / Shared mailbox / Room / Equipment
  * Mail-enabled security group or distribution group
  * Mail user  *  Mail contact

Important rules:
  * A permission is only migrated when BOTH the mailbox owner
    AND the delegate have a row in this file.
  * DL delegates are automatically expanded to their members at
    collection time -- include individual members, not the DL itself.
  * onmicrosoft.com addresses are fine as Target values.
"@
    $ttMap.SetToolTip($radCsvMapping, $csvTip)
    $ttMap.SetToolTip($btnBrowseMap,  $csvTip)

    # Manual option
    $radManualEntry         = New-Object System.Windows.Forms.RadioButton
    $radManualEntry.Text    = "Enter source/target pairs manually  (no CSV needed -- useful for single-user migrations)"
    $radManualEntry.Location = New-Object System.Drawing.Point(10, 62)
    $radManualEntry.Size    = New-Object System.Drawing.Size(754, 20)
    $grpMap.Controls.Add($radManualEntry)

    $txtManualSrc                 = New-Object System.Windows.Forms.TextBox
    $txtManualSrc.Location        = New-Object System.Drawing.Point(10, 86)
    $txtManualSrc.Size            = New-Object System.Drawing.Size(334, 22)
    $txtManualSrc.PlaceholderText = "source@domain.com"
    $txtManualSrc.Enabled         = $false
    $grpMap.Controls.Add($txtManualSrc)

    $lblArrow                = New-Object System.Windows.Forms.Label
    $lblArrow.Text           = "->"
    $lblArrow.TextAlign      = 'MiddleCenter'
    $lblArrow.Location       = New-Object System.Drawing.Point(350, 88)
    $lblArrow.Size           = New-Object System.Drawing.Size(16, 18)
    $grpMap.Controls.Add($lblArrow)

    $txtManualTgt                 = New-Object System.Windows.Forms.TextBox
    $txtManualTgt.Location        = New-Object System.Drawing.Point(372, 86)
    $txtManualTgt.Size            = New-Object System.Drawing.Size(312, 22)
    $txtManualTgt.PlaceholderText = "target@domain.com"
    $txtManualTgt.Enabled         = $false
    $grpMap.Controls.Add($txtManualTgt)

    $btnAddPair             = New-Object System.Windows.Forms.Button
    $btnAddPair.Text        = "Add"
    $btnAddPair.Location    = New-Object System.Drawing.Point(692, 84)
    $btnAddPair.Size        = New-Object System.Drawing.Size(62, 26)
    $btnAddPair.Enabled     = $false
    $grpMap.Controls.Add($btnAddPair)

    $lvPairs                = New-Object System.Windows.Forms.ListView
    $lvPairs.View           = 'Details'
    $lvPairs.FullRowSelect  = $true
    $lvPairs.GridLines      = $true
    $lvPairs.Location       = New-Object System.Drawing.Point(10, 116)
    $lvPairs.Size           = New-Object System.Drawing.Size(754, 62)
    $lvPairs.Enabled        = $false
    $lvPairs.Columns.Add("Source SMTP", 370) | Out-Null
    $lvPairs.Columns.Add("Target SMTP", 370) | Out-Null
    $grpMap.Controls.Add($lvPairs)

    $btnRemovePair          = New-Object System.Windows.Forms.Button
    $btnRemovePair.Text     = "Remove Selected"
    $btnRemovePair.Location = New-Object System.Drawing.Point(10, 184)
    $btnRemovePair.Size     = New-Object System.Drawing.Size(126, 24)
    $btnRemovePair.Enabled  = $false
    $grpMap.Controls.Add($btnRemovePair)

    $lblMapCount            = New-Object System.Windows.Forms.Label
    $lblMapCount.Text       = ""
    $lblMapCount.ForeColor  = [System.Drawing.Color]::DarkGreen
    $lblMapCount.Location   = New-Object System.Drawing.Point(144, 186)
    $lblMapCount.Size       = New-Object System.Drawing.Size(600, 16)
    $grpMap.Controls.Add($lblMapCount)

    # --- 1. Connections ---
    $grpConn           = New-Object System.Windows.Forms.GroupBox
    $grpConn.Text      = "1. Connections  (connect to both tenants first to validate credentials)"
    $grpConn.Location  = New-Object System.Drawing.Point(12, 46)
    $grpConn.Size      = New-Object System.Drawing.Size(676, 100)
    $form.Controls.Add($grpConn)

    $btnConnSrc              = New-Object System.Windows.Forms.Button
    $btnConnSrc.Text         = "Connect Source"
    $btnConnSrc.Location     = New-Object System.Drawing.Point(10, 24)
    $btnConnSrc.Size         = New-Object System.Drawing.Size(148, 30)
    $grpConn.Controls.Add($btnConnSrc)

    $lblSrcStatus            = New-Object System.Windows.Forms.Label
    $lblSrcStatus.Text       = "Not connected"
    $lblSrcStatus.ForeColor  = [System.Drawing.Color]::Gray
    $lblSrcStatus.Location   = New-Object System.Drawing.Point(168, 30)
    $lblSrcStatus.Size       = New-Object System.Drawing.Size(496, 18)
    $grpConn.Controls.Add($lblSrcStatus)

    $btnConnTgt              = New-Object System.Windows.Forms.Button
    $btnConnTgt.Text         = "Connect Target"
    $btnConnTgt.Location     = New-Object System.Drawing.Point(10, 62)
    $btnConnTgt.Size         = New-Object System.Drawing.Size(148, 30)
    $grpConn.Controls.Add($btnConnTgt)

    $lblTgtStatus            = New-Object System.Windows.Forms.Label
    $lblTgtStatus.Text       = "Not connected"
    $lblTgtStatus.ForeColor  = [System.Drawing.Color]::Gray
    $lblTgtStatus.Location   = New-Object System.Drawing.Point(168, 68)
    $lblTgtStatus.Size       = New-Object System.Drawing.Size(496, 18)
    $grpConn.Controls.Add($lblTgtStatus)

    # --- 3. Scope ---
    $grpScope           = New-Object System.Windows.Forms.GroupBox
    $grpScope.Text      = "3. Scope"
    $grpScope.Location  = New-Object System.Drawing.Point(12, 376)
    $grpScope.Size      = New-Object System.Drawing.Size(676, 72)
    $form.Controls.Add($grpScope)

    $radAll                  = New-Object System.Windows.Forms.RadioButton
    $radAll.Text             = "All mapped identities"
    $radAll.Checked          = $true
    $radAll.Location         = New-Object System.Drawing.Point(10, 22)
    $radAll.Size             = New-Object System.Drawing.Size(200, 22)
    $grpScope.Controls.Add($radAll)

    $radSingle               = New-Object System.Windows.Forms.RadioButton
    $radSingle.Text          = "Single mailbox:"
    $radSingle.Location      = New-Object System.Drawing.Point(10, 48)
    $radSingle.Size          = New-Object System.Drawing.Size(120, 22)
    $grpScope.Controls.Add($radSingle)

    $txtSingleMbx                 = New-Object System.Windows.Forms.TextBox
    $txtSingleMbx.Location        = New-Object System.Drawing.Point(136, 46)
    $txtSingleMbx.Size            = New-Object System.Drawing.Size(524, 22)
    $txtSingleMbx.Enabled         = $false
    $txtSingleMbx.PlaceholderText = "source@domain.com"
    $grpScope.Controls.Add($txtSingleMbx)

    # --- 4. Permission Types ---
    $grpPerms           = New-Object System.Windows.Forms.GroupBox
    $grpPerms.Text      = "4. Permission Types"
    $grpPerms.Location  = New-Object System.Drawing.Point(12, 456)
    $grpPerms.Size      = New-Object System.Drawing.Size(676, 56)
    $form.Controls.Add($grpPerms)

    $chkFullAccess            = New-Object System.Windows.Forms.CheckBox
    $chkFullAccess.Text       = "Full Access"
    $chkFullAccess.Checked    = $true
    $chkFullAccess.Location   = New-Object System.Drawing.Point(10, 22)
    $chkFullAccess.Size       = New-Object System.Drawing.Size(120, 22)
    $grpPerms.Controls.Add($chkFullAccess)

    $chkSendAs                = New-Object System.Windows.Forms.CheckBox
    $chkSendAs.Text           = "Send As"
    $chkSendAs.Checked        = $true
    $chkSendAs.Location       = New-Object System.Drawing.Point(220, 22)
    $chkSendAs.Size           = New-Object System.Drawing.Size(120, 22)
    $grpPerms.Controls.Add($chkSendAs)

    $chkSendOnBehalf          = New-Object System.Windows.Forms.CheckBox
    $chkSendOnBehalf.Text     = "Send on Behalf"
    $chkSendOnBehalf.Checked  = $true
    $chkSendOnBehalf.Location = New-Object System.Drawing.Point(430, 22)
    $chkSendOnBehalf.Size     = New-Object System.Drawing.Size(160, 22)
    $grpPerms.Controls.Add($chkSendOnBehalf)

    # --- 5. Options ---
    $grpOpts           = New-Object System.Windows.Forms.GroupBox
    $grpOpts.Text      = "5. Options"
    $grpOpts.Location  = New-Object System.Drawing.Point(12, 520)
    $grpOpts.Size      = New-Object System.Drawing.Size(676, 90)
    $form.Controls.Add($grpOpts)

    $chkResolveDLs            = New-Object System.Windows.Forms.CheckBox
    $chkResolveDLs.Text       = "Resolve DL delegates to members"
    $chkResolveDLs.Checked    = $true
    $chkResolveDLs.Location   = New-Object System.Drawing.Point(10, 22)
    $chkResolveDLs.Size       = New-Object System.Drawing.Size(240, 22)
    $grpOpts.Controls.Add($chkResolveDLs)

    $chkNestedDLs             = New-Object System.Windows.Forms.CheckBox
    $chkNestedDLs.Text        = "Recurse nested DLs"
    $chkNestedDLs.Checked     = $true
    $chkNestedDLs.Location    = New-Object System.Drawing.Point(264, 22)
    $chkNestedDLs.Size        = New-Object System.Drawing.Size(180, 22)
    $grpOpts.Controls.Add($chkNestedDLs)

    $lblOutPath               = New-Object System.Windows.Forms.Label
    $lblOutPath.Text          = "Log output path:"
    $lblOutPath.Location      = New-Object System.Drawing.Point(10, 58)
    $lblOutPath.Size          = New-Object System.Drawing.Size(112, 22)
    $grpOpts.Controls.Add($lblOutPath)

    $txtOutPath               = New-Object System.Windows.Forms.TextBox
    $txtOutPath.Text          = (Get-Location).Path
    $txtOutPath.Location      = New-Object System.Drawing.Point(126, 56)
    $txtOutPath.Size          = New-Object System.Drawing.Size(434, 22)
    $grpOpts.Controls.Add($txtOutPath)

    $btnBrowseOut             = New-Object System.Windows.Forms.Button
    $btnBrowseOut.Text        = "Browse..."
    $btnBrowseOut.Location    = New-Object System.Drawing.Point(570, 54)
    $btnBrowseOut.Size        = New-Object System.Drawing.Size(94, 26)
    $grpOpts.Controls.Add($btnBrowseOut)

    # --- Action Buttons ---
    $btnPreview               = New-Object System.Windows.Forms.Button
    $btnPreview.Text          = "Preview"
    $btnPreview.Location      = New-Object System.Drawing.Point(12, 620)
    $btnPreview.Size          = New-Object System.Drawing.Size(130, 34)
    $btnPreview.Enabled       = $false
    $form.Controls.Add($btnPreview)

    $btnRun                   = New-Object System.Windows.Forms.Button
    $btnRun.Text              = "Run Migration"
    $btnRun.Location          = New-Object System.Drawing.Point(260, 620)
    $btnRun.Size              = New-Object System.Drawing.Size(180, 34)
    $btnRun.BackColor         = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnRun.ForeColor         = [System.Drawing.Color]::White
    $btnRun.FlatStyle         = 'Flat'
    $btnRun.Enabled           = $false
    $form.Controls.Add($btnRun)

    $btnExportLog             = New-Object System.Windows.Forms.Button
    $btnExportLog.Text        = "Export Log"
    $btnExportLog.Location    = New-Object System.Drawing.Point(558, 620)
    $btnExportLog.Size        = New-Object System.Drawing.Size(130, 34)
    $btnExportLog.Enabled     = $false
    $form.Controls.Add($btnExportLog)

    # --- Status Log ---
    $grpLog           = New-Object System.Windows.Forms.GroupBox
    $grpLog.Text      = "Status Log"
    $grpLog.Location  = New-Object System.Drawing.Point(12, 664)
    $grpLog.Size      = New-Object System.Drawing.Size(676, 152)
    $form.Controls.Add($grpLog)

    $rtbLog                   = New-Object System.Windows.Forms.RichTextBox
    $rtbLog.Location          = New-Object System.Drawing.Point(8, 18)
    $rtbLog.Size              = New-Object System.Drawing.Size(660, 126)
    $rtbLog.ReadOnly          = $true
    $rtbLog.BackColor         = [System.Drawing.Color]::FromArgb(18, 18, 28)
    $rtbLog.ForeColor         = [System.Drawing.Color]::Silver
    $rtbLog.Font              = New-Object System.Drawing.Font("Consolas", 8.5)
    $rtbLog.ScrollBars        = 'Vertical'
    $rtbLog.BorderStyle       = 'None'
    $grpLog.Controls.Add($rtbLog)

    #endregion

    #region ---- Event Handlers ----

    # Mapping mode: CSV vs Manual
    $radCsvMapping.Add_CheckedChanged({
        $btnBrowseMap.Enabled  = $radCsvMapping.Checked
        $lblMapPath.Enabled    = $radCsvMapping.Checked
        $txtManualSrc.Enabled  = -not $radCsvMapping.Checked
        $txtManualTgt.Enabled  = -not $radCsvMapping.Checked
        $btnAddPair.Enabled    = -not $radCsvMapping.Checked
        $lvPairs.Enabled       = -not $radCsvMapping.Checked
        $btnRemovePair.Enabled = -not $radCsvMapping.Checked
        if ($radCsvMapping.Checked) {
            $lvPairs.Items.Clear()
            $script:PMMappingTable = @{}
            $lblMapCount.Text = ""
        }
    })
    $radManualEntry.Add_CheckedChanged({
        $btnBrowseMap.Enabled  = -not $radManualEntry.Checked
        $lblMapPath.Enabled    = -not $radManualEntry.Checked
        $txtManualSrc.Enabled  = $radManualEntry.Checked
        $txtManualTgt.Enabled  = $radManualEntry.Checked
        $btnAddPair.Enabled    = $radManualEntry.Checked
        $lvPairs.Enabled       = $radManualEntry.Checked
        $btnRemovePair.Enabled = $radManualEntry.Checked
        if ($radManualEntry.Checked) {
            $script:PMMappingTable = @{}
            $lblMapPath.Text      = "No file loaded"
            $lblMapPath.ForeColor = [System.Drawing.Color]::Gray
            $lblMapCount.Text     = "0 pair(s) entered"
        }
    })

    # Add a manual pair
    $btnAddPair.Add_Click({
        $src = $txtManualSrc.Text.Trim()
        $tgt = $txtManualTgt.Text.Trim()
        if (-not $src -or -not $tgt) {
            [System.Windows.Forms.MessageBox]::Show("Enter both a source and target SMTP address.", "Missing Address", 'OK', 'Warning') | Out-Null
            return
        }
        $item = New-Object System.Windows.Forms.ListViewItem($src)
        $item.SubItems.Add($tgt) | Out-Null
        $lvPairs.Items.Add($item) | Out-Null
        $script:PMMappingTable[$src.ToLower()] = $tgt
        $txtManualSrc.Text = ''
        $txtManualTgt.Text = ''
        $lblMapCount.Text  = "$($lvPairs.Items.Count) pair(s) entered"
        Write-PMLog "Pair added: $src -> $tgt" ([System.Drawing.Color]::LimeGreen)
        $txtManualSrc.Focus() | Out-Null
    })

    # Enter key in target field triggers Add
    $txtManualTgt.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Return) { $btnAddPair.PerformClick() }
    })

    # Remove selected pair(s)
    $btnRemovePair.Add_Click({
        $selected = @($lvPairs.SelectedItems)
        if (-not $selected.Count) { return }
        foreach ($item in $selected) { $lvPairs.Items.Remove($item) }
        $script:PMMappingTable = @{}
        foreach ($item in $lvPairs.Items) {
            $script:PMMappingTable[$item.Text.ToLower()] = $item.SubItems[1].Text
        }
        $lblMapCount.Text = "$($lvPairs.Items.Count) pair(s) entered"
        Write-PMLog "Pair(s) removed. $($lvPairs.Items.Count) remaining." ([System.Drawing.Color]::DarkGoldenrod)
    })

    # Scope radio buttons
    $radSingle.Add_CheckedChanged({ $txtSingleMbx.Enabled = $radSingle.Checked })
    $radAll.Add_CheckedChanged({ $txtSingleMbx.Enabled = -not $radAll.Checked })

    # DL checkbox dependency
    $chkResolveDLs.Add_CheckedChanged({ $chkNestedDLs.Enabled = $chkResolveDLs.Checked })

    # Browse mapping CSV
    $btnBrowseMap.Add_Click({
        $dlg        = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title  = "Select Source-to-Target Identity Mapping CSV"
        $dlg.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $rows = Import-Csv $dlg.FileName
        if (-not ($rows | Get-Member -Name 'Source') -or -not ($rows | Get-Member -Name 'Target')) {
            [System.Windows.Forms.MessageBox]::Show(
                "CSV is missing required columns.`n`nRequired: Source, Target",
                "Invalid CSV", 'OK', 'Error') | Out-Null
            return
        }
        $script:PMMappingTable = @{}
        foreach ($row in $rows) {
            if ($row.Source -and $row.Target) {
                $script:PMMappingTable[$row.Source.Trim().ToLower()] = $row.Target.Trim()
            }
        }
        $lblMapPath.Text      = $dlg.FileName
        $lblMapPath.ForeColor = [System.Drawing.Color]::DarkGreen
        $lblMapCount.Text     = "$($script:PMMappingTable.Count) identities loaded"
        Write-PMLog "Mapping CSV loaded: $($script:PMMappingTable.Count) entries" ([System.Drawing.Color]::LimeGreen)
    })

    # Browse output folder
    $btnBrowseOut.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Select log output folder"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtOutPath.Text = $dlg.SelectedPath
        }
    })

    # Connect Source -- auth validation only, no data collection
    $btnConnSrc.Add_Click({
        Set-PMStatusLabel $lblSrcStatus "Connecting..." ([System.Drawing.Color]::DarkGoldenrod)
        $form.UseWaitCursor = $true
        $btnConnSrc.Enabled = $false
        try {
            Write-PMLog "Connecting to source tenant..." ([System.Drawing.Color]::Silver)
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            $script:PMSourceConnected = $true
            Set-PMStatusLabel $lblSrcStatus "Connected" ([System.Drawing.Color]::DarkGreen)
            Write-PMLog "Source connected. Load your mapping, then use Preview or Run Migration." ([System.Drawing.Color]::LimeGreen)
            $btnPreview.Enabled = $true
            if ($script:PMTargetConnected) { $btnRun.Enabled = $true }
        } catch {
            Set-PMStatusLabel $lblSrcStatus "Connection failed" ([System.Drawing.Color]::Tomato)
            Write-PMLog "ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Tomato)
        } finally {
            $form.UseWaitCursor = $false
            $btnConnSrc.Enabled = $true
        }
    })

    # Connect Target -- auth validation only
    $btnConnTgt.Add_Click({
        Set-PMStatusLabel $lblTgtStatus "Connecting..." ([System.Drawing.Color]::DarkGoldenrod)
        $form.UseWaitCursor = $true
        $btnConnTgt.Enabled = $false
        try {
            Write-PMLog "Connecting to target tenant..." ([System.Drawing.Color]::Silver)
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            $script:PMTargetConnected = $true
            Set-PMStatusLabel $lblTgtStatus "Connected" ([System.Drawing.Color]::DarkGreen)
            Write-PMLog "Target connected." ([System.Drawing.Color]::LimeGreen)
            if ($script:PMSourceConnected) { $btnRun.Enabled = $true }
        } catch {
            Set-PMStatusLabel $lblTgtStatus "Connection failed" ([System.Drawing.Color]::Tomato)
            Write-PMLog "ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Tomato)
        } finally {
            $form.UseWaitCursor = $false
            $btnConnTgt.Enabled = $true
        }
    })

    # Preview -- re-connects to source, collects, then shows what would be applied
    $btnPreview.Add_Click({
        if (-not $script:PMMappingTable.Count) {
            $msg = if ($radCsvMapping.Checked) { "Load a mapping CSV first." } else { "Add at least one source/target pair first." }
            [System.Windows.Forms.MessageBox]::Show($msg, "No Mapping", 'OK', 'Warning') | Out-Null
            return
        }
        if (-not $chkFullAccess.Checked -and -not $chkSendAs.Checked -and -not $chkSendOnBehalf.Checked) {
            [System.Windows.Forms.MessageBox]::Show("Select at least one permission type.", "No Permission Types", 'OK', 'Warning') | Out-Null
            return
        }
        $mailboxList = if ($radSingle.Checked) {
            $addr = $txtSingleMbx.Text.Trim()
            if (-not $addr) {
                [System.Windows.Forms.MessageBox]::Show("Enter a source mailbox SMTP address.", "No Address", 'OK', 'Warning') | Out-Null
                return
            }
            @($addr)
        } else { @($script:PMMappingTable.Keys) }

        $btnPreview.Enabled = $false
        $form.UseWaitCursor = $true
        try {
            Write-PMLog "Re-connecting to source for permission read..." ([System.Drawing.Color]::DimGray)
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            Set-PMStatusLabel $lblSrcStatus "Connected - reading $($mailboxList.Count) mailbox(es)..." ([System.Drawing.Color]::DarkOrange)

            Collect-PMSourcePermissions -MailboxList $mailboxList
            Set-PMStatusLabel $lblSrcStatus "Connected - $($script:PMSourceData.Count) record(s) read" ([System.Drawing.Color]::DarkGreen)

            Write-PMLog "---- PREVIEW ($($script:PMSourceData.Count) records) ----" ([System.Drawing.Color]::CornflowerBlue)
            $wouldApply = 0; $wouldSkip = 0
            foreach ($perm in $script:PMSourceData) {
                $mbxT = $script:PMMappingTable[$perm.Mailbox.ToLower()]
                $delT = $script:PMMappingTable[$perm.Delegate.ToLower()]
                if ($mbxT -and $delT) {
                    Write-PMLog "  APPLY [$($perm.PermissionType)] $delT -> $mbxT" ([System.Drawing.Color]::LimeGreen)
                    $wouldApply++
                } else {
                    $miss = if (-not $mbxT) { "mailbox '$($perm.Mailbox)'" } else { "delegate '$($perm.Delegate)'" }
                    Write-PMLog "  SKIP  [$($perm.PermissionType)] No mapping for $miss" ([System.Drawing.Color]::DarkGoldenrod)
                    $wouldSkip++
                }
            }
            Write-PMLog "---- Preview: $wouldApply would apply, $wouldSkip would skip ----" ([System.Drawing.Color]::CornflowerBlue)
        } catch {
            Write-PMLog "ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Tomato)
        } finally {
            $btnPreview.Enabled = $true
            $form.UseWaitCursor = $false
        }
    })

    # Run Migration -- re-connects to source to collect, then to target to apply
    $btnRun.Add_Click({
        if (-not $script:PMMappingTable.Count) {
            $msg = if ($radCsvMapping.Checked) { "Load a mapping CSV first." } else { "Add at least one source/target pair first." }
            [System.Windows.Forms.MessageBox]::Show($msg, "No Mapping", 'OK', 'Warning') | Out-Null
            return
        }
        if (-not $chkFullAccess.Checked -and -not $chkSendAs.Checked -and -not $chkSendOnBehalf.Checked) {
            [System.Windows.Forms.MessageBox]::Show("Select at least one permission type.", "No Permission Types", 'OK', 'Warning') | Out-Null
            return
        }
        $mailboxList = if ($radSingle.Checked) {
            $addr = $txtSingleMbx.Text.Trim()
            if (-not $addr) {
                [System.Windows.Forms.MessageBox]::Show("Enter a source mailbox SMTP address.", "No Address", 'OK', 'Warning') | Out-Null
                return
            }
            @($addr)
        } else { @($script:PMMappingTable.Keys) }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "This will connect to both tenants, read permissions for $($mailboxList.Count) source mailbox(es), and apply them to the target tenant.`n`nProceed?",
            "Confirm Migration",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $btnRun.Enabled = $false
        $form.UseWaitCursor = $true
        try {
            Write-PMLog "Re-connecting to source..." ([System.Drawing.Color]::DimGray)
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            Set-PMStatusLabel $lblSrcStatus "Connected - reading $($mailboxList.Count) mailbox(es)..." ([System.Drawing.Color]::DarkOrange)

            Collect-PMSourcePermissions -MailboxList $mailboxList
            Set-PMStatusLabel $lblSrcStatus "Connected - $($script:PMSourceData.Count) record(s) read" ([System.Drawing.Color]::DarkGreen)

            Write-PMLog "Re-connecting to target..." ([System.Drawing.Color]::DimGray)
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            Set-PMStatusLabel $lblTgtStatus "Connected - applying permissions..." ([System.Drawing.Color]::DarkOrange)

            Write-PMLog "---- MIGRATION STARTED ($($script:PMSourceData.Count) records) ----" ([System.Drawing.Color]::CornflowerBlue)
            Apply-PMTargetPermissions

            $cSuccess = @($script:PMResultLog | Where-Object { $_.Status -eq 'Success' }).Count
            $cSkipped = @($script:PMResultLog | Where-Object { $_.Status -eq 'Skipped' }).Count
            $cDupe    = @($script:PMResultLog | Where-Object { $_.Status -eq 'AlreadyExists' }).Count
            $cFailed  = @($script:PMResultLog | Where-Object { $_.Status -eq 'Failed' }).Count
            Write-PMLog "---- COMPLETE: $cSuccess applied | $cSkipped skipped | $cDupe already existed | $cFailed failed ----" ([System.Drawing.Color]::CornflowerBlue)
            Set-PMStatusLabel $lblTgtStatus "Connected - migration complete" ([System.Drawing.Color]::DarkGreen)
            $btnExportLog.Enabled = $true
        } catch {
            Write-PMLog "FATAL: $($_.Exception.Message)" ([System.Drawing.Color]::Tomato)
        } finally {
            $btnRun.Enabled = $true
            $form.UseWaitCursor = $false
        }
    })

    # Export Log
    $btnExportLog.Add_Click({
        $outPath = $txtOutPath.Text.Trim()
        if (-not (Test-Path $outPath)) {
            Write-PMLog "Output path not found: $outPath" ([System.Drawing.Color]::Tomato)
            return
        }
        $file = Join-Path $outPath "PermissionMigrationLog_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
        try {
            $script:PMResultLog | Export-Csv $file -NoTypeInformation -Encoding UTF8
            Write-PMLog "Log exported: $file" ([System.Drawing.Color]::LimeGreen)
        } catch {
            Write-PMLog "Export failed: $($_.Exception.Message)" ([System.Drawing.Color]::Tomato)
        }
    })

    #endregion

    Write-PMLog "Ready. Connect to both tenants first to validate credentials, then load your mapping." ([System.Drawing.Color]::Silver)
    Write-PMLog "Note: Preview and Run Migration re-authenticate to source (to read) then target (to apply)." ([System.Drawing.Color]::DimGray)

    $form.ShowDialog() | Out-Null
    $form.Dispose()
}

function Invoke-FrankensteinDLMigrator {
    <#
    .SYNOPSIS
        GUI tool for migrating distribution group membership from M365 source to M365 or on-prem Exchange target.
    #>
    [CmdletBinding()]
    Param()

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    #region ---- State ----
    $script:DLSourceConnected    = $false
    $script:DLTargetConnected    = $false
    $script:DLOnPremSession      = $null
    $script:DLSourceUpn          = ''
    $script:DLSourceOrg          = ''
    $script:DLSourceConnectionId = $null
    $script:DLTargetUpn          = ''
    $script:DLTargetOrg          = ''
    $script:DLTargetConnectionId = $null
    $script:DLMappingTable       = @{}
    $script:DLProxyIndex         = @{}   # proxySmtp.ToLower() -> targetSmtp (all aliases of every mapped source)
    $script:DLMappingCsvPath     = ''
    $script:DLGroupMeta          = @{}   # source.ToLower() -> {DisplayName, Alias, Type, DLType}
    $script:DLSourceData         = [System.Collections.Generic.List[PSCustomObject]]::new()
    $script:DLResultLog          = [System.Collections.Generic.List[PSCustomObject]]::new()
    $script:DLLiveLogPath        = ''
    #endregion

    #region ---- Inner Functions ----

    function Write-DLLog ([string]$msg, [System.Drawing.Color]$color) {
        $ts = (Get-Date).ToString('HH:mm:ss')
        $rtbLog.SelectionStart  = $rtbLog.TextLength
        $rtbLog.SelectionLength = 0
        $rtbLog.SelectionColor  = $color
        $rtbLog.AppendText("[$ts] $msg`n")
        $rtbLog.ScrollToCaret()
        $rtbLog.Refresh()
        if ($script:DLLiveLogPath) {
            try { "[$ts] $msg" | Out-File -Append -FilePath $script:DLLiveLogPath -Encoding UTF8 } catch {}
        }
    }

    function Update-DLProgress ([int]$done, [int]$total) {
        if ($total -le 0) { $progressBar.Value = 0; return }
        $pct = [math]::Min(100, [math]::Round(($done / $total) * 100))
        $progressBar.Value = $pct
        $progressBar.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
    }

    function Set-DLStatusLabel ([System.Windows.Forms.Label]$lbl, [string]$text, [System.Drawing.Color]$color) {
        $lbl.Text      = $text
        $lbl.ForeColor = $color
    }

    function Find-DLMapping ([object]$recipient) {
        # Check primary SMTP first, then all proxy addresses (handles onmicrosoft.com as CSV key)
        $primary = $recipient.PrimarySmtpAddress.ToLower()
        if ($script:DLMappingTable.ContainsKey($primary)) { return $script:DLMappingTable[$primary] }
        if ($recipient.EmailAddresses) {
            foreach ($addr in $recipient.EmailAddresses) {
                if ($addr.ToString() -match '^smtp:(.+)$') {
                    $proxyKey = $Matches[1].ToLower()
                    if ($script:DLMappingTable.ContainsKey($proxyKey)) { return $script:DLMappingTable[$proxyKey] }
                }
            }
        }
        return $null
    }

    function Add-TargetDLMember ([string]$TargetGroup, [string]$TargetMember, [string]$GroupType) {
        if ($GroupType -eq 'GroupMailbox') {
            Add-UnifiedGroupLinks -Identity $TargetGroup -LinkType Members -Links $TargetMember -ErrorAction Stop
        } elseif ($radOnPrem.Checked -and $script:DLOnPremSession) {
            Invoke-Command -Session $script:DLOnPremSession -ScriptBlock {
                param($g, $m)
                Add-DistributionGroupMember -Identity $g -Member $m -BypassSecurityGroupManagerCheck -ErrorAction Stop
            } -ArgumentList $TargetGroup, $TargetMember
        } else {
            Add-DistributionGroupMember -Identity $TargetGroup -Member $TargetMember -ErrorAction Stop
        }
    }

    function New-TargetDLGroup ([string]$Name, [string]$DisplayName, [string]$Alias, [string]$PrimarySmtp, [string]$DLType, [string]$GroupType, [string]$OU) {
        if ($GroupType -eq 'GroupMailbox') {
            New-UnifiedGroup -DisplayName $DisplayName -Alias $Alias -PrimarySmtpAddress $PrimarySmtp -SuppressWarmupMessage:$chkSuppressWelcome.Checked -ErrorAction Stop | Out-Null
        } elseif ($radOnPrem.Checked -and $script:DLOnPremSession) {
            Invoke-Command -Session $script:DLOnPremSession -ScriptBlock {
                param($n, $dn, $a, $s, $t, $o)
                $p = @{ Name = $n; DisplayName = $dn; Alias = $a; PrimarySmtpAddress = $s; Type = $t }
                if ($o) { $p['OrganizationalUnit'] = $o }
                New-DistributionGroup @p -ErrorAction Stop
            } -ArgumentList $Name, $DisplayName, $Alias, $PrimarySmtp, $DLType, $OU
        } else {
            $p = @{ Name = $Name; DisplayName = $DisplayName; Alias = $Alias; PrimarySmtpAddress = $PrimarySmtp; Type = $DLType }
            if ($OU) { $p['OrganizationalUnit'] = $OU }
            New-DistributionGroup @p -ErrorAction Stop | Out-Null
        }
    }

    function Set-TargetDLProperties ([string]$TargetIdentity, [hashtable]$Meta, [string]$GroupType, [string]$DefaultOwner) {
        function Map-ListToTarget ([string[]]$sourceList) {
            $seen = @{}
            @($sourceList | ForEach-Object {
                $key = $_.ToLower()
                $t   = $script:DLMappingTable[$key]
                if (-not $t) { $t = $script:DLProxyIndex[$key] }
                if ($t -and -not $seen[$t]) { $seen[$t] = $true; $t }
            })
        }

        $copyManagedBy    = $chkPropManagedBy.Checked
        $copyHidden       = $chkPropHiddenGAL.Checked
        $copyReqAuth      = $chkPropRequireSenderAuth.Checked
        $copySendOnBehalf = $chkPropGrantSendOnBehalf.Checked
        $copyModeration   = $chkPropModeration.Checked
        $copyAcceptFrom   = $chkPropAcceptFrom.Checked
        $copyRejectFrom   = $chkPropRejectFrom.Checked

        $tgtManagedBy    = if ($copyManagedBy)    { Map-ListToTarget $Meta.ManagedBy }       else { @() }
        if ($copyManagedBy -and $Meta.ManagedBy.Count) {
            $unmapped = @($Meta.ManagedBy | Where-Object {
                $_ -and
                -not $script:DLMappingTable.ContainsKey($_.ToLower()) -and
                -not $script:DLProxyIndex.ContainsKey($_.ToLower())
            })
            foreach ($u in $unmapped) {
                Write-DLLog "    WARNING: Owner '$u' not in mapping CSV -- skipped (add a user row for this address to copy it)" ([System.Drawing.Color]::DarkGoldenrod)
            }
        }
        if ($copyManagedBy -and $DefaultOwner -and $tgtManagedBy -notcontains $DefaultOwner) { $tgtManagedBy += $DefaultOwner }
        $tgtModBy        = if ($copyModeration)   { Map-ListToTarget $Meta.ModeratedBy }     else { @() }
        $tgtSendOnBehalf = if ($copySendOnBehalf) { Map-ListToTarget $Meta.GrantSendOnBehalf } else { @() }
        $tgtAcceptFrom   = if ($copyAcceptFrom)   { Map-ListToTarget $Meta.AcceptFrom }      else { @() }
        $tgtRejectFrom   = if ($copyRejectFrom)   { Map-ListToTarget $Meta.RejectFrom }      else { @() }

        if ($GroupType -eq 'GroupMailbox') {
            $p = @{ Identity = $TargetIdentity }
            if ($copyHidden)                      { $p['HiddenFromAddressListsEnabled'] = $Meta.HiddenFromGAL }
            if ($tgtManagedBy.Count)              { $p['Owners']                        = $tgtManagedBy }
            if ($tgtSendOnBehalf.Count)           { $p['GrantSendOnBehalfTo']           = $tgtSendOnBehalf }
            Set-UnifiedGroup @p -ErrorAction Stop
        } elseif ($radOnPrem.Checked -and $script:DLOnPremSession) {
            $applyHidden   = $copyHidden
            $applyModEn    = $copyModeration
            $applyReqAuth  = $copyReqAuth
            $hiddenVal     = $Meta.HiddenFromGAL
            $modEnabledVal = $Meta.ModerationEnabled
            $reqAuthVal    = $Meta.RequireSenderAuth
            Invoke-Command -Session $script:DLOnPremSession -ScriptBlock {
                param($id, $managedBy, $modBy, $sob, $accept, $reject,
                      $applyHidden, $hiddenVal, $applyModEn, $modEnabledVal, $applyReqAuth, $reqAuthVal)
                $p = @{ Identity = $id }
                if ($applyHidden)  { $p['HiddenFromAddressListsEnabled']      = $hiddenVal }
                if ($applyModEn)   { $p['ModerationEnabled']                  = $modEnabledVal }
                if ($applyReqAuth) { $p['RequireSenderAuthenticationEnabled'] = $reqAuthVal }
                if ($managedBy.Count) { $p['ManagedBy']                              = $managedBy }
                if ($modBy.Count)     { $p['ModeratedBy']                            = $modBy }
                if ($sob.Count)       { $p['GrantSendOnBehalfTo']                    = $sob }
                if ($accept.Count)    { $p['AcceptMessagesOnlyFromSendersOrMembers'] = $accept }
                if ($reject.Count)    { $p['RejectMessagesFromSendersOrMembers']     = $reject }
                Set-DistributionGroup @p -ErrorAction Stop
            } -ArgumentList $TargetIdentity, $tgtManagedBy, $tgtModBy, $tgtSendOnBehalf, $tgtAcceptFrom, $tgtRejectFrom,
                            $applyHidden, $hiddenVal, $applyModEn, $modEnabledVal, $applyReqAuth, $reqAuthVal
        } else {
            $p = @{ Identity = $TargetIdentity }
            if ($copyHidden)        { $p['HiddenFromAddressListsEnabled']      = $Meta.HiddenFromGAL }
            if ($copyModeration)    { $p['ModerationEnabled']                  = $Meta.ModerationEnabled }
            if ($copyReqAuth)       { $p['RequireSenderAuthenticationEnabled'] = $Meta.RequireSenderAuth }
            if ($tgtManagedBy.Count)    { $p['ManagedBy']                              = $tgtManagedBy }
            if ($tgtModBy.Count)        { $p['ModeratedBy']                            = $tgtModBy }
            if ($tgtSendOnBehalf.Count) { $p['GrantSendOnBehalfTo']                    = $tgtSendOnBehalf }
            if ($tgtAcceptFrom.Count)   { $p['AcceptMessagesOnlyFromSendersOrMembers'] = $tgtAcceptFrom }
            if ($tgtRejectFrom.Count)   { $p['RejectMessagesFromSendersOrMembers']     = $tgtRejectFrom }
            Set-DistributionGroup @p -ErrorAction Stop
        }
    }

    function Collect-DLSourceData {
        $script:DLSourceData.Clear()
        $script:DLGroupMeta.Clear()
        $dlTypes = @('MailUniversalDistributionGroup','MailUniversalSecurityGroup','GroupMailbox')

        function Resolve-ToSmtp ([string]$identity) {
            if (-not $identity) { return $null }
            $r = Get-Recipient -Identity $identity -ErrorAction SilentlyContinue
            if ($r) { return $r.PrimarySmtpAddress } else { return $null }
        }

        function Get-SmtpList ([object[]]$identities) {
            @($identities | Where-Object { $_ } | ForEach-Object { Resolve-ToSmtp "$_" } | Where-Object { $_ })
        }

        Write-DLLog "Scanning mapping for source groups..." ([System.Drawing.Color]::DimGray)
        $sourceGroups = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($srcSmtp in @($script:DLMappingTable.Keys)) {
            $grp = Get-DistributionGroup -Identity $srcSmtp -ErrorAction SilentlyContinue
            if ($grp) {
                $typeDetail = $grp.RecipientTypeDetails
                $typeAllowed = ($typeDetail -eq 'MailUniversalDistributionGroup' -and $chkTypeDL.Checked) -or
                               ($typeDetail -eq 'MailUniversalSecurityGroup'     -and $chkTypeMailSec.Checked)
                if (-not $typeAllowed) {
                    Write-DLLog "  SKIP $srcSmtp ($typeDetail) -- type filter" ([System.Drawing.Color]::DimGray)
                    continue
                }
                $script:DLGroupMeta[$srcSmtp.ToLower()] = @{
                    Name              = $grp.Name
                    DisplayName       = $grp.DisplayName
                    Alias             = $grp.Alias
                    Type              = $typeDetail
                    DLType            = if ($typeDetail -eq 'MailUniversalSecurityGroup') { 'Security' } else { 'Distribution' }
                    RequireSenderAuth = $grp.RequireSenderAuthenticationEnabled
                    HiddenFromGAL     = $grp.HiddenFromAddressListsEnabled
                    ModerationEnabled = $grp.ModerationEnabled
                    ManagedBy         = Get-SmtpList $grp.ManagedBy
                    ModeratedBy       = Get-SmtpList $grp.ModeratedBy
                    GrantSendOnBehalf = Get-SmtpList $grp.GrantSendOnBehalfTo
                    AcceptFrom        = Get-SmtpList $grp.AcceptMessagesOnlyFromSendersOrMembers
                    RejectFrom        = Get-SmtpList $grp.RejectMessagesFromSendersOrMembers
                }
                $sourceGroups.Add([PSCustomObject]@{ Smtp = $srcSmtp; Type = $typeDetail })
                continue
            }
            $ug = Get-UnifiedGroup -Identity $srcSmtp -ErrorAction SilentlyContinue
            if ($ug) {
                if (-not $chkTypeM365.Checked) {
                    Write-DLLog "  SKIP $srcSmtp (GroupMailbox) -- type filter" ([System.Drawing.Color]::DimGray)
                    continue
                }
                $ugOwners = @(Get-UnifiedGroupLinks -Identity $ug.Identity -LinkType Owners -ErrorAction SilentlyContinue |
                    ForEach-Object { $_.PrimarySmtpAddress } | Where-Object { $_ })
                $script:DLGroupMeta[$srcSmtp.ToLower()] = @{
                    Name              = $ug.DisplayName
                    DisplayName       = $ug.DisplayName
                    Alias             = $ug.Alias
                    Type              = 'GroupMailbox'
                    DLType            = 'Distribution'
                    RequireSenderAuth = $ug.RequireSenderAuthenticationEnabled
                    HiddenFromGAL     = $ug.HiddenFromAddressListsEnabled
                    ModerationEnabled = $false
                    ManagedBy         = $ugOwners
                    ModeratedBy       = @()
                    GrantSendOnBehalf = Get-SmtpList $ug.GrantSendOnBehalfTo
                    AcceptFrom        = @()
                    RejectFrom        = @()
                    HasTeam           = ($ug.ResourceProvisioningOptions -contains 'Team')
                }
                $sourceGroups.Add([PSCustomObject]@{ Smtp = $srcSmtp; Type = 'GroupMailbox' })
            }
        }

        $total = $sourceGroups.Count
        Write-DLLog "Found $total source group(s). Reading members..." ([System.Drawing.Color]::DimGray)
        $idx = 0

        foreach ($grpEntry in $sourceGroups) {
            $idx++
            $srcSmtp = $grpEntry.Smtp
            $srcType = $grpEntry.Type
            $tgtSmtp = $script:DLMappingTable[$srcSmtp.ToLower()]
            Write-DLLog "  [$idx/$total] $srcSmtp ($srcType)" ([System.Drawing.Color]::DimGray)

            $members = if ($srcType -eq 'GroupMailbox') {
                @(Get-UnifiedGroupLinks -Identity $srcSmtp -LinkType Members -ResultSize Unlimited -ErrorAction SilentlyContinue)
            } else {
                @(Get-DistributionGroupMember -Identity $srcSmtp -ResultSize Unlimited -ErrorAction SilentlyContinue)
            }

            foreach ($m in $members) {
                $mSmtp   = $m.PrimarySmtpAddress
                $mType   = $m.RecipientTypeDetails
                $mTarget = Find-DLMapping $m

                if ($mType -in $dlTypes) {
                    if ($null -ne $mTarget) {
                        # $mTarget may be '' if this nested group is pending creation -- Phase 2 will re-resolve
                        $script:DLSourceData.Add([PSCustomObject]@{
                            SourceGroup     = $srcSmtp
                            TargetGroup     = $tgtSmtp
                            SourceMember    = $mSmtp
                            TargetMember    = $mTarget
                            MemberType      = $mType
                            ExpandedFrom    = ''
                            SourceGroupType = $srcType
                        })
                    } else {
                        Write-DLLog "    Nested '$mSmtp' unmapped -- expanding members" ([System.Drawing.Color]::DarkCyan)
                        $expanded = @(Expand-FrankensteinDLMembers -Identity $mSmtp -Recurse)
                        foreach ($em in $expanded) {
                            $emTarget = Find-DLMapping $em
                            if ($emTarget) {
                                $script:DLSourceData.Add([PSCustomObject]@{
                                    SourceGroup     = $srcSmtp
                                    TargetGroup     = $tgtSmtp
                                    SourceMember    = $em.PrimarySmtpAddress
                                    TargetMember    = $emTarget
                                    MemberType      = $em.RecipientTypeDetails
                                    ExpandedFrom    = $mSmtp
                                    SourceGroupType = $srcType
                                })
                            } else {
                                Write-DLLog "      SKIP $($em.PrimarySmtpAddress) (expanded from $mSmtp) -- not mapped" ([System.Drawing.Color]::DarkGoldenrod)
                            }
                        }
                    }
                } else {
                    if ($mTarget) {
                        $script:DLSourceData.Add([PSCustomObject]@{
                            SourceGroup     = $srcSmtp
                            TargetGroup     = $tgtSmtp
                            SourceMember    = $mSmtp
                            TargetMember    = $mTarget
                            MemberType      = $mType
                            ExpandedFrom    = ''
                            SourceGroupType = $srcType
                        })
                    } else {
                        Write-DLLog "    SKIP $mSmtp -- not mapped" ([System.Drawing.Color]::DarkGoldenrod)
                    }
                }
            }
        }
        Write-DLLog "Collection complete: $($script:DLSourceData.Count) member record(s) across $total group(s)." ([System.Drawing.Color]::LimeGreen)

        # Build proxy address index: maps every SMTP alias of each mapped source identity → its target.
        # Needed because CSV may use onmicrosoft.com addresses while ManagedBy/etc. resolve to vanity-domain SMTPs.
        $script:DLProxyIndex.Clear()
        $mappedSources = @($script:DLMappingTable.Keys | Where-Object { $script:DLMappingTable[$_] })
        if ($mappedSources.Count) {
            Write-DLLog "Building address alias index for $($mappedSources.Count) mapped identities..." ([System.Drawing.Color]::DimGray)
            foreach ($srcKey in $mappedSources) {
                $tgtSmtpVal = $script:DLMappingTable[$srcKey]
                $r = Get-Recipient -Identity $srcKey -ErrorAction SilentlyContinue
                if ($r -and $r.EmailAddresses) {
                    foreach ($addr in $r.EmailAddresses) {
                        if ($addr.ToString() -match '^smtp:(.+)$') {
                            $proxyKey = $Matches[1].ToLower()
                            if (-not $script:DLProxyIndex.ContainsKey($proxyKey)) {
                                $script:DLProxyIndex[$proxyKey] = $tgtSmtpVal
                            }
                        }
                    }
                }
            }
            Write-DLLog "Address alias index built ($($script:DLProxyIndex.Count) entries)." ([System.Drawing.Color]::DimGray)
        }
    }

    function Apply-DLTargetData {
        $script:DLResultLog.Clear()
        $createMode = $radCreateGroups.Checked
        $prefix     = $txtPrefix.Text.Trim()
        $newDomain  = $txtNewDomain.Text.Trim().TrimStart('@')

        # Phase 1: Create groups
        if ($createMode) {
            Write-DLLog "---- Phase 1: Creating groups ----" ([System.Drawing.Color]::CornflowerBlue)
            $p1Total = [math]::Max(1, $script:DLGroupMeta.Count); $p1Done = 0
            foreach ($srcSmtp in @($script:DLGroupMeta.Keys)) {
                $p1Done++
                Update-DLProgress ([int]($p1Done / $p1Total * 40)) 100
                $meta    = $script:DLGroupMeta[$srcSmtp]
                $srcType = $meta.Type

                # Mixed CSV: if a non-blank target was already in the mapping, skip creation
                $existingTarget = $script:DLMappingTable[$srcSmtp.ToLower()]
                if ($existingTarget) {
                    Write-DLLog "  EXISTING $srcSmtp -> $existingTarget (target already mapped -- skipping creation)" ([System.Drawing.Color]::DarkCyan)
                    continue
                }

                if ($srcType -eq 'GroupMailbox' -and $radOnPrem.Checked) {
                    Write-DLLog "  SKIP $srcSmtp -- M365 Groups cannot be created on-prem" ([System.Drawing.Color]::DarkGoldenrod)
                    $script:DLResultLog.Add([PSCustomObject][ordered]@{
                        Operation = 'CreateGroup'; SourceGroup = $srcSmtp; TargetGroup = 'N/A'
                        SourceMember = ''; TargetMember = ''; ExpandedFrom = ''
                        Status = 'Skipped'; Details = 'M365 Group not supported on on-prem target'
                        Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                    })
                    continue
                }

                $pfxAlias       = if ($chkPrefixAlias.Checked       -and $prefix) { $prefix } else { '' }
                $pfxName        = if ($chkPrefixName.Checked        -and $prefix) { $prefix } else { '' }
                $pfxDisplayName = if ($chkPrefixDisplayName.Checked -and $prefix) { $prefix } else { '' }
                $tgtAlias       = ($pfxAlias + $meta.Alias).ToLower() -replace '[^a-z0-9\-_]',''
                $tgtName        = "$pfxName$($meta.Name)"
                $tgtDisplayName = "$pfxDisplayName$($meta.DisplayName)"
                $tgtSmtp        = "$tgtAlias@$newDomain"

                try {
                    New-TargetDLGroup -Name $tgtName -DisplayName $tgtDisplayName -Alias $tgtAlias -PrimarySmtp $tgtSmtp -DLType $meta.DLType -GroupType $srcType -OU $txtOU.Text.Trim()
                    $script:DLMappingTable[$srcSmtp.ToLower()] = $tgtSmtp
                    Write-DLLog "  CREATED $tgtDisplayName ($tgtSmtp)" ([System.Drawing.Color]::LimeGreen)
                    try {
                        Set-TargetDLProperties -TargetIdentity $tgtSmtp -Meta $meta -GroupType $srcType -DefaultOwner $txtDefaultOwner.Text.Trim()
                        Write-DLLog "    Properties applied (Hidden=$($meta.HiddenFromGAL) ExtSenders=$(-not $meta.RequireSenderAuth) Moderated=$($meta.ModerationEnabled) Owners=$($meta.ManagedBy.Count))" ([System.Drawing.Color]::DimGray)
                    } catch {
                        Write-DLLog "    WARNING: Group created but properties failed -- $($_.Exception.Message)" ([System.Drawing.Color]::DarkGoldenrod)
                    }
                    # Team provisioning for M365 Groups
                    if ($srcType -eq 'GroupMailbox' -and -not $chkSkipTeams.Checked -and $meta.HasTeam) {
                        if (Get-Command New-Team -ErrorAction SilentlyContinue) {
                            try {
                                $newGroup = Get-UnifiedGroup -Identity $tgtSmtp -ErrorAction Stop
                                New-Team -GroupId $newGroup.ExternalDirectoryObjectId -ErrorAction Stop | Out-Null
                                Write-DLLog "    Team provisioned for $tgtSmtp" ([System.Drawing.Color]::LimeGreen)
                            } catch {
                                Write-DLLog "    WARNING: Team provisioning failed -- $($_.Exception.Message)" ([System.Drawing.Color]::DarkGoldenrod)
                            }
                        } else {
                            Write-DLLog "    WARNING: Source group has a Team but MicrosoftTeams module not loaded -- skipping Team provisioning for $tgtSmtp" ([System.Drawing.Color]::DarkGoldenrod)
                        }
                    }
                    $script:DLResultLog.Add([PSCustomObject][ordered]@{
                        Operation = 'CreateGroup'; SourceGroup = $srcSmtp; TargetGroup = $tgtSmtp
                        SourceMember = ''; TargetMember = ''; ExpandedFrom = ''
                        Status = 'Created'; Details = ''
                        Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                    })
                } catch {
                    $err    = $_.Exception.Message
                    $status = if ($err -match 'already|exists|conflict') { 'Conflict' } else { 'Failed' }
                    $color  = if ($status -eq 'Conflict') { [System.Drawing.Color]::SteelBlue } else { [System.Drawing.Color]::Tomato }
                    Write-DLLog "  $status $tgtName -- $err" $color
                    $script:DLResultLog.Add([PSCustomObject][ordered]@{
                        Operation = 'CreateGroup'; SourceGroup = $srcSmtp; TargetGroup = $tgtSmtp
                        SourceMember = ''; TargetMember = ''; ExpandedFrom = ''
                        Status = $status; Details = $err
                        Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                    })
                }
            }
        }

        # Write computed targets back to the source CSV so it can be reused in Add Members mode
        if ($createMode -and $script:DLMappingCsvPath) {
            Write-DLLog "---- Writing computed targets back to CSV ----" ([System.Drawing.Color]::CornflowerBlue)
            try {
                $csvRows     = Import-Csv -Path $script:DLMappingCsvPath
                $writeCount  = 0
                foreach ($row in $csvRows) {
                    if ($row.Source -and -not $row.Target) {
                        $computed = $script:DLMappingTable[$row.Source.Trim().ToLower()]
                        if ($computed) {
                            $row.Target = $computed
                            $writeCount++
                        }
                    }
                }
                $csvRows | Export-Csv -Path $script:DLMappingCsvPath -NoTypeInformation -Force
                Write-DLLog "  $writeCount computed target(s) written back to: $script:DLMappingCsvPath" ([System.Drawing.Color]::LimeGreen)
            } catch {
                Write-DLLog "  WARNING: Could not update CSV -- $($_.Exception.Message)" ([System.Drawing.Color]::DarkGoldenrod)
            }
        }

        # Phase 2: Add members
        $p2Total = $script:DLSourceData.Count; $p2Done = 0
        Write-DLLog "---- Phase 2: Adding members ($p2Total record(s)) ----" ([System.Drawing.Color]::CornflowerBlue)
        foreach ($rec in $script:DLSourceData) {
            $p2Done++
            Update-DLProgress (40 + [int]($p2Done / [math]::Max(1,$p2Total) * 60)) 100
            $tgtGroup  = $script:DLMappingTable[$rec.SourceGroup.ToLower()]
            if (-not $tgtGroup) { $tgtGroup = $rec.TargetGroup }
            # Re-resolve TargetMember in case it was a nested group pending creation at collection time
            $tgtMember = $rec.TargetMember
            if (-not $tgtMember -and $rec.SourceMember) {
                $tgtMember = $script:DLMappingTable[$rec.SourceMember.ToLower()]
            }
            $srcGroupType = $rec.SourceGroupType

            $log = [ordered]@{
                Operation    = 'AddMember'
                SourceGroup  = $rec.SourceGroup
                TargetGroup  = $tgtGroup
                SourceMember = $rec.SourceMember
                TargetMember = $tgtMember
                ExpandedFrom = $rec.ExpandedFrom
                Status       = ''
                Details      = ''
                Timestamp    = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            }

            if (-not $tgtGroup) {
                $log.Status  = 'Skipped'
                $log.Details = 'No target group identity -- group was not created or target mapping is missing'
                Write-DLLog "  SKIP  $tgtMember -- target group for '$($rec.SourceGroup)' unknown" ([System.Drawing.Color]::DarkGoldenrod)
                $script:DLResultLog.Add([PSCustomObject]$log)
                continue
            }

            if (-not $tgtMember) {
                $log.Status  = 'Skipped'
                $log.Details = 'Nested group target could not be resolved -- creation may have failed'
                Write-DLLog "  SKIP  $($rec.SourceMember) -> $tgtGroup -- nested group was not created" ([System.Drawing.Color]::DarkGoldenrod)
                $script:DLResultLog.Add([PSCustomObject]$log)
                continue
            }

            if ($srcGroupType -eq 'GroupMailbox' -and $radOnPrem.Checked) {
                $log.Status  = 'Skipped'
                $log.Details = 'M365 Group membership not supported on on-prem target'
                Write-DLLog "  SKIP  $tgtMember -> $tgtGroup (M365 Group on on-prem)" ([System.Drawing.Color]::DarkGoldenrod)
                $script:DLResultLog.Add([PSCustomObject]$log)
                continue
            }

            try {
                Add-TargetDLMember -TargetGroup $tgtGroup -TargetMember $tgtMember -GroupType $srcGroupType
                $log.Status  = 'Success'
                $log.Details = 'Member added'
                Write-DLLog "  OK    $tgtMember -> $tgtGroup$(if ($rec.ExpandedFrom) { " (from $($rec.ExpandedFrom))" })" ([System.Drawing.Color]::LimeGreen)
            } catch {
                $err = $_.Exception.Message
                if ($err -match 'already|member|exists') {
                    $log.Status  = 'AlreadyMember'
                    $log.Details = 'Already a member'
                    Write-DLLog "  DUP   $tgtMember already in $tgtGroup" ([System.Drawing.Color]::SteelBlue)
                } else {
                    $log.Status  = 'Failed'
                    $log.Details = $err
                    Write-DLLog "  ERR   $tgtMember -> $tgtGroup -- $err" ([System.Drawing.Color]::Tomato)
                }
            }
            $script:DLResultLog.Add([PSCustomObject]$log)
        }
        $progressBar.Value = 100
        $progressBar.Refresh()
    }

    #endregion

    #region ---- Build Form ----

    $form                 = New-Object System.Windows.Forms.Form
    $form.Text            = "Frankenstein - DL Migrator"
    $form.ClientSize      = New-Object System.Drawing.Size(800, 994)
    $form.StartPosition   = 'CenterScreen'
    $form.FormBorderStyle = 'FixedSingle'
    $form.MaximizeBox     = $false
    $form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.BackColor       = [System.Drawing.Color]::FromArgb(245, 245, 248)

    $lblHeader           = New-Object System.Windows.Forms.Label
    $lblHeader.Text      = "Frankenstein - Distribution Group Migrator"
    $lblHeader.Font      = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
    $lblHeader.ForeColor = [System.Drawing.Color]::FromArgb(28, 28, 72)
    $lblHeader.Location  = New-Object System.Drawing.Point(14, 12)
    $lblHeader.Size      = New-Object System.Drawing.Size(660, 28)
    $form.Controls.Add($lblHeader)

    # --- 1. Connections ---
    $grpConn           = New-Object System.Windows.Forms.GroupBox
    $grpConn.Text      = "1. Connections  (source is always M365 -- connect to both first to validate credentials)"
    $grpConn.Location  = New-Object System.Drawing.Point(12, 46)
    $grpConn.Size      = New-Object System.Drawing.Size(776, 148)
    $form.Controls.Add($grpConn)

    $btnConnSrc              = New-Object System.Windows.Forms.Button
    $btnConnSrc.Text         = "Connect Source (M365)"
    $btnConnSrc.Location     = New-Object System.Drawing.Point(10, 22)
    $btnConnSrc.Size         = New-Object System.Drawing.Size(168, 30)
    $grpConn.Controls.Add($btnConnSrc)

    $lblSrcStatus            = New-Object System.Windows.Forms.Label
    $lblSrcStatus.Text       = "Not connected"
    $lblSrcStatus.ForeColor  = [System.Drawing.Color]::Gray
    $lblSrcStatus.Location   = New-Object System.Drawing.Point(188, 28)
    $lblSrcStatus.Size       = New-Object System.Drawing.Size(576, 18)
    $grpConn.Controls.Add($lblSrcStatus)

    $lblTargetType           = New-Object System.Windows.Forms.Label
    $lblTargetType.Text      = "Target type:"
    $lblTargetType.Location  = New-Object System.Drawing.Point(10, 64)
    $lblTargetType.Size      = New-Object System.Drawing.Size(78, 18)
    $grpConn.Controls.Add($lblTargetType)

    $radM365                 = New-Object System.Windows.Forms.RadioButton
    $radM365.Text            = "M365"
    $radM365.Checked         = $true
    $radM365.Location        = New-Object System.Drawing.Point(92, 62)
    $radM365.Size            = New-Object System.Drawing.Size(58, 20)
    $grpConn.Controls.Add($radM365)

    $radOnPrem               = New-Object System.Windows.Forms.RadioButton
    $radOnPrem.Text          = "On-prem Exchange"
    $radOnPrem.Location      = New-Object System.Drawing.Point(156, 62)
    $radOnPrem.Size          = New-Object System.Drawing.Size(148, 20)
    $grpConn.Controls.Add($radOnPrem)

    $txtOnPremUri                 = New-Object System.Windows.Forms.TextBox
    $txtOnPremUri.Location        = New-Object System.Drawing.Point(10, 86)
    $txtOnPremUri.Size            = New-Object System.Drawing.Size(754, 22)
    $txtOnPremUri.PlaceholderText = "On-prem URI: https://mailserver.domain.com/PowerShell/  (leave blank to use current Exchange session)"
    $txtOnPremUri.Enabled         = $false
    $grpConn.Controls.Add($txtOnPremUri)

    $btnConnTgt              = New-Object System.Windows.Forms.Button
    $btnConnTgt.Text         = "Connect Target (M365)"
    $btnConnTgt.Location     = New-Object System.Drawing.Point(10, 114)
    $btnConnTgt.Size         = New-Object System.Drawing.Size(168, 30)
    $grpConn.Controls.Add($btnConnTgt)

    $lblTgtStatus            = New-Object System.Windows.Forms.Label
    $lblTgtStatus.Text       = "Not connected"
    $lblTgtStatus.ForeColor  = [System.Drawing.Color]::Gray
    $lblTgtStatus.Location   = New-Object System.Drawing.Point(188, 120)
    $lblTgtStatus.Size       = New-Object System.Drawing.Size(576, 18)
    $grpConn.Controls.Add($lblTgtStatus)

    # --- 2. Identity Mapping ---
    $grpMap           = New-Object System.Windows.Forms.GroupBox
    $grpMap.Text      = "2. Identity Mapping"
    $grpMap.Location  = New-Object System.Drawing.Point(12, 202)
    $grpMap.Size      = New-Object System.Drawing.Size(776, 214)
    $form.Controls.Add($grpMap)

    $radCsvMapping          = New-Object System.Windows.Forms.RadioButton
    $radCsvMapping.Text     = "Load CSV mapping file  (Source + Target columns -- include both groups AND their members)"
    $radCsvMapping.Checked  = $true
    $radCsvMapping.Location = New-Object System.Drawing.Point(10, 20)
    $radCsvMapping.Size     = New-Object System.Drawing.Size(670, 20)
    $grpMap.Controls.Add($radCsvMapping)

    $btnBrowseMap           = New-Object System.Windows.Forms.Button
    $btnBrowseMap.Text      = "Browse..."
    $btnBrowseMap.Location  = New-Object System.Drawing.Point(682, 16)
    $btnBrowseMap.Size      = New-Object System.Drawing.Size(82, 26)
    $grpMap.Controls.Add($btnBrowseMap)

    $lblMapPath             = New-Object System.Windows.Forms.Label
    $lblMapPath.Text        = "No file loaded"
    $lblMapPath.ForeColor   = [System.Drawing.Color]::Gray
    $lblMapPath.Location    = New-Object System.Drawing.Point(10, 42)
    $lblMapPath.Size        = New-Object System.Drawing.Size(754, 16)
    $grpMap.Controls.Add($lblMapPath)

    $ttDLMap = New-Object System.Windows.Forms.ToolTip
    $ttDLMap.AutoPopDelay = 12000
    $ttDLMap.InitialDelay = 350
    $ttDLMap.ReshowDelay  = 200
    $ttDLMap.ShowAlways   = $true
    $dlCsvTip = @"
CSV Requirements
------------------------------------------------
Column headers (exact, case-insensitive):
  Source  -- identity in the SOURCE tenant (M365)
  Target  -- corresponding identity in the TARGET

Include BOTH types of rows in one file:
  * Group rows:  source DL/SG SMTP -> target DL/SG SMTP
  * Member rows: source user SMTP  -> target user SMTP

Create Groups mode -- group rows support three patterns:
  1. Blank target  -> group doesn't exist yet; target SMTP is
     computed from prefix + alias + new domain at run time.
  2. Target present -> group already exists in the target tenant;
     creation is skipped and the provided target is used directly.
  3. Mixed CSV     -> some group rows blank, some with targets.
     Each is handled independently: blank = create, present = use.
  Blank-target rows are flagged with a warning in
  Add Members mode (where targets are required).

Nested groups in the mapping: target equivalent is
  added as a member (structure preserved).
Nested groups NOT in the mapping: members expanded,
  and those members in the mapping are added individually.

Accepted formats: SMTP, UPN, or any Exchange-resolvable identity.
Supported group types: MailUniversalDistributionGroup,
  MailUniversalSecurityGroup, GroupMailbox (M365 Groups).
"@
    $ttDLMap.SetToolTip($radCsvMapping, $dlCsvTip)
    $ttDLMap.SetToolTip($btnBrowseMap,  $dlCsvTip)

    $radManualEntry         = New-Object System.Windows.Forms.RadioButton
    $radManualEntry.Text    = "Enter source/target pairs manually"
    $radManualEntry.Location = New-Object System.Drawing.Point(10, 62)
    $radManualEntry.Size    = New-Object System.Drawing.Size(754, 20)
    $grpMap.Controls.Add($radManualEntry)

    $txtManualSrc                 = New-Object System.Windows.Forms.TextBox
    $txtManualSrc.Location        = New-Object System.Drawing.Point(10, 86)
    $txtManualSrc.Size            = New-Object System.Drawing.Size(334, 22)
    $txtManualSrc.PlaceholderText = "source@domain.com"
    $txtManualSrc.Enabled         = $false
    $grpMap.Controls.Add($txtManualSrc)

    $lblArrow                = New-Object System.Windows.Forms.Label
    $lblArrow.Text           = "->"
    $lblArrow.TextAlign      = 'MiddleCenter'
    $lblArrow.Location       = New-Object System.Drawing.Point(350, 88)
    $lblArrow.Size           = New-Object System.Drawing.Size(16, 18)
    $grpMap.Controls.Add($lblArrow)

    $txtManualTgt                 = New-Object System.Windows.Forms.TextBox
    $txtManualTgt.Location        = New-Object System.Drawing.Point(372, 86)
    $txtManualTgt.Size            = New-Object System.Drawing.Size(312, 22)
    $txtManualTgt.PlaceholderText = "target@domain.com"
    $txtManualTgt.Enabled         = $false
    $grpMap.Controls.Add($txtManualTgt)

    $btnAddPair             = New-Object System.Windows.Forms.Button
    $btnAddPair.Text        = "Add"
    $btnAddPair.Location    = New-Object System.Drawing.Point(692, 84)
    $btnAddPair.Size        = New-Object System.Drawing.Size(62, 26)
    $btnAddPair.Enabled     = $false
    $grpMap.Controls.Add($btnAddPair)

    $lvPairs                = New-Object System.Windows.Forms.ListView
    $lvPairs.View           = 'Details'
    $lvPairs.FullRowSelect  = $true
    $lvPairs.GridLines      = $true
    $lvPairs.Location       = New-Object System.Drawing.Point(10, 116)
    $lvPairs.Size           = New-Object System.Drawing.Size(754, 62)
    $lvPairs.Enabled        = $false
    $lvPairs.Columns.Add("Source SMTP", 370) | Out-Null
    $lvPairs.Columns.Add("Target SMTP", 370) | Out-Null
    $grpMap.Controls.Add($lvPairs)

    $btnRemovePair          = New-Object System.Windows.Forms.Button
    $btnRemovePair.Text     = "Remove Selected"
    $btnRemovePair.Location = New-Object System.Drawing.Point(10, 184)
    $btnRemovePair.Size     = New-Object System.Drawing.Size(126, 24)
    $btnRemovePair.Enabled  = $false
    $grpMap.Controls.Add($btnRemovePair)

    $lblMapCount            = New-Object System.Windows.Forms.Label
    $lblMapCount.Text       = ""
    $lblMapCount.ForeColor  = [System.Drawing.Color]::DarkGreen
    $lblMapCount.Location   = New-Object System.Drawing.Point(144, 186)
    $lblMapCount.Size       = New-Object System.Drawing.Size(600, 16)
    $grpMap.Controls.Add($lblMapCount)

    # --- 3. Operation ---
    $grpOp           = New-Object System.Windows.Forms.GroupBox
    $grpOp.Text      = "3. Operation"
    $grpOp.Location  = New-Object System.Drawing.Point(12, 424)
    $grpOp.Size      = New-Object System.Drawing.Size(776, 294)
    $form.Controls.Add($grpOp)

    # Group type filter -- always active, both modes
    $lblGroupTypes          = New-Object System.Windows.Forms.Label
    $lblGroupTypes.Text     = "Group types:"
    $lblGroupTypes.Location = New-Object System.Drawing.Point(10, 18)
    $lblGroupTypes.Size     = New-Object System.Drawing.Size(80, 18)
    $grpOp.Controls.Add($lblGroupTypes)

    $chkAllTypes          = New-Object System.Windows.Forms.CheckBox
    $chkAllTypes.Text     = "All"
    $chkAllTypes.Checked  = $true
    $chkAllTypes.Location = New-Object System.Drawing.Point(94, 16)
    $chkAllTypes.Size     = New-Object System.Drawing.Size(44, 20)
    $grpOp.Controls.Add($chkAllTypes)

    $chkTypeDL          = New-Object System.Windows.Forms.CheckBox
    $chkTypeDL.Text     = "Distribution List"
    $chkTypeDL.Checked  = $true
    $chkTypeDL.Location = New-Object System.Drawing.Point(142, 16)
    $chkTypeDL.Size     = New-Object System.Drawing.Size(130, 20)
    $grpOp.Controls.Add($chkTypeDL)

    $chkTypeMailSec          = New-Object System.Windows.Forms.CheckBox
    $chkTypeMailSec.Text     = "Mail-Enabled Security"
    $chkTypeMailSec.Checked  = $true
    $chkTypeMailSec.Location = New-Object System.Drawing.Point(276, 16)
    $chkTypeMailSec.Size     = New-Object System.Drawing.Size(152, 20)
    $grpOp.Controls.Add($chkTypeMailSec)

    $chkTypeM365          = New-Object System.Windows.Forms.CheckBox
    $chkTypeM365.Text     = "M365 Group"
    $chkTypeM365.Checked  = $false
    $chkTypeM365.Location = New-Object System.Drawing.Point(432, 16)
    $chkTypeM365.Size     = New-Object System.Drawing.Size(96, 20)
    $grpOp.Controls.Add($chkTypeM365)

    # M365-specific options -- visible only when chkTypeM365 is checked, enabled only in Create mode
    $chkSuppressWelcome          = New-Object System.Windows.Forms.CheckBox
    $chkSuppressWelcome.Text     = "Suppress welcome email"
    $chkSuppressWelcome.Checked  = $true
    $chkSuppressWelcome.Location = New-Object System.Drawing.Point(432, 38)
    $chkSuppressWelcome.Size     = New-Object System.Drawing.Size(164, 20)
    $chkSuppressWelcome.Visible  = $false
    $chkSuppressWelcome.Enabled  = $false
    $grpOp.Controls.Add($chkSuppressWelcome)

    $chkSkipTeams          = New-Object System.Windows.Forms.CheckBox
    $chkSkipTeams.Text     = "Skip Team provisioning"
    $chkSkipTeams.Checked  = $true
    $chkSkipTeams.Location = New-Object System.Drawing.Point(600, 38)
    $chkSkipTeams.Size     = New-Object System.Drawing.Size(60, 20)
    $chkSkipTeams.Visible  = $false
    $chkSkipTeams.Enabled  = $false
    $grpOp.Controls.Add($chkSkipTeams)

    $radAddMembers          = New-Object System.Windows.Forms.RadioButton
    $radAddMembers.Text     = "Add members to existing groups  (groups must already exist on target)"
    $radAddMembers.Checked  = $true
    $radAddMembers.Location = New-Object System.Drawing.Point(10, 62)
    $radAddMembers.Size     = New-Object System.Drawing.Size(654, 20)
    $grpOp.Controls.Add($radAddMembers)

    $radCreateGroups          = New-Object System.Windows.Forms.RadioButton
    $radCreateGroups.Text     = "Create groups, then add members  (name/alias conflicts are skipped and logged)"
    $radCreateGroups.Location = New-Object System.Drawing.Point(10, 84)
    $radCreateGroups.Size     = New-Object System.Drawing.Size(654, 20)
    $grpOp.Controls.Add($radCreateGroups)

    $lblPrefix               = New-Object System.Windows.Forms.Label
    $lblPrefix.Text          = "Prefix:"
    $lblPrefix.Location      = New-Object System.Drawing.Point(28, 112)
    $lblPrefix.Size          = New-Object System.Drawing.Size(46, 18)
    $lblPrefix.Enabled       = $false
    $grpOp.Controls.Add($lblPrefix)

    $txtPrefix                    = New-Object System.Windows.Forms.TextBox
    $txtPrefix.Location           = New-Object System.Drawing.Point(78, 110)
    $txtPrefix.Size               = New-Object System.Drawing.Size(120, 22)
    $txtPrefix.PlaceholderText    = "e.g. MIGR-"
    $txtPrefix.Enabled            = $false
    $grpOp.Controls.Add($txtPrefix)

    $lblNewDomain            = New-Object System.Windows.Forms.Label
    $lblNewDomain.Text       = "New SMTP domain:"
    $lblNewDomain.Location   = New-Object System.Drawing.Point(216, 112)
    $lblNewDomain.Size       = New-Object System.Drawing.Size(112, 18)
    $lblNewDomain.Enabled    = $false
    $grpOp.Controls.Add($lblNewDomain)

    $txtNewDomain                 = New-Object System.Windows.Forms.TextBox
    $txtNewDomain.Location        = New-Object System.Drawing.Point(332, 110)
    $txtNewDomain.Size            = New-Object System.Drawing.Size(344, 22)
    $txtNewDomain.PlaceholderText = "target.com"
    $txtNewDomain.Enabled         = $false
    $grpOp.Controls.Add($txtNewDomain)

    $lblPrefixApply          = New-Object System.Windows.Forms.Label
    $lblPrefixApply.Text     = "Apply prefix to:"
    $lblPrefixApply.Location = New-Object System.Drawing.Point(28, 139)
    $lblPrefixApply.Size     = New-Object System.Drawing.Size(96, 18)
    $lblPrefixApply.Enabled  = $false
    $grpOp.Controls.Add($lblPrefixApply)

    $chkPrefixName               = New-Object System.Windows.Forms.CheckBox
    $chkPrefixName.Text          = "Name"
    $chkPrefixName.Location      = New-Object System.Drawing.Point(128, 137)
    $chkPrefixName.Size          = New-Object System.Drawing.Size(62, 20)
    $chkPrefixName.Enabled       = $false
    $chkPrefixName.Checked       = $true
    $grpOp.Controls.Add($chkPrefixName)

    $chkPrefixDisplayName               = New-Object System.Windows.Forms.CheckBox
    $chkPrefixDisplayName.Text          = "Display Name"
    $chkPrefixDisplayName.Location      = New-Object System.Drawing.Point(194, 137)
    $chkPrefixDisplayName.Size          = New-Object System.Drawing.Size(106, 20)
    $chkPrefixDisplayName.Enabled       = $false
    $chkPrefixDisplayName.Checked       = $true
    $grpOp.Controls.Add($chkPrefixDisplayName)

    $chkPrefixAlias               = New-Object System.Windows.Forms.CheckBox
    $chkPrefixAlias.Text          = "Alias"
    $chkPrefixAlias.Location      = New-Object System.Drawing.Point(304, 137)
    $chkPrefixAlias.Size          = New-Object System.Drawing.Size(62, 20)
    $chkPrefixAlias.Enabled       = $false
    $chkPrefixAlias.Checked       = $true
    $grpOp.Controls.Add($chkPrefixAlias)

    $lblOU               = New-Object System.Windows.Forms.Label
    $lblOU.Text          = "Organizational Unit:"
    $lblOU.Location      = New-Object System.Drawing.Point(28, 166)
    $lblOU.Size          = New-Object System.Drawing.Size(124, 18)
    $lblOU.Enabled       = $false
    $grpOp.Controls.Add($lblOU)

    $txtOU                    = New-Object System.Windows.Forms.TextBox
    $txtOU.Location           = New-Object System.Drawing.Point(156, 164)
    $txtOU.Size               = New-Object System.Drawing.Size(606, 22)
    $txtOU.PlaceholderText    = "OU=Groups,DC=domain,DC=com  (optional -- leave blank for Exchange default)"
    $txtOU.Enabled            = $false
    $grpOp.Controls.Add($txtOU)

    $lblDefaultOwner          = New-Object System.Windows.Forms.Label
    $lblDefaultOwner.Text     = "Default Owner:"
    $lblDefaultOwner.Location = New-Object System.Drawing.Point(28, 192)
    $lblDefaultOwner.Size     = New-Object System.Drawing.Size(96, 18)
    $lblDefaultOwner.Enabled  = $false
    $grpOp.Controls.Add($lblDefaultOwner)

    $txtDefaultOwner                 = New-Object System.Windows.Forms.TextBox
    $txtDefaultOwner.Location        = New-Object System.Drawing.Point(128, 190)
    $txtDefaultOwner.Size            = New-Object System.Drawing.Size(634, 22)
    $txtDefaultOwner.PlaceholderText = "admin@target.com  (added as owner on all created groups)"
    $txtDefaultOwner.Enabled         = $false
    $grpOp.Controls.Add($txtDefaultOwner)

    # Properties to copy (create mode only)
    $lblCopyProps          = New-Object System.Windows.Forms.Label
    $lblCopyProps.Text     = "Copy properties:"
    $lblCopyProps.Location = New-Object System.Drawing.Point(10, 220)
    $lblCopyProps.Size     = New-Object System.Drawing.Size(106, 18)
    $lblCopyProps.Enabled  = $false
    $grpOp.Controls.Add($lblCopyProps)

    $chkAllProps          = New-Object System.Windows.Forms.CheckBox
    $chkAllProps.Text     = "Select All"
    $chkAllProps.Checked  = $true
    $chkAllProps.Location = New-Object System.Drawing.Point(120, 218)
    $chkAllProps.Size     = New-Object System.Drawing.Size(86, 20)
    $chkAllProps.Enabled  = $false
    $grpOp.Controls.Add($chkAllProps)

    $chkPropManagedBy          = New-Object System.Windows.Forms.CheckBox
    $chkPropManagedBy.Text     = "Owners"
    $chkPropManagedBy.Checked  = $true
    $chkPropManagedBy.Location = New-Object System.Drawing.Point(28, 240)
    $chkPropManagedBy.Size     = New-Object System.Drawing.Size(72, 20)
    $chkPropManagedBy.Enabled  = $false
    $grpOp.Controls.Add($chkPropManagedBy)

    $chkPropHiddenGAL          = New-Object System.Windows.Forms.CheckBox
    $chkPropHiddenGAL.Text     = "Hidden from GAL"
    $chkPropHiddenGAL.Checked  = $true
    $chkPropHiddenGAL.Location = New-Object System.Drawing.Point(106, 240)
    $chkPropHiddenGAL.Size     = New-Object System.Drawing.Size(124, 20)
    $chkPropHiddenGAL.Enabled  = $false
    $grpOp.Controls.Add($chkPropHiddenGAL)

    $chkPropRequireSenderAuth          = New-Object System.Windows.Forms.CheckBox
    $chkPropRequireSenderAuth.Text     = "External Senders"
    $chkPropRequireSenderAuth.Checked  = $true
    $chkPropRequireSenderAuth.Location = New-Object System.Drawing.Point(236, 240)
    $chkPropRequireSenderAuth.Size     = New-Object System.Drawing.Size(126, 20)
    $chkPropRequireSenderAuth.Enabled  = $false
    $grpOp.Controls.Add($chkPropRequireSenderAuth)

    $chkPropGrantSendOnBehalf          = New-Object System.Windows.Forms.CheckBox
    $chkPropGrantSendOnBehalf.Text     = "Send on Behalf"
    $chkPropGrantSendOnBehalf.Checked  = $true
    $chkPropGrantSendOnBehalf.Location = New-Object System.Drawing.Point(368, 240)
    $chkPropGrantSendOnBehalf.Size     = New-Object System.Drawing.Size(116, 20)
    $chkPropGrantSendOnBehalf.Enabled  = $false
    $grpOp.Controls.Add($chkPropGrantSendOnBehalf)

    $chkPropModeration          = New-Object System.Windows.Forms.CheckBox
    $chkPropModeration.Text     = "Moderation"
    $chkPropModeration.Checked  = $true
    $chkPropModeration.Location = New-Object System.Drawing.Point(28, 262)
    $chkPropModeration.Size     = New-Object System.Drawing.Size(92, 20)
    $chkPropModeration.Enabled  = $false
    $grpOp.Controls.Add($chkPropModeration)

    $chkPropAcceptFrom          = New-Object System.Windows.Forms.CheckBox
    $chkPropAcceptFrom.Text     = "Accept Restrictions"
    $chkPropAcceptFrom.Checked  = $true
    $chkPropAcceptFrom.Location = New-Object System.Drawing.Point(126, 262)
    $chkPropAcceptFrom.Size     = New-Object System.Drawing.Size(146, 20)
    $chkPropAcceptFrom.Enabled  = $false
    $grpOp.Controls.Add($chkPropAcceptFrom)

    $chkPropRejectFrom          = New-Object System.Windows.Forms.CheckBox
    $chkPropRejectFrom.Text     = "Reject Restrictions"
    $chkPropRejectFrom.Checked  = $true
    $chkPropRejectFrom.Location = New-Object System.Drawing.Point(278, 262)
    $chkPropRejectFrom.Size     = New-Object System.Drawing.Size(146, 20)
    $chkPropRejectFrom.Enabled  = $false
    $grpOp.Controls.Add($chkPropRejectFrom)

    # --- 4. Options ---
    $grpOpts           = New-Object System.Windows.Forms.GroupBox
    $grpOpts.Text      = "4. Options"
    $grpOpts.Location  = New-Object System.Drawing.Point(12, 726)
    $grpOpts.Size      = New-Object System.Drawing.Size(776, 52)
    $form.Controls.Add($grpOpts)

    $lblOutPath               = New-Object System.Windows.Forms.Label
    $lblOutPath.Text          = "Log output path:"
    $lblOutPath.Location      = New-Object System.Drawing.Point(10, 18)
    $lblOutPath.Size          = New-Object System.Drawing.Size(112, 22)
    $grpOpts.Controls.Add($lblOutPath)

    $txtOutPath               = New-Object System.Windows.Forms.TextBox
    $txtOutPath.Text          = (Get-Location).Path
    $txtOutPath.Location      = New-Object System.Drawing.Point(126, 16)
    $txtOutPath.Size          = New-Object System.Drawing.Size(534, 22)
    $grpOpts.Controls.Add($txtOutPath)

    $btnBrowseOut             = New-Object System.Windows.Forms.Button
    $btnBrowseOut.Text        = "Browse..."
    $btnBrowseOut.Location    = New-Object System.Drawing.Point(670, 14)
    $btnBrowseOut.Size        = New-Object System.Drawing.Size(94, 26)
    $grpOpts.Controls.Add($btnBrowseOut)

    # --- Action Buttons ---
    $btnPreview               = New-Object System.Windows.Forms.Button
    $btnPreview.Text          = "Preview"
    $btnPreview.Location      = New-Object System.Drawing.Point(12, 788)
    $btnPreview.Size          = New-Object System.Drawing.Size(130, 34)
    $btnPreview.Enabled       = $false
    $form.Controls.Add($btnPreview)

    $btnRun                   = New-Object System.Windows.Forms.Button
    $btnRun.Text              = "Run Migration"
    $btnRun.Location          = New-Object System.Drawing.Point(260, 788)
    $btnRun.Size              = New-Object System.Drawing.Size(180, 34)
    $btnRun.BackColor         = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnRun.ForeColor         = [System.Drawing.Color]::White
    $btnRun.FlatStyle         = 'Flat'
    $btnRun.Enabled           = $false
    $form.Controls.Add($btnRun)

    $btnExportLog             = New-Object System.Windows.Forms.Button
    $btnExportLog.Text        = "Export Log"
    $btnExportLog.Location    = New-Object System.Drawing.Point(658, 788)
    $btnExportLog.Size        = New-Object System.Drawing.Size(130, 34)
    $btnExportLog.Enabled     = $false
    $form.Controls.Add($btnExportLog)

    # --- Progress Bar ---
    $progressBar               = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location      = New-Object System.Drawing.Point(12, 828)
    $progressBar.Size          = New-Object System.Drawing.Size(776, 18)
    $progressBar.Minimum       = 0
    $progressBar.Maximum       = 100
    $progressBar.Value         = 0
    $progressBar.Style         = 'Continuous'
    $form.Controls.Add($progressBar)

    # --- Status Log ---
    $grpLog           = New-Object System.Windows.Forms.GroupBox
    $grpLog.Text      = "Status Log"
    $grpLog.Location  = New-Object System.Drawing.Point(12, 852)
    $grpLog.Size      = New-Object System.Drawing.Size(776, 132)
    $form.Controls.Add($grpLog)

    $rtbLog                   = New-Object System.Windows.Forms.RichTextBox
    $rtbLog.Location          = New-Object System.Drawing.Point(8, 18)
    $rtbLog.Size              = New-Object System.Drawing.Size(760, 106)
    $rtbLog.ReadOnly          = $true
    $rtbLog.BackColor         = [System.Drawing.Color]::FromArgb(18, 18, 28)
    $rtbLog.ForeColor         = [System.Drawing.Color]::Silver
    $rtbLog.Font              = New-Object System.Drawing.Font("Consolas", 8.5)
    $rtbLog.ScrollBars        = 'Vertical'
    $rtbLog.BorderStyle       = 'None'
    $grpLog.Controls.Add($rtbLog)

    #endregion

    #region ---- Event Handlers ----

    # Target type selection
    $radM365.Add_CheckedChanged({
        $txtOnPremUri.Enabled = -not $radM365.Checked
        $btnConnTgt.Text      = if ($radM365.Checked) { "Connect Target (M365)" } else { "Connect Target (On-prem)" }
    })
    $radOnPrem.Add_CheckedChanged({
        $txtOnPremUri.Enabled = $radOnPrem.Checked
        $btnConnTgt.Text      = if ($radOnPrem.Checked) { "Connect Target (On-prem)" } else { "Connect Target (M365)" }
    })

    # Mapping mode: CSV vs Manual
    $radCsvMapping.Add_CheckedChanged({
        $btnBrowseMap.Enabled  = $radCsvMapping.Checked
        $lblMapPath.Enabled    = $radCsvMapping.Checked
        $txtManualSrc.Enabled  = -not $radCsvMapping.Checked
        $txtManualTgt.Enabled  = -not $radCsvMapping.Checked
        $btnAddPair.Enabled    = -not $radCsvMapping.Checked
        $lvPairs.Enabled       = -not $radCsvMapping.Checked
        $btnRemovePair.Enabled = -not $radCsvMapping.Checked
        if ($radCsvMapping.Checked) {
            $lvPairs.Items.Clear()
            $script:DLMappingTable = @{}
            $lblMapCount.Text = ""
        }
    })
    $radManualEntry.Add_CheckedChanged({
        $btnBrowseMap.Enabled  = -not $radManualEntry.Checked
        $lblMapPath.Enabled    = -not $radManualEntry.Checked
        $txtManualSrc.Enabled  = $radManualEntry.Checked
        $txtManualTgt.Enabled  = $radManualEntry.Checked
        $btnAddPair.Enabled    = $radManualEntry.Checked
        $lvPairs.Enabled       = $radManualEntry.Checked
        $btnRemovePair.Enabled = $radManualEntry.Checked
        if ($radManualEntry.Checked) {
            $script:DLMappingTable   = @{}
            $script:DLMappingCsvPath = ''
            $lblMapPath.Text         = "No file loaded"
            $lblMapPath.ForeColor    = [System.Drawing.Color]::Gray
            $lblMapCount.Text        = "0 pair(s) entered"
        }
    })

    $btnAddPair.Add_Click({
        $src = $txtManualSrc.Text.Trim()
        $tgt = $txtManualTgt.Text.Trim()
        if (-not $src -or -not $tgt) {
            [System.Windows.Forms.MessageBox]::Show("Enter both a source and target address.", "Missing Address", 'OK', 'Warning') | Out-Null
            return
        }
        $item = New-Object System.Windows.Forms.ListViewItem($src)
        $item.SubItems.Add($tgt) | Out-Null
        $lvPairs.Items.Add($item) | Out-Null
        $script:DLMappingTable[$src.ToLower()] = $tgt
        $txtManualSrc.Text = ''
        $txtManualTgt.Text = ''
        $lblMapCount.Text  = "$($lvPairs.Items.Count) pair(s) entered"
        Write-DLLog "Pair added: $src -> $tgt" ([System.Drawing.Color]::LimeGreen)
        $txtManualSrc.Focus() | Out-Null
    })
    $txtManualTgt.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Return) { $btnAddPair.PerformClick() }
    })

    $btnRemovePair.Add_Click({
        $selected = @($lvPairs.SelectedItems)
        if (-not $selected.Count) { return }
        foreach ($item in $selected) { $lvPairs.Items.Remove($item) }
        $script:DLMappingTable = @{}
        foreach ($item in $lvPairs.Items) { $script:DLMappingTable[$item.Text.ToLower()] = $item.SubItems[1].Text }
        $lblMapCount.Text = "$($lvPairs.Items.Count) pair(s) entered"
        Write-DLLog "Pair(s) removed. $($lvPairs.Items.Count) remaining." ([System.Drawing.Color]::DarkGoldenrod)
    })

    # Operation mode
    $radAddMembers.Add_CheckedChanged({
        $c = -not $radAddMembers.Checked
        $lblPrefix.Enabled                = $c
        $txtPrefix.Enabled                = $c
        $lblNewDomain.Enabled             = $c
        $txtNewDomain.Enabled             = $c
        $lblPrefixApply.Enabled           = $c
        $chkPrefixName.Enabled            = $c
        $chkPrefixDisplayName.Enabled     = $c
        $chkPrefixAlias.Enabled           = $c
        $lblOU.Enabled                    = $c
        $txtOU.Enabled                    = $c
        $lblDefaultOwner.Enabled          = $c
        $txtDefaultOwner.Enabled          = $c
        $lblCopyProps.Enabled             = $c
        $chkAllProps.Enabled              = $c
        $chkPropManagedBy.Enabled         = $c
        $chkPropHiddenGAL.Enabled         = $c
        $chkPropRequireSenderAuth.Enabled = $c
        $chkPropGrantSendOnBehalf.Enabled = $c
        $chkPropModeration.Enabled        = $c
        $chkPropAcceptFrom.Enabled        = $c
        $chkPropRejectFrom.Enabled        = $c
        # M365 sub-options: only enabled in create mode
        $m365create = $c -and $chkTypeM365.Checked
        $chkSuppressWelcome.Enabled = $m365create
        $chkSkipTeams.Enabled       = $m365create
    })
    $radCreateGroups.Add_CheckedChanged({
        $c = $radCreateGroups.Checked
        $lblPrefix.Enabled                = $c
        $txtPrefix.Enabled                = $c
        $lblNewDomain.Enabled             = $c
        $txtNewDomain.Enabled             = $c
        $lblPrefixApply.Enabled           = $c
        $chkPrefixName.Enabled            = $c
        $chkPrefixDisplayName.Enabled     = $c
        $chkPrefixAlias.Enabled           = $c
        $lblOU.Enabled                    = $c
        $txtOU.Enabled                    = $c
        $lblDefaultOwner.Enabled          = $c
        $txtDefaultOwner.Enabled          = $c
        $lblCopyProps.Enabled             = $c
        $chkAllProps.Enabled              = $c
        $chkPropManagedBy.Enabled         = $c
        $chkPropHiddenGAL.Enabled         = $c
        $chkPropRequireSenderAuth.Enabled = $c
        $chkPropGrantSendOnBehalf.Enabled = $c
        $chkPropModeration.Enabled        = $c
        $chkPropAcceptFrom.Enabled        = $c
        $chkPropRejectFrom.Enabled        = $c
        # M365 sub-options: only enabled in create mode
        $m365create = $c -and $chkTypeM365.Checked
        $chkSuppressWelcome.Enabled = $m365create
        $chkSkipTeams.Enabled       = $m365create
    })

    # M365 sub-options visibility toggle
    $chkTypeM365.Add_CheckedChanged({
        $chkSuppressWelcome.Visible = $chkTypeM365.Checked
        $chkSkipTeams.Visible       = $chkTypeM365.Checked
        # Enable only when also in create mode
        $m365create = $chkTypeM365.Checked -and $radCreateGroups.Checked
        $chkSuppressWelcome.Enabled = $m365create
        $chkSkipTeams.Enabled       = $m365create
    })

    # Select All - Group Types
    $script:DLTypesUpdating = $false
    $chkAllTypes.Add_CheckedChanged({
        if ($script:DLTypesUpdating) { return }
        $script:DLTypesUpdating = $true
        $chkTypeDL.Checked      = $chkAllTypes.Checked
        $chkTypeMailSec.Checked = $chkAllTypes.Checked
        $chkTypeM365.Checked    = $chkAllTypes.Checked
        $script:DLTypesUpdating = $false
    })
    $updateAllTypes = {
        if ($script:DLTypesUpdating) { return }
        $script:DLTypesUpdating = $true
        $chkAllTypes.Checked = ($chkTypeDL.Checked -and $chkTypeMailSec.Checked -and $chkTypeM365.Checked)
        $script:DLTypesUpdating = $false
    }
    $chkTypeDL.Add_CheckedChanged($updateAllTypes)
    $chkTypeMailSec.Add_CheckedChanged($updateAllTypes)
    $chkTypeM365.Add_CheckedChanged($updateAllTypes)

    # Select All - Copy Properties
    $script:DLPropsUpdating = $false
    $chkAllProps.Add_CheckedChanged({
        if ($script:DLPropsUpdating) { return }
        $script:DLPropsUpdating = $true
        $chkPropManagedBy.Checked         = $chkAllProps.Checked
        $chkPropHiddenGAL.Checked         = $chkAllProps.Checked
        $chkPropRequireSenderAuth.Checked = $chkAllProps.Checked
        $chkPropGrantSendOnBehalf.Checked = $chkAllProps.Checked
        $chkPropModeration.Checked        = $chkAllProps.Checked
        $chkPropAcceptFrom.Checked        = $chkAllProps.Checked
        $chkPropRejectFrom.Checked        = $chkAllProps.Checked
        $script:DLPropsUpdating = $false
    })
    $updateAllProps = {
        if ($script:DLPropsUpdating) { return }
        $script:DLPropsUpdating = $true
        $chkAllProps.Checked = ($chkPropManagedBy.Checked -and $chkPropHiddenGAL.Checked -and
                                $chkPropRequireSenderAuth.Checked -and $chkPropGrantSendOnBehalf.Checked -and
                                $chkPropModeration.Checked -and $chkPropAcceptFrom.Checked -and
                                $chkPropRejectFrom.Checked)
        $script:DLPropsUpdating = $false
    }
    $chkPropManagedBy.Add_CheckedChanged($updateAllProps)
    $chkPropHiddenGAL.Add_CheckedChanged($updateAllProps)
    $chkPropRequireSenderAuth.Add_CheckedChanged($updateAllProps)
    $chkPropGrantSendOnBehalf.Add_CheckedChanged($updateAllProps)
    $chkPropModeration.Add_CheckedChanged($updateAllProps)
    $chkPropAcceptFrom.Add_CheckedChanged($updateAllProps)
    $chkPropRejectFrom.Add_CheckedChanged($updateAllProps)

    # Browse mapping CSV
    $btnBrowseMap.Add_Click({
        $dlg        = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title  = "Select Source-to-Target Mapping CSV"
        $dlg.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $rows = Import-Csv $dlg.FileName
        if (-not ($rows | Get-Member -Name 'Source') -or -not ($rows | Get-Member -Name 'Target')) {
            [System.Windows.Forms.MessageBox]::Show("CSV must have 'Source' and 'Target' columns.", "Invalid CSV", 'OK', 'Error') | Out-Null
            return
        }
        $script:DLMappingTable = @{}
        $blankTargetCount = 0
        foreach ($row in $rows) {
            if ($row.Source) {
                $script:DLMappingTable[$row.Source.Trim().ToLower()] = if ($row.Target) { $row.Target.Trim() } else { '' }
                if (-not $row.Target) { $blankTargetCount++ }
            }
        }
        $script:DLMappingCsvPath = $dlg.FileName
        $lblMapPath.Text         = $dlg.FileName
        $lblMapPath.ForeColor    = [System.Drawing.Color]::DarkGreen
        $lblMapCount.Text        = "$($script:DLMappingTable.Count) entries loaded$(if ($blankTargetCount) { " ($blankTargetCount with no target -- OK for Create mode)" })"
        Write-DLLog "Mapping CSV loaded: $($script:DLMappingTable.Count) entries$(if ($blankTargetCount) { ", $blankTargetCount with blank target (valid in Create Groups mode)" })" ([System.Drawing.Color]::LimeGreen)
    })

    # Browse output folder
    $btnBrowseOut.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Select log output folder"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtOutPath.Text = $dlg.SelectedPath }
    })

    # Connect Source
    $btnConnSrc.Add_Click({
        Set-DLStatusLabel $lblSrcStatus "Connecting..." ([System.Drawing.Color]::DarkGoldenrod)
        $form.UseWaitCursor = $true
        $btnConnSrc.Enabled = $false
        try {
            Write-DLLog "Connecting to source M365 tenant..." ([System.Drawing.Color]::Silver)
            $beforeIds = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | ForEach-Object { $_.ConnectionId })
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            $allConns = @(Get-ConnectionInformation -ErrorAction SilentlyContinue)
            $newConn  = $allConns | Where-Object { $_.ConnectionId -notin $beforeIds } | Select-Object -First 1
            if (-not $newConn) { $newConn = $allConns | Select-Object -First 1 }
            if ($newConn) {
                $script:DLSourceUpn          = [string]$newConn.UserPrincipalName
                $script:DLSourceOrg          = [string]$newConn.Organization
                $script:DLSourceConnectionId = $newConn.ConnectionId
            }
            if (-not $script:DLSourceOrg) {
                $script:DLSourceOrg = [string](Get-OrganizationConfig -ErrorAction SilentlyContinue).Name
            }
            $srcDisplay = @($script:DLSourceOrg, $script:DLSourceUpn) | Where-Object { $_ }
            $script:DLSourceConnected = $true
            Set-DLStatusLabel $lblSrcStatus "Connected  --  $($srcDisplay -join '  |  ')" ([System.Drawing.Color]::DarkGreen)
            Write-DLLog "Source connected: $($srcDisplay -join ' | ')" ([System.Drawing.Color]::LimeGreen)
            $btnPreview.Enabled = $true
            if ($script:DLTargetConnected) { $btnRun.Enabled = $true }
        } catch {
            Set-DLStatusLabel $lblSrcStatus "Connection failed" ([System.Drawing.Color]::Tomato)
            Write-DLLog "ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Tomato)
        } finally {
            $form.UseWaitCursor = $false
            $btnConnSrc.Enabled = $true
        }
    })

    # Connect Target
    $btnConnTgt.Add_Click({
        Set-DLStatusLabel $lblTgtStatus "Connecting..." ([System.Drawing.Color]::DarkGoldenrod)
        $form.UseWaitCursor = $true
        $btnConnTgt.Enabled = $false
        try {
            if ($radM365.Checked) {
                Write-DLLog "Connecting to target M365 tenant..." ([System.Drawing.Color]::Silver)
                $beforeIds = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | ForEach-Object { $_.ConnectionId })
                Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
                $allConns = @(Get-ConnectionInformation -ErrorAction SilentlyContinue)
                $newConn  = $allConns | Where-Object { $_.ConnectionId -notin $beforeIds } | Select-Object -First 1
                if (-not $newConn) { $newConn = $allConns | Select-Object -First 1 }
                if ($newConn) {
                    $script:DLTargetUpn          = [string]$newConn.UserPrincipalName
                    $script:DLTargetOrg          = [string]$newConn.Organization
                    $script:DLTargetConnectionId = $newConn.ConnectionId
                }
                if (-not $script:DLTargetOrg) {
                    $script:DLTargetOrg = [string](Get-OrganizationConfig -ErrorAction SilentlyContinue).Name
                }
                $tgtDisplay = @($script:DLTargetOrg, $script:DLTargetUpn) | Where-Object { $_ }
                Set-DLStatusLabel $lblTgtStatus "Connected  --  $($tgtDisplay -join '  |  ')" ([System.Drawing.Color]::DarkGreen)
                Write-DLLog "Target M365 connected: $($tgtDisplay -join ' | ')" ([System.Drawing.Color]::LimeGreen)
            } else {
                $uri = $txtOnPremUri.Text.Trim()
                if ($uri) {
                    Write-DLLog "Connecting to on-prem Exchange: $uri" ([System.Drawing.Color]::Silver)
                    $script:DLOnPremSession = New-PSSession -ConfigurationName Microsoft.Exchange `
                        -ConnectionUri $uri -Authentication Kerberos -ErrorAction Stop
                    Set-DLStatusLabel $lblTgtStatus "Connected (on-prem via URI)" ([System.Drawing.Color]::DarkGreen)
                    Write-DLLog "On-prem session established." ([System.Drawing.Color]::LimeGreen)
                } else {
                    Set-DLStatusLabel $lblTgtStatus "Using current Exchange session" ([System.Drawing.Color]::DarkGreen)
                    Write-DLLog "Using current Exchange PS session for on-prem target." ([System.Drawing.Color]::LimeGreen)
                    Write-DLLog "Note: After connecting to source EXO above, verify on-prem cmdlets are still active before running." ([System.Drawing.Color]::DarkGoldenrod)
                }
            }
            $script:DLTargetConnected = $true
            if ($script:DLSourceConnected) { $btnRun.Enabled = $true }
        } catch {
            Set-DLStatusLabel $lblTgtStatus "Connection failed" ([System.Drawing.Color]::Tomato)
            Write-DLLog "ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Tomato)
        } finally {
            $form.UseWaitCursor = $false
            $btnConnTgt.Enabled = $true
        }
    })

    # Returns the current EXO org name, or $null if not connected
    function Get-DLCurrentOrg {
        try {
            if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
                $conn = Get-ConnectionInformation -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($conn) { return [string]$conn.Organization }
            }
            $org = Get-OrganizationConfig -ErrorAction SilentlyContinue
            if ($org) { return [string]$org.Name }
        } catch {}
        return $null
    }

    # Ensure we are connected to source; use Set-ConnectionContext for silent switching
    function Switch-DLToSource {
        # Best path: Set-ConnectionContext (EXO 3.2+) is truly silent
        if ($script:DLSourceConnectionId -and (Get-Command Set-ConnectionContext -ErrorAction SilentlyContinue)) {
            try {
                Set-ConnectionContext -ConnectionId $script:DLSourceConnectionId -ErrorAction Stop
                return
            } catch {}
        }
        # Fast path: single active connection and it is already source
        if ($script:DLSourceOrg -and (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
            $active = @(Get-ConnectionInformation -ErrorAction SilentlyContinue)
            if ($active.Count -eq 1 -and [string]$active[0].Organization -eq $script:DLSourceOrg) { return }
        }
        # Reconnect -- pass stored UPN so MSAL can serve the cached token silently
        Write-DLLog "Switching to source$(if ($script:DLSourceOrg) { " ($script:DLSourceOrg)" })..." ([System.Drawing.Color]::DimGray)
        if ($script:DLSourceUpn) {
            Connect-ExchangeOnline -UserPrincipalName $script:DLSourceUpn -ShowBanner:$false -ErrorAction Stop
        } else {
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        }
    }

    # Ensure we are connected to target; use Set-ConnectionContext for silent switching
    function Switch-DLToTarget {
        # Best path: Set-ConnectionContext (EXO 3.2+) is truly silent
        if ($script:DLTargetConnectionId -and (Get-Command Set-ConnectionContext -ErrorAction SilentlyContinue)) {
            try {
                Set-ConnectionContext -ConnectionId $script:DLTargetConnectionId -ErrorAction Stop
                return
            } catch {}
        }
        # Fast path: single active connection and it is already target
        if ($script:DLTargetOrg -and (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
            $active = @(Get-ConnectionInformation -ErrorAction SilentlyContinue)
            if ($active.Count -eq 1 -and [string]$active[0].Organization -eq $script:DLTargetOrg) { return }
        }
        # Reconnect -- pass stored UPN so MSAL can serve the cached token silently
        Write-DLLog "Switching to target$(if ($script:DLTargetOrg) { " ($script:DLTargetOrg)" })..." ([System.Drawing.Color]::DimGray)
        if ($script:DLTargetUpn) {
            Connect-ExchangeOnline -UserPrincipalName $script:DLTargetUpn -ShowBanner:$false -ErrorAction Stop
        } else {
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        }
    }

    # Shared validation -- returns $true if ready to proceed
    function Confirm-DLReadyState {
        if (-not $script:DLMappingTable.Count) {
            $msg = if ($radCsvMapping.Checked) { "Load a mapping CSV first." } else { "Add at least one source/target pair first." }
            [System.Windows.Forms.MessageBox]::Show($msg, "No Mapping", 'OK', 'Warning') | Out-Null
            return $false
        }
        if ($radCreateGroups.Checked -and -not $txtNewDomain.Text.Trim()) {
            [System.Windows.Forms.MessageBox]::Show("Enter a new SMTP domain for group creation.", "Missing Domain", 'OK', 'Warning') | Out-Null
            return $false
        }
        if ($radAddMembers.Checked) {
            $emptyTargets = @($script:DLMappingTable.GetEnumerator() | Where-Object { -not $_.Value })
            if ($emptyTargets.Count) {
                $result = [System.Windows.Forms.MessageBox]::Show(
                    "$($emptyTargets.Count) mapping entry(s) have no target value and will be skipped.`n`nEntries with blank targets are only valid in 'Create Groups' mode.`n`nProceed anyway?",
                    "Blank Target Entries",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
                if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return $false }
            }
        }
        return $true
    }

    # Preview
    $btnPreview.Add_Click({
        if (-not (Confirm-DLReadyState)) { return }
        $btnPreview.Enabled = $false
        $form.UseWaitCursor = $true
        $progressBar.Value  = 0
        $outPath = $txtOutPath.Text.Trim()
        $script:DLLiveLogPath = if ($outPath -and (Test-Path $outPath)) {
            Join-Path $outPath "DLPreviewLog_$((Get-Date).ToString('yyyyMMdd_HHmmss')).txt"
        } else { '' }
        try {
            Switch-DLToSource
            Set-DLStatusLabel $lblSrcStatus "Connected - scanning groups..." ([System.Drawing.Color]::DarkOrange)

            Collect-DLSourceData
            Set-DLStatusLabel $lblSrcStatus "Connected - $($script:DLSourceData.Count) member record(s)" ([System.Drawing.Color]::DarkGreen)

            Write-DLLog "---- PREVIEW ----" ([System.Drawing.Color]::CornflowerBlue)

            if ($radCreateGroups.Checked) {
                $prefix    = $txtPrefix.Text.Trim()
                $newDomain = $txtNewDomain.Text.Trim().TrimStart('@')
                Write-DLLog "  [Create Phase -- $($script:DLGroupMeta.Count) group(s)]" ([System.Drawing.Color]::CornflowerBlue)
                foreach ($srcSmtp in @($script:DLGroupMeta.Keys)) {
                    $meta           = $script:DLGroupMeta[$srcSmtp]
                    $existingTarget = $script:DLMappingTable[$srcSmtp.ToLower()]
                    if ($existingTarget) {
                        Write-DLLog "  EXISTING $srcSmtp -> $existingTarget (already mapped -- skipping creation)" ([System.Drawing.Color]::DarkCyan)
                        continue
                    }
                    $pfxAlias       = if ($chkPrefixAlias.Checked       -and $prefix) { $prefix } else { '' }
                    $pfxName        = if ($chkPrefixName.Checked        -and $prefix) { $prefix } else { '' }
                    $pfxDisplayName = if ($chkPrefixDisplayName.Checked -and $prefix) { $prefix } else { '' }
                    $tgtAlias       = ($pfxAlias + $meta.Alias).ToLower() -replace '[^a-z0-9\-_]',''
                    $tgtName        = "$pfxName$($meta.Name)"
                    $tgtDisplayName = "$pfxDisplayName$($meta.DisplayName)"
                    $tgtSmtp        = "$tgtAlias@$newDomain"
                    if ($meta.Type -eq 'GroupMailbox' -and $radOnPrem.Checked) {
                        Write-DLLog "  SKIP   $srcSmtp -- M365 Group cannot be created on-prem" ([System.Drawing.Color]::DarkGoldenrod)
                    } else {
                        $ouInfo = if ($txtOU.Text.Trim()) { "  OU=$($txtOU.Text.Trim())" } else { '' }
                        Write-DLLog "  CREATE $tgtDisplayName ($tgtSmtp)  [Name: $tgtName | Alias: $tgtAlias | $($meta.Type)]$ouInfo" ([System.Drawing.Color]::LimeGreen)
                    }
                }
            }

            $wouldAdd = 0; $wouldSkip = 0
            $total = $script:DLSourceData.Count
            Write-DLLog "  [Member Phase -- $total record(s)]" ([System.Drawing.Color]::CornflowerBlue)
            $idx = 0
            foreach ($rec in $script:DLSourceData) {
                $idx++
                $tgt = $script:DLMappingTable[$rec.SourceGroup.ToLower()]
                if (-not $tgt) { $tgt = $rec.TargetGroup }
                Write-DLLog "  ADD   $($rec.TargetMember) -> $tgt$(if ($rec.ExpandedFrom) { " (from $($rec.ExpandedFrom))" })" ([System.Drawing.Color]::LimeGreen)
                $wouldAdd++
                Update-DLProgress $idx $total
            }
            $progressBar.Value = 100
            Write-DLLog "---- Preview: $wouldAdd member(s) would be added ----" ([System.Drawing.Color]::CornflowerBlue)
        } catch {
            Write-DLLog "ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Tomato)
        } finally {
            $btnPreview.Enabled = $true
            $form.UseWaitCursor = $false
        }
    })

    # Run Migration
    $btnRun.Add_Click({
        if (-not (Confirm-DLReadyState)) { return }
        $modeDesc = if ($radCreateGroups.Checked) { "create groups and add members" } else { "add members to existing groups" }
        $confirm  = [System.Windows.Forms.MessageBox]::Show(
            "This will connect to both tenants and $modeDesc for $($script:DLGroupMeta.Count) group(s).`n`nProceed?",
            "Confirm Migration",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $btnRun.Enabled     = $false
        $form.UseWaitCursor = $true
        $progressBar.Value  = 0
        $outPath = $txtOutPath.Text.Trim()
        $script:DLLiveLogPath = if ($outPath -and (Test-Path $outPath)) {
            Join-Path $outPath "DLMigrationLog_$((Get-Date).ToString('yyyyMMdd_HHmmss'))_live.txt"
        } else { '' }
        if ($script:DLLiveLogPath) { Write-DLLog "Live log: $script:DLLiveLogPath" ([System.Drawing.Color]::DimGray) }
        try {
            Switch-DLToSource
            Set-DLStatusLabel $lblSrcStatus "Connected - scanning groups..." ([System.Drawing.Color]::DarkOrange)
            Collect-DLSourceData
            Set-DLStatusLabel $lblSrcStatus "Connected - $($script:DLSourceData.Count) member record(s) read" ([System.Drawing.Color]::DarkGreen)

            if ($radM365.Checked) {
                Switch-DLToTarget
            } elseif ($script:DLOnPremSession -and $script:DLOnPremSession.State -ne 'Opened') {
                $uri = $txtOnPremUri.Text.Trim()
                if ($uri) {
                    Write-DLLog "Re-establishing on-prem session..." ([System.Drawing.Color]::DimGray)
                    $script:DLOnPremSession = New-PSSession -ConfigurationName Microsoft.Exchange `
                        -ConnectionUri $uri -Authentication Kerberos -ErrorAction Stop
                }
            }
            Set-DLStatusLabel $lblTgtStatus "Connected - applying..." ([System.Drawing.Color]::DarkOrange)

            Apply-DLTargetData

            $cCreated = @($script:DLResultLog | Where-Object { $_.Status -eq 'Created' }).Count
            $cSuccess = @($script:DLResultLog | Where-Object { $_.Status -eq 'Success' }).Count
            $cDupe    = @($script:DLResultLog | Where-Object { $_.Status -eq 'AlreadyMember' }).Count
            $cConf    = @($script:DLResultLog | Where-Object { $_.Status -eq 'Conflict' }).Count
            $cFailed  = @($script:DLResultLog | Where-Object { $_.Status -eq 'Failed' }).Count
            Write-DLLog "---- COMPLETE: $cCreated created | $cSuccess added | $cDupe already members | $cConf conflicts | $cFailed failed ----" ([System.Drawing.Color]::CornflowerBlue)

            # Flag any groups that were not created
            $notCreated = @($script:DLResultLog | Where-Object {
                $_.Operation -eq 'CreateGroup' -and $_.Status -in @('Conflict','Failed')
            })
            if ($notCreated.Count) {
                Write-DLLog "!!!! $($notCreated.Count) GROUP(S) NOT CREATED -- members for these groups were skipped !!!!" ([System.Drawing.Color]::OrangeRed)
                foreach ($nc in $notCreated) {
                    $tag = if ($nc.Status -eq 'Conflict') { 'CONFLICT' } else { 'FAILED' }
                    Write-DLLog "  [$tag]  $($nc.SourceGroup)  ->  $($nc.TargetGroup)" ([System.Drawing.Color]::OrangeRed)
                    if ($nc.Details) {
                        Write-DLLog "           $($nc.Details)" ([System.Drawing.Color]::DarkSalmon)
                    }
                }
                Write-DLLog "!!!! Adjust prefix/domain or fix conflicts, then re-run !!!!" ([System.Drawing.Color]::OrangeRed)
            }
            Set-DLStatusLabel $lblTgtStatus "Connected - migration complete" ([System.Drawing.Color]::DarkGreen)
            $btnExportLog.Enabled = $true
        } catch {
            Write-DLLog "FATAL: $($_.Exception.Message)" ([System.Drawing.Color]::Tomato)
        } finally {
            $btnRun.Enabled = $true
            $form.UseWaitCursor = $false
        }
    })

    # Export Log
    $btnExportLog.Add_Click({
        $outPath = $txtOutPath.Text.Trim()
        if (-not (Test-Path $outPath)) {
            Write-DLLog "Output path not found: $outPath" ([System.Drawing.Color]::Tomato)
            return
        }
        $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')

        # Status log (full RichTextBox content as plain text)
        $txtFile = Join-Path $outPath "DLStatusLog_$stamp.txt"
        try {
            $rtbLog.Text | Out-File $txtFile -Encoding UTF8 -Force
            Write-DLLog "Status log exported: $txtFile" ([System.Drawing.Color]::LimeGreen)
        } catch {
            Write-DLLog "ERROR exporting status log: $($_.Exception.Message)" ([System.Drawing.Color]::Tomato)
        }

        # Result log (CSV of every operation)
        $logFile = Join-Path $outPath "DLMigrationLog_$stamp.csv"
        try {
            $script:DLResultLog | Export-Csv $logFile -NoTypeInformation -Encoding UTF8
            Write-DLLog "Operation log exported: $logFile" ([System.Drawing.Color]::LimeGreen)
        } catch {
            Write-DLLog "ERROR exporting log: $($_.Exception.Message)" ([System.Drawing.Color]::Tomato)
        }

        # Group map -- Source/Target pairs for all processed groups with final (post-creation) targets
        if ($script:DLGroupMeta.Count) {
            $mapFile = Join-Path $outPath "DLGroupMap_$stamp.csv"
            try {
                $mapRows = foreach ($srcSmtp in ($script:DLGroupMeta.Keys | Sort-Object)) {
                    $tgt = $script:DLMappingTable[$srcSmtp.ToLower()]
                    [PSCustomObject]@{ Source = $srcSmtp; Target = if ($tgt) { $tgt } else { '' } }
                }
                $mapRows | Export-Csv $mapFile -NoTypeInformation -Encoding UTF8
                Write-DLLog "Group map exported: $mapFile" ([System.Drawing.Color]::LimeGreen)
            } catch {
                Write-DLLog "ERROR exporting group map: $($_.Exception.Message)" ([System.Drawing.Color]::Tomato)
            }
        }
    })

    #endregion

    Write-DLLog "Ready. Connect to both tenants first, then load your mapping." ([System.Drawing.Color]::Silver)
    Write-DLLog "Tip: Include both group rows AND member rows in the same mapping CSV." ([System.Drawing.Color]::DimGray)

    $form.ShowDialog() | Out-Null
    $form.Dispose()
    if ($script:DLOnPremSession) { Remove-PSSession $script:DLOnPremSession -ErrorAction SilentlyContinue }
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
