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

    4) Install-ExchangeOnline
       Installs and configures Exchange Online PowerShell requirements.

    5) Connect-All
       Connects to AzureAD and Exchange Online PS Sessions.
       Switches: [-NoMFA]
       Note: AzureAD/MSOnline modules are deprecated. Migrate to Microsoft Graph where possible.

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

    10) Get-FrankensteinAzureDiscovery
        Outputs Azure/MSOL discovery information.
        Switches: [-CSV] [-UseCurrentSession]
        Note: Requires legacy AzureAD/MSOnline modules (deprecated).

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

function Connect-All {
    [CmdletBinding()]
    Param (
        [Switch]$NoMFA
    )

    Write-Warning "AzureAD and MSOnline modules are deprecated. Consider migrating to Microsoft Graph (Connect-MgGraph)."

    if ($NoMFA) {
        $AdminUsername    = Read-Host -Prompt "Azure/Office 365 Admin User Account"
        $AdminPassword    = Read-Host -Prompt "Password" -AsSecureString
        $adminCredentials = New-Object System.Management.Automation.PSCredential($AdminUsername, $AdminPassword)

        Connect-AzureAD         -Credential $adminCredentials
        Connect-MSOLService     -Credential $adminCredentials
        Connect-ExchangeOnline  -Credential $adminCredentials
    }
    else {
        Connect-AzureAD
        Connect-MSOLService
        Connect-ExchangeOnline
    }
}

#endregion

#region Installation

function Install-ExchangeOnline {
    [CmdletBinding()]
    Param()

    Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Install-PackageProvider -Name NuGet -Force
    Install-Module -Name PowerShellGet -Force
    Update-Module  -Name PowerShellGet
    Install-Module -Name ExchangeOnlineManagement -Confirm:$false
    Import-Module ExchangeOnlineManagement
}

function Install-All {
    [CmdletBinding()]
    Param()

    Write-Warning "MSOnline and AzureAD modules are deprecated. Consider using the Microsoft Graph PowerShell SDK (Install-Module Microsoft.Graph) instead."
    Install-ExchangeOnline
    Install-Module MSOnline -Scope CurrentUser
    Install-Module AzureAD  -AllowClobber -Scope CurrentUser
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

function Get-FrankensteinAzureDiscovery {
    [CmdletBinding()]
    Param (
        [Switch]$CSV,
        [Switch]$UseCurrentSession
    )

    Write-Warning "This function uses the deprecated AzureAD and MSOnline modules. Consider migrating to Microsoft Graph."

    if (-not $UseCurrentSession) {
        Connect-AzureAD
        Connect-MsolService
    }

    $DateStamp = (Get-Date).ToString('MMddyy')
    $OutputDir = ".\FrankensteinAzureDiscovery_$DateStamp"
    New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
    Push-Location $OutputDir

    Start-Transcript ".\Get-FrankensteinAzureDiscovery_$DateStamp.txt"

    $MSOLUser = Get-MsolUser -All
    $Device   = Get-MSOLDevice -All
    $Licenses = Get-MsolAccountSku

    Get-Linebreak
    Write-Host "Get-MsolUser ($($MSOLUser.Count) users discovered)" -ForegroundColor Cyan
    if ($CSV) {
        $MSOLUser | Select-Object `
            @{Name="AlternateEmailAddresses";            Expression={$_.AlternateEmailAddresses -join ";"}},
            @{Name="AlternateMobilePhones";              Expression={$_.AlternateMobilePhones -join ";"}},
            @{Name="AlternativeSecurityIds";             Expression={$_.AlternativeSecurityIds -join ";"}},
            BlockCredential, City, CloudExchangeRecipientType, Country, Department,
            @{Name="DirSyncProvisioningErrors";          Expression={$_.DirSyncProvisioningErrors -join ";"}},
            DisplayName, Errors, Fax, FirstName, ImmutableID,
            @{Name="IndirectLicenseErrors";              Expression={$_.IndirectLicenseErrors -join ";"}},
            IsBlackberryUser, IsLicensed, LastDirSynced, LastName, LastPasswordChangeTimestamp,
            @{Name="LicenseAssignmentDetails";           Expression={$_.LicenseAssignmentDetails -join ";"}},
            LicenseReconciliationNeeded,
            @{Name="Licenses";                           Expression={$_.Licenses -join ";"}},
            LiveId, MSExchRecipientTypeDetails, MSRtcSipDeploymentLocator,
            MSRtcSipPrimaryUserAddress, MobilePhone, ObjectId, Office,
            OverallProvisioningStatus, PasswordNeverExpires,
            PasswordResetNotRequiredDuringActivate, PhoneNumber, PortalSettings,
            PostalCode, PreferredDataLocation, PreferredLanguage,
            @{Name="ProxyAddresses";                     Expression={$_.ProxyAddresses -join ";"}},
            ReleaseTrack,
            @{Name="ServiceInformation";                 Expression={$_.ServiceInformation -join ";"}},
            SignInName, SoftDeletionTimestamp, State, StreetAddress,
            @{Name="StrongAuthenticationMethods";        Expression={$_.StrongAuthenticationMethods -join ";"}},
            @{Name="StrongAuthenticationPhoneAppDetails";Expression={$_.StrongAuthenticationPhoneAppDetails -join ";"}},
            @{Name="StrongAuthenticationProofupTime";    Expression={$_.StrongAuthenticationProofupTime -join ";"}},
            @{Name="StrongAuthenticationRequirements";   Expression={$_.StrongAuthenticationRequirements -join ";"}},
            @{Name="StrongAuthenticationUserDetails";    Expression={$_.StrongAuthenticationUserDetails -join ";"}},
            StrongPasswordRequired, StsRefreshTokensValidFrom, Title, UsageLocation,
            UserLandingPageIdentifierForO365Shell, UserPrincipalName,
            UserThemeIdentifierForO365Shell, UserType, ValidationStatus, WhenCreated |
            Export-Csv ".\MSOLUsers_$DateStamp.csv" -NoTypeInformation
    }

    Get-Linebreak
    Write-Host "Get-MsolCompanyInformation" -ForegroundColor Cyan
    Get-MsolCompanyInformation

    Get-Linebreak
    Write-Host "Get-MsolAccountSku" -ForegroundColor Cyan
    $Licenses | Select-Object AccountSkuID, ActiveUnits, WarningUnits, ConsumedUnits
    if ($CSV) {
        $Licenses | Select-Object AccountName, AccountSkuID, ActiveUnits, ConsumedUnits,
            LockedOutUnits, SKUID, SkuPartNumber, TargetClass, SuspendedUnits, WarningUnits |
            Export-Csv ".\MSOLLicenses_$DateStamp.csv" -NoTypeInformation
    }

    Get-Linebreak
    Write-Host "Get-MsolDevice ($($Device.Count) devices discovered)" -ForegroundColor Cyan
    if ($CSV) {
        $Device | Select-Object Enabled, ObjectID, DeviceID, DisplayName, DeviceObjectVersion,
            DeviceOSType, DeviceOSVersion, DeviceTrustType, DeviceTrustLevel,
            @{Name="DevicePhysicalIds";      Expression={$_.DevicePhysicalIds -join ";"}},
            ApproximateLastLogonTimestamp,
            @{Name="AlternativeSecurityIds"; Expression={$_.AlternativeSecurityIds -join ";"}},
            DirSyncEnabled, LastDirSyncTime, RegisteredOwners,
            @{Name="GraphDeviceObject";      Expression={$_.GraphDeviceObject -join ";"}} |
            Export-Csv ".\MSOLDevices_$DateStamp.csv" -NoTypeInformation
    }

    Get-Linebreak
    Write-Host "Get-MSOLDirSyncFeatures" -ForegroundColor Cyan
    Get-MSOLDirSyncFeatures

    Get-Linebreak
    Write-Host "Get-AzureADExtensionProperty" -ForegroundColor Cyan
    Get-AzureADExtensionProperty | Format-List
    if ($CSV) {
        Get-AzureADExtensionProperty | Select-Object Name, ObjectID, AppDisplayName, DataType,
            IsSyncedFromOnPremises,
            @{Name="TargetObjects"; Expression={$_.TargetObjects -join ";"}} |
            Export-Csv ".\AzureADExtensions_$DateStamp.csv" -NoTypeInformation
    }

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
