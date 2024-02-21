


<#
.SYNOPSIS
    Creation of Eric Frank. Discovers Exchange On-Premises and Online Information.

.DESCRIPTION
    This module contains a series of functions used to collect and export data in preparation from an Exchange to Exchange Online migration.

.PARAMETER 
    Get-FrankensteinHelp: View all Functions in this module

.EXAMPLE
    Get-FrankensteinExchangeDiscovery -Online -CSV -UseCurrentSession -PublicFolders
    Get-FrankensteinGSuiteDiscovery -CSV


.INPUTS
    

.OUTPUTS
    CSV and .txt files
    

.NOTES
    Author:  Eric D. Frank
    11/07/23 - Updated to use GitHub as repository
  
#>
 
function Get-FrankensteinHelp {    
    [CmdletBinding()]
    Param (
        )    
        Write-Host "
        
        Frankenstein offers several functions to assist in the Exchange, Azure and GSuite discovery processes. Below represents a brief explanation of each:

        1) Get-FrankensteinExchangeDiscovery: Provides Exchange on-premises discovery information and outputs a transcript along with optional CSV outputs. The default is on-premises unless the -Online switch is specified. 

            [-Virtualdirectories] [-CSV] [-UseCurrentSession] [-Online] [-PublicFolders]

        2) Get-FrankensteinPublicFolderDiscovery: Provides CSV outputs for Exchange Public Folder information.

        3) Get-FrankensteinGSuiteDiscovery: Outputs G Suite discovery CSV files. 
                
            Prerequisites: PSGsuite https://psgsuite.io/       
        

        4) Install-ExchangeOnline: Will install and configure Exchange Online PowerShell requirements to run Connect-ExchangeOnline

        5) Connect-All: Will connect to MSOL, AzureAD and ExO PS Sessions

            [-noMFA]

        6) Connect-OnPremServer: Connects to on-premises Exchange server using FQDN

        7) Get-FrankesnteinRecipientCounts: Displays summary of all recipient types

         
        "
        }
function Get-Linebreak {
    [CmdletBinding()]
    Param (
    )
        Write-Host "
        
################################################################################################
        
        "
}
function Connect-ExchangeOnPremServer {    
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [String]$ExchangeServerFQDN
        )    
    $UserCredential = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServerFQDN/PowerShell/ -Authentication Kerberos -Credential $UserCredential
    Import-PSSession $Session -DisableNameChecking
}
function Get-FrankensteinVirtualDirectories {    
    [CmdletBinding()]
    Param (
    [Switch]$CSV
    )
      
       
        Get-Linebreak
        "Get-VirtualDirectories"
        if($CSV){       
        $ClientAccess = Get-ClientAccessServer
        $ClientAccess | ForEach-Object{Get-AutoDiscoverVirtualDirectory -ADPropertiesOnly | Select-Object server,name,internalurl,externalurl,@{Name="Internalauthenticationmethods";Expression={$_.Internalauthenticationmethods -join “;”}},@{Name="Externalauthenticationmethods";Expression={$_.Externalauthenticationmethods -join “;”}},IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        $ClientAccess | ForEach-Object{Get-OwaVirtualDirectory -ADPropertiesOnly | Select-Object server,name,internalurl,externalurl,@{Name="Internalauthenticationmethods";Expression={$_.Internalauthenticationmethods -join “;”}},@{Name="Externalauthenticationmethods";Expression={$_.Externalauthenticationmethods -join “;”}},IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation -Append
        $ClientAccess | ForEach-Object{Get-ECPVirtualDirectory -ADPropertiesOnly | Select-Object server,name,internalurl,externalurl,@{Name="Internalauthenticationmethods";Expression={$_.Internalauthenticationmethods -join “;”}},@{Name="Externalauthenticationmethods";Expression={$_.Externalauthenticationmethods -join “;”}},IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation -Append
        $ClientAccess | ForEach-Object{Get-MAPIVirtualDirectory -ADPropertiesOnly | Select-Object server,name,internalurl,externalurl,@{Name="Internalauthenticationmethods";Expression={$_.Internalauthenticationmethods -join “;”}},@{Name="Externalauthenticationmethods";Expression={$_.Externalauthenticationmethods -join “;”}},IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation -Append
        $ClientAccess | ForEach-Object{Get-ActiveSyncVirtualDirectory -ADPropertiesOnly | Select-Object server,name,internalurl,externalurl,@{Name="Internalauthenticationmethods";Expression={$_.Internalauthenticationmethods -join “;”}},@{Name="Externalauthenticationmethods";Expression={$_.Externalauthenticationmethods -join “;”}},IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation -Append
        $ClientAccess | ForEach-Object{Get-WebServicesVirtualDirectory -ADPropertiesOnly | Select-Object server,name,internalurl,externalurl,@{Name="Internalauthenticationmethods";Expression={$_.Internalauthenticationmethods -join “;”}},@{Name="Externalauthenticationmethods";Expression={$_.Externalauthenticationmethods -join “;”}},IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation -Append
        $ClientAccess | ForEach-Object{Get-OABVirtualDirectory -ADPropertiesOnly | Select-Object server,name,internalurl,externalurl,@{Name="Internalauthenticationmethods";Expression={$_.Internalauthenticationmethods -join “;”}},@{Name="Externalauthenticationmethods";Expression={$_.Externalauthenticationmethods -join “;”}},IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation -Append
        $ClientAccess | ForEach-Object{Get-OutlookAnywhere -ADPropertiesOnly | Select-Object server,name,internalurl,externalurl,@{Name="Internalauthenticationmethods";Expression={$_.Internalauthenticationmethods -join “;”}},@{Name="Externalauthenticationmethods";Expression={$_.Externalauthenticationmethods -join “;”}},IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation -Append
        }
        else {
            $ClientAccess | ForEach-Object{Get-AutoDiscoverVirtualDirectory -ADPropertiesOnly | Select-Object server,name,internalurl,externalurl,@{Name="Internalauthenticationmethods";Expression={$_.Internalauthenticationmethods -join “;”}},@{Name="Externalauthenticationmethods";Expression={$_.Externalauthenticationmethods -join “;”}},IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod}
            $ClientAccess | ForEach-Object{Get-OwaVirtualDirectory -ADPropertiesOnly | Select-Object server,name,internalurl,externalurl,@{Name="Internalauthenticationmethods";Expression={$_.Internalauthenticationmethods -join “;”}},@{Name="Externalauthenticationmethods";Expression={$_.Externalauthenticationmethods -join “;”}},IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} 
            $ClientAccess | ForEach-Object{Get-ECPVirtualDirectory -ADPropertiesOnly | Select-Object server,name,internalurl,externalurl,@{Name="Internalauthenticationmethods";Expression={$_.Internalauthenticationmethods -join “;”}},@{Name="Externalauthenticationmethods";Expression={$_.Externalauthenticationmethods -join “;”}},IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} 
            $ClientAccess | ForEach-Object{Get-MAPIVirtualDirectory -ADPropertiesOnly | Select-Object server,name,internalurl,externalurl,@{Name="Internalauthenticationmethods";Expression={$_.Internalauthenticationmethods -join “;”}},@{Name="Externalauthenticationmethods";Expression={$_.Externalauthenticationmethods -join “;”}},IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} 
            $ClientAccess | ForEach-Object{Get-ActiveSyncVirtualDirectory -ADPropertiesOnly | Select-Object server,name,internalurl,externalurl,@{Name="Internalauthenticationmethods";Expression={$_.Internalauthenticationmethods -join “;”}},@{Name="Externalauthenticationmethods";Expression={$_.Externalauthenticationmethods -join “;”}},IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} 
            $ClientAccess | ForEach-Object{Get-WebServicesVirtualDirectory -ADPropertiesOnly | Select-Object server,name,internalurl,externalurl,@{Name="Internalauthenticationmethods";Expression={$_.Internalauthenticationmethods -join “;”}},@{Name="Externalauthenticationmethods";Expression={$_.Externalauthenticationmethods -join “;”}},IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} 
            $ClientAccess | ForEach-Object{Get-OABVirtualDirectory -ADPropertiesOnly | Select-Object server,name,internalurl,externalurl,@{Name="Internalauthenticationmethods";Expression={$_.Internalauthenticationmethods -join “;”}},@{Name="Externalauthenticationmethods";Expression={$_.Externalauthenticationmethods -join “;”}},IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} 
            $ClientAccess | ForEach-Object{Get-OutlookAnywhere -ADPropertiesOnly | Select-Object server,name,internalurl,externalurl,@{Name="Internalauthenticationmethods";Expression={$_.Internalauthenticationmethods -join “;”}},@{Name="Externalauthenticationmethods";Expression={$_.Externalauthenticationmethods -join “;”}},IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} 
                
        }

}
function Install-ExchangeOnline {    
    [CmdletBinding()]
    Param (
    
    )
   
        Set-ExecutionPolicy RemoteSigned
        #winrm set winrm/config/client/auth @{Basic="true"}
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Install-PackageProvider -Name NuGet -Force
        Install-Module -Name PowerShellGet -Force
        Update-Module -Name PowerShellGet
        Install-Module -Name ExchangeOnlineManagement -Confirm:$false
        Import-Module ExchangeOnlineManagement       

}

function Install-All {
    [CmdletBinding()]
    Param (
    
    )
        Install-ExchangeOnline
        Install-Module msonline
        Install-Module azureAD -AllowClobber
}
function Get-FrankensteinRecipientCounts {
    [CmdletBinding()]
    Param (
    )   

      #Define Variables
      "Building variables for recipient collection..."
      $AllMailboxes = Get-Mailbox -ResultSize Unlimited
      $AllDistGroups = Get-DistributionGroup -ResultSize Unlimited
      $CASMailbox = Get-CASMailbox -ResultSize Unlimited
      
      "Exchange Recipient Count
      
      "  
      $TotalMBXCount = ($AllMailboxes).count 
      Write-Host "$TotalMBXCount Total Mailboxes"

      $UserMBXCount = ($AllMailboxes | Where-Object{$_.recipienttypedetails -eq "UserMailbox"} | Measure-Object).count
      Write-Host "$UserMBXCount User Mailboxes"    
      
      $SharedMBXCount = ($AllMailboxes | Where-Object{$_.recipienttypedetails -eq "SharedMailbox"}| Measure-Object).count
      Write-Host "$SharedMBXCount Shared Mailboxes"
      
      $RoomMBXCount = ($AllMailboxes | Where-Object{$_.recipienttypedetails -eq "RoomMailbox"} | Measure-Object).count
      Write-Host "$RoomMBXCount Room Mailboxes"
    
      $EquipmentMBXCount = ($AllMailboxes | Where-Object{$_.recipienttypedetails -eq "EquipmentMailbox"} | Measure-Object).count
      Write-Host "$EquipmentMBXCount Equipment Mailboxes"

      $MailUserCount = (Get-MailUser -ResultSize Unlimited | Measure-Object).count 
      Write-Host "$MailUserCount MailUsers"

      $MailContactCount = (Get-MailContact -ResultSize Unlimited | Measure-Object).count 
      Write-Host "$MailContactCount Mail Contacts"

      $DistributionGroupCount = ($AllDistGroups | Measure-Object).count 
      Write-Host "$DistributionGroupCount Distribution Groups"

      $DynamicDistributionGroup = (Get-DynamicDistributionGroup -ResultSize Unlimited | Measure-Object).count 
      Write-Host "$DynamicDistributionGroup DynamicDistribution Groups"
      
    $UnifiedGroup = (Get-UnifiedGroup -ResultSize unlimited -ErrorAction SilentlyContinue).count
    Write-Host "$UnifiedGroup Unified Groups"    

      $LitHoldCount = ($AllMailboxes | Where-Object{$_.LitigationHoldEnabled -eq $TRUE} | Measure-Object).count 
      Write-Host "$LitHoldCount Mailboxes on Litigation Hold"

      $RetentionHoldCount = ($AllMailboxes | Where-Object{$_.RetentionHoldEnabled -eq $TRUE} | Measure-Object).count
      Write-Host "$RetentionHoldCount Mailboxes on Retention Hold"

      $GetPublicFolder = (Get-PublicFolder -recurse -ErrorAction SilentlyContinue | Measure-Object).count
      Write-Host "$GetPublicFolder Public Folders"

      $GetMailPublicFolder = (Get-MailPublicFolder -Resultsize Unlimited -ErrorAction SilentlyContinue | Measure-Object).count
      Write-Host "$GetMailPublicFolder Mail Public Folders"

      $GetPublicFolderMailbox = (Get-Mailbox -ResultSize unlimited -PublicFolder -ErrorAction SilentlyContinue | Measure-Object).count
      Write-Host "$GetPublicFolderMailbox Public Folder Mailboxes"

      $POP = ($CASMailbox | Where-Object{$_.popenabled -eq $true} | Measure-Object).count 
      Write-Host "$POP Mailboxes with POP3 Enabled"
      
      $IMAP = ($CASMailbox | Where-Object{$_.imapenabled -eq $true} | Measure-Object).count 
      Write-Host "$IMAP Mailboxes with IMAP Enabled"
      
      $MAPI = ($CASMailbox | Where-Object{$_.mapienabled -eq $true} | Measure-Object).count 
      Write-Host "$MAPI Mailboxes with MAPI Enabled"
      
      $ActiveSync = ($CASMailbox | Where-Object{$_.activesyncenabled -eq $true} | Measure-Object).count 
      Write-Host "$ActiveSync Mailboxes with ActiveSync Enabled"
      
      $OWA = ($CASMailbox | Where-Object{$_.owaenabled -eq $true} | Measure-Object).count 
      Write-Host "$OWA Mailboxes with OWA Enabled" 
      
      $ADPDisabled = ($AllMailboxes | Where-Object{$_.EmailAddressPolicyEnabled -eq $false} | Measure-Object).count 
      Write-Host "$ADPDisabled Mailboxes with Email Address Policy Disabled"       
            
}
function Connect-All {    
    [CmdletBinding()]
    Param (
    [Switch]$NoMFA
    ) 

    if($NoMFA)    {
        $AdminUsername = Read-Host -Prompt "Azure/Office 365 Admin User Account"
        $AdminPassword = Read-Host -Prompt "Password" -AsSecureString
        $adminCredentials = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $AdminUsername, $AdminPassword
    
        Connect-AzureAD -Credential $adminCredentials
        Connect-MSOLService -Credential $adminCredentials
        Connect-ExchangeOnline -Credential $adminCredentials
    }

    else {
        Connect-AzureAD 
        Connect-MSOLService 
        Connect-ExchangeOnline 
    }
} 
function Get-FrankensteinAzureDiscovery {    
    [CmdletBinding()]
    Param (
    [Switch]$CSV,
    [Switch]$UseCurrentSession
    )
   
    if($UseCurrentSession){

    }
    else {
    Connect-AzureAD
    Connect-MsolService
    }
        
    #Define Variables
    $MSOLUser = Get-MsolUser -All
    $Device = Get-MSOLDevice -all
    $Licenses = Get-MsolAccountSku

    mkdir .\FrankensteinAzureDiscovery_$((Get-Date).ToString('MMddyy'))
    Set-Location  .\FrankensteinAzureDiscovery_$((Get-Date).ToString('MMddyy'))

    Start-Transcript .\Get-FrankensteinAzureDiscovery_$((Get-Date).ToString('MMddyy')).txt

    Get-Linebreak
    "Get-MsolUser"
    if($CSV)    {
    Write-Host $MSOLUser.count "user's discovered"
    $MsolUser | Select-Object @{Name="AlternateEmailAddresses";Expression={$_.AlternateEmailAddresses -join “;”}},@{Name="AlternateMobilePhones";Expression={$_.AlternateMobilePhones -join “;”}},@{Name="AlternativeSecurityIds";Expression={$_.AlternativeSecurityIds -join “;”}},BlockCredential,City,CloudExchangeRecipientType,Country,Department,@{Name="DirSyncProvisioningErrors";Expression={$_.DirSyncProvisioningErrors -join “;”}},DisplayName,Errors,Fax,FirstName,ImmutableID,@{Name="IndirectLicenseErrors";Expression={$_.IndirectLicenseErrors -join “;”}},IsBlackberryUser,IsLicensed,LastDirSynced,LastName,LastPasswordChangeTimestamp,@{Name="LicenseAssignmentDetails";Expression={$_.LicenseAssignmentDetails -join “;”}},LicenseReconciliationNeeded,@{Name="Licenses";Expression={$_.Licenses -join “;”}},LiveId,MSExchRecipientTypeDetails,MSRtcSipDeploymentLocator,MSRtcSipPrimaryUserAddress,MobilePhone,ObjectId,Office,OverallProvisioningStatus,PasswordNeverExpires,PasswordResetNotRequiredDuringActivate,PhoneNumber,PortalSettings,PostalCode,PreferredDataLocation,PreferredLanguage,@{Name="ProxyAddresses";Expression={$_.ProxyAddresses -join “;”}},ReleaseTrack,@{Name="ServiceInformation";Expression={$_.ServiceInformation -join “;”}},SignInName,SoftDeletionTimestamp,State,StreetAddress,@{Name="StrongAuthenticationMethods";Expression={$_.StrongAuthenticationMethods -join “;”}},@{Name="StrongAuthenticationPhoneAppDetails";Expression={$_.StrongAuthenticationPhoneAppDetails -join “;”}},@{Name="StrongAuthenticationProofupTime";Expression={$_.StrongAuthenticationProofupTime -join “;”}},@{Name="StrongAuthenticationRequirements";Expression={$_.StrongAuthenticationRequirements -join “;”}},@{Name="StrongAuthenticationUserDetails";Expression={$_.StrongAuthenticationUserDetails -join “;”}},StrongPasswordRequired,StsRefreshTokensValidFrom,Title,UsageLocation,UserLandingPageIdentifierForO365Shell,UserPrincipalName,UserThemeIdentifierForO365Shell,UserType,ValidationStatus,WhenCreated  | Export-Csv .\MSOLUsers_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
    }
    else {
        Write-Host $MSOLUser.count "user's discovered"    
    }

    Get-Linebreak
    "Get-MsolCompanyInformation"
    Get-MsolCompanyInformation

    Get-Linebreak
    "Get-MsolAccountSku"
    if($CSV) {
    $Licenses | Select-Object AccountSkuID,ActiveUnits,WarningUnits,ConsumedUnits
    $Licenses | Select-Object AccountName,AccountSkuID,ActiveUnits,ConsumedUnits,LockedOutUnits,SKUID,SkuPartNumber,TargetClass,SuspendedUnits,WarningUnits | Export-Csv .\MSOLLicenses_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
    }
    else {
        $Licenses | Select-Object AccountSkuID,ActiveUnits,WarningUnits,ConsumedUnits 
    }

    Get-Linebreak
    "Get-MsolDevice"
    if($CSV){    
    Write-Host $Device.count "device's discovered"
    $Device |Select-Object Enabled,ObjectID,DeviceID,DisplayName,DeviceObjectVersion,DeviceOSType,DeviceOSVersion,DeviceTrustType,DeviceTrustLevel,@{Name="DevicePhysicalIds";Expression={$_.DevicePhysicalIds -join “;”}},ApproximateLastLogonTimestamp,@{Name="AlternativeSecurityIds";Expression={$_.AlternativeSecurityIds -join “;”}},DirSyncEnabled,LastDirSyncTime,RegisteredOwners,@{Name="GraphDeviceObject";Expression={$_.GraphDeviceObject -join “;”}}  | Export-Csv -NoTypeInformation .\MSOLDevices_$((Get-Date).ToString('MMddyy')).csv
    }
    else {
        Write-Host $Device.count "device's discovered"
    }

    Get-Linebreak
    "Get-MSOLDirSyncFeatures"
    Get-MSOLDirSyncFeatures

    Get-Linebreak
    if($CSV) {
    "Get-AzureADExtensionProperty"
    Get-AzureADExtensionProperty | Format-List
    Get-AzureADExtensionProperty | Select-Object Name,ObjectID,AppDisplayName,DataType,IsSyncedFromOnPremises,@{Name="TargetObjects";Expression={$_.TargetObjects -join “;”}} | Export-Csv .\AzureADExtensions_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
    }
    else {
        "Get-AzureADExtensionProperty"
        Get-AzureADExtensionProperty | Format-List
    }

    Stop-Transcript
}
function Get-FrankensteinGSuiteDiscovery {
    [CmdletBinding()]
    Param (
    [Switch]$CSV
    )

    if (Get-InstalledModule -Name PSGsuite -ErrorAction SilentlyContinue ) {
        Write-Host "PSGSuite Module detected, continuing with discovery"
        Start-Sleep -s 2
        Show-PSGSuiteConfig
        Start-Sleep -s 2
        
    } 
    else {        
        Write-Host "Foolish human. You must install the PSGsuite PowerShell Module to continue: https://psgsuite.io/"
        exit
    }

    Get-Linebreak
    "Processing GSUser user report....."
    $GSUserImport = Get-GSUserList

    mkdir .\GSuiteDiscovery_$((Get-Date).ToString('MMddyy')) 
    Set-Location  .\GSuiteDiscovery_$((Get-Date).ToString('MMddyy'))
    
    $GSUserImport | ForEach-Object{Get-GSUser -User $_.User} | Select-object User,PrimaryEmail,AgreedToTerms,@{Name="Aliases";Expression={$_.Aliases -join “;”}},Archived,ChangePasswordAtNextLogin,CreationTime,DeletionTime,Id,IncludeInGlobalAddressList,IpWhitelisted,IsAdmin,IsDelegate,IsEnforced,IsEnrolledIn2Sv,IsMailboxSetup,LastLoginTime,@{Name="NonEditableAliases";Expression={$_.NonEditableAliases -join “;”}},OrgUnitPath,@{Name="Organizations";Expression={$_.Organizations -join “;”}},@{Name="Phones";Expression={$_.Phones -join “;”}},RecoveryEmail,Suspended,SuspensionReason | Export-csv .\PSGsuiteUsers_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation

    Get-Linebreak
    "Processing GSUser alias report....."
    $GSUserImport | ForEach-Object{Get-GSUserAlias -user $_.User} | Select-object AliasValue,PrimaryEmail | Export-CSV .\PSGsuiteAlias_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation

    Get-Linebreak
    "Processing GSUser delegates....."
    $GSUserImport | ForEach-Object{Get-GSGmailDelegates -user $_.User -WarningAction SilentlyContinue} | Export-CSV .\PSGsuiteDelegates_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation   

    Get-Linebreak
    "Processing GSUser SendAS settings....."
    $GSUserImport | ForEach-Object{Get-GSGmailSendAsSettings -user $_.User -WarningAction SilentlyContinue} | Where-Object{$_.ISPrimary -ne $true} | Select-object User,DisplayName,IsDefault,IsPrimary,SendAsEmail,ReplyToAddress | Export-CSV .\PSGsuiteSendAs_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation   

    Get-Linebreak
    "Processing GSUser Auto Forwarding report....."
    $GSUserImport | ForEach-Object{Get-GSGmailAutoForwardingSettings -user $_.User} | Where-Object{$_.EmailAddress -ne $null} | Select-object User,Disposition,EmailAddress,Enabled | Export-CSV .\PSGsuiteAutoForwardSettings_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation

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

    if($UseCurrentSession){
    }
    elseif ($Online) {
        Connect-ExchangeOnline
    }
   else {
       Connect-ExchangeOnPremServer
   }
  
   if($Online){        
        mkdir .\Frankenstein_ExchangeOnline_Discovery_$((Get-Date).ToString('MMddyy'))
        Set-Location  .\Frankenstein_ExchangeOnline_Discovery_$((Get-Date).ToString('MMddyy'))        
        Start-Transcript -Path .\ExchangeOnline_DiscoveryTranscript_$((Get-Date).ToString('MMddyy')).txt
   }
   else {
        mkdir .\Frankenstein_ExchangeOnPrem_Discovery_$((Get-Date).ToString('MMddyy'))
        Set-Location  .\Frankenstein_ExchangeOnPrem_Discovery_$((Get-Date).ToString('MMddyy'))        
        Start-Transcript -Path .\ExchangeOnPrem_DiscoveryTranscript_$((Get-Date).ToString('MMddyy')).txt
   }
        
        Get-Linebreak
        Get-FrankensteinRecipientCounts                     

        
        if($online){
        }
        elseif($CSV){
        Get-Linebreak
        "Get-ExchangeServer"
        $ExchangeServers = Get-ExchangeServer
        $ExchangeServers|Format-List  
        $ExchangeServers|Select-Object Name,Domain,Edition,FQDN,IsHubTransportServer,IsClientAccessServer,IsEdgeServer,IsMailboxServer,IsUnifiedMessagingServer,IsFrontendTransportServer,OrganizationalUnit,AdminDisplayVersion,Site,ServerRole | Export-Csv .\ExchangeServers_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }        
        else {
            "Get-ExchangeServer"
            $ExchangeServers = Get-ExchangeServer
            $ExchangeServers|Format-List  
        }

        
        if($online){
        }
        elseif($CSV){
        Get-Linebreak
        "Get-ExchangeServerDatabase" 
        Get-MailboxDatabase
        Get-MailboxDatabase | Format-List
        Get-MailboxDatabase | Select-Object Name,Server,MailboxRetention,ProhibitSendReceiveQuota,ProhibitSendQuota,RecoverableItemsQuota,RecoverableItemsWarningQuota,IsExcludedFromProvisioning,ReplicationType,DeletedItemRetention,
        CircularLoggingEnabled, AdminDisplayVersion | Export-Csv .\Databases_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }        
        else {
            "Get-ExchangeServerDatabase" 
            Get-MailboxDatabase
            Get-MailboxDatabase | Format-List
        }        
        
        if ($online) {
           
        }
        elseif($CSV){
        Get-Linebreak
        "Get-DatabaseAvailabilityGroup"
        Get-DatabaseAvailabilityGroup
        Get-DatabaseAvailabilityGroup | Format-List
        Get-DatabaseAvailabilityGroup | Format-List | Export-Csv .\DAG__$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }
        else {
            "Get-DatabaseAvailabilityGroup"
            Get-DatabaseAvailabilityGroup
            Get-DatabaseAvailabilityGroup | Format-List 
        }
        
        Get-Linebreak
        "Get-RetentionPolicy"
        if($CSV){
        Get-RetentionPolicy
        Get-RetentionPolicy | Format-List
        Get-RetentionPolicy | Select-Object name,@{Name="RetentionPolicyTagLinks";Expression={$_.RetentionPolicyTagLinks -join “;”}} | Export-Csv .\RetentionPolicies_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }
        else {
            Get-RetentionPolicy
            Get-RetentionPolicy | Format-List
        }

        Get-Linebreak
        "Get-RetentionPolicyTag"
        if($CSV) {
        Get-RetentionPolicyTag
        Get-RetentionPolicyTag | Format-List
        Get-RetentionPolicyTag | Select-Object name,type,agelimitforretention,retentionaction | Export-Csv .\RetentionPoliciesTag_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }
        else {
            Get-RetentionPolicyTag
            Get-RetentionPolicyTag | Format-List  
        }

        Get-Linebreak
        "Get-JournalRule"
        if($CSV){
        Get-JournalRule
        Get-JournalRule | Format-List
        Get-JournalRule | Select-Object Name,Recipient,JournalEmailAddress,Scope,Enabled | Export-Csv .\JournalRules_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }
        else {
            Get-JournalRule
            Get-JournalRule | Format-List 
        }

        Get-Linebreak
        "Get-AcceptedDomain"
        if($CSV){
        $AcceptedDomain = Get-AcceptedDomain
        $AcceptedDomain
        $AcceptedDomain | Format-List
        $AcceptedDomain | Select-Object name,domainname,domaintype,default | Export-Csv -Path .\AcceptedDomains_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        Get-Linebreak
        "Domain MX Records"
        foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type MX}
        Get-Linebreak
        "Domain TXT Records"
        foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type TXT}
        Get-Linebreak
        "Domain CNAME Records"
        foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type CNAME}
        }
        else {
            $AcceptedDomain = Get-AcceptedDomain
            $AcceptedDomain
            $AcceptedDomain | Format-List
            "Domain MX Records"
            foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type MX}
            Get-Linebreak
            "Domain TXT Records"
            foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type TXT}
            Get-Linebreak
            "Domain CNAME Records"
            foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type CNAME}
            
        } 

        Get-Linebreak
        "Get-RemoteDomain"
        if($CSV){
        Get-RemoteDomain
        Get-RemoteDomain | Format-List
        Get-RemoteDomain | Select-Object name,domainname,allowedooftype | Export-Csv -Path .\RemoteDomains_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }
        else {
            Get-RemoteDomain
            Get-RemoteDomain | Format-List   
        }

        Get-Linebreak
        "Get-EmailAddressPolicy"
        if($CSV){
        Get-EmailAddressPolicy
        Get-EmailAddressPolicy | Format-List
        Get-EmailAddressPolicy | Select-Object Name,Priority,IncludedRecipients,@{Name="EnabledEmailAddressTemplates";Expression={$_.EnabledEmailAddressTemplates -join “;”}},RecipientFilterApplied | Export-Csv -Path .\EmailAddressPolicies_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }
        else {
            Get-EmailAddressPolicy
            Get-EmailAddressPolicy | Format-List   
        }
      
        Get-Linebreak
        "Get-TransportRule"
        if($CSV){
        Get-TransportRule
        Get-TransportRule | Format-List
        Get-TransportRule | Select-Object Name,Description, State, Priority | Export-Csv -Path .\TransportRules_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        $file = Export-TransportRuleCollection
        Set-Content -Path ".\Rules.xml" -Value $file.FileData -Encoding Byte
        }
        else {
            Get-TransportRule
            Get-TransportRule | Format-List
            
        }

        Get-Linebreak        
        if($Online -and $CSV){
        "Get-OutboundConnector"
        Get-OutboundConnector
        Get-OutboundConnector | Format-List
        Get-OutboundConnector | Select-Object name,@{Name="SmartHosts";Expression={$_.SmartHosts -join “;”}},Enabled,@{Name="AddressSpaces";Expression={$_.AddressSpaces -join “;”}},@{Name="SourceTransportServers";Expression={$_.SourceTransportServers -join “;”}},FQDN,MaxMessageSize,ProtocolLoggingLevel,RequireTLS |Export-Csv -Path .\OutboundConnectors_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }
        elseif($Online){
            "Get-OutboundConnector"
            Get-OutboundConnector
            Get-OutboundConnector | Format-List
        }
        elseif($CSV) {
        "Get-SendConnector"
        Get-SendConnector
        Get-SendConnector | Format-List
        Get-SendConnector | Select-Object name,@{Name="SmartHosts";Expression={$_.SmartHosts -join “;”}},Enabled,@{Name="AddressSpaces";Expression={$_.AddressSpaces -join “;”}},@{Name="SourceTransportServers";Expression={$_.SourceTransportServers -join “;”}},FQDN,MaxMessageSize,ProtocolLoggingLevel,RequireTLS |Export-Csv -Path .\SendConnectors_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }
        else {
            "Get-SendConnector"
            Get-SendConnector
            Get-SendConnector | Format-List
        }

        Get-Linebreak
        if($Online -and $CSV){
        "Get-InboundConnector"
        Get-InboundConnector
        Get-InboundConnector | Format-List
        Get-InboundConnector | Select-Object name,authmechanism,@{Name="Bindings";Expression={$_.Bindings -join “;”}},enabled,@{Name="RemoteIPRanges";Expression={$_.RemoteIPRanges -join “;”}},requireTLS,originatingserver | Export-Csv -Path .\InboundConnectors_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }
        elseif ($Online) {
            "Get-InboundConnector"
            Get-InboundConnector
            Get-InboundConnector | Format-List            
        }
        elseif($CSV){
        "Get-ReceiveConnector"
        Get-ReceiveConnector
        Get-ReceiveConnector | Format-List
        Get-ReceiveConnector | Select-Object name,authmechanism,@{Name="Bindings";Expression={$_.Bindings -join “;”}},enabled,@{Name="RemoteIPRanges";Expression={$_.RemoteIPRanges -join “;”}},requireTLS,originatingserver | Export-Csv -Path .\ReceiveConnectors_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }
        else {
            Get-ReceiveConnector
            Get-ReceiveConnector | Format-List
        }
            
        
        if($Online){
        }
        else{
        Get-Linebreak
        "Get-TransportAgent"
        Get-TransportAgent
        Get-TransportAgent | Format-List
        }

        if($Online){

        }
        else{
        Get-Linebreak
        "Get-AddressList"
        Get-AddressList
        Get-AddressBookPolicy
        Start-Sleep -s 5
        }    

        Get-Linebreak
        "Get-OrganizationConfig"
        Get-OrganizationConfig | Format-List

        Get-Linebreak
        "Get-FederationTrust"
        Get-FederationTrust
        Get-FederationTrust | Format-List
        Get-Linebreak

        "Get-OrganizationRelationship"
        if($CSV){
        Get-OrganizationRelationship
        Get-OrganizationRelationship | Format-List
        Get-OrganizationRelationship | Select-Object name,@{Name="DomainNames";Expression={$_.DomainNames -join “;”}},targetautodiscoverepr,targetowaurl,targetsharingepr,targetapplicationuri,enabled |Export-Csv -Path .\OrganizationRelationships_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }
        else {
            Get-OrganizationRelationship
            Get-OrganizationRelationship | Format-List 
        }

        Get-Linebreak
        "Get-IntraOrganizationConnector"
        Get-IntraOrganizationConnector | Format-List
        "Get-IntraOrganizationConfiguration"
        Get-IntraOrganizationConfiguration
               
        if($Online){

        }
        elseif($CSV){
        Get-Linebreak 
        "Get-ExchangeCertificate"
        Get-ExchangeCertificate
        Get-ExchangeCertificate | Format-List
        Get-ExchangeCertificate | Select-Object subject,Issuer,Thumbprint,FriendlyName,NotAfter | Export-Csv .\ExchangeCertificates_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }
        else {
            "Get-ExchangeCertificate"
            Get-ExchangeCertificate
            Get-ExchangeCertificate | Format-List
        }

        
        if($Online){

        }
        else{
        Get-Linebreak
        "Get-HybridConfiguration"
        $Hybrid = Get-HybridConfiguration 
        if($Hybrid -ne $null)
        {
            foreach($result in $Hybrid)
            {
                "Hybrid configuration detected"
                $Hybrid 
            }
        }
            else {
                "No hybrid configuration detected"
            }
        }        


        Get-Linebreak

        
#Call Functions        
        if($VirtualDirectories){
        Get-FrankensteinVirtualDirectories
        }

        if($PublicFolders){
        Get-FrankensteinPublicFolderDiscovery
        }
  
        Stop-Transcript
}
function Get-FrankensteinPublicFolderDiscovery {    
    [CmdletBinding()]
    Param (
    
    )
    Get-Linebreak
    "Getting Public Folders..."
    $PF = Get-PublicFolder -Recurse -ErrorAction SilentlyContinue -ErrorVariable ProcessError
    $PF | Select-Object RunspaceId,Identity,Name,MailEnabled,MailRecipientGuid,ParentPath,LostAndFoundFolderOriginalPath,ContentMailboxName,ContentMailboxGuid,PerUserReadStateEnabled,EntryId,DumpsterEntryId,ParentFolder,OrganizationId,AgeLimit,RetainDeletedItemsFor,ProhibitPostQuota,IssueWarningQuota,MaxItemSize,LastMovedTime,AdminFolderFlags,FolderSize,HasSubfolders,FolderClass,FolderPath,AssociatedDumpsterFolders,DefaultFolderType,ExtendedFolderFlags,MailboxOwnerId,IsValid,ObjectState | Export-CSV .\Get_PublicFolder_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
    

    Get-Linebreak
    "Getting Mail Public Folders..."
    $MPF = Get-MailPublicFolder -ResultSize unlimited -ErrorAction SilentlyContinue  
    $MPF | Select-Object RunspaceId,DisplayName,PrimarySmtpAddress,@{Name="EmailAddresses";Expression={$_.EmailAddresses -join “;”}},Contacts,ContentMailbox,DeliverToMailboxAndForward,ExternalEmailAddress,OnPremisesObjectId,IgnoreMissingFolderLink,ForwardingAddress,AcceptMessagesOnlyFrom,AcceptMessagesOnlyFromDLMembers,AcceptMessagesOnlyFromSendersOrMembers,GrantSendOnBehalfTo,AddressListMembership,AdministrativeUnits,Alias,ArbitrationMailbox,BypassModerationFromSendersOrMembers,OrganizationalUnit,HiddenFromAddressListsEnabled,LastExchangeChangedTime,LegacyExchangeDN,MaxSendSize,MaxReceiveSize,ModerationEnabled,ModeratedBy,EmailAddressPolicyEnabled,RequireSenderAuthenticationEnabled,WindowsEmailAddress,WhenChanged,WhenCreated,ExchangeObjectId,Guid | Export-CSV .\Get_MailPublicFolder_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation

    Get-Linebreak
    "Getting Public Folder Mailboxes..."
    $PFM = Get-Mailbox -PublicFolder -ResultSize Unlimited -ErrorAction SilentlyContinue -ErrorVariable ProcessError
    $PFM | Select-Object RunspaceId,DisplayName,PrimarySmtpAddress,LegacyExchangeDN,Database,DeliverToMailboxAndForward,IsHierarchyReady,IsHierarchySyncEnabled,LitigationHoldEnabled,SingleItemRecoveryEnabled,RetentionHoldEnabled,EndDateForRetentionHold,StartDateForRetentionHold,LitigationHoldDate,LitigationHoldOwner,LitigationHoldDuration,ComplianceTagHoldApplied,DelayHoldApplied,RetentionPolicy,AddressBookPolicy,ExchangeGuid,@{Name="MailboxLocations";Expression={$_.MailboxLocations -join “;”}},ExchangeUserAccountControl,AdminDisplayVersion,ForwardingAddress,ForwardingSmtpAddress,RetainDeletedItemsFor,IsMailboxEnabled,ProhibitSendQuota,ProhibitSendReceiveQuota,RecoverableItemsQuota,RecoverableItemsWarningQuota,CalendarLoggingQuota,RecipientLimits,ImListMigrationCompleted,IsRootPublicFolderMailbox,LinkedMasterAccount,SamAccountName,UserPrincipalName,RoleAssignmentPolicy,SharingPolicy,@{Name="EmailAddresses";Expression={$_.EmailAddresses -join “;”}},MaxSendSize,MaxReceiveSize,ModerationEnabled,ModeratedBy,RecipientTypeDetails,WhenChanged,WhenCreated | Export-CSV .\Get_MailboxPF_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
  

}



#Accept input paramenters

function Get-FrankensteinMailboxPermissions {    
    [CmdletBinding()]
    Param (
        [switch]$FullAccess,
        [switch]$SendAs,
        [switch]$SendOnBehalf,
        [switch]$UserMailboxOnly,
        [switch]$AdminsOnly,
        [string]$MBNamesFile,
        [Switch]$UseCurrentSession,
        [Switch]$Help
        )
        
        if($Help)
        {Write-Host
        
        "
        .SYNOPSIS
            Retreives Full Access, SendAS and SendOnBehalf permissions.
        
        .DESCRIPTION
            Retreives Full Access, SendAS and SendOnBehalf permissions.
            Requires minimum of Exchange Reader. Global Reader will not work.
        
        .PARAMETER 
            
        
        .EXAMPLE
            .\FrankensteinPermissions.ps1 -UseCurrentSession -FullAccess -SendAs -SendOnBehalf 
        
        
        .INPUTS
            CSV - Must Include DisplayName header
        
        .OUTPUTS
            CSV
            
        
        .NOTES
            Author:  Eric D. Frank
            09/13/23 - Updated to use GitHub as repository
          
        "
        }
        else
        {
        
        function Print_Output
        {
        
        #Connect to Exchange Online
            if($UseCurrentSession){
        }
        
        else {
            Connect-ExchangeOnline
        }
        
            #Mailbox type based filter
            if(($UserMailboxOnly.IsPresent) -and ($MBType -ne "UserMailbox"))
            { 
            $Print=0 
            }
        
            #Admin Role based filter
            if(($AdminsOnly.IsPresent) -and ($RolesAssigned -eq "No roles"))
            { 
            $Print=0 
            }
        
            #Print Output
            if($Print -eq 1)
            {
            $Result = @{'DisplayName'=$_.Displayname;'UserPrinciPalName'=$upn;'MailboxType'=$MBType;'AccessType'=$AccessType;'UserWithAccess'=$userwithAccess;'Roles'=$RolesAssigned} 
            $Results = New-Object PSObject -Property $Result 
            $Results |select-object DisplayName,UserPrinciPalName,MailboxType,AccessType,UserWithAccess,Roles | Export-Csv -Path $ExportCSV -Notype -Append 
            }
        }

   
        
        #Getting Mailbox permission
        function Get_MBPermission
            {
            $upn=$_.UserPrincipalName
            $DisplayName=$_.Displayname
            $MBType=$_.RecipientTypeDetails
            $Print=0
            Write-Progress -Activity "`n     Processed mailbox count: $MBUserCount "`n"  Currently Processing: $DisplayName"
        
            #Getting delegated Fullaccess permission for mailbox
            if(($FilterPresent -eq 'False') -or ($FullAccess.IsPresent))
            {
            $FullAccessPermissions=(Get-MailboxPermission -Identity $upn | Where-Object{ ($_.AccessRights -contains "FullAccess") -and ($_.IsInherited -eq $false) -and -not ($_.User -match "NT AUTHORITY" -or $_.User -match "S-1-5-21") }).User
            if([string]$FullAccessPermissions -ne "")
            {
            $Print=1
            $UserWithAccess=""
            $AccessType="FullAccess"
            foreach($FullAccessPermission in $FullAccessPermissions)
            {
                $UserWithAccess=$UserWithAccess+$FullAccessPermission
                if($FullAccessPermissions.indexof($FullAccessPermission) -lt (($FullAccessPermissions.count)-1))
                {
                $UserWithAccess=$UserWithAccess+","
                }
            }
            Print_Output
            }
            }
        
            #Getting delegated SendAs permission for mailbox
            if(($FilterPresent -eq 'False') -or ($SendAs.IsPresent))
            {
            $SendAsPermissions=(Get-RecipientPermission -Identity $upn | Where-Object{ -not (($_.Trustee -match "NT AUTHORITY") -or ($_.Trustee -match "S-1-5-21"))}).Trustee
            if([string]$SendAsPermissions -ne "")
            {
            $Print=1
            $UserWithAccess=""
            $AccessType="SendAs"
            foreach($SendAsPermission in $SendAsPermissions)
            {
                $UserWithAccess=$UserWithAccess+$SendAsPermission
                if($SendAsPermissions.indexof($SendAsPermission) -lt (($SendAsPermissions.count)-1))
                {
                $UserWithAccess=$UserWithAccess+","
                }
            }
            Print_Output
            }
            }
            
            #Getting delegated SendOnBehalf permission for mailbox
            if(($FilterPresent -eq 'False') -or ($SendOnBehalf.IsPresent))
            {
            $SendOnBehalfPermissions=$_.GrantSendOnBehalfTo
            if([string]$SendOnBehalfPermissions -ne "")
            {
                $Print=1
                $UserWithAccess=""
                $AccessType="SendOnBehalf"
                foreach($SendOnBehalfPermissionDN in $SendOnBehalfPermissions)
                {
                $SendOnBehalfPermission=(Get-Mailbox -Identity $SendOnBehalfPermissionDN).UserPrincipalName
                $UserWithAccess=$UserWithAccess+$SendOnBehalfPermission
                if($SendOnBehalfPermissions.indexof($SendOnBehalfPermission) -lt (($SendOnBehalfPermissions.count)-1))
                {
                $UserWithAccess=$UserWithAccess+","
                }
                }
                Print_Output
            }
            }
            }
        
        
        function main{
            #Connect AzureAD and Exchange Online from PowerShell
            #Get-PSSession | Remove-PSSession
        
            #Check for MSOnline module
            $Modules=Get-Module -Name MSOnline -ListAvailable 
            if($Modules.count -eq 0)
            {
            Write-Host  Please install MSOnline module using below command: -ForegroundColor yellow 
            Write-Host Install-Module MSOnline  
            Exit
            }
            
            #Set output file
            $ExportCSV=".\MBPermission_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
            $Result="" 
            $Results=@()
            $MBUserCount=0
            $RolesAssigned=""
        
            #Check for AccessType filter
            if(($FullAccess.IsPresent) -or ($SendAs.IsPresent) -or ($SendOnBehalf.IsPresent))
            {}
            else
            {
            $FilterPresent='False'
            }
        
            #Check for input file
            if ($MBNamesFile -ne "") 
            { 
            #We have an input file, read it into memory 
            $MBs=@()
            $MBs=Import-Csv -Header "DisplayName" $MBNamesFile
            foreach($item in $MBs)
            {
            Get-Mailbox -Identity $item.displayname | ForEach-Object{
            $MBUserCount++
            Get_MBPermission
            }
            }
            }
            #Getting all User mailbox
            else
            {
            Get-mailbox -ResultSize Unlimited | Where-Object{$_.DisplayName -notlike "Discovery Search Mailbox"} | ForEach-Object{
            $MBUserCount++
            Get_MBPermission}
            }
        
            
            #Open output file after execution 
        Write-Host `nScript executed successfully
        if((Test-Path -Path $ExportCSV) -eq "True")
        {
            Write-Host "Detailed report available in: $ExportCSV" 
            $Prompt = New-Object -ComObject wscript.shell  
            $UserInput = $Prompt.popup("Do you want to export results to .CSV?",`
            0,"Open Output File",4)  
            If ($UserInput -eq 6)  
            {  
            Invoke-Item "$ExportCSV"  
            } 
        }
        Else
        {
            Write-Host No mailbox found that matches your criteria.
        }
        
        }
            . main
}            
        }