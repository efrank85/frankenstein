


<#
.SYNOPSIS
    Test creation of Eric Frank

.DESCRIPTION
    This module contains a series of functions used to collect and export data in preparation from an Exchange to Exchange Online migration.

.PARAMETER 
    Get-FrankensteinHelp: View all Functions in this module

.EXAMPLE
    Get-FrankensteinExchangeOnlineDiscovery -VirtualDirectories 


.INPUTS
    

.OUTPUTS
    CSV and .txt files
    

.NOTES
    Author:  Eric D. Frank
    12/6/21 - Updated to use GitHub as repository
  
#>
 

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

function Get-FrankensteinHelp {    
    [CmdletBinding()]
    Param (
        )    
        Write-Host "
        
        Frankenstein offers several modules to assist in Exchange and Azure discovery processes. Below represents a brief explanation of each:

        1) Get-FrankensteinExchangeDiscovery: Provides Exchange on-premises discovery information and outputs a transcript along with various CSV outputs. 
        You must be connected to Exchange PowerShell prior to launching this module.

        [-virtualdirectories]

        2) Get-FrankensteinExchangeOnlineDiscovery: Provides Exchange Online discovery information and outputs a transcript along with various CSV outputs. 
        This function will automatically attempt to connect to Exchange Online and prompt for credentials.

        [-virtualdirectories]

        3) Install-ExchangeOnline: Will install and configure Exchange Online PowerShell requirements to run Connect-ExchangeOnline

        4) Connect-All: Will connect to MSOL, AzureAD and ExO PS Sessions

        [-noMFA]

        5) Connect-OnPremServer: Connects to on-premises Exchange server using FQDN
                
                "
        }


function Get-FrankensteinVirtualDirectories {    
    [CmdletBinding()]
    Param (
    )
      
       
        Get-Linebreak

        Write-Host "Get-VirtualDirectories"
       
        $ClientAccess = Get-ClientAccessService
        $ClientAccess | ForEach-Object{Get-AutoDiscoverVirtualDirectory | Select-Object server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories.csv -NoTypeInformation
        $ClientAccess | ForEach-Object{Get-OwaVirtualDirectory | Select-Object server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories.csv -NoTypeInformation -Append
        $ClientAccess | ForEach-Object{Get-ECPVirtualDirectory | Select-Object server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories.csv -NoTypeInformation -Append
        $ClientAccess | ForEach-Object{Get-MAPIVirtualDirectory | Select-Object server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories.csv -NoTypeInformation -Append
        $ClientAccess | ForEach-Object{Get-ActiveSyncVirtualDirectory | Select-Object server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories.csv -NoTypeInformation -Append
        $ClientAccess | ForEach-Object{Get-WebServicesVirtualDirectory | Select-Object server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories.csv -NoTypeInformation -Append
        $ClientAccess | ForEach-Object{Get-OutlookAnywhere | Select-Object server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories.csv -NoTypeInformation -Append
        

        }


function Get-FrankensteinExchangeDiscovery {    
    [CmdletBinding()]
    Param (
    [Switch]$virtualDirectories
    
    )
   
        
        Start-Transcript -Path .\ExchangeDiscoveryTranscript.txt
        
        Write-Host Exchange Recipient Count
        
        #Define Variables
        $AllMailboxes = Get-Mailbox -ResultSize Unlimited -IgnoreDefaultScope
        $AllDistGroups = Get-DistributionGroup -ResultSize Unlimited -IgnoreDefaultScope 
        $ExchangeServers = Get-ExchangeServer
        #$ClientAccess = Get-ClientAccessService
   
        
        #Get Recipient Types
        $TotalMBXCount = ($AllMailboxes).count 
        Write-Host "$TotalMBXCount Total Mailboxes"

        $UserMBXCount = (Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox).count 
        Write-Host "$UserMBXCount User Mailboxes"        
        
        $SharedMBXCount = (Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox).count 
        Write-Host "$SharedMBXCount Shared Mailboxes"
        
        $RoomMBXCount = (Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails RoomMailbox).count 
        Write-Host "$RoomMBXCount Room Mailboxes"
      
        $EquipmentMBXCount = (Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails equipmentmailbox).count
        Write-Host "$EquipmentMBXCount Equipment Mailboxes"

        $LinkedMbxCount = (Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails LinkedMailbox).count 
        Write-Host "$LinkedMbxCount Linked Mailboxes"

        $RemoteMbxCount = (Get-RemoteMailbox -ResultSize Unlimited).count 
        Write-Host "$RemoteMbxCount Remote Mailboxes"

        $MailUserCount = (Get-MailUser -ResultSize Unlimited).count 
        Write-Host "$MailUserCount MailUsers"

        $MailContactCount = (Get-MailContact -ResultSize Unlimited).count 
        Write-Host "$MailContactCount Mail Contacts"

        $DistributionGroupCount = ($AllDistGroups).count 
        Write-Host "$DistributionGroupCount Distribution Groups"

        $DynamicDistributionGroup = (Get-DynamicDistributionGroup -ResultSize Unlimited).count 
        Write-Host "$DynamicDistributionGroup DynamicDistribution Groups"

        $LitHoldCount = ($AllMailboxes | Where-Object{$_.LitigationHoldEnabled -eq $TRUE}).count 
        Write-Host "$LitHoldCount Mailboxes on Litigation Hold"

        $RetentionHoldCount = ($AllMailboxes | Where-Object{$_.RetentionHoldEnabled -eq $TRUE}).count
        Write-Host "$RetentionHoldCount Mailboxes on Retention Hold"

        $GetPublicFolder = (Get-PublicFolder -recurse).count
        Write-Host "$GetPublicFolder Public Folders"

        $GetMailPublicFolder = (Get-MailPublicFolder).count
        Write-Host "$GetMailPublicFolder Mail Public Folders"

        $GetPublicFolderMailbox = (Get-Mailbox -ResultSize unlimited -PublicFolder -IgnoreDefaultScope).count
        Write-Host "$GetPublicFolderMailbox Public Folder Mailboxes"

        $POP = ($CASMailbox | Where-Object{$_.popenabled -eq $true}).count 
        Write-Host "$POP Mailboxes with POP3 Enabled"
        
        $IMAP = ($CASMailbox | Where-Object{$_.imapenabled -eq $true}).count 
        Write-Host "$IMAP Mailboxes with IMAP Enabled"
        
        $MAPI = ($CASMailbox | Where-Object{$_.mapienabled -eq $true}).count 
        Write-Host "$MAPI Mailboxes with MAPI Enabled"
        
        $ActiveSync = ($CASMailbox | Where-Object{$_.activesyncenabled -eq $true}).count 
        Write-Host "$ActiveSync Mailboxes with ActiveSync Enabled"
        
        $OWA = ($CASMailbox | Where-Object{$_.owaenabled -eq $true}).count 
        Write-Host "$OWA Mailboxes with OWA Enabled" 
        
        $ADPDisabled = ($AllMailboxes | Where-Object{$_.EmailAddressPolicyEnabled -eq $false}).count 
        Write-Host "$ADPDisabled Mailboxes with Email Address Policy Disabled"     
                

        Get-Linebreak
        "Get-ExchangeServer" 
        $ExchangeServers
        $ExchangeServers|Format-List

        Get-Linebreak
        "Get-ExchangeServerDatabase" 
        Get-MailboxDatabase
        Get-MailboxDatabase | Format-List
        Get-MailboxDatabase | Select-Object Name,Server,MailboxRetention,ProhibitSendReceiveQuota,ProhibitSendQuota,RecoverableItemsQuota,RecoverableItemsWarningQuota,IsExcludedFromProvisioning,ReplicationType,DeletedItemRetention,
        CircularLoggingEnabled, AdminDisplayVersion | Export-Csv .\Databases.csv -NoTypeInformation
        
        Get-Linebreak
        "Get-DatabaseAvailabilityGroup"
        Get-DatabaseAvailabilityGroup
        Get-DatabaseAvailabilityGroup | Format-List
        
        Get-Linebreak
        "Get-RetentionPolicy"
        Get-RetentionPolicy
        Get-RetentionPolicy | Format-List
        Get-RetentionPolicy | Select-Object name,@{Name="RetentionPolicyTagLinks";Expression={$_.RetentionPolicyTagLinks -join “;”}} | Export-Csv .\RetentionPolicies.csv -NoTypeInformation
        
        Get-Linebreak
        "Get-RetentionPolicyTag"
        Get-RetentionPolicyTag
        Get-RetentionPolicyTag | Format-List
        Get-RetentionPolicyTag | Select-Object name,type,agelimitforretention,retentionaction | Export-Csv .\RetentionPoliciesTag.csv -NoTypeInformation

        Get-Linebreak
        "Get-JournalRule"
        Get-JournalRule
        Get-JournalRule | Format-List

        Get-Linebreak
        "Get-AcceptedDomain"
        $AcceptedDomain = Get-AcceptedDomain
        $AcceptedDomain
        $AcceptedDomain | Format-List
        $AcceptedDomain | Select-Object name,domainname,domaintype,default | Export-Csv -Path .\AcceptedDomains.csv -NoTypeInformation
        foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type MX}
        foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type TXT}
        foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type CNAME} 

        Get-Linebreak
        "Get-EmailAddressPolicy"
        Get-EmailAddressPolicy
        Get-EmailAddressPolicy | Format-List
        Get-EmailAddressPolicy | Select-Object Name,Priority,IncludedRecipients,@{Name="EnabledEmailAddressTemplates";Expression={$_.EnabledEmailAddressTemplates -join “;”}},RecipientFilterApplied | Export-Csv -Path .\EmailAddressPolicies.csv -NoTypeInformation
        
      
        Get-Linebreak
        "Get-TransportRule"
        Get-TransportRule
        Get-TransportRule | Format-List
        Get-TransportRule | Select-Object Name,Description, State, Priority | Export-Csv -Path .\TransportRules.csv -NoTypeInformation
        $file = Export-TransportRuleCollection
        Set-Content -Path ".\Rules.xml" -Value $file.FileData -Encoding Byte

        Get-Linebreak
        "Get-SendConnector"
        Get-SendConnector
        Get-SendConnector | Format-List
        Get-SendConnector | Select-Object name,@{Name="SmartHosts";Expression={$_.SmartHosts -join “;”}},Enabled,@{Name="AddressSpaces";Expression={$_.AddressSpaces -join “;”}},@{Name="SourceTransportServers";Expression={$_.SourceTransportServers -join “;”}},FQDN,MaxMessageSize,ProtocolLoggingLevel,RequireTLS |Export-Csv -Path .\SendConnectors.csv -NoTypeInformation

        Get-Linebreak
        "Get-ReceiveConnector"
        Get-ReceiveConnector
        Get-ReceiveConnector | Format-List
        Get-ReceiveConnector | Select-Object name,authmechanism,@{Name="Bindings";Expression={$_.Bindings -join “;”}},enabled,@{Name="RemoteIPRanges";Expression={$_.RemoteIPRanges -join “;”}},requireTLS,originatingserver | Export-Csv -Path .\ReceiveConnectors.csv -NoTypeInformation

        Get-Linebreak
        "Get-TransportAgent"
        Get-TransportAgent
        Get-TransportAgent | Format-List

        Get-Linebreak
        "Get-AddressList"
        Get-AddressList
        Get-AddressBookPolicy
        Start-Sleep -s 5

        Get-Linebreak
        "Get-PublicFolder"
        Get-PublicFolder -Recurse
        Start-Sleep -s 5
        "Get-MailPublicFolder"
        Get-MailPublicFolder -ResultSize unlimited
        Start-Sleep -s 5
        "Get-PublicFolderMailbox"
        Get-Mailbox -PublicFolder -ResultSize unlimited
        Start-Sleep -s 5


        Get-Linebreak
        "Get-OrganizationConfig"
        Get-OrganizationConfig
        Start-Sleep -s 5

        Get-Linebreak
        "Get-FederationTrust"
        Get-FederationTrust
        Get-FederationTrust | Format-List
        Get-Linebreak

        "Get-OrganizationRelationship"
        Get-OrganizationRelationship
        Get-OrganizationRelationship | Format-List
        Get-OrganizationRelationship | Select-Object name,@{Name="DomainNames";Expression={$_.DomainNames -join “;”}},targetautodiscoverepr,targetowaurl,targetsharingepr,targetapplicationuri,enabled |Export-Csv -Path .\OrganizationRelationships.csv -NoTypeInformation

        Get-Linebreak
        "Get-RemoteDomain"
        Get-RemoteDomain
        Get-RemoteDomain | Format-List
        Get-RemoteDomain | Select-Object name,domainname,allowedooftype | Export-Csv -Path .\RemoteDomains.csv -NoTypeInformation

        Get-Linebreak
        "Get-ExchangeCertificate"
        Get-ExchangeCertificate
        Get-ExchangeCertificate | Format-List
        Get-ExchangeCertificate | Select-Object subject,Issuer,Thumbprint,FriendlyName,NotAfter | Export-Csv .\ExchangeCertificates.csv -NoTypeInformation

        Get-Linebreak
        "Get-HybridConfiguration"
        $Hybrid = Get-HybridConfiguration 
        if($Hybrid -ne $null)
        {
            foreach($result in $Hybrid)
            {
                $Hybrid 
            }
        }
            else {
                "No hybrid configuration detected"
            }
        

        Start-Sleep -s 5

        Get-Linebreak

        
#Call Functions        
        if($VirtualDirectories){
        Get-FrankensteinVirtualDirectories
        }
  
        Stop-Transcript
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


function Get-FrankensteinExchangeOnlineDiscovery {    
    [CmdletBinding()]
    Param (
    
    )
   
        #Define Variables
        $AllMailboxes = Get-Mailbox -ResultSize Unlimited
        $AllDistGroups = Get-DistributionGroup -ResultSize Unlimited
        $CASMailbox = Get-CASMailbox
        
        Connect-ExchangeOnline

        Start-Transcript -Path .\ExchangeOnlineDiscoveryTranscript.txt
        
        Write-Host Exchange Recipient Count
        

        
        #Get Recipient Types
        $TotalMBXCount = ($AllMailboxes).count 
        Write-Host "$TotalMBXCount Total Mailboxes"

        $UserMBXCount = (Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox).count 
        Write-Host "$UserMBXCount User Mailboxes"        
        
        $SharedMBXCount = (Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox).count 
        Write-Host "$SharedMBXCount Shared Mailboxes"
        
        $RoomMBXCount = (Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails RoomMailbox).count 
        Write-Host "$RoomMBXCount Room Mailboxes"
      
        $EquipmentMBXCount = (Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails equipmentmailbox).count
        Write-Host "$EquipmentMBXCount Equipment Mailboxes"

        $MailUserCount = (Get-MailUser -ResultSize Unlimited).count 
        Write-Host "$MailUserCount MailUsers"

        $MailContactCount = (Get-MailContact -ResultSize Unlimited).count 
        Write-Host "$MailContactCount Mail Contacts"

        $DistributionGroupCount = ($AllDistGroups).count 
        Write-Host "$DistributionGroupCount Distribution Groups"

        $DynamicDistributionGroup = (Get-DynamicDistributionGroup -ResultSize Unlimited).count 
        Write-Host "$DynamicDistributionGroup DynamicDistribution Groups"

        $LitHoldCount = ($AllMailboxes | Where-Object{$_.LitigationHoldEnabled -eq $TRUE}).count 
        Write-Host "$LitHoldCount Mailboxes on Litigation Hold"

        $RetentionHoldCount = ($AllMailboxes | Where-Object{$_.RetentionHoldEnabled -eq $TRUE}).count
        Write-Host "$RetentionHoldCount Mailboxes on Retention Hold"

        $GetPublicFolder = (Get-PublicFolder -recurse).count
        Write-Host "$GetPublicFolder Public Folders"

        $GetMailPublicFolder = (Get-MailPublicFolder).count
        Write-Host "$GetMailPublicFolder Mail Public Folders"

        $GetPublicFolderMailbox = (Get-Mailbox -ResultSize unlimited -PublicFolder).count
        Write-Host "$GetPublicFolderMailbox Public Folder Mailboxes"

        $POP = ($CASMailbox | Where-Object{$_.popenabled -eq $true}).count 
        Write-Host "$POP Mailboxes with POP3 Enabled"
        
        $IMAP = ($CASMailbox | Where-Object{$_.imapenabled -eq $true}).count 
        Write-Host "$IMAP Mailboxes with IMAP Enabled"
        
        $MAPI = ($CASMailbox | Where-Object{$_.mapienabled -eq $true}).count 
        Write-Host "$MAPI Mailboxes with MAPI Enabled"
        
        $ActiveSync = ($CASMailbox | Where-Object{$_.activesyncenabled -eq $true}).count 
        Write-Host "$ActiveSync Mailboxes with ActiveSync Enabled"
        
        $OWA = ($CASMailbox | Where-Object{$_.owaenabled -eq $true}).count 
        Write-Host "$OWA Mailboxes with OWA Enabled" 
        
        $ADPDisabled = ($AllMailboxes | Where-Object{$_.EmailAddressPolicyEnabled -eq $false}).count 
        Write-Host "$ADPDisabled Mailboxes with Email Address Policy Disabled"     
                

        Get-Linebreak
        "Get-RetentionPolicy"
        Get-RetentionPolicy
        Get-RetentionPolicy | Format-List
        Get-RetentionPolicy | Select-Object name,@{Name="RetentionPolicyTagLinks";Expression={$_.RetentionPolicyTagLinks -join “;”}} | Export-Csv .\RetentionPolicies.csv -NoTypeInformation
        
        Get-Linebreak
        "Get-RetentionPolicyTag"
        Get-RetentionPolicyTag
        Get-RetentionPolicyTag | Format-List
        Get-RetentionPolicyTag | Select-Object
        Get-RetentionPolicyTag | Select-Object name,type,agelimitforretention,retentionaction | Export-Csv .\EXORetentionPoliciesTag.csv -NoTypeInformation

        Get-Linebreak
        "Get-JournalRule"
        Get-JournalRule
        Get-JournalRule | Format-List

        Get-Linebreak
        "Get-AcceptedDomain"
        $AcceptedDomain = Get-AcceptedDomain
        $AcceptedDomain | Format-List
        $AcceptedDomain | Select-Object name,domainname,domaintype,default | Export-Csv -Path .\EXOAcceptedDomains.csv -NoTypeInformation
        foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type MX}
        foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type TXT}
        foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type CNAME} 

        Get-Linebreak
        "Get-EmailAddressPolicy"
        Get-EmailAddressPolicy
        Get-EmailAddressPolicy | Format-List
        Get-EmailAddressPolicy | Select-Object Name,Priority,IncludedRecipients,@{Name="EnabledEmailAddressTemplates";Expression={$_.EnabledEmailAddressTemplates -join “;”}},RecipientFilterApplied | Export-Csv -Path .\EmailAddressPolicies.csv -NoTypeInformation
       
        Get-Linebreak
        "Get-TransportRule"
        Get-TransportRule
        Get-TransportRule | Format-List
        Get-TransportRule | Select-Object Name,Description, State, Priority | Export-Csv -Path .\EXOTransportRules.csv -NoTypeInformation
        $file = Export-TransportRuleCollection
        Set-Content -Path ".\EXORules.xml" -Value $file.FileData -Encoding Byte

        Get-Linebreak
        "Get-OutboundConnector"
        Get-OutboundConnector
        Get-OutboundConnector | Format-List
        Get-OutboundConnector | Select-Object name,@{Name="SmartHosts";Expression={$_.SmartHosts -join “;”}},Enabled,@{Name="AddressSpaces";Expression={$_.AddressSpaces -join “;”}},@{Name="SourceTransportServers";Expression={$_.SourceTransportServers -join “;”}},FQDN,MaxMessageSize,ProtocolLoggingLevel,RequireTLS |Export-Csv -Path .\SendConnectors.csv -NoTypeInformation

        Get-Linebreak
        "Get-InboundConnector"
        Get-InboundConnector
        Get-InboundConnector | Format-List
        Get-InboundConnector | Select-Object name,authmechanism,@{Name="Bindings";Expression={$_.Bindings -join “;”}},enabled,@{Name="RemoteIPRanges";Expression={$_.RemoteIPRanges -join “;”}},requireTLS,originatingserver | Export-Csv -Path .\ReceiveConnectors.csv -NoTypeInformation

        Get-Linebreak
        "Get-AddressBookPolicy"
        Get-AddressBookPolicy

        Get-Linebreak
        "Get-PublicFolder"
        Get-PublicFolder -Recurse -
        "Get-MailPublicFolder"
        Get-MailPublicFolder -ResultSize unlimited
        "Get-PublicFolderMailbox"
        Get-Mailbox -PublicFolder -ResultSize unlimited


        Get-Linebreak
        "Get-OrganizationConfig"
        Get-OrganizationConfig

        Get-Linebreak
        "Get-FederationTrust"
        Get-FederationTrust
        Get-FederationTrust | Format-List
        "Get-OrganizationRelationship"
        Get-OrganizationRelationship
        Get-OrganizationRelationship | Format-List
        Get-OrganizationRelationship | Select-Object name,@{Name="DomainNames";Expression={$_.DomainNames -join “;”}},targetautodiscoverepr,targetowaurl,targetsharingepr,targetapplicationuri,enabled |Export-Csv -Path .\OrganizationRelationships.csv -NoTypeInformation

        Get-Linebreak
        "Get-RemoteDomain"
        Get-RemoteDomain
        Get-RemoteDomain | Format-List
        Get-RemoteDomain | Select-Object name,domainname,allowedooftype | Export-Csv -Path .\EXORemoteDomains.csv -NoTypeInformation

      

  
        Stop-Transcript
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

    
 

