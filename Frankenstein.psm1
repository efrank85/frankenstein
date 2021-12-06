


<#
.SYNOPSIS
    Test creation of Eric Frank

.DESCRIPTION
    This module contains a series of functions used to collect and export data in preparation from an Exchange to Exchange Online migration.

.PARAMETER 
    Get-FrankensteinExchangeDiscovery: Retreives Exchange on-premises settings and outputs valid CSV tables used in Exchange design docs along with a transcript of relevant Exchange settings
        [SWITCH] VirtualDirectories: Using this switch will retrieve virtual directory information and output a CSV table
    
    Get-FrankensteinExchangeOnlineDiscovery: Retreives Exchange Online settings and outputs valid CSV tables used in Exchange design docs along with a transcript of relevant Exchange Online settings
    
    Install-ExchangeOnline: Installs prerequisites for Exchange V2 module

    Connect-ExchangeOnPremServer - Connects to Exchange on premises server by FQDN. Function will prompt for FQDN

.EXAMPLE
    Get-FrankensteinExchangeDiscovery -VirtualDirectories 


.INPUTS
    

.OUTPUTS
    CSV and .txt files
    

.NOTES
    Author:  Eric D. Frank
  
#>
 

function Insert-Linebreak {
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
    )
 
      
       
        Insert-Linebreak

        Write-Host "Get-VirtualDirectories"
       
      
        $ClientAccess | ForEach-Object{Get-AutoDiscoverVirtualDirectory | select server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories.csv -NoTypeInformation
        $ClientAccess | ForEach-Object{Get-OwaVirtualDirectory | select server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories.csv -NoTypeInformation -Append
        $ClientAccess | ForEach-Object{Get-ECPVirtualDirectory | select server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories.csv -NoTypeInformation -Append
        $ClientAccess | ForEach-Object{Get-MAPIVirtualDirectory | select server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories.csv -NoTypeInformation -Append
        $ClientAccess | ForEach-Object{Get-ActiveSyncVirtualDirectory | select server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories.csv -NoTypeInformation -Append
        $ClientAccess | ForEach-Object{Get-WebServicesVirtualDirectory | select server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories.csv -NoTypeInformation -Append
        $ClientAccess | ForEach-Object{Get-OutlookAnywhere | select server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories.csv -NoTypeInformation -Append
        

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
        $ClientAccess = Get-ClientAccessService
   
        
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

        $LitHoldCount = ($AllMailboxes | ?{$_.LitigationHoldEnabled -eq $TRUE}).count 
        Write-Host "$LitHoldCount Mailboxes on Litigation Hold"

        $RetentionHoldCount = ($AllMailboxes | ?{$_.RetentionHoldEnabled -eq $TRUE}).count
        Write-Host "$RetentionHoldCount Mailboxes on Retention Hold"

        $GetPublicFolder = (Get-PublicFolder -recurse).count
        Write-Host "$GetPublicFolder Public Folders"

        $GetMailPublicFolder = (Get-MailPublicFolder).count
        Write-Host "$GetMailPublicFolder Mail Public Folders"

        $GetPublicFolderMailbox = (Get-Mailbox -ResultSize unlimited -PublicFolder -IgnoreDefaultScope).count
        Write-Host "$GetPublicFolderMailbox Public Folder Mailboxes"

        $POP = ($CASMailbox | ?{$_.popenabled -eq $true}).count 
        Write-Host "$POP Mailboxes with POP3 Enabled"
        
        $IMAP = ($CASMailbox | ?{$_.imapenabled -eq $true}).count 
        Write-Host "$IMAP Mailboxes with IMAP Enabled"
        
        $MAPI = ($CASMailbox | ?{$_.mapienabled -eq $true}).count 
        Write-Host "$MAPI Mailboxes with MAPI Enabled"
        
        $ActiveSync = ($CASMailbox | ?{$_.activesyncenabled -eq $true}).count 
        Write-Host "$ActiveSync Mailboxes with ActiveSync Enabled"
        
        $OWA = ($CASMailbox | ?{$_.owaenabled -eq $true}).count 
        Write-Host "$OWA Mailboxes with OWA Enabled" 
        
        $ADPDisabled = ($AllMailboxes | ?{$_.EmailAddressPolicyEnabled -eq $false}).count 
        Write-Host "$ADPDisabled Mailboxes with Email Address Policy Disabled"     
                

        Insert-Linebreak
        "Get-ExchangeServer" 
        $ExchangeServers
        $ExchangeServers|FL

        Insert-Linebreak
        "Get-ExchangeServerDatabase" 
        Get-MailboxDatabase
        Get-MailboxDatabase | fl
        Get-MailboxDatabase | select Name,Server,MailboxRetention,ProhibitSendReceiveQuota,ProhibitSendQuota,RecoverableItemsQuota,RecoverableItemsWarningQuota,IsExcludedFromProvisioning,ReplicationType,DeletedItemRetention,
        CircularLoggingEnabled, AdminDisplayVersion | Export-Csv .\Databases.csv -NoTypeInformation
        
        Insert-Linebreak
        "Get-DatabaseAvailabilityGroup"
        Get-DatabaseAvailabilityGroup
        Get-DatabaseAvailabilityGroup | fl
        
        Insert-Linebreak
        "Get-RetentionPolicy"
        Get-RetentionPolicy
        Get-RetentionPolicy | FL
        Get-RetentionPolicy | select name,retentionpolicytaglinks | Export-Csv .\RetentionPolicies.csv -NoTypeInformation
        
        Insert-Linebreak
        "Get-RetentionPolicyTag"
        Get-RetentionPolicyTag
        Get-RetentionPolicyTag | FL
        Get-RetentionPolicyTag | select name,type,agelimitforretention,retentionaction | Export-Csv .\RetentionPoliciesTag.csv -NoTypeInformation

        Insert-Linebreak
        "Get-JournalRule"
        Get-JournalRule
        Get-JournalRule | FL

        Insert-Linebreak
        "Get-AcceptedDomain"
        $AcceptedDomain = Get-AcceptedDomain
        $AcceptedDomain
        $AcceptedDomain | FL
        $AcceptedDomain | select name,domainname,domaintype,default | Export-Csv -Path .\AcceptedDomains.csv -NoTypeInformation

        Insert-Linebreak
        "Get-EmailAddressPolicy"
        Get-EmailAddressPolicy
        Get-EmailAddressPolicy | fl
        Get-EmailAddressPolicy | Select Name,Priority,IncludedRecipients,EnabledEmailAddressTemplates,RecipientFilterApplied | Export-Csv -Path .\EmailAddressPolicies.csv -NoTypeInformation
        
      
        Insert-Linebreak
        "Get-TransportRule"
        Get-TransportRule
        Get-TransportRule | fl
        Get-TransportRule | Select Name,Description, State, Priority | Export-Csv -Path .\TransportRules.csv -NoTypeInformation
        $file = Export-TransportRuleCollection
        Set-Content -Path ".\Rules.xml" -Value $file.FileData -Encoding Byte

        Insert-Linebreak
        "Get-SendConnector"
        Get-SendConnector
        Get-SendConnector | fl
        Get-SendConnector | select name,SmartHosts,Enabled,AddressSpaces,SourceTransportServers,FQDN,MaxMessageSize,ProtocolLoggingLevel,RequireTLS |Export-Csv -Path .\SendConnectors.csv -NoTypeInformation

        Insert-Linebreak
        "Get-ReceiveConnector"
        Get-ReceiveConnector
        Get-ReceiveConnector | fl
        Get-ReceiveConnector | select name,authmechanism,bindings,enabled,remoteIPRanges,requireTLS,originatingserver | Export-Csv -Path .\ReceiveConnectors.csv -NoTypeInformation

        Insert-Linebreak
        "Get-TransportAgent"
        Get-TransportAgent
        Get-TransportAgent | fl

        Insert-Linebreak
        "Get-AddressList"
        Get-AddressList
        Get-AddressBookPolicy
        Start-Sleep -s 5

        Insert-Linebreak
        "Get-PublicFolder"
        Get-PublicFolder -Recurse
        Start-Sleep -s 5
        "Get-MailPublicFolder"
        Get-MailPublicFolder -ResultSize unlimited
        Start-Sleep -s 5
        "Get-PublicFolderMailbox"
        Get-Mailbox -PublicFolder -ResultSize unlimited
        Start-Sleep -s 5


        Insert-Linebreak
        "Get-OrganizationConfig"
        Get-OrganizationConfig
        Start-Sleep -s 5

        Insert-Linebreak
        "Get-FederationTrust"
        Get-FederationTrust
        Get-FederationTrust | fl
        Insert-Linebreak
        "Get-OrganizationRelationship"
        Get-OrganizationRelationship
        Get-OrganizationRelationship | fl
        Get-OrganizationRelationship | select name,domainnames,targetautodiscoverepr,targetowaurl,targetsharingepr,targetapplicationuri,enabled |Export-Csv -Path .\OrganizationRelationships.csv -NoTypeInformation


        Insert-Linebreak
        "Get-RemoteDomain"
        Get-RemoteDomain
        Get-RemoteDomain | fl
        Get-RemoteDomain | select name,domainname,allowedooftype | Export-Csv -Path .\RemoteDomains.csv -NoTypeInformation

        Insert-Linebreak
        "Get-ExchangeCertificate"
        Get-ExchangeCertificate
        Get-ExchangeCertificate | fl
        Get-ExchangeCertificate | select subject,Issuer,Thumbprint,FriendlyName,NotAfter | Export-Csv .\ExchangeCertificates.csv -NoTypeInformation

        Insert-Linebreak
        "Get-HybridConfiguration"
        Get-HybridConfiguration
        Start-Sleep -s 5

        Insert-Linebreak

        
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

        $LitHoldCount = ($AllMailboxes | ?{$_.LitigationHoldEnabled -eq $TRUE}).count 
        Write-Host "$LitHoldCount Mailboxes on Litigation Hold"

        $RetentionHoldCount = ($AllMailboxes | ?{$_.RetentionHoldEnabled -eq $TRUE}).count
        Write-Host "$RetentionHoldCount Mailboxes on Retention Hold"

        $GetPublicFolder = (Get-PublicFolder -recurse).count
        Write-Host "$GetPublicFolder Public Folders"

        $GetMailPublicFolder = (Get-MailPublicFolder).count
        Write-Host "$GetMailPublicFolder Mail Public Folders"

        $GetPublicFolderMailbox = (Get-Mailbox -ResultSize unlimited -PublicFolder).count
        Write-Host "$GetPublicFolderMailbox Public Folder Mailboxes"

        $POP = ($CASMailbox | ?{$_.popenabled -eq $true}).count 
        Write-Host "$POP Mailboxes with POP3 Enabled"
        
        $IMAP = ($CASMailbox | ?{$_.imapenabled -eq $true}).count 
        Write-Host "$IMAP Mailboxes with IMAP Enabled"
        
        $MAPI = ($CASMailbox | ?{$_.mapienabled -eq $true}).count 
        Write-Host "$MAPI Mailboxes with MAPI Enabled"
        
        $ActiveSync = ($CASMailbox | ?{$_.activesyncenabled -eq $true}).count 
        Write-Host "$ActiveSync Mailboxes with ActiveSync Enabled"
        
        $OWA = ($CASMailbox | ?{$_.owaenabled -eq $true}).count 
        Write-Host "$OWA Mailboxes with OWA Enabled" 
        
        $ADPDisabled = ($AllMailboxes | ?{$_.EmailAddressPolicyEnabled -eq $false}).count 
        Write-Host "$ADPDisabled Mailboxes with Email Address Policy Disabled"     
                

        Insert-Linebreak
        "Get-RetentionPolicy"
        Get-RetentionPolicy
        Get-RetentionPolicy | FL
        Get-RetentionPolicy | select name,retentionpolicytaglinks | Export-Csv .\EXORetentionPolicies.csv -NoTypeInformation
        
        Insert-Linebreak
        "Get-RetentionPolicyTag"
        Get-RetentionPolicyTag
        Get-RetentionPolicyTag | FL
        Get-RetentionPolicyTag | Select
        Get-RetentionPolicyTag | select name,type,agelimitforretention,retentionaction | Export-Csv .\EXORetentionPoliciesTag.csv -NoTypeInformation

        Insert-Linebreak
        "Get-JournalRule"
        Get-JournalRule
        Get-JournalRule | FL

        Insert-Linebreak
        "Get-AcceptedDomain"
        $AcceptedDomain = Get-AcceptedDomain
        $AcceptedDomain | FL
        $AcceptedDomain | select name,domainname,domaintype,default | Export-Csv -Path .\EXOAcceptedDomains.csv -NoTypeInformation

        Insert-Linebreak
        "Get-EmailAddressPolicy"
        Get-EmailAddressPolicy
        Get-EmailAddressPolicy | fl
        Get-EmailAddressPolicy | Select Name,Priority,IncludedRecipients,EnabledEmailAddressTemplates,RecipientFilterApplied | Export-Csv -Path .\EXOEmailAddressPolicies.csv -NoTypeInformation
        $file = Export-TransportRuleCollection
        Set-Content -Path ".\EXORules.xml" -Value $file.FileData -Encoding Byte

      
        Insert-Linebreak
        "Get-TransportRule"
        Get-TransportRule
        Get-TransportRule | fl
        Get-TransportRule | Select Name,Description, State, Priority | Export-Csv -Path .\EXOTransportRules.csv -NoTypeInformation


        Insert-Linebreak
        "Get-OutboundConnector"
        Get-OutboundConnector
        Get-OutboundConnector | fl
        Get-OutboundConnector | select name,SmartHosts,Enabled,AddressSpaces,SourceTransportServers,FQDN,MaxMessageSize,ProtocolLoggingLevel,RequireTLS |Export-Csv -Path .\EXOOutboundConnectors.csv -NoTypeInformation

        Insert-Linebreak
        "Get-InboundConnector"
        Get-InboundConnector
        Get-InboundConnector | fl
        Get-InboundConnector | select name,ConnectorType, SenderDomains, Requiretls, TlsSenderCertificateName | Export-Csv -Path .\EXOInboundConnectors.csv -NoTypeInformation

        Insert-Linebreak
        "Get-AddressBookPolicy"
        Get-AddressBookPolicy

        Insert-Linebreak
        "Get-PublicFolder"
        Get-PublicFolder -Recurse -
        "Get-MailPublicFolder"
        Get-MailPublicFolder -ResultSize unlimited
        "Get-PublicFolderMailbox"
        Get-Mailbox -PublicFolder -ResultSize unlimited


        Insert-Linebreak
        "Get-OrganizationConfig"
        Get-OrganizationConfig

        Insert-Linebreak
        "Get-FederationTrust"
        Get-FederationTrust
        Get-FederationTrust | fl
        "Get-OrganizationRelationship"
        Get-OrganizationRelationship
        Get-OrganizationRelationship | fl
        Get-OrganizationRelationship | select name,domainnames,targetautodiscoverepr,targetowaurl,targetsharingepr,targetapplicationuri,enabled |Export-Csv -Path .\EXOOrganizationRelationships.csv -NoTypeInformation


        Insert-Linebreak
        "Get-RemoteDomain"
        Get-RemoteDomain
        Get-RemoteDomain | fl
        Get-RemoteDomain | select name,domainname,allowedooftype | Export-Csv -Path .\EXORemoteDomains.csv -NoTypeInformation

      

  
        Stop-Transcript
}