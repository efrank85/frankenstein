


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

        1) Get-FrankensteinExchangeDiscovery: Provides Exchange on-premises discovery information and outputs a transcript along with optional CSV outputs. 
        You must be connected to Exchange PowerShell prior to launching this module.

        [-virtualdirectories] [-CSV] [-UseCurrentSession]

        2) Get-FrankensteinExchangeOnlineDiscovery: Provides Exchange Online discovery information and outputs a transcript along with optional CSV outputs. 
        This function will automatically attempt to connect to Exchange Online and prompt for credentials.

        [-CSV] [-UseCurrentSession]

        3) Install-ExchangeOnline: Will install and configure Exchange Online PowerShell requirements to run Connect-ExchangeOnline

        4) Connect-All: Will connect to MSOL, AzureAD and ExO PS Sessions

        [-noMFA]

        5) Connect-OnPremServer: Connects to on-premises Exchange server using FQDN

        6) Get-FrankesnteinRecipientCounts: Displays summary of all recipient types
                
                "
        }

function Get-FrankensteinVirtualDirectories {    
    [CmdletBinding()]
    Param (
    [Switch]$CSV
    )
      
       
        Get-Linebreak
        "Get-VirtualDirectories"
        if($CSV){       
        $ClientAccess = Get-ClientAccessService
        $ClientAccess | ForEach-Object{Get-AutoDiscoverVirtualDirectory | Select-Object server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        $ClientAccess | ForEach-Object{Get-OwaVirtualDirectory | Select-Object server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation -Append
        $ClientAccess | ForEach-Object{Get-ECPVirtualDirectory | Select-Object server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation -Append
        $ClientAccess | ForEach-Object{Get-MAPIVirtualDirectory | Select-Object server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation -Append
        $ClientAccess | ForEach-Object{Get-ActiveSyncVirtualDirectory | Select-Object server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation -Append
        $ClientAccess | ForEach-Object{Get-WebServicesVirtualDirectory | Select-Object server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation -Append
        $ClientAccess | ForEach-Object{Get-OutlookAnywhere | Select-Object server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} | Export-Csv .\VirtualDirectories$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation -Append
        }
        else {
            $ClientAccess | ForEach-Object{Get-AutoDiscoverVirtualDirectory | Select-Object server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod}
            $ClientAccess | ForEach-Object{Get-OwaVirtualDirectory | Select-Object server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} 
            $ClientAccess | ForEach-Object{Get-ECPVirtualDirectory | Select-Object server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} 
            $ClientAccess | ForEach-Object{Get-MAPIVirtualDirectory | Select-Object server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} 
            $ClientAccess | ForEach-Object{Get-ActiveSyncVirtualDirectory | Select-Object server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} 
            $ClientAccess | ForEach-Object{Get-WebServicesVirtualDirectory | Select-Object server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} 
            $ClientAccess | ForEach-Object{Get-OutlookAnywhere | Select-Object server,name,internalurl,externalurl,internalauthenticationmethods,externalauthenticationmethods,IISauthenticationmethods,internalhostname,externalhostname,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod} 
                
        }

    }

function Get-FrankensteinExchangeDiscovery {    
    [CmdletBinding()]
    Param (
    [Switch]$virtualDirectories,
    [Switch]$CSV,
    [Switch]$UseCurrentSession
    
    )

    if($UseCurrentSession){

    }
   else {
       Connect-ExchangeOnPremServer
   }
        
        mkdir .\FrankensteinEXDiscovery_$((Get-Date).ToString('MMddyy'))
        Set-Location  .\FrankensteinEXDiscovery_$((Get-Date).ToString('MMddyy'))
        
        Start-Transcript -Path .\ExchangeDiscoveryTranscript_$((Get-Date).ToString('MMddyy')).txt
        
        Get-Linebreak
        Get-FrankensteinRecipientCounts                     

        Get-Linebreak
        "Get-ExchangeServer"
        if($CSV){
        $ExchangeServers = Get-ExchangeServer
        $ExchangeServers|Format-List$ExchangeServers|Select-Object Name,Domain,Edition,FQDN,IsHubTransportServer,IsClientAccessServer,IsEdgeServer,IsMailboxServer,IsUnifiedMessagingServer,IsFrontendTransportServer,OrganizationalUnit,AdminDisplayVersion,Site,ServerRole | Export-Csv .\ExchangeServers_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        
        }
        else {
            $ExchangeServers = Get-ExchangeServer
            $ExchangeServers|Format-List  
        }

        Get-Linebreak
        "Get-ExchangeServerDatabase" 
        if($CSV){
        Get-MailboxDatabase
        Get-MailboxDatabase | Format-List
        Get-MailboxDatabase | Select-Object Name,Server,MailboxRetention,ProhibitSendReceiveQuota,ProhibitSendQuota,RecoverableItemsQuota,RecoverableItemsWarningQuota,IsExcludedFromProvisioning,ReplicationType,DeletedItemRetention,
        CircularLoggingEnabled, AdminDisplayVersion | Export-Csv .\Databases_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }
        else {
            Get-MailboxDatabase
            Get-MailboxDatabase | Format-List
            
        }
        
        Get-Linebreak
        "Get-DatabaseAvailabilityGroup"
        if($CSV){
        Get-DatabaseAvailabilityGroup
        Get-DatabaseAvailabilityGroup | Format-List
        Get-DatabaseAvailabilityGroup | Format-List | Export-Csv .\DAG__$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }
        else {
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
        foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type MX}
        foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type TXT}
        foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type CNAME}
        }
        else {
            $AcceptedDomain = Get-AcceptedDomain
            $AcceptedDomain
            $AcceptedDomain | Format-List
            foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type MX}
            foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type TXT}
            foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type CNAME}
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
        "Get-SendConnector"
        if($CSV) {
        Get-SendConnector
        Get-SendConnector | Format-List
        Get-SendConnector | Select-Object name,@{Name="SmartHosts";Expression={$_.SmartHosts -join “;”}},Enabled,@{Name="AddressSpaces";Expression={$_.AddressSpaces -join “;”}},@{Name="SourceTransportServers";Expression={$_.SourceTransportServers -join “;”}},FQDN,MaxMessageSize,ProtocolLoggingLevel,RequireTLS |Export-Csv -Path .\SendConnectors_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }
        else {
            Get-SendConnector
            Get-SendConnector | Format-List
        }

        Get-Linebreak
        "Get-ReceiveConnector"
        if($CSV){
        Get-ReceiveConnector
        Get-ReceiveConnector | Format-List
        Get-ReceiveConnector | Select-Object name,authmechanism,@{Name="Bindings";Expression={$_.Bindings -join “;”}},enabled,@{Name="RemoteIPRanges";Expression={$_.RemoteIPRanges -join “;”}},requireTLS,originatingserver | Export-Csv -Path .\ReceiveConnectors_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }
        else {
            Get-ReceiveConnector
            Get-ReceiveConnector | Format-List
        }
            
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
        "Get-ExchangeCertificate"
        if($CSV){
        Get-ExchangeCertificate
        Get-ExchangeCertificate | Format-List
        Get-ExchangeCertificate | Select-Object subject,Issuer,Thumbprint,FriendlyName,NotAfter | Export-Csv .\ExchangeCertificates_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }
        else {
            Get-ExchangeCertificate
            Get-ExchangeCertificate | Format-List
        }

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

function Get-FrankensteinRecipientCounts {
    [CmdletBinding()]
    Param (
    )   

      #Define Variables
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

      $LitHoldCount = ($AllMailboxes | Where-Object{$_.LitigationHoldEnabled -eq $TRUE} | Measure-Object).count 
      Write-Host "$LitHoldCount Mailboxes on Litigation Hold"

      $RetentionHoldCount = ($AllMailboxes | Where-Object{$_.RetentionHoldEnabled -eq $TRUE} | Measure-Object).count
      Write-Host "$RetentionHoldCount Mailboxes on Retention Hold"

      $GetPublicFolder = (Get-PublicFolder -recurse | Measure-Object).count
      Write-Host "$GetPublicFolder Public Folders"

      $GetMailPublicFolder = (Get-MailPublicFolder -Resultsize Unlimited | Measure-Object).count
      Write-Host "$GetMailPublicFolder Mail Public Folders"

      $GetPublicFolderMailbox = (Get-Mailbox -ResultSize unlimited -PublicFolder | Measure-Object).count
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

function Get-FrankensteinExchangeOnlineDiscovery {    
    [CmdletBinding()]
    Param (
    [Switch]$CSV,
    [Switch]$UseCurrentSession
    )
   
        if($UseCurrentSession){

        }
        else {
            Connect-ExchangeOnline
        }

        mkdir .\FrankensteinEXODiscovery_$((Get-Date).ToString('MMddyy'))
        Set-Location  .\FrankensteinEXODiscovery_$((Get-Date).ToString('MMddyy'))

        Start-Transcript -Path .\ExchangeOnlineDiscoveryTranscript_$((Get-Date).ToString('MMddyy')).txt
        

        Get-Linebreak
        Get-FrankensteinRecipientCounts                

        Get-Linebreak
        "Get-RetentionPolicy"
        if($CSV){
        Get-RetentionPolicy
        Get-RetentionPolicy | Format-List
        Get-RetentionPolicy | Select-Object name,@{Name="RetentionPolicyTagLinks";Expression={$_.RetentionPolicyTagLinks -join “;”}} | Export-Csv .\EXORetentionPolicies_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }
        else {
            Get-RetentionPolicy
            Get-RetentionPolicy | Format-List  
        }

        Get-Linebreak
        "Get-RetentionPolicyTag"
        if($CSV){
        Get-RetentionPolicyTag
        Get-RetentionPolicyTag | Format-List
        Get-RetentionPolicyTag | Select-Object
        Get-RetentionPolicyTag | Select-Object name,type,agelimitforretention,retentionaction | Export-Csv .\EXORetentionPoliciesTag_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }
        else {
            Get-RetentionPolicyTag
            Get-RetentionPolicyTag | Format-List
            Get-RetentionPolicyTag | Select-Object 
        }

        Get-Linebreak
        "Get-JournalRule"
        Get-JournalRule
        Get-JournalRule | Format-List
        

        Get-Linebreak
        "Get-AcceptedDomain"
        if($CSV){
        $AcceptedDomain = Get-AcceptedDomain
        $AcceptedDomain | Format-List
        $AcceptedDomain | Select-Object name,domainname,domaintype,default | Export-Csv -Path .\EXOAcceptedDomains_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type MX}
        foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type TXT}
        foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type CNAME}
        }
        else {
            $AcceptedDomain = Get-AcceptedDomain
            $AcceptedDomain | Format-List
            foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type MX}
            foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type TXT}
            foreach($domain in $AcceptedDomain) {Resolve-DnsName -Name  $domain -type CNAME}
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
        Get-TransportRule | Select-Object Name,Description, State, Priority | Export-Csv -Path .\EXOTransportRules_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        $file = Export-TransportRuleCollection
        Set-Content -Path ".\EXORules.xml" -Value $file.FileData -Encoding Byte
        }
        else {
            Get-TransportRule
            Get-TransportRule | Format-List
           
        }

        Get-Linebreak
        "Get-OutboundConnector"
        if($CSV){
        Get-OutboundConnector
        Get-OutboundConnector | Format-List
        Get-OutboundConnector | Select-Object name,@{Name="SmartHosts";Expression={$_.SmartHosts -join “;”}},Enabled,@{Name="AddressSpaces";Expression={$_.AddressSpaces -join “;”}},@{Name="SourceTransportServers";Expression={$_.SourceTransportServers -join “;”}},FQDN,MaxMessageSize,ProtocolLoggingLevel,RequireTLS |Export-Csv -Path .\EXOInboundConnectors_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation

        Get-Linebreak
        "Get-InboundConnector"
        if($CSV){
        Get-InboundConnector
        Get-InboundConnector | Format-List
        Get-InboundConnector | Select-Object name,authmechanism,@{Name="Bindings";Expression={$_.Bindings -join “;”}},enabled,@{Name="RemoteIPRanges";Expression={$_.RemoteIPRanges -join “;”}},requireTLS,originatingserver | Export-Csv -Path .\EXOOutboundConnectors_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }
        else {
            Get-InboundConnector
            Get-InboundConnector | Format-List 
        }

        Get-Linebreak
        "Get-AddressBookPolicy"
        Get-AddressBookPolicy

        Get-Linebreak
        "Get-PublicFolder"
        Get-PublicFolder -Recurse -
        "Get-MailPublicFolder"
        Get-MailPublicFolder -ResultSize unlimited

        Get-Linebreak
        "Get-OrganizationConfig"
        Get-OrganizationConfig

        Get-Linebreak
        "Get-FederationTrust"
        Get-FederationTrust
        Get-FederationTrust | Format-List

        Get-Linebreak
        "Get-OrganizationRelationship"
        if($CSV){
        Get-OrganizationRelationship
        Get-OrganizationRelationship | Format-List
        Get-OrganizationRelationship | Select-Object name,@{Name="DomainNames";Expression={$_.DomainNames -join “;”}},targetautodiscoverepr,targetowaurl,targetsharingepr,targetapplicationuri,enabled |Export-Csv -Path .\EXOOrganizationRelationships_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation
        }
        else {
            Get-OrganizationRelationship
            Get-OrganizationRelationship | Format-List
        }

        Get-Linebreak
        "Get-RemoteDomain"
        if($CSV){
        Get-RemoteDomain
        Get-RemoteDomain | Format-List
        Get-RemoteDomain | Select-Object name,domainname,allowedooftype | Export-Csv -Path .\EXORemoteDomains_$((Get-Date).ToString('MMddyy')).csv -NoTypeInformation     
        }
        else {
            Get-RemoteDomain
            Get-RemoteDomain | Format-List
        }
  
        Stop-Transcript
    }
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
    Stop-Transcript
}

