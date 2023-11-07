<#
.SYNOPSIS
    Retreives Full Access, SendAS and SendOnBehalf permissions.

.DESCRIPTION
    Retreives Full Access, SendAS and SendOnBehalf permissions.
    Requires minimum of Exchange Reader. Global Reader will not work.

.PARAMETER 
    

.EXAMPLE
    .\FrankensteinPermissions.ps1 -UseCurrentSession -FullAccess -SendAs -SendOnBehalf 


.INPUTS
    CSV - Must Include "DisplayName" header

.OUTPUTS
    CSV
    

.NOTES
    Author:  Eric D. Frank
    09/13/23 - Updated to use GitHub as repository
  
#>
 

#Accept input paramenters
param(
[switch]$FullAccess,
[switch]$SendAs,
[switch]$SendOnBehalf,
[switch]$UserMailboxOnly,
[switch]$AdminsOnly,
[string]$MBNamesFile,
[Switch]$UseCurrentSession
)


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
 