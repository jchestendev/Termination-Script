<#
.NOTES
=============================================================
###########################################
#                                         #
#             Chesten Jones               #
#                Jan2017                  #
#                                         #
###########################################

==============================================================

.DESCRIPTION
==============================================================
This script is for the purpose of off-boarding users
 
The script does the following;

1. Disables the specified user account
2. Updates the user description
3. Moves the account to the Termination OU based off location 
4. Changes Password


#The data in this script is pulled from a CSV(Comma Delimited) file

#>
import-module ActiveDirectory
$c = Get-Credential 
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $c -Authentication Basic -AllowRedirection
import-pssession $session
$Users = Import-Csv -Path "Path to CSV"

      
foreach ($User in $Users)           	
{
#Setting Variables 
$SAM = $User.NewID
$EmployeeDetails = Get-ADUser $SAM -properties *
$User = $EmployeeDetails.Name
$Domain = "Domain" 
$Email = $SAM + "@company.com"
$Date = Get-Date -Format 'MM/dd/yyyy' 
$Manager = Get-ADUser $EmployeeDetails.Manager | Select -ExpandProperty Name
$ManagerEmail = Get-ADUser $EmployeeDetails.Manager | Select -ExpandProperty UserPrincipalName






write-host ****************"Disabling & Updating $SAM's ActiveDirectory ACCOUNT"****************
write-host " "

   
# Disable User AD Account
Disable-ADAccount -Identity $SAM
write-host "$SAM Account Disabled" -foregroundcolor "GREEN"
            
# CHANGE DISPLAYNAME AND DESCRIPTION TO DISPLAY TERMINATED - $USER
Get-Aduser -Identity $SAM | Set-ADObject -Description " Left $Date Forwarding Email to $Manager"
write-host " Profile Details Updated" -foregroundcolor "GREEN"
write-host " "

write-host ****************"Making the Another Security Group Primary"****************
write-host " "
#CHANGE NOLOGIN Group to Primary Group
$group = Get-ADGroup "NOLOGIN"
$groupSid = $group.sid
$groupSid
[int]$GroupID = $groupSid.Value.Substring($groupSid.Value.LastIndexOf("-")+1)
        
Get-ADUser "$SAM" | Set-ADObject -Replace @{primaryGroupID="$GroupID"}
	sleep -Seconds 5

#Remove From All Groups

$ADgroups = Get-ADPrincipalGroupMembership -Identity $SAM | where {$_.name -ne "NOLOGIN"}
Remove-ADPrincipalGroupMembership -Identity "$SAM" -MemberOf $ADgroups -Confirm:$false

write-host "Removing All AD Groups, Please Wait" -foregroundcolor "RED"
	sleep -Seconds 5

# CHANGE USER PASSWORD TO "Solong123!"
$password = "Solong123!" | ConvertTo-SecureString -AsPlainText -Force

Set-ADAccountPassword -NewPassword $password -Identity $SAM -Reset
set-adaccountcontrol $SAM -passwordneverexpires $true -cannotchangepassword $true

write-host ****************"Password has been changed "****************

               
 write-host ****************"Moving OU to the Termination Folder"****************
# Moving OU to the Termination Folder
        $termu = Get-Aduser $SAM | select -ExpandProperty DistinguishedName 
        $location = $EmployeeDetails.Office


	if ($location -eq "London")
	{
        Move-ADObject -Identity $termu -targetpath "OU=Termination,OU=LONDON,DC=domain,DC=company,DC=com"
        write-host "MOVING TO LONDON TERMINATION FOLDER"
        	sleep -Seconds 20
    }
    elseif ($location -eq "San Francisco")
    {

        Move-ADObject -Identity $termu -targetpath "OU=Termination,OU=SAN FRANCISCO,DC=domain,DC=company,DC=com"
        write-host "MOVING TO SAN FRANCISCO TERMINATION FOLDER"
            sleep -Seconds 20
    }   
   
    else
    {
        $null
    } 

	Write-host "Removing phone number and office location"
    write-host " "
    Set-ADUser $SAM -OfficePhone $null -Office $null
	if ($?){
	Write-Host "Done" -foregroundcolor "Green"
	}
	Else {
	Write-Host "Phone and Location Not Removed" -foregroundcolor "Red"
	}


write-host " ******************************** DISABLING $USER EXCHANGE ACCOUNT ******************************** "
write-host " "

Write-Host "Step 1. Convert $User Mailbox to Shared" -ForegroundColor Yellow
write-host " "
Set-mailbox -identity $Email -type shared


Write-Host "Step 2. Disabling POP,IMAP, OWA, MAPI, and ActiveSync access for $User" -ForegroundColor Yellow
write-host " "
Set-CasMailbox -Identity $Email -OWAEnabled $false -POPEnabled $false -MAPI $false -ImapEnabled $false -ActiveSyncEnabled $false

Write-Host "Step 3. Set Email Forwarding to $Manager" -ForegroundColor Yellow
write-host " "
set-mailbox -identity $Email -forwardingaddress $ManagerEmail

Write-Host "Step 4. Sending Confirmation E-mail To Employee's Manager." -ForegroundColor Yellow

            $ME = "youremail@company.com"
            $To = $ManagerEmail
            $Subject = "$User Off-Boarding" 
            $Body = "$Manager - As a part of $User off-boarding, we have set email forwarding to you."
            $smtpServer = "SMTP Server"

       
        ############################################################################

		Send-MailMessage -SmtpServer $smtpServer -From $ME -To $ManagerEmail -Subject $Subject -Body $Body -Priority High -dno OnSuccess, OnFailure



}

Remove-PSSession $Session

