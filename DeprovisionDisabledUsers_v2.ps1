$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange.domain.local/PowerShell/ -Authentication Kerberos
Import-PSSession $Session

$DeleteAfterDays = 7

$CurrentTime = Get-Date
$StrDateTime = get-date -uformat "%m%d%Y-%I%M%S%p" #08112010-83805PM = 08/11/2010 8:38:05 PM
$ScriptLogPath = "C:\AdminScripts\Deprovision\DisabledUsers\Logs"
$LogFile = "$ScriptLogPath\$StrDateTime.txt"
start-transcript -path $LogFile

#------------------------------------------------------------------------------------------
#                    First Stage
# Load Disabled Users. Determine if they have CustomAttribute14 written and if not set it
# to equal "WhenChanged" value.
#------------------------------------------------------------------------------------------
Write-Host "+++++++++++++++++++++++++++++++++++++++++++`r" -foregroundcolor "yellow"
Write-Host "Executing First Stage`r" -foregroundcolor "yellow"
Write-Host "+++++++++++++++++++++++++++++++++++++++++++`r" -foregroundcolor "yellow"
$Date = Get-Date
$DisabledUsers = Get-ADUser -Properties * -Filter * -SearchScope Subtree -SearchBase "OU=Disabled,DC=domain,DC=local"

foreach ($DisabledUser in $DisabledUsers)
{
	$SamAccountName = $DisabledUser.SamAccountName
	$ExtensionAttribute14 = $DisabledUser.extensionAttribute14
	$WhenChanged = $DisabledUser.whenChanged.ToString("MM/dd/yyyy")
	$DistinguishedName = $DisabledUser.DistinguishedName
	Write-Host "------------------------------------------`r" -foregroundcolor "yellow"
	Write-Host "Looking at ... $DistinguishedName`r" -foregroundcolor "yellow"
	if ($ExtensionAttribute14 -eq $NULL)
	{
		Write-Host "Null extensionAttribute14. Setting extensionAttribute14 to $WhenChanged`r"
		Set-ADUser $SamAccountName -Add @{extensionAttribute14=$WhenChanged}
	}
}

#------------------------------------------------------------------------------------------
#                    Second Stage
# Reload the User Data to get fresh information.
# Determine how old the object is and if it older than $DeleteAfterDays then it is removed.
# Otherwise the Description field of the object is updated to provide a count down time.
# Also detects if an account is enabled and disables it.
#------------------------------------------------------------------------------------------
Write-Host "+++++++++++++++++++++++++++++++++++++++++++`r" -foregroundcolor "yellow"
Write-Host "Executing Second Stage`r" -foregroundcolor "yellow"
Write-Host "+++++++++++++++++++++++++++++++++++++++++++`r" -foregroundcolor "yellow"
$DisabledUsers = Get-ADUser -Properties * -Filter * -SearchScope Subtree -SearchBase "OU=Disabled,DC=domain,DC=local"

foreach ($DisabledUser in $DisabledUsers)
{
	$IsEnabled = $DisabledUser.Enabled
	$GALEnabled = $DisabledUser.msExchHideFromAddressLists
	$homeMDB = $DisabledUser.homeMDB
	$ProfileFLDR =  $DisabledUser.profilePath
	$SamAccountName = $DisabledUser.SamAccountName
	$ExtensionAttribute14 = [datetime]$DisabledUser.extensionAttribute14
	$Calculation = $Date - $ExtensionAttribute14
	$TotalDays = ($Calculation.TotalDays)
	$DistinguishedName = $DisabledUser.DistinguishedName
	$Description = $DisabledUser.Description
	$DaysToDeletion = $DeleteAfterDays - $TotalDays
	$DaysToDeletion = "{0:N0}" -f $DaysToDeletion
	Write-Host "Looking at ... $DistinguishedName`r" -foregroundcolor "yellow"
	if ($IsEnabled -eq $true)
	{
		Write-Host "   Disabling account.`r"
		Disable-ADAccount $SamAccountName -Confirm:$false
	}
	if ($GALEnabled -eq $false -Or $GALEnabled -eq $NULL)
	{
		Write-Host "   Hiding from GAL.`r"
		Set-ADUser $SamAccountName -Replace @{msExchHideFromAddressLists = $true}
	}
	
		#Added to clean up user profile folders
	if ($ProfileFLDR -eq $null)
	{
		Write-Host "Did not detect a user profile folder, skipping deletion...`r" -foregroundcolor "yellow"
	}
	else
	{
		$TestProfilePath = test-path $ProfileFLDR
		if ($TestProfilePath -eq $true)
		{
			Write-Host "Detected user profile folder attempting to remove.`r" -foregroundcolor "yellow"
			Remove-Item $ProfileFLDR -Recurse -Force
		}
		
		$ProfileFLDRV2 = $ProfileFLDR + ".V2"
		$TestProfilePathV2 = test-path $ProfileFLDRV2
		if ($TestProfilePathV2 -eq $true)
		{
			Write-Host "Detected V2 user profile folder attempting to remove.`r" -foregroundcolor "yellow"
			Remove-Item $ProfileFLDRV2 -Recurse -Force
		}		
	}
	
	#Added to clean up FullAccess AD Permissions on accounts. This prevents Outlook on other users to display logon prompts on
	#accounts that no longer exist once it is purged from the system. This steps happens on the first day the user is detected.
	$FullAccessPerms = get-user $SamAccountName  | Get-MailboxPermission | where {$_.IsInherited -eq $false -And $_.Deny -eq $false -And $_.AccessRights -eq "FullAccess"}
	foreach ($FullAccessPerm in $FullAccessPerms)
	{
		$User = $FullAccessPerm.User
		If ($User -Like "*Admin*" -OR $User -Like "*CRM*" -OR $User -Like "*archive*")
		{
			#Write-Host "   Doing nothing on $User because it doesn't match`r"
		}
		else
		{
			Write-Host "   Removing $User FullAccess permissions from $SamAccountName's mailbox.`r"
			Remove-MailboxPermission -Identity $SamAccountName -AccessRights "FullAccess" -User $User -Confirm:$false
		}
	}
	
	if ($TotalDays -ge $DeleteAfterDays)
	{
		if ($homeMDB -eq $NULL)
		{
			Write-Host "   Account is being deleted because it is $TotalDays days old which is greater than $DeleteAfterDays.`r"
			
			#Tests for any left over sub objects that will prevent Remove ADUser from working
			$TestADObject = Get-AdObject -Filter * -SearchScope oneLevel -SearchBase $DistinguishedName
			if($TestADObject -ne $NULL)
			{
				Write-Host "     Found sub objects, removing them.`r"
				Get-AdObject -Filter * -SearchScope oneLevel -SearchBase $DistinguishedName | Remove-AdObject -Recursive -Confirm:$false
			}

			Remove-ADUser $SamAccountName -Confirm:$false
		}
		else
		{
			Write-Host "   Disconnecting mailbox first and deleting user object afterwards. It is $TotalDays days old which is greater than $DeleteAfterDays.`r"

			#Checks for ActiveSync Devices attached to the account; which will prevent Remove-ADUser from working.
			$TestForActiveSync = Get-ActiveSyncDevice -Mailbox $SamAccountName
			if($TestForActiveSync -ne $NULL)
			{
				Write-Host "     Found Active Sync Devices attempting to remove.`r"
				Get-ActiveSyncDevice -Mailbox $SamAccountName | Remove-ActiveSyncDevice -Confirm:$false
			}
			
			#Removes Exchange Attributes from Account, this is done because homeMDB was detected.
			Write-Host "     Disabling mailbox.`r"
			Disable-Mailbox -Identity $SamAccountName -Confirm:$false
			
			#Tests for any left over sub objects that will prevent Remove ADUser from working
			$TestADObject = Get-AdObject -Filter * -SearchScope oneLevel -SearchBase $DistinguishedName
			if($TestADObject -ne $NULL)
			{
				Write-Host "     Found sub objects, removing them.`r"
				Get-AdObject -Filter * -SearchScope oneLevel -SearchBase $DistinguishedName | Remove-AdObject -Recursive -Confirm:$false
			}
			
			Remove-ADUser $SamAccountName -Confirm:$false
		}
	}
	else
	{
		$NewDescription = "$DaysToDeletion Days Till Deletion"
		$CompareDescriptions = Compare-Object $NewDescription $Description
		if ($CompareDescriptions -ne $NULL)
		{
			Write-Host "   Updating Description field to new count down time.`r"
			Set-ADUser $SamAccountName -Clear Description
			Set-ADUser $SamAccountName -Add @{Description="$NewDescription"}			
		}
	}
}
stop-transcript