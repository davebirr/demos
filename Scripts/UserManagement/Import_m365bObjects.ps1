#This script should be run one time to import users into demo on-premises AD
#It it intended to be run from C:\M365bLab directory
#It relies on a properly formatted CSV file

$GlobalAdminUsername = Read-Host 'Please enter the M365 Tenant Global Admin username'

$mydir = "C:\M365bLab"
$inputCsvFile = "m365bObjects.csv"

$parentOU = Read-Host 'Please enter the On-Premises Domain Parent OU for objects (e.g. Contoso)'
$defaultPassword = Read-Host 'Please enter the On-Premises default user password'
$smtpDomain = $GlobalAdminUsername.split("@")[1]
$Logfile = "$mydir\Importm365bObjects_Log.txt"


$domain = Get-ADDomain
$upnSuffix = $domain.DNSRoot

#Add users to on-premises AD
$domainObjects = ipcsv -Path $inputCsvFile
$ouPath = $upnsuffix.split(".").foreach({"DC=" + $_}) -join ","
$topOU = New-ADOrganizationalUnit -Name $parentOU -Path $ouPath -PassThru

$userOU = New-ADOrganizationalUnit -Name Users -Path $topOU -PassThru
$groupOU = New-ADOrganizationalUnit -Name Groups -Path $topOU -PassThru
$contactOU = New-ADOrganizationalUnit -Name Contacts -Path $topOU -PassThru


$domainUsers = $domainObjects | ? {$_.objType -eq "user"}
$domainContacts = $domainObjects | ? {$_.objType -eq "contact"}
$domainGroups = $domainObjects | ? {$_.objType -eq "group"}

$j = @($domainUsers).count
$i=0
ForEach ($user in $domainUsers) {
	$i++
	Write-Progress -ID 1 -Activity "Adding user $i to AD" -Status 'Progress->' -PercentComplete (($i/$j)*100)
	Try {
	    $props = @{
	        Name              = $user.DisplayName 
	        GivenName         = $user.FirstName
	        Surname           = $user.Lastname 
	        DisplayName       = $user.Displayname 
	        sAMAccountName    = $user.alias
	        UserPrincipalName = $($user.alias,$smtpDomain -join "@") 
	        Department        = $user.Department
			Office            = $user.Office
			OfficePhone       = $user.PhoneNumber
	        Company           = $user.Company
			Title             = $user.Title
	        AccountPassword   = (ConvertTo-SecureString $defaultPassword -AsPlainText -Force)
			StreetAddress     = $user.StreetAddress
			City              = $user.City
			State             = $user.State
			PostalCode        = $user.PostalCode
			#Country           = $user.Country
	        Enabled           = $true 
			PasswordNeverExpires	= [Boolean][Int]$Item.PasswordNeverExpires
			Path              = $userOU
			ErrorAction       = 'Stop'
		}
		
		New-ADUser @props -Passthru
    	Write-Progress -ID 2 -Activity "Processing user $($_.alias)"
		Start-Sleep -Seconds 1
	}
	Catch {
		$ErrorMessage = $_.Exception.Message
		$FailedItem = $_.Exception.ItemName
		Write-Output "Could not create user $($_.alias)"
		Add-Content $LogFile ((Get-Date | Out-String) + "There was an error creating the account $FailedItem, $ErrorMessage")
	}
	Finally {
	}
}

#Add contacts to on-premises AD
$j = @($domainContacts).count
$i=0
ForEach ($contact in $domainContacts) {
	$i++
	Write-Progress -ID 1 -Activity "Adding contact $i to AD" -Status 'Progress->' -PercentComplete (($i/$j)*100)
	Try {

		New-ADObject -Type Contact -Name $contact.DisplayName -Path $contactOU -ErrorAction 'stop' -OtherAttributes `
			@{'displayName'=$contact.DisplayName;'mail'=$contact.EmailAddress}
		Write-Progress -ID 2 -Activity "Processing contact $($_.alias)"
		Start-Sleep -Seconds 1
	}
	Catch {
		$ErrorMessage = $_.Exception.Message
		$FailedItem = $_.Exception.ItemName
		Write-Output "Could not create contact $($_.alias)"
		Add-Content $LogFile ((Get-Date | Out-String) + "There was an error creating the contact account $FailedItem, $ErrorMessage")
	}
	Finally {
	}
    
}

#Add groups to on-premises AD
$j = @($domainGroups).count
$i=0
ForEach ($group in $domainGroups) {
	$i++
	Write-Progress -ID 1 -Activity "Adding group $i to AD" -Status 'Progress->' -PercentComplete (($i/$j)*100)
	if ($group.groupType -eq "MailEnabledSecurity"){
			$type = "Security"
		} elseif ($group.groupType -eq "DistributionList"){
			$type = "Distribution"
		}
	Try {
	    $props = @{
	        Name              = $group.DisplayName 
	        DisplayName       = $group.Displayname 
	        sAMAccountName    = $group.alias
			Path              = $groupOU
			ErrorAction       = 'Stop'
			GroupCategory     = $type
			GroupScope        = "Global"
		}
		
		$objGroup = New-ADGroup @props -Passthru
    	Write-Progress -ID 2 -Activity "Processing group $($_.alias)"
		Start-Sleep -Seconds 1
		$grpMembers = @()
		$grpMembers = $group.groupMembership.split(";")
		Add-ADGroupMember -Identity $objGroup -Members $grpMembers
	}
	Catch {
		$ErrorMessage = $_.Exception.Message
		$FailedItem = $_.Exception.ItemName
		Write-Output "Could not create group $($_.alias)"
		Add-Content $LogFile ((Get-Date | Out-String) + "There was an error creating the group $FailedItem, $ErrorMessage")
	}
	Finally {
	}
}

$configContainer = $domain.SubordinateReferences -match "Configuration"
$recycleBinFeature = "CN=Recycle Bin Feature,CN=Optional Features,CN=Directory Service,CN=Windows NT,CN=Services,$configContainer"
Enable-ADOptionalFeature –Identity $recycleBinFeature –Scope ForestOrConfigurationSet –Target $upnSuffix
