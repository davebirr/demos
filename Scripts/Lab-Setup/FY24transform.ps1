#FY24 CSP Masters tenant transform script
$mydir = Get-Location
$inputCsvFile = "M365 Business Tech Series Staging Pool.csv"
$Logfile = "$mydir\FY24CspLabTransform_Log.txt"

$stagingCreds = ipcsv -Path $inputCsvFile
$appID = "aff75787-e598-43f9-a0ea-7a0ca00ababc"
$agentIdentifier = 'Microsoft-{CSPMasters}-{TenantTranformer}/{1.0.0}'

$testTenant = $stagingCreds[0].TenantName
$testPrefix = $stagingCreds[0].TenantPrefix
$testUser = $stagingCreds[0].'Administrative Username'
$testPass = $stagingCreds[0].'Administrative Password'

$secureSecret = ConvertTo-SecureString -String $testPass  -AsPlainText -Force
$testCredential = New-Object System.Management.Automation.PSCredential ($testUser, $secureSecret)

#Let's build a Graph auth body
$scope = 'https://graph.microsoft.com/.default'
$AuthBody = @{
    client_id     = $appID
    scope         = $Scope
    username      = $testUser
    password      = $testPass
    grant_type    = 'password'
}

#Lets get an access Token using that auth body
$AccessToken = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$($testTenant)/oauth2/v2.0/token" -Body $Authbody -ErrorAction Stop

$AADAuthBody = @{
    client_id     = $appID
    scope         = 'https://graph.windows.net/.default'
    username      = $testUser
    password      = $testPass
    grant_type    = 'password'
}
$AADGraphToken = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$($testTenant)/oauth2/v2.0/token" -Body $AADAuthBody -ErrorAction Stop

$headers = @{ Authorization = "Bearer $($AccessToken.access_token)";'User-Agent' = $agentIdentifier }
$AADheaders = @{ Authorization = "Bearer $($AADGraphToken.access_token)";'User-Agent' = $agentIdentifier }

$TeamsHeader = @{  
    Authorization = "Bearer $($AADGraphToken.access_token)";'User-Agent' = $agentIdentifier  
    "Content-Type"= "application/json"  
    'Content-Range' = "bytes 0-$($fileLength-1)/$fileLength"	
}  

$uri = "https://graph.microsoft.com/v1.0/groups/$team.GroupID/sites/root/weburl"

$ReturnSPSiteID = 

$ReturnedData = try {
        $Data = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/organization" -Method GET -Headers $headers -ContentType 'application/json; charset=utf-8'
        if ($data.value) { $data.value } else { ($Data) }
    }
    catch {
            $_.Exception.message
    }
    
$params = @{
    Id                = $ReturnedData.Id
    DisplayName       = $("Adventure Works",$testPrefix -join " ")
    defaultDomainName = $ReturnedData.defaultDomainName
}

$bodyToPatch = '{"displayName":"' + $("Adventure Works",$testPrefix -join " ") + '"}'
$patchTenant = (Invoke-RestMethod -Method PATCH -Uri "https://graph.microsoft.com/v1.0/organization/$($ReturnedData.Id)" -Body $bodyToPatch -ContentType 'application/json' -Headers $headers -ErrorAction Stop)

$ReturnedData = try {
    $Data = Invoke-RestMethod -Uri "https://graph.windows.net/$($testTenant)/tenantDetails?api-version=1.6" -Method GET -Headers $AADheaders -ContentType 'application/json; charset=utf-8'
    if ($data.value) { $data.value } else { ($Data) }
}
catch {
        $_.Exception.message
}
$bodyToPatch = '{"displayName":"' + $("Adventure Works",$testPrefix -join " ") + '"}'
$patchTenant = (Invoke-RestMethod -Method PATCH -Uri "https://graph.windows.net/$($testTenant)/tenantDetails?api-version=1.6" -Body $bodyToPatch -ContentType 'application/json' -Headers $AADheaders -ErrorAction Stop)



    #PATCH https://graph.microsoft.com/v1.0/organization/84841066-274d-4ec0-a5c1-276be684bdd3
    #Content-type: application/json
    
    #{
    #  "marketingNotificationEmails" : ["marketing@contoso.com"],
    #  "privacyProfile" :
    #    {
    #      "contactEmail":"alice@contoso.com",
    #      "statementUrl":"https://contoso.com/privacyStatement"
    #    },
    #  "securityComplianceNotificationMails" : ["security@contoso.com"],
    #  "securityComplianceNotificationPhones" : ["(123) 456-7890"],
    #  "technicalNotificationMails" : ["tech@contoso.com"]
    #}

-tenantid $stagingCreds[0].TenantName -AsROPC $true -AppID $appID -ReturnRefresh $false -userName $stagingCreds[0].'Administrativer Username' -userPassword $stagingCreds[0].'Administrativer Password'

function Get-GraphToken($tenantid, $scope, $AsROPC, $AppID, $refreshToken, $ReturnRefresh, $userName, $userPassword) {
    if (!$scope) { $scope = 'https://graph.microsoft.com/.default' }

    $AuthBody = @{
        client_id     = $env:ApplicationID
        client_secret = $env:ApplicationSecret
        scope         = $Scope
        refresh_token = $env:RefreshToken
        grant_type    = 'refresh_token'
    }
    if ($asROPC -eq $true) {
        $AuthBody = @{
            client_id     = $appID
            scope         = $Scope
            username      = $UserName
            password      = $userPassword
            grant_type    = 'password'
        }
    }

    if ($null -ne $AppID -and $null -ne $refreshToken) {
        $AuthBody = @{
            client_id     = $appid
            refresh_token = $RefreshToken
            scope         = $Scope
            grant_type    = 'refresh_token'
        }
    }

    if (!$tenantid) { $tenantid = $env:TenantID }

    try {
        $AccessToken = (Invoke-RestMethod -Method post -Uri "https://login.microsoftonline.com/$($tenantid)/oauth2/v2.0/token" -Body $Authbody -ErrorAction Stop)
        if ($ReturnRefresh) { $header = $AccessToken } else { $header = @{ Authorization = "Bearer $($AccessToken.access_token)";'User-Agent' = $agentIdentifier } }
        return $header
        Write-Host $header['Authorization']
    }
    #This is CIPP code for tracking API failures in a database. Super useful, but commented out for now since this example doesn't have a DB.
    #I'll recode this section with something simpler that doesn't have requirement for Azure DB (e.g. local hash table or a file)
    catch {
        # Track consecutive Graph API failures
        #$TenantsTable = Get-CippTable -tablename Tenants
        #$Filter = "PartitionKey eq 'Tenants' and (defaultDomainName eq '{0}' or customerId eq '{0}')" -f $tenantid
        #$Tenant = Get-AzDataTableEntity @TenantsTable -Filter $Filter
        #if (!$Tenant.RowKey) {
        #    $donotset = $true
        #    $Tenant = [pscustomobject]@{
        #        GraphErrorCount     = $null
        #        LastGraphTokenError = $null
        #        LastGraphError      = $null
        #        PartitionKey        = 'TenantFailed'
        #        RowKey              = 'Failed'
        #    }
        #}
        #$Tenant.LastGraphError = if ( $_.ErrorDetails.Message) {
        #    $msg = $_.ErrorDetails.Message | ConvertFrom-Json
        #    "$($msg.error):$($msg.error_description)"
        #}
        #else {
            $_.Exception.message
        #}
        #$Tenant.GraphErrorCount++

       # if (!$donotset) { Update-AzDataTableEntity @TenantsTable -Entity $Tenant }
       # throw "$($Tenant.LastGraphError)"
    }
}

function New-GraphGetRequest {
    Param(
        [String] $uri,
        [String] $tenantid,
        [String] $scope,
        $AsROPC,
        [GUID] $AppID,
        [String] $userName,
        [String] $userPass,
        $noPagination,
        $NoAuthCheck,
        [switch]$ComplexFilter,
        [switch]$CountOnly
    )

    if ($scope -eq 'ExchangeOnline') {
        $AccessToken = Get-ClassicAPIToken -resource 'https://outlook.office365.com' -Tenantid $tenantid
        $headers = @{ Authorization = "Bearer $($AccessToken.access_token)" }
        $headers = @{ 
            Authorization = "Bearer $($AccessToken.access_token)"
           'User-Agent' = $agentIdentifier
        }
    }
    else {
        Write-Host "Authorizing with Graph API $uri"
        Write-Host "TenantID: $tenantid"
        Write-Host "Scope: $scope"
        Write-Host "AsROPC: $asROPC"
        Write-Host "AppID: $appid"
        Write-Host "UserName: $userName"
        Write-Host "UserPassword: $userPass"
        $headers = Get-GraphToken -tenantid $tenantid -scope $scope -AsROPC $asROPC -AppID $appid -userName $userName -userPass $userPass
        $headers | Out-String | Write-Host
    }

    if ($ComplexFilter) {
        $headers['ConsistencyLevel'] = 'eventual'
    }
    Write-Verbose "Using $($uri) as url"
    $nextURL = $uri

    # Track consecutive Graph API failures
    #$TenantsTable = Get-CippTable -tablename Tenants
    #$Filter = "PartitionKey eq 'Tenants' and (defaultDomainName eq '{0}' or customerId eq '{0}')" -f $tenantid
    #$Tenant = Get-AzDataTableEntity @TenantsTable -Filter $Filter
    #if (!$Tenant) {
    #    $Tenant = @{
    #        GraphErrorCount = 0
    #        LastGraphError  = $null
    #        PartitionKey    = 'TenantFailed'
    #        RowKey          = 'Failed'
    #    }
    #}
    $ReturnedData = do {
        try {
            $Data = (Invoke-RestMethod -Uri $nextURL -Method GET -Headers $headers -ContentType 'application/json; charset=utf-8')
            if ($CountOnly) {
                $Data.'@odata.count'
                $nextURL = $null
            }
            else {
                if ($data.value) { $data.value } else { ($Data) }
                if ($noPagination) { $nextURL = $null } else { $nextURL = $data.'@odata.nextLink' }
            }
        }
        catch {
            $Message = ($_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue).error.message
            if ($Message -eq $null) { $Message = $($_.Exception.Message) }
            if ($Message -ne 'Request not applicable to target tenant.') {
                #$Tenant.LastGraphError = $Message
                #$Tenant.GraphErrorCount++
                #Update-AzDataTableEntity @TenantsTable -Entity $Tenant
            }
            throw $Message
        }
    } until ($null -eq $NextURL)
    #$Tenant.LastGraphError = ''
    #Update-AzDataTableEntity @TenantsTable -Entity $Tenant
    return $ReturnedData
}

#Get Access Token
#POST {tenant}/oauth2/v2.0/token
#Host: login.microsoftonline.com
#Content-Type: application/x-www-form-urlencoded

#client_id=6731de76-14a6-49ae-97bc-6eba6914391e
#&scope=user.read%20openid%20profile%20offline_access
#&username=MyUsername@myTenant.com
#&password=SuperS3cret
#&grant_type=password
$AuthBody = @{
    client_id     = $appID
    scope         = $Scope
    username      = $env:Username
    password      = $env:Password
    refresh_token = $env:RefreshToken
    grant_type    = 'password'
}

$headers = Get-GraphToken -tenantid $tenantid -scope $scope -AsROPC $true -AppID $appID -refreshToken $refreshToken -ReturnRefresh $false -userName $env:Username -userPassword $env:Password

$uri1 =  "https://graph.windows.net/myorganization"
$uri2 = "https://graph.microsoft.com/beta/organization"

New-GraphGetRequest -uri "https://graph.windows.net/myorganization" -tenantid $stagingCreds[0].TenantName -AsROPC $true -AppID $appID -ReturnRefresh $false -userName $stagingCreds[0].'Administrativer Username' -userPassword $stagingCreds[0].'Administrativer Password'

New-GraphGetRequest -uri "https://graph.microsoft.com/beta/organization" -tenantid $stagingCreds[0].TenantName -AsROPC $true -AppID $appID -ReturnRefresh $false -userName $stagingCreds[0].'Administrativer Username' -userPassword $stagingCreds[0].'Administrativer Password'


#Rename Tenant
#
#Import-Module Microsoft.Graph.Identity.DirectoryManagement
#$params = @{
#	"@odata.type" = "#microsoft.graph.organization"
#	mobileDeviceManagementAuthority = "intune"
#}
#Update-MgOrganization -OrganizationId $organizationId -BodyParameter $params

$j = @($stagingCreds).count
$i=0

ForEach ($credential in $stagingCreds) {
	$i++
	Write-Progress -ID 1 -Activity "Transforming tenant $i" -Status 'Progress->' -PercentComplete (($i/$j)*100)
	Try {
	    $params = @{
	        "@odata.type" = "#microsoft.graph.organization"
            DisplayName         = $("Adventure Works",$credential.TenantPrefix -join " ")
			ErrorAction       = 'Stop'
		}
        Connect-MgGraph
		Update-MgOrganization -OrganizationId $organizationId -BodyParameter $params

		#New-ADUser @props -Passthru
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