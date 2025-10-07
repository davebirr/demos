# FILEPATH: delete_configprofiles.ps1
#FY24 CSP Masters tenant transform script
$global:localDir = "C:\CSPMastersSourceFiles"
$global:mydir = Get-Location
$global:inputCsvFile = "M365 Business Tech Series Staging Pool.csv"
$global:Logfile = "$mydir\FY24CspLabTransform_Log.txt"
$global:appID = "aff75787-e598-43f9-a0ea-7a0ca00ababc" #This is a CDX specific appID with permissions to GRAPH API
$global:agentIdentifier = 'Microsoft-{CSPMasters}-{TenantTranformer}/{1.0.0}' #This is a self-identifier for GRAPH API actions

Function Write_Log
	{
		param(
		$Message_Type,	
		$Message
		)
		
		$MyDate = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
        Switch ($Message_Type)
        {
            "INFO" { $color = "Yellow" }
            "SUCCESS" { $color = "Green" }
            "ERROR" { $color = "Red" }
            Default { $color = "White" }
        }

		Write-host -ForegroundColor $color "$MyDate - $Message_Type : $Message"			
	}

function GetAccessToken {
    # function to return an Oauth access token

    <# Uses global values applicable for the application used to connect to the Graph
    $tenantName
    $appID
    $User
    $Pass
    #>
    
    # Construct URI and body needed for authentication
    $Uri = "https://login.microsoftonline.com/$tenantName/oauth2/v2.0/token"
    $refreshTokenExpiry = 3600
    $Body = @{
        client_id     = $appID
        scope         = "https://graph.microsoft.com/.default"
        username      = $User
        password      = $Pass
        grant_type    = 'password'
    }
    
    #Lets get an OAuth 2.0 Token using that auth body
    Write_Log -Message_Type "INFO" -Message "Getting Access Token"
    Try
        {
            #$tokenRequest = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$($tenantName)/oauth2/v2.0/token" -Body $Authbody -ErrorAction Stop
            $tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
            $Global:TokenCreationDate = (Get-Date)
            $Global:TokenExpiredDate = (Get-date).AddSeconds($refreshTokenExpiry)
            $Token = ($tokenRequest.Content | ConvertFrom-Json).access_token
            # Unpack Access Token
            Write_Log -Message_Type "SUCCESS" -Message "Getting Access Token"
    
            $Global:headers = @{
                        'Content-Type'  = "application\json"
                        'User-Agent' = $agentIdentifier
                        'Authorization' = "Bearer $Token" 
                        'ConsistencyLevel' = "eventual" 
                    }
        }
    Catch
        {
            $ex = $_.Exception
            $errorResponse = $ex.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd();
            Write_Log -Message_Type "ERROR" -Message "Response content:`n$responseBody"
            Write_Log -Message_Type "ERROR" -Message "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
            $message = $_
            Write_Log -Message_Type "ERROR" -Message "Getting Access Token"	
            Write_Log -Message_Type "ERROR" -Message $message
            $message = $null
            #EXIT
        }
    Return $Token
}
Function Check-AccessToken {
    # Function to check if the access token needs to be refreshed. If it does, request a new token
    # This often needs to happen when the script processes more than a few thousands groups
    $TimeNow = (Get-Date)
    if($TimeNow -ge $TokenExpiredDate) {
        $Global:Token = GetAccessToken
        $Global:TokenExpiredDate = (Get-Date).AddMinutes($TimeToRefreshToken) 
        Write_Log -Message_Type "INFO" -Message "Requested new access token - expiration at $TokenExpiredDate" 
    } else {
        Write_Log -Message_Type "INFO" -Message "Access token still valid - expiration at $TokenExpiredDate"
    }
    #Return $Token
}

    $stagingCreds = ipcsv -Path $inputCsvFile

foreach ($tenant in $stagingCreds){
    $tenantName = $tenant.TenantName
    $User = $tenant.'Administrative Username'
    $Pass = $tenant.'Administrative Password'

    # Check if $tenant.'Administrative Password' starts with '= and remove the single quotes if it does
    if ($Pass.StartsWith('''=')) {
        $Pass = $Pass.Replace('''=', '=')
    }

    $tenantCount++
    Write-Progress -Id 0 -Activity "Deleting existing Intune Device Profiles" -Status "Tenant $($tenant.TenantPrefix): $tenantCount of $($stagingCreds.Count)" -PercentComplete $([Math]::Ceiling($tenantCount/$stagingCreds.Count*100))

    <#Testing Creds
    $tenantName = $stagingCreds[0].TenantName
    $user = $stagingCreds[0].'Administrative Username'
    $pass = $stagingCreds[0].'Administrative Password'
    #>
    
    $Global:Token = GetAccessToken

    #Let's get the Device Configurations
    $graphApiVersion = "Beta"
    $DCP_resource = "deviceManagement/deviceConfigurations"

    $uri = "https://graph.microsoft.com/$graphApiVersion/$($DCP_resource)"
    Write_Log -Message_Type "INFO" -Message "Getting Device Management object for tenant $tenantName"	
    $count=0
    $deviceConfigs  = $null
    do {
        $count++
        Try
        {
            # Call the function to create a new team with specified owners
            $deviceConfigs = (Invoke-RestMethod -Uri $uri -Headers $headers -TimeoutSec 60).Value
            Write_Log -Message_Type "SUCCESS" -Message "Getting Device Configurations"		
        }
        Catch
        {
            $message = $_
            Write_Log -Message_Type "ERROR" -Message "Getting Device Configurations"	
            Write_Log -Message_Type "ERROR" -Message $message
            $message = $null
            #EXIT
        }
        if ($deviceConfigs -eq $null){
            Write_Log -Message_Type "INFO" -Message "Device Configurations not ready, retrying in 5 seconds"
            Start-Sleep -Seconds 5
        }
    } while ($deviceConfigs -eq $null -and $count -lt 10)

    foreach ($config in $deviceConfigs){
        Write-Progress -Id 1 -ParentId 0 -Activity "Deleting device configs: " -Status "$($config.displayName)"

        $graphApiVersion = "Beta"
        $DCP_resource = "deviceManagement/deviceConfigurations"

        try {
            $uri = "https://graph.microsoft.com/$graphApiVersion/$($DCP_resource)/$($config.Id)"
            Invoke-RestMethod -Uri $uri -Headers $headers -Method DELETE -TimeoutSec 60 | Out-Null
            Write_Log -Message_Type "SUCCESS" -Message "Deleting Device Configuration $($config.Id)"		
        }
        catch {
            $ex = $_.Exception
            $errorResponse = $ex.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd();
            Write_Log -Message_Type "ERROR" -Message "Response content:`n$responseBody"
            Write_Log -Message_Type "ERROR" -Message "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
            $message = $_
            Write_Log -Message_Type "ERROR" -Message "Deleting Device Configuration $($config.Id)"	
            Write_Log -Message_Type "ERROR" -Message $message
            $message = $null
            #EXIT
        }
}



}