# FILEPATH: get_apipermissions.ps1
#FY24 CSP Masters tenant transform script
$localDir = "C:\CSPMastersSourceFiles"
$mydir = Get-Location
$inputCsvFile = "M365 Business Tech Series Staging Pool.csv"
$Logfile = "$mydir\FY24CspLabTransform_Log.txt"
$appID = "aff75787-e598-43f9-a0ea-7a0ca00ababc" #This is a CDX specific appID with permissions to GRAPH API
$agentIdentifier = 'Microsoft-{CSPMasters}-{TenantTranformer}/{1.0.0}' #This is a self-identifier for GRAPH API actions

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

    $stagingCreds = ipcsv -Path $inputCsvFile
    <#Testing Creds
    $tenantName = $stagingCreds[0].TenantName
    $user = $stagingCreds[0].'Administrative Username'
    $pass = $stagingCreds[0].'Administrative Password'
    #>
    
    $Global:Token = GetAccessToken


    #Let's get the app
    #GET /servicePrincipals(appId='{appId}')
    #GET https://graph.microsoft.com/v1.0/servicePrincipals/7408235b-7540-4850-82fe-a5f15ed019e2?$select=id,appId,displayName,appRoles,oauth2PermissionScopes,resourceSpecificApplicationPermissions
    $graphApiVersion = "beta"
    $resource = "servicePrincipals"
    $uri = "https://graph.microsoft.com/$graphApiVersion/$($resource)(appId='$appID')"
    #$uri = "https://graph.microsoft.com/$graphApiVersion/$($resource)(appId='$appID')?`$select=id,appId,displayName,appRoles,oauth2PermissionScopes,resourceSpecificApplicationPermissions"
    $serviceprincipal = Invoke-RestMethod -Uri $uri -Headers $headers -TimeoutSec 60
    $spId = $serviceprincipal.id

    $uri = "https://graph.microsoft.com/$graphApiVersion/$($resource)/$spId/oauth2PermissionGrants"
    $oAuth = Invoke-RestMethod -Uri $uri -Headers $headers -TimeoutSec 60

    $resource = "applications"
    $uri = "https://graph.microsoft.com/$graphApiVersion/$($resource)/$spId"
    $application = Invoke-RestMethod -Uri $uri -Headers $headers -TimeoutSec 60

    #$uri = "https://graph.microsoft.com/$graphApiVersion/$($resource)(appId='",$appID,''')?$select=id,appId,displayName,requiredResourceAccess' -join("")
    $uri = "https://graph.microsoft.com/$graphApiVersion/$($resource)"
    #GET https://graph.microsoft.com/beta/applications(appId='$appID')?$select=id,appId,displayName,requiredResourceAccess

    $count=0
    $application  = $null
    do {
        $count++
        Try
        {
            # Call the function to create a new team with specified owners
            $application = (Invoke-RestMethod -Uri $uri -Headers $headers -TimeoutSec 60).Value
            Write_Log -Message_Type "SUCCESS" -Message "Getting Application"		
        }
        Catch
        {
            $message = $_
            Write_Log -Message_Type "ERROR" -Message "Getting Application"	
            Write_Log -Message_Type "ERROR" -Message $message
            $message = $null
            #EXIT
        }
        if ($application -eq $null){
            Write_Log -Message_Type "INFO" -Message "Application not ready, retrying in 5 seconds"
            Start-Sleep -Seconds 5
        }
    } while ($application -eq $null -and $count -lt 10)

