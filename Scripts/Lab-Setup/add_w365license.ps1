# FILEPATH: add_w365license.ps1
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
    
function Retry-Command {
    [CmdletBinding()]
    Param(
        [Parameter(Position=0, Mandatory=$true)]
        [scriptblock]$ScriptBlock,

        [Parameter(Position=1, Mandatory=$false)]
        [int]$Maximum = 5,

        [Parameter(Position=2, Mandatory=$false)]
        [int]$Delay = 100,

        [Parameter(Position=2, Mandatory=$false)]
        [string]$Message = ""      
    )

    Begin {
        $cnt = 0
        Write_Log -Message_Type "INFO" -Message "Begin $Message"	
    }

    Process {
        do {
            $cnt++
            try {
                # If you want messages from the ScriptBlock
                # Invoke-Command -Command $ScriptBlock
                # Otherwise use this command which won't display underlying script messages
                Write_Log -Message_Type "SUCCESS" -Message $Message
                $ScriptBlock.Invoke()
                return
            } catch {
                $errorMessage = $_
                Write_Log -Message_Type "ERROR" -Message $errorMessage
                Write_Log -Message_Type "ERROR" $_.Exception.InnerException.Message -ErrorAction Continue
                Start-Sleep -Milliseconds $Delay
            }
        } while ($cnt -lt $Maximum)

        # Throw an error after $Maximum unsuccessful invocations. Doesn't need
        # a condition, since the function returns upon successful invocation.
        throw 'Execution failed.'
    }
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
    Write-Progress -id 0 -Activity "Assinging W365 licenses" -Status "Tenant $($tenant.TenantPrefix): $tenantCount of $($stagingCreds.Count)" -PercentComplete ($tenantCount/$stagingCreds.Count*100)

    <#Test User
    $tenantName = $stagingCreds[0].TenantName
    $User = $stagingCreds[0].'Administrative Username'
    $Pass = $stagingCreds[0].'Administrative Password'
    #>

    $Global:Token = GetAccessToken

    #Get list of SKUs
    #https://graph.microsoft.com/v1.0/subscribedSkus
    Write_Log -Message_Type "INFO" -Message "Retriving SKUs for $tenantName"	

    #$uri = 'https://graph.microsoft.com/v1.0/groups?$filter=displayName eq ',"'$groupName'" -join("")
    $uri = 'https://graph.microsoft.com/v1.0/subscribedSkus?$select=skuPartNumber,skuId'
    Retry-Command -ScriptBlock {
        $script:SubscribedSkus = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
    } -Maximum 3 -Message "Retriving SKUs"

    #Find the SkuID
    $Usernames = "PattiF", "NestorW", "IsaiahL", "LidiaH", "admin"
    $sku = ($SubscribedSkus.value | ? {$_.skuPartNumber -like "*CPC_E_2C_8GB_256GB*"}).skuID

    #Assign to users
    foreach ($Username in $Usernames) {
        # Assign Windows 365 license to user
        #$LicenseAssignment = New-MsolLicenseOptions -AccountSkuId $LicenseSKU
        $upn = $userName,$tenantName -join("@")
        Write_Log -Message_Type "INFO" -Message "Adding $sku to $upn"
        $URLtoassignLicense = "https://graph.microsoft.com/v1.0/users/$upn/assignLicense" 
        $params = @{
            addLicenses = @(
                @{
                    "skuId" = $sku
                }
            )
            "removeLicenses" = @()
        }
        $BodyJsontoassignLicense = $params | ConvertTo-Json -Depth 5 -Compress
        Invoke-RestMethod -Headers $headers -Body $BodyJsontoassignLicense -Uri $URLtoassignLicense -Method POST -ContentType "application/json" | Out-Null
        Write_Log -Message_Type "SUCCESS" -Message "Assigned Windows 365 license to user: $Username"
    }
}

