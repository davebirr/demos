# FILEPATH: create_app.ps1
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
    
    $stagingCreds = ipcsv -Path $inputCsvFile

    foreach ($tenant in $stagingCreds){
        $tenantName = $tenant.TenantName
        $User = $tenant.'Administrative Username'
        $Pass = $tenant.'Administrative Password'
        $tenantPrefix = $tenant.TenantPrefix
        
        # Check if $tenant.'Administrative Password' starts with '= and remove the single quotes if it does
        if ($Pass.StartsWith('''=')) {
            $Pass = $Pass.Replace('''=', '=')
        }

        #Test User
        #$tenantName = $stagingCreds[0].TenantName
        #$User = $stagingCreds[0].'Administrative Username'
        #$Pass = $stagingCreds[0].'Administrative Password'
        #$tenantPrefix = $stagingCreds[0].TenantPrefix
    
        #Let's build a Graph auth body
        $scope = 'https://graph.microsoft.com/.default'
        $AuthBody = @{
            client_id     = $appID
            scope         = $Scope
            username      = $User
            password      = $Pass
            grant_type    = 'password'
        }
    
        #Lets get an access Token using that auth body
        Write_Log -Message_Type "INFO" -Message "Getting Access Token"	
    
        Try
            {
                $AccessToken = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$($tenantName)/oauth2/v2.0/token" -Body $Authbody -TimeoutSec 5
                Write_Log -Message_Type "SUCCESS" -Message "Getting Access Token"		
            }
        Catch
            {
                $message = $_
                Write_Log -Message_Type "ERROR" -Message "Getting Access Token"	
                Write_Log -Message_Type "ERROR" -Message $message
                $message = $null
                EXIT
            }
    
        $headers = @{ Authorization = "Bearer $($AccessToken.access_token)";'User-Agent' = $agentIdentifier }


    #Create an app registraiton
    #POST https://graph.microsoft.com/v1.0/applications
    #Content-type: application/json
    #
    #{
    #  "displayName": "Display name"
    #}
    Write_Log -Message_Type "INFO" -Message "Creating app registraiton for $tenantName"	
    $params = @{
        displayName = "CSP Masters app"
    }
    $uri = "https://graph.microsoft.com/v1.0/applications"
    $body = $params | ConvertTo-Json -Depth 5 -Compress
    Retry-Command -ScriptBlock {
        $script:AppRegistration = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $body -ContentType "application/json" -TimeoutSec 15
    } -Maximum 3 -Message "Creating app registraiton"

    $tenanidID = Invoke-RestMethod -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id' -Headers $headers -TimeoutSec 5

    $output = @{
        AadClientId = $AppRegistration.appId
        AadSecret = "YOUR_CLIENT_SECRET_HERE"
        AadTenantDomain = $AppRegistration.PublisherDomain
        AadTenantId =  $($tenanidID.value.id)
    }

    $output | ConvertTo-Json | Out-File $("$mydir",$($tenantPrefix,"txt"-join(".")) -join("\")) -encoding ascii
}
