# FILEPATH: create_w365.ps1
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
#Create W365 Provisioning Profile
foreach ($tenant in $stagingCreds){
    $tenantName = $tenant.TenantName
    $User = $tenant.'Administrative Username'
    $Pass = $tenant.'Administrative Password'
    
    # Check if $tenant.'Administrative Password' starts with '= and remove the single quotes if it does
    if ($Pass.StartsWith('''=')) {
        $Pass = $Pass.Replace('''=', '=')
    }

    $tenantCount++
    Write-Progress -id 0 -Activity "Creating W365 provisiong policy" -Status "Tenant $($tenant.TenantPrefix): $tenantCount of $($stagingCreds.Count)" -PercentComplete ($tenantCount/$stagingCreds.Count*100)

    <#Test User
    $tenantName = $stagingCreds[0].TenantName
    $User = $stagingCreds[0].'Administrative Username'
    $Pass = $stagingCreds[0].'Administrative Password'
    #>

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
            $AccessToken = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$($tenantName)/oauth2/v2.0/token" -Body $Authbody -ErrorAction Stop
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
    
    $regions = @(
        "australia",
        "canada",
        "usCentral",
        "usEast",
        "usWest",
        "europeUnion"
    )

    $params = @{
        "@odata.type" = "#microsoft.graph.cloudPcProvisioningPolicy"
        displayName = "CSP Masters Policy"
        description = "Provisioning Policy to create Cloud PCs for members for the CSP Masters"
        provisioningType = "dedicated"
        managedBy = "windows365"
        domainJoinConfiguration = @{
            domainJoinType = "azureADJoin"
            regionGroup = Get-Random -InputObject $regions
            regionName = "automatic"
        }
        enableSingleSignOn = $true
        imageDisplayName = "Windows 11 Enterprise + Microsoft 365 Apps 23H2"
        imageId = "microsoftwindowsdesktop_windows-ent-cpc_win11-23h2-ent-cpc-m365"
        imageType = "gallery"
        microsoftManagedDesktop = @{
            type ="notManaged"
            profile = ""
        }
        windowsSettings = @{
            language = "en-US"
        }
        cloudPcNamingTemplate = "%USERNAME:7%PC-%RAND:5%"
    }
    $BodyJSON = $params | ConvertTo-Json -Compress  

    #Lets create the provisioning policy
    Write_Log -Message_Type "INFO" -Message "Creating W365 Provisioning Policy for $tenantName in $($params.domainJoinConfiguration.regionGroup)"	
    Retry-Command -ScriptBlock {
        $script:W365Policy = Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies" -Headers $headers -Body $BodyJSON -ContentType "application/json"
    } -Maximum 3 -Message "Creating W365 Provisioning Policy for $tenantName"

    #Try
    #    {
    #        $W365Policy = Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies" -Headers $headers -Body $BodyJSON -ContentType "application/json"
    #        Write_Log -Message_Type "SUCCESS" -Message "Creating W365 Provisioning Policy"	
    #    }
    #Catch
    #    {
    #        $message = $_
    #        Write_Log -Message_Type "ERROR" -Message "Creating W365 Provisioning Policy"
    #        Write_Log -Message_Type "ERROR" -Message $message
    #        $message = $null
    #        #EXIT
    #    }
    
    $owner = "admin"
    $members = "admin", "PattiF", "NestorW", "IsaiahL", "LidiaH"
    #Let's get the Device Configurations
    $graphApiVersion = "v1.0"
    $User_resource = "users"
    $groupmembers = @()
    foreach ($member in $members) {
        $upn = $member,$tenantName -join("@")
        $uri = "https://graph.microsoft.com/$graphApiVersion/$($User_resource)/$upn"

        Write_Log -Message_Type "INFO" -Message "Getting $member UserId for $tenantName"	
        $count=0
        $userId  = $null
        do {
            $count++
            Try
            {
                # Call the function to create a new team with specified owners
                $userId = (Invoke-RestMethod -Uri $uri -Headers $headers -TimeoutSec 60).Id
                Write_Log -Message_Type "SUCCESS" -Message "Getting User"		
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
                Write_Log -Message_Type "ERROR" -Message "Getting User"
                Write_Log -Message_Type "ERROR" -Message $message
                $message = $null
                #EXIT
            }
            if ($userId -eq $null){
                Write_Log -Message_Type "INFO" -Message "User not ready, retrying in 5 seconds"
                Start-Sleep -Seconds 5
            }
        } while ($userId -eq $null -and $count -lt 10)

        $groupmembers += "https://graph.microsoft.com/$graphApiVersion/users/$($userId)"

    }

    $ownerID = @()
    $groupowners = @()
    $upn = $owner,$tenantName -join("@")
    $uri = "https://graph.microsoft.com/$graphApiVersion/$($User_resource)/$upn"
    Write_Log -Message_Type "INFO" -Message "Getting admin user id for $tenantName"	
    $count=0
    do {
        $count++
        Try
        {
            # Call the function to create a new team with specified owners
            $ownerID += (Invoke-RestMethod -Uri $uri -Headers $headers -TimeoutSec 60).Id
            $groupowners += "https://graph.microsoft.com/v1.0/users/$($ownerID)"
            Write_Log -Message_Type "SUCCESS" -Message "Getting User"		
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
            Write_Log -Message_Type "ERROR" -Message "Getting User"
            Write_Log -Message_Type "ERROR" -Message $message
            $message = $null
            #EXIT
        }
        if ($groupowners.count -eq 0){
            Write_Log -Message_Type "INFO" -Message "User not ready, retrying in 5 seconds"
            Start-Sleep -Seconds 5
        }
    } while ($groupowners.count -eq 0 -and $count -lt 10)

    #Create a group
    #POST https://graph.microsoft.com/v1.0/groups
    #Content-type: application/json
    $params = @{
        displayName = "CSP Masters W365"
        description = "CSP Masters W365 Users"
        mailEnabled = $false
        mailNickname = "CSPMastersW365"
        securityEnabled = $true
        "owners@odata.bind" = $groupowners
        "members@odata.bind" = $groupmembers
    }
    #Let's get the Device Configurations
    $body = $params | ConvertTo-Json -Compress
    $graphApiVersion = "v1.0"
    $resource = "groups"

    $uri = "https://graph.microsoft.com/$graphApiVersion/$($resource)"
    Write_Log -Message_Type "INFO" -Message "Creating group $($params.displayName) for tenant $tenantName"	
    $count=0
    $W365groupID  = $null
    do {
        $count++
        Try
        {
            # Call the function to create a new team with specified owners
            $W365groupID = (Invoke-RestMethod -Uri $uri -Headers $headers -Method POST -Body $body -ContentType "application/json" -TimeoutSec 60).Id
            Write_Log -Message_Type "SUCCESS" -Message "Creating Group"		
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
            Write_Log -Message_Type "ERROR" -Message "Deleting Device Configuration $configId"	
            Write_Log -Message_Type "ERROR" -Message $message
            $message = $null
            $message = $_
            Write_Log -Message_Type "ERROR" -Message "Creating Group"		
            Write_Log -Message_Type "ERROR" -Message $message
            $message = $null
            EXIT
        }
        if ($W365groupID -eq $null){
            Write_Log -Message_Type "INFO" -Message "Group not ready, retrying in 5 seconds"
            Start-Sleep -Seconds 5
        }
    } while ($W365groupID -eq $null -and $count -lt 10)

    #$groupName = "CSP Masters Team"
    #Let's find the group ID for the CSP Masters Team
    #GET https://graph.microsoft.com/v1.0/groups?$search="displayName:CSP Masters Team"
    #ConsistencyLevel: eventual
    #Write_Log -Message_Type "INFO" -Message "Looking for groupID for $groupName"
    #$uri = 'https://graph.microsoft.com/v1.0/groups?$filter=displayName eq ',"'$groupName'" -join("")
    #Retry-Command -ScriptBlock {
    #    $script:group = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers        
    #} -Maximum 3 -Message "Finding groupID for $groupName"
    #Try
    #    {
    #        $uri = 'https://graph.microsoft.com/v1.0/groups?$filter=displayName eq ',"'$groupName'" -join("")
    #        $group = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
    #        Write_Log -Message_Type "SUCCESS" -Message "Found Group"	
    #    }
    #Catch
    #    {
    #        $message = $_
    #        Write_Log -Message_Type "ERROR" -Message "Finding Group"
    #        Write_Log -Message_Type "ERROR" -Message $message
    #        $message = $null
    #        EXIT
    #    }
    
    #Lets assign the provisioning policy
    #POST https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies/{id}/assign
    #Content-Type: application/json
    
    $params = @{
        assignments = @(
            @{
    
                target = @{
                    "@odata.type" = "microsoft.graph.cloudPcManagementGroupAssignmentTarget"
                    groupId = $W365groupID
                }
            }
        )
    }
        $BodyJSON = $params | ConvertTo-Json -Depth 5 -Compress  
    
        Write_Log -Message_Type "INFO" -Message "Assigning W365 Provisioning Policy for $tenantName to $($group.value.displayName)"
        $uri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies/$($W365Policy.id)/assign"
        Retry-Command -ScriptBlock {
            Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $BodyJSON -ContentType "application/json"
        } -Maximum 3 -Message "Assigning W365 Provisioning Policy"


        #Try
        #    {
        #        $uri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies/$($W365Policy.id)/assign"
        #        Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $BodyJSON -ContentType "application/json"
        #        Write_Log -Message_Type "SUCCESS" -Message "Assigned W365 Provisioning Policy"	
        #    }
        #Catch
        #    {
        #        $message = $_
        #        Write_Log -Message_Type "ERROR" -Message "Assigning W365 Provisioning Policy"
        #        Write_Log -Message_Type "ERROR" -Message $message
        #        $message = $null
        #        EXIT
        #    }
        


}