# FILEPATH: create_team.ps1
#FY24 CSP Masters tenant transform script
$global:localDir = "C:\CSPMastersSourceFiles"
$global:mydir = Get-Location
$global:inputCsvFile = "M365 Business Tech Series Staging Pool.csv"
$global:Logfile = "$mydir\FY24CspLabTransform_Log.txt"
$global:appID = "aff75787-e598-43f9-a0ea-7a0ca00ababc" #MOD Demo Platform UnifiedApiConsumer
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
                    $ScriptBlock.Invoke()
                    Write_Log -Message_Type "SUCCESS" -Message $Message
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

# Function to create a new team with specified owners
function Create-TeamWithOwners {
    param (
        [Parameter(Mandatory=$true)]
        [string]$TeamName,
        [Parameter(Mandatory=$true)]
        [string[]]$Owners,
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$adminCredential
    )

    #Lets connect to Teams
    Write_Log -Message_Type "INFO" -Message "Connect-MicrosofTeams using $($adminCredential.Username)"	

    Try
        {
            Connect-MicrosoftTeams -Credential $adminCredential
            Write_Log -Message_Type "SUCCESS" -Message "Connect-MicrosofTeams"		
        }
    Catch
        {
            $message = $_
            Write_Log -Message_Type "ERROR" -Message "Connect-MicrosofTeams"
            Write_Log -Message_Type "ERROR" -Message $message
            $message = $null
            EXIT
        }

    #Check if team exists
    Write_Log -Message_Type "INFO" -Message "Checking if Team $TeamName exists"
    Try {
        $ProgressPreference = 'SilentlyContinue' #Prevent 'Fetching teams' progress indicators from showing up
        $private:existingTeam = Get-Team -DisplayName $TeamName
        if ($existingTeam) {
            Write_Log -Message_Type "SUCCESS" -Message "Team found"
            return $existingTeam
        } 		
    }
    Catch {
        $message = $_
        Write_Log -Message_Type "ERROR" -Message "Checking if Team exists"	
        Write_Log -Message_Type "ERROR" -Message $message
        $message = $null
        #EXIT
    }

    if ($existingTeam -eq $null) {
        Write_Log -Message_Type "INFO" -Message "Team $TeamName does not exist"
        #Create a new team
        Write_Log -Message_Type "INFO" -Message "Creating Team $TeamName"
        Try
            {
                $private:newteam = New-Team -DisplayName $TeamName
                Write_Log -Message_Type "SUCCESS" -Message "Creating Team"		
            }
        Catch
            {
                $message = $_
                Write_Log -Message_Type "ERROR" -Message "Creating Team"
                Write_Log -Message_Type "ERROR" -Message $message
                $message = $null
                EXIT
            }


        # Add owners to the team
        #Change this to GRAPH, MSOL is deprecated
        foreach ($owner in $Owners) {
            Write_Log -Message_Type "INFO" -Message "Adding $owner to Team"
            $upn = $owner,$tenantName -join("@")
            Try
            {
                Add-TeamUser -GroupId $newteam.GroupId -User $upn -Role Owner
                Write_Log -Message_Type "SUCCESS" -Message "Adding Owner"		
            }
            Catch
            {
                $message = $_
                Write_Log -Message_Type "ERROR" -Message "Adding Owner"
                Write_Log -Message_Type "ERROR" -Message $message
                $message = $null
                EXIT
            }
        }

        Write_Log -Message_Type "SUCCESS" -Message "Team '$TeamName' created successfully with owners: $($Owners -join ', ')"
        return $newteam
    } else {
        Write_Log -Message_Type "INFO" -Message "Team $TeamName exists no action taken"
    }
    
    
    # Disconnect from Teams
    Write_Log -Message_Type "INFO" -Message "Disconnecting from MicrosoftTeams"
    Try
        {
        Disconnect-MicrosoftTeams
        Write_Log -Message_Type "SUCCESS" -Message "Disconnected from MicrosoftTeams"		
        }
    Catch
        {
            $message = $_
            Write_Log -Message_Type "ERROR" -Message "Error disconnecting from MicrosoftTeams"	
            Write_Log -Message_Type "ERROR" -Message $message
            $message = $null
            #EXIT
        }
}


$stagingCreds = ipcsv -Path $inputCsvFile
$tenantCount = 0
foreach ($tenant in $stagingCreds){
    $tenantName = $tenant.TenantName
    $User = $tenant.'Administrative Username'
    $Pass = $tenant.'Administrative Password'

    # Check if $tenant.'Administrative Password' starts with '= and remove the single quotes if it does
    if ($Pass.StartsWith('''=')) {
        $Pass = $Pass.Replace('''=', '=')
    }

    $tenantCount++
    Write-Progress -Id 0 -Activity "Creating CSP Master Teams" -Status "Tenant $($tenant.TenantPrefix): $tenantCount of $($stagingCreds.Count)" -PercentComplete $([Math]::Ceiling($tenantCount/$stagingCreds.Count*100))
    <#Testing Creds
    $tenantName = $stagingCreds[0].TenantName
    $user = $stagingCreds[0].'Administrative Username'
    $pass = $stagingCreds[0].'Administrative Password'
    $secureSecret = ConvertTo-SecureString -String $Pass -AsPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential ($User, $secureSecret)
    Connect-MicrosoftTeams -Credential $Credential
    $team = Get-Team -DisplayName "CSP Masters Team"
    #>
    $secureSecret = ConvertTo-SecureString -String $Pass -AsPlainText -Force

    #Lets create a Credential
    Write_Log -Message_Type "INFO" -Message "Creating Credential for $tenantName"	

    Try
        {
            $Credential = New-Object System.Management.Automation.PSCredential ($User, $secureSecret)
            Write_Log -Message_Type "SUCCESS" -Message "Creating Credential"		
        }
    Catch
        {
            $message = $_
            Write_Log -Message_Type "ERROR" -Message "Creating Credential"	
            Write_Log -Message_Type "ERROR" -Message $message
            $message = $null
            EXIT
        }

    #Create Team
    Write_Log -Message_Type "INFO" -Message "Creating Team with owners for $tenantName"	

    Try
        {
        # Call the function to create a new team with specified owners
        $script:team = Create-TeamWithOwners -TeamName "CSP Masters Team" -Owners @("NestorW", "PattiF", "LidiaH", "IsaiahL") -adminCredential $Credential
        Write_Log -Message_Type "SUCCESS" -Message "Creating Team with owners"		
        }
    Catch
        {
            $message = $_
            Write_Log -Message_Type "ERROR" -Message "Creating Team with owners"	
            Write_Log -Message_Type "ERROR" -Message $message
            $message = $null
            EXIT
        }

    $Global:Token = GetAccessToken
    <#
    MOD Demo Platform UnifiedApiConsumer doesn't currently have API permission to create a Team via Graph
  
    $teamName = "CSP Masters Team"
    #Let's check if the Team exists
    Write_Log -Message_Type "INFO" -Message "Checking if Team $teamName exists"
    #$uri = "https://graph.microsoft.com/v1.0/groups/?`$filter=displayName eq '$teamName'"
    $uri = 'https://graph.microsoft.com/v1.0/teams?$filter=displayName eq ',$teamName,'&$select=id,description' -join("'")
    $count=0
    $teamExists = $null
    do {
        $count++
        Try
        {
            # Call the function to create a new team with specified owners
            $teamExists = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers -TimeoutSec 5
            Write_Log -Message_Type "SUCCESS" -Message "Checking if Team exists"		
        }
        Catch
        {
            $message = $_
            Write_Log -Message_Type "ERROR" -Message "Checking if Team exists"	
            Write_Log -Message_Type "ERROR" -Message $message
            $message = $null
            #EXIT
        }
        if ($teamExists -eq $null){
            Write_Log -Message_Type "INFO" -Message "Team not ready, retrying in 5 seconds"
            Start-Sleep -Seconds 5
        }
    } while ($teamExists -eq $null -and $count -lt 10)
    #>

    #Let's get the SharePoint URL for our Team site
    Write_Log -Message_Type "INFO" -Message "Getting SharePoint URL for GroupID $($team.GroupID)"	
    $script:uri = "https://graph.microsoft.com/v1.0/groups/$($team.GroupID)/sites/root/weburl"
    Write_Log -Message_Type "INFO" -Message "Getting SharePoint URL using $uri"	
    $count=0
    $SP_WebUrl = $null
    do {
        $count++
        Try
        {
            # Call the function to create a new team with specified owners
            $script:SP_WebUrl = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers -TimeoutSec 5
            Write_Log -Message_Type "SUCCESS" -Message "Getting SharePoint URL"		
        }
        Catch
        {
            $message = $_
            Write_Log -Message_Type "ERROR" -Message "Getting SharePoint URL"	
            Write_Log -Message_Type "ERROR" -Message $message
            $message = $null
            #EXIT
        }
        if ($SP_WebUrl -eq $null){
            Write_Log -Message_Type "INFO" -Message "SharePoint URL not ready, retrying in 5 seconds"
            Start-Sleep -Seconds 5
        }
    } while ($SP_WebUrl -eq $null -and $count -lt 10)

    $SP_Hostname = $SP_webURL.value.Split("/")[2]
    $SP_Server_Relative_Path = $SP_webURL.value.Split("/")[3],$SP_webURL.value.Split("/")[4] -join("/")
    #$Sp_docs = $sp_WebURL,"Shared%20Documents" -join("/")

    #Let get SP site info
    $uri = "https://graph.microsoft.com/v1.0/sites/",$SP_Hostname,":/",$SP_Server_Relative_Path -join("")
    Write_Log -Message_Type "INFO" -Message "Getting SharePoint site info for $($SP_Hostname,':/',$SP_Server_Relative_Path -join(''))"
    Retry-Command -ScriptBlock {
        $script:spSite = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers -TimeoutSec 5
    } -Maximum 3 -Message "Getting SharePoint site info"

    $siteId = $spSite.id.split(",")[1]

    $SharePoint_Graph_URL = "https://graph.microsoft.com/v1.0/sites/$siteId/drives"
    #$BodyJSON = $Body | ConvertTo-Json -Compress  

    Write_Log -Message_Type "INFO" -Message "Getting SharePoint drive info"
    Retry-Command -ScriptBlock {
        $script:Result = Invoke-RestMethod -Uri $SharePoint_Graph_URL -Method 'GET' -Headers $headers -ContentType "application/json" -TimeoutSec 5
    } -Maximum 3 -Message "Getting SharePoint drive"

    $DriveID = $Result.value | Select-Object id -ExpandProperty id
    $filesToUpload = Get-ChildItem -Path $localDir -Recurse | Where-Object {$_.PSIsContainer -eq $false}
    [Console]::TreatControlCAsInput = $True
    Start-Sleep -Seconds 1
    $Host.UI.RawUI.FlushInputBuffer()

    foreach ($file in $FilesToUpload){
        If ($Host.UI.RawUI.KeyAvailable -and ($Key = $Host.UI.RawUI.ReadKey("AllowCtrlC,NoEcho,IncludeKeyUp"))) {
            If ([Int]$Key.Character -eq 3) {
                Write-Host ""
                Write-Warning "CTRL-C was used - Shutting down any running jobs before exiting the script."
                [Console]::TreatControlCAsInput = $False
                Exit -HardExit $true
            }
            # Flush the key buffer again for the next loop.
            $Host.UI.RawUI.FlushInputBuffer()
            }
    
        

        $SharePoint_ExportFolder = "General/$($file.Directory.Name)" # folder where to upload file

        #Check to see if the file exists
        $fileExists = $null
        $uri = "https://graph.microsoft.com/v1.0/sites/$siteId/drives/$DriveID/root:/$SharePoint_ExportFolder/$($file.Name)"
        Write_Log -Message_Type "INFO" -Message "Checking if file $($file.Name) exists"
        try
        {
            $fileExists = Invoke-RestMethod -Uri $uri -Method 'GET' -Headers $headers -ContentType "application/json" -TimeoutSec 10 -ErrorAction SilentlyContinue
            if ($fileExists) {
                Write_Log -Message_Type "SUCCESS" -Message "File $($file.Name) already exists, skipping"
                continue
            } 
        }
        catch
        {
            Write_Log -Message_Type "INFO" -Message "File $($file.Name) does not exist, uploading"
        }
        
        

        $createUploadSessionUri = "https://graph.microsoft.com/v1.0/sites/$siteId/drives/$DriveID/root:/$SharePoint_ExportFolder/$($file.Name):/createUploadSession"
        $response = $null
        #{
        #"@microsoft.graph.conflictBehavior": "fail (default) | replace | rename",
        #"description": "description",
        #"fileSize": 1234,
        #"name": "filename.txt"
        #}
        $uploadBody = @{
            item = @{
                "@microsoft.graph.conflictBehavior" = "replace"
                description = "CSP Masters Lab File"
                fileSize = $file.Length
                name = $file.Name
            }
            deferCommit = $true
        } | ConvertTo-Json -Compress

        $uploadBody = @{
            item = @{
                "@microsoft.graph.conflictBehavior" = "replace"
            }
            deferCommit = $true
        } | ConvertTo-Json -Compress

        #POST https://sn3302.up.1drv.com/up/fe6987415ace7X4e1eF866337
        #Content-Length: 0

        #Prepare for file upload
        Write_Log -Message_Type "INFO" -Message "Uploading the file $($File.Name)"
        do
        {
            $count++
            Check-AccessToken
            Try
                {
                    $retry = $false
                    $uploadSession = Invoke-RestMethod -Uri $createUploadSessionUri -Method 'POST' -Headers $headers -body $uploadBody -ContentType "application/json" -TimeoutSec 5
                    Write_Log -Message_Type "SUCCESS" -Message "Preparing the file for the upload"
                    #Read File
                    $fileInBytes = [System.IO.File]::ReadAllBytes($file.FullName)
                    Write_Log -Message_Type "SUCCESS" -Message "Reading the file"
                    $fileLength = $fileInBytes.Length
                    $timeoutSec = $fileLength/307200 +10 # 3KB/sec miniumum upload speed plus buffer
                    
                    if ($fileLength -lt 4194304)
                    {
                        $Uploadheaders = @{'Content-Range' = "bytes 0-$($fileLength-1)/$fileLength"}
                        #Upload the file
                        Write_Log -Message_Type "INFO" -Message "Uploading file $($file.FullName)"
                        $response = Invoke-RestMethod -Method 'Put' -Uri $uploadSession.uploadUrl -Body $fileInBytes -Headers $Uploadheaders -TimeoutSec $timeoutSec
                        Write_Log -Message_Type "SUCCESS" -Message "Uploading the file"		
                    } else {
                        #$partSizeBytes = 320 * 1024 * 4  #Uploads 1.31MiB at a time.
                        $partSizeBytes = 1024 * 1024 * 10  #Uploads 10MiB at a time.
                        $index = 0
                        $start = 0
                        $end = 0
                        
                        $maxloops = [Math]::Round([Math]::Ceiling($fileLength / $partSizeBytes))

                        while ($fileLength -gt ($end + 1)) {
                            $start = $index * $partSizeBytes
                            if (($start + $partSizeBytes - 1 ) -lt $fileLength) {
                                $end = ($start + $partSizeBytes - 1)
                            }
                            else {
                                $end = ($start + ($fileLength - ($index * $partSizeBytes)) - 1)
                            }
                            [byte[]]$bodyBytes = $fileInBytes[$start..$end]
                            $Uploadheaders = @{    
                                'Content-Range' = "bytes $start-$end/$fileLength"
                            }
                            #Write_Log -Activity "INFO" -Message "bytes $start-$end/$fileLength | Index: $index and ChunkSize: $partSizeBytes"
                            $response = Invoke-WebRequest -Method Put -Uri $uploadSession.uploadUrl -Body $bodyBytes -Headers $Uploadheaders -TimeoutSec $timeoutSec
                            $index++
                            #Write_Log -Message_Type "SUCCESS" -Message "Percentage Complete: $([Math]::Ceiling($index/$maxloops*100)) %"
                            Write-Progress -Id 1 -ParentId 0 -Activity "Upload in Progress: $($File.Name)" -Status "$([Math]::Ceiling($index/$maxloops*100))% Complete" -PercentComplete $([Math]::Ceiling($index/$maxloops*100));
                        }
                        Write-Progress -Id 1 -ParentId 0 -Activity "Upload in Progress: $($File.Name)" -Status "100% Complete" -Completed
                    }
        
                }
            Catch
                {
                    $message = $_
                    Write_Log -Message_Type "ERROR" -Message "Uploading the file"
                    Write_Log -Message_Type "ERROR" -Message $message
                    Write_Log -Message_Type "CODE:" -Message $_.Exception.Response.StatusCode.value__
                    Write_Log -Message_Type "DESCRIPTION:" -Message $_.Exception.Response.StatusDescription
                    Write_Log -Message_Type "RESPONSE" -Message $response
                    $retry= $true
                    $message = $null
                    #EXIT
                }
            Finally
                {
                    if ($retry -eq $true) {
                        Write_Log -Message_Type "INFO" -Message "File not uploaded, retrying in 5 seconds"
                        Start-Sleep -Seconds 5
                    } else {
                        Write_Log -Message_Type "SUCCESS" -Message "Committing the file"
                        Invoke-RestMethod -Method POST -Uri $uploadSession.uploadUrl -Headers @{'Content-Length'=0} -TimeoutSec $timeoutSec | Out-Null
                    }
                }
        } while ($retry -eq $true -and $count -lt 10)
        }
}




