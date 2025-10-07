<#Test Creds
$tenant = 'a12635bc-203a-474b-91b7-731bb17331ec'
$ApplicationId = 'c2c179af-4c51-4946-a8c3-89aeedd7e5d5'
$appSecret = 'vQJutsXfs7IUM7axiX1hWbgVgWV4nhIHYvYglNuRppQ='
$connectionSub = '@lab.CloudSubscription.Id'
#>



#Credentials for M365 Subscription - Do Not Modify
$tenant = '@lab.CloudSubscription.TenantId'
$ApplicationId = '@lab.CloudSubscription.AppId'
$appSecret = '@lab.CloudSubscription.AppSecret'
$connectionSub = '@lab.CloudSubscription.Id'

$SecuredPassword = ConvertTo-Securestring "$appSecret" -AsPlainText -Force
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($ApplicationId, $SecuredPassword)

Connect-AzAccount -Credential $Credential -TenantId $tenant -ServicePrincipal -SubscriptionId $connectionSub | Out-Null

$azureUri = "https://login.microsoftonline.com/$tenant/oauth2/token" 
$body = @{"grant_Type" = "client_credentials"
"client_id" = "$ApplicationId"
"client_secret" = "$appSecret"
"resource" = "https://graph.microsoft.com"}

$AuthRequest = Invoke-RestMethod -Uri $azureUri -Body $body -Method Post
$authToken = $AuthRequest.access_token
$graphPass = ConvertTo-Securestring "$authToken" -AsPlainText -Force

Connect-MgGraph -AccessToken $graphPass -NoWelcome

# Base License Options
$m365e3Sku = Get-MgSubscribedSku -All | Where SkuPartNumber -eq 'Microsoft_365_E3_(no_Teams)'
$m365e5Sku = Get-MgSubscribedSku -All | Where SkuPartNumber -eq 'Microsoft_365_E5_(no_Teams)'

# Addon License Options
$copilotSku = Get-MgSubscribedSku -All | Where SkuPartNumber -eq 'Microsoft_365_Copilot'
$PBIPremSku = Get-MgSubscribedSku -All | Where SkuPartNumber -eq 'PBI_PREMIUM_P1_ADDON'    
$teamsEntSku = Get-MgSubscribedSku -All | Where SkuPartNumber -eq 'Microsoft_Teams_Enterprise_New'  

# License Detection
$e3License = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Virtual Machine\External' -Name LabTag_994 -ErrorAction Ignore
$e5License = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Virtual Machine\External' -Name LabTag_995 -ErrorAction Ignore
$CopilotLicense = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Virtual Machine\External' -Name LabTag_996 -ErrorAction Ignore
$PBIPremLicense = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Virtual Machine\External' -Name LabTag_997 -ErrorAction Ignore
$teamsEnt = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Virtual Machine\External' -Name LabTag_1022 -ErrorAction Ignore
$addLicenses=@() 

# Determine Licenses
if ($null -eq $e5License) {
     if ($null -eq $e3License) {
        Write-Error "No Base License Detected"
    }else{
        $365license = $m365e3sku
    }
}else{
    $365license = $M365e5sku
}


# Disable Apps
$disabledPlans = $365license.ServicePlans | Where ServicePlanName -in ("YAMMER_ENTERPRISE", "VIVAENGAGE_CORE", "KAIZALA_O365_P3") | Select -ExpandProperty ServicePlanId
$disabledTeamsPlans = $teamsEntSku.ServicePlans | Where ServicePlanName -in ("MCOIMP") | Select -ExpandProperty ServicePlanId

$addLicenses+=@{
    SkuId = $365license.SkuId
    DisabledPlans = $disabledPlans
}

#if ($teamsEnt -ne "false"){
if ($null -ne $teamsEnt){
    $addLicenses+=@{
        SkuId = $teamsEntSku.SkuId
        DisabledPlans = $disabledTeamsPlans
    }
}

if ($null -ne $CopilotLicense){
    $addLicenses+=@{SkuId = $copilotSku.SkuId}
}
    
if ($null -ne $PBIPremLicense){
    $addLicenses+=@{SkuId = $PBIPremSku.SkuId}
} 


# set user and location
$userUPN="@lab.CloudPortalCredential(User1).Username"
$userLoc="US"

# update the users location
Update-MgUser -UserId $userUPN -UsageLocation $userLoc

# Assign selected license to user

Set-MgUserLicense -UserId $userUPN -AddLicenses $addLicenses -RemoveLicenses @()