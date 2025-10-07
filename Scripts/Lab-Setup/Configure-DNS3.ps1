[CmdletBinding()]
Param(
    [parameter (Mandatory=$true,position=0)]
    $tenantAdminUser,
    [parameter (Mandatory=$true,position=1)]
    $tenantAdminPassword,
    [parameter (Mandatory=$true,position=2)]
    $tenantVanityDomain
)

$myAzTenantId = "72f988bf-86f1-41af-91ab-2d7cd011db47"
$myAzApplicantionId = "6ad39a5d-0e9c-4c11-b9e9-65fc4b990332"
$myDnsResourceGroupName = "rg-m365master"
#$myClientPrincipalName = "lods-secret"
#$myClientSecret = "1ouQ-_.m2LURELtL4G4GGC.uiaYfDE289-"
$myClientPrincipalName = "lods-secretFY22"
$myClientSecret = "RWj.JzxTvwCmwxC6.mk.P.g2LRfCIzb8v0"



#Make sure AzureRM is installed
#Install-Module -Name AzureRM -AllowClobber

#function Check-Prerequsites {
#    try {
#        $module = Find-Module -Name AzureRM.Dns -ErrorAction STOP
#        if ($module.Version -le $moduleDnsVersionRequired) {
#            Throw $moduleDnsVersionRequired_Error
#        }
#        $module = Find-Module -Name AzureRM.Network -ErrorAction STOP
#        if ($module.Version -le $moduleNetworkVersionRequired) {
#            Throw $moduleNetworkVersionRequired_Error
#        }
#    }
#    catch {
#        Write-Host $PSItem.Exception.Message -ForegroundColor RED
#    }
#    finally {
#        $Error.Clear()
#    }
#}

$secureSecret = ConvertTo-SecureString -String $myClientSecret  -AsPlainText -Force
$myAzCredential = New-Object System.Management.Automation.PSCredential ($myAzApplicantionId, $secureSecret)

#Connect to Azure and stay connected 
Connect-AzAccount -Credential $myAzCredential -ServicePrincipal -TenantId $myAzTenantId

Class demoTenant{
    [string] $Name;
    [string] $adminUser;
    [string] $labSuffix = "m365master.com";
    [string] $mxSuffix = "mail.protection.outlook.com"
    [PSCredential] $psCredential;
    [string] $verifyTxt;
    [boolean] $isVerified = $False;
    [string] $dnsResourceGroupName = "rg-m365master";
    [string] $customDomain;
    [string] $domainGUID;
    [string] $initialDomain;
    [string] $mxRecord;
    [string] $spfValue = "v=spf1 include:spf.protection.outlook.com -all";
    [string] $autodiscoverCname = "autodiscover.outlook.com";
    [string] $sipCname = "sipdir.online.lync.com";
    [string] $lyncdiscoverCname = "webdir.online.lync.com";
    [string] $sipSrv = "sipdir.online.lync.com";
    [string] $sipSrvProtocol = "_tls";
    [string] $sipSrvPort = 443;
    [string] $sipSrvWeight = 1;
    [string] $sipSrvPriority = 100;
    [string] $sipfederationtlsSrv = "sipfed.online.lync.com";
    [string] $sipfederationtlsSrvProtocol = "_tcp";
    [string] $sipfederationtlsSrvPort = 5061;
    [string] $sipfederationtlsSrvWeight = 1;
    [string] $sipfederationtlsSrvPriority = 100;
    [string] $enterpriseregistrationCname = "enterpriseregistration.windows.net";
    [string] $enterpriseenrollmentCname = "enterpriseenrollment.manage.microsoft.com";
    [string] $selector1Cname;
    [string] $selector2Cname;
    [string] $dmarcTxt;
    demoTenant([string]$adminUserIn, [string]$adminPasswordIn, [string]$vanityDomainIn) {
        $this.Name = $vanityDomainIn;
        $this.adminUser = $adminUserIn;
        #$this.psCredential = [PSCredential]::New($adminUserIn,(ConvertTo-SecureString -String $adminPasswordIn -AsPlainText -Force))
        $this.psCredential = Get-Credential -UserName $adminUserIn
        $this.customDomain = $vanityDomainIn,".",$this.labSuffix -join ''
        $this.domainGUID = $this.customDomain -replace "\.","-"
        $this.mxRecord = $this.domainGUID,$this.mxSuffix -join '.'
    }
    [string] labName(){
        $a = $null
        $a = $this.Name,".", $this.labSuffix -join ''
        return $a
    }
    [void] addCustomDomain(){
       Connect-MsolService -Credential $this.psCredential
       #Set initial domain
       $this.initialDomain = (Get-MsolDomain | ? {$_.IsInitial -eq $true}).Name
       #Add code to Check if domain exists
       #
       if(Get-MsolDomain -DomainName $this.labName()){
        $this.verifyTxt = (Get-MsolDomainVerificationDns -DomainName $this.labName() -Mode DnsTxtRecord).Text
        }else {
        New-MsolDomain -Name $this.labName()
        $this.verifyTxt = (Get-MsolDomainVerificationDns -DomainName $this.labName() -Mode DnsTxtRecord).Text
       }
       $this.selector1Cname = ("selector1",$this.domainGUID -join "-"),"_domainkey",$this.initialDomain -join "."
       $this.selector2Cname = ("selector2",$this.domainGUID -join "-"),"_domainkey",$this.initialDomain -join "."
       $this.dmarcTxt = "v=DMARC1; p=reject; pct=100; rua=mailto:",(($this.adminUser -split ('@'))[0],$this.customDomain -join "@"),"; ruf=mailto:",$this.adminUser,"; fo=1" -join ''
    }
}
Function addVerifyTxt{
<#
    .SYNOPSIS
    
    .DESCRIPTION
    
    .PARAMETER
    
    .EXPAMPLE
#>

param (
    [parameter (Mandatory=$true,position=0)]
    $rgName,
    [parameter (Mandatory=$true,position=1)]
    $msolTenant
)
    #Check to see if recordset already exists
    #
    if (Get-AzDnsRecordSet -Name $msolTenant.Name -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType TXT) {
        $rs = Get-AzDnsRecordSet -Name $msolTenant.Name -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType TXT
        $rs.records[0].Value = $msolTenant.VerifyTxt
        Set-AzDnsRecordSet -RecordSet $rs
    }else {
        New-AzDnsRecordSet -Name $msolTenant.Name -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -Ttl 3600 -RecordType TXT -DnsRecords (New-AzDnsRecordConfig -Value $msolTenant.VerifyTxt)
    }
}
Function mxRecord{
    <#
        .SYNOPSIS
        
        .DESCRIPTION
        
        .PARAMETER
        
        .EXPAMPLE
    #>
    
    param (
        [parameter (Mandatory=$true,position=0)]
        $rgName,
        [parameter (Mandatory=$true,position=1)]
        $msolTenant,
        [parameter (Mandatory=$false,position=2)]
        [Switch]$Remove
    )
        #Check to see if recordset already exists
        if (Get-AzDnsRecordSet -Name $msolTenant.Name -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType MX) {
            $rs = Get-AzDnsRecordSet -Name $msolTenant.Name -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType MX
            $rs.records[0].Preference = 0
            $rs.records[0].Exchange = $msolTenant.mxRecord
            #If -Remove then remove Recordset
            if ($remove) {
                Remove-AzDnsRecordSet -RecordSet $rs
            #Otherswise update Recordset
            } else {
                Set-AzDnsRecordSet -RecordSet $rs
            }
         
        }
        #If RecordSet does not exist and we didn't call -Remove switch  
        elseif (-not $remove) {
            New-AzDnsRecordSet -Name $msolTenant.Name -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -Ttl 3600 -RecordType MX -DnsRecords (New-AzDnsRecordConfig -Exchange $msolTenant.mxRecord -Preference 0)
        }
}

Function autodiscoverCname{
    <#
        .SYNOPSIS
        
        .DESCRIPTION
        
        .PARAMETER
        
        .EXPAMPLE
    #>
    
    param (
        [parameter (Mandatory=$true,position=0)]
        $rgName,
        [parameter (Mandatory=$true,position=1)]
        $msolTenant,
        [parameter (Mandatory=$false,position=2)]
        [Switch]$Remove
    )
        #Check to see if recordset already exists
        if (Get-AzDnsRecordSet -Name $("autodiscover",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType CNAME) {
            $rs = Get-AzDnsRecordSet -Name $("autodiscover",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType CNAME
            $rs.records[0].Cname = $msolTenant.autodiscoverCname
            if ($remove) {
                Remove-AzDnsRecordSet -RecordSet $rs
            #Otherswise update Recordset
            } else {
                Set-AzDnsRecordSet -RecordSet $rs
            }
        }
        #If RecordSet does not exist and we didn't call -Remove switch  
        elseif (-not $remove) {
            New-AzDnsRecordSet -Name $("autodiscover",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -Ttl 3600 -RecordType CNAME -DnsRecords (New-AzDnsRecordConfig -Cname $msolTenant.autodiscoverCname)
        }
}
Function sipCname{
    <#
        .SYNOPSIS
        
        .DESCRIPTION
        
        .PARAMETER
        
        .EXPAMPLE
    #>
    
    param (
        [parameter (Mandatory=$true,position=0)]
        $rgName,
        [parameter (Mandatory=$true,position=1)]
        $msolTenant,
        [parameter (Mandatory=$false,position=2)]
        [Switch]$Remove
    )
        #Check to see if recordset already exists
        #
        if (Get-AzDnsRecordSet -Name $("sip",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType CNAME) {
            $rs = Get-AzDnsRecordSet -Name $("sip",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType CNAME
            $rs.records[0].Cname = $msolTenant.sipCname
            if ($remove) {
                Remove-AzDnsRecordSet -RecordSet $rs
            #Otherswise update Recordset
            } else {
                Set-AzDnsRecordSet -RecordSet $rs
            }
        }
        #If RecordSet does not exist and we didn't call -Remove switch
        elseif (-not $remove) {
            New-AzDnsRecordSet -Name $("sip",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -Ttl 3600 -RecordType CNAME -DnsRecords (New-AzDnsRecordConfig -Cname $msolTenant.sipCname)
        }
}
Function lyncDiscoverCname{
    <#
        .SYNOPSIS
        
        .DESCRIPTION
        
        .PARAMETER
        
        .EXPAMPLE
    #>
    
    param (
        [parameter (Mandatory=$true,position=0)]
        $rgName,
        [parameter (Mandatory=$true,position=1)]
        $msolTenant,
        [parameter (Mandatory=$false,position=2)]
        [Switch]$Remove
    )
        #Check to see if recordset already exists
        #
        if (Get-AzDnsRecordSet -Name $("lyncdiscover",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType CNAME) {
            $rs = Get-AzDnsRecordSet -Name $("lyncdiscover",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType CNAME
            $rs.records[0].Cname = $msolTenant.lyncdiscoverCname
            if ($remove) {
                Remove-AzDnsRecordSet -RecordSet $rs
            #Otherswise update Recordset
            } else {
                Set-AzDnsRecordSet -RecordSet $rs
            }
        }
        #If RecordSet does not exist and we didn't call -Remove switch
        elseif (-not $remove) {
            New-AzDnsRecordSet -Name $("lyncdiscover",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -Ttl 3600 -RecordType CNAME -DnsRecords (New-AzDnsRecordConfig -Cname $msolTenant.lyncdiscoverCname)
        }
}
Function enterpriseregistrationCname{
    <#
        .SYNOPSIS
        
        .DESCRIPTION
        
        .PARAMETER
        
        .EXPAMPLE
    #>
    
    param (
        [parameter (Mandatory=$true,position=0)]
        $rgName,
        [parameter (Mandatory=$true,position=1)]
        $msolTenant,
        [parameter (Mandatory=$false,position=2)]
        [Switch]$Remove
    )
        #Check to see if recordset already exists
        #
        if (Get-AzDnsRecordSet -Name $("enterpriseregistration",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType CNAME) {
            $rs = Get-AzDnsRecordSet -Name $("enterpriseregistration",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType CNAME
            $rs.records[0].Cname = $msolTenant.enterpriseregistrationCname
            if ($remove) {
                Remove-AzDnsRecordSet -RecordSet $rs
            #Otherswise update Recordset
            } else {
            Set-AzDnsRecordSet -RecordSet $rs
            }
        }
        #If RecordSet does not exist and we didn't call -Remove switch
        elseif (-not $remove) {
            New-AzDnsRecordSet -Name $("enterpriseregistration",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -Ttl 3600 -RecordType CNAME -DnsRecords (New-AzDnsRecordConfig -Cname $msolTenant.enterpriseregistrationCname)
        }
}
Function enterpriseenrollmentCname{
    <#
        .SYNOPSIS
        
        .DESCRIPTION
        
        .PARAMETER
        
        .EXPAMPLE
    #>
    
    param (
        [parameter (Mandatory=$true,position=0)]
        $rgName,
        [parameter (Mandatory=$true,position=1)]
        $msolTenant,
        [parameter (Mandatory=$false,position=2)]
        [Switch]$Remove
    )
        #Check to see if recordset already exists
        #
        if (Get-AzDnsRecordSet -Name $("enterpriseenrollment",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType CNAME) {
            $rs = Get-AzDnsRecordSet -Name $("enterpriseenrollment",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType CNAME
            $rs.records[0].Cname = $msolTenant.enterpriseenrollmentCname
            if ($remove) {
                Remove-AzDnsRecordSet -RecordSet $rs
            #Otherswise update Recordset
            } else {
            Set-AzDnsRecordSet -RecordSet $rs
            }
        }
        #If RecordSet does not exist and we didn't call -Remove switch
        elseif (-not $remove) {
            New-AzDnsRecordSet -Name $("enterpriseenrollment",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -Ttl 3600 -RecordType CNAME -DnsRecords (New-AzDnsRecordConfig -Cname $msolTenant.enterpriseenrollmentCname)
        }
}
Function selector1Cname{
    <#
        .SYNOPSIS
        
        .DESCRIPTION
        
        .PARAMETER
        
        .EXPAMPLE
    #>
    
    param (
        [parameter (Mandatory=$true,position=0)]
        $rgName,
        [parameter (Mandatory=$true,position=1)]
        $msolTenant,
        [parameter (Mandatory=$false,position=2)]
        [Switch]$Remove
    )
        #Check to see if recordset already exists
        #
        if (Get-AzDnsRecordSet -Name $("selector1._domainkey",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType CNAME) {
            $rs = Get-AzDnsRecordSet -Name $("selector1._domainkey",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType CNAME
            $rs.records[0].Cname = $msolTenant.selector1Cname
            if ($remove) {
                Remove-AzDnsRecordSet -RecordSet $rs
            #Otherswise update Recordset
            } else {
            Set-AzDnsRecordSet -RecordSet $rs
            }
        }
        #If RecordSet does not exist and we didn't call -Remove switch
        elseif (-not $remove) {
            New-AzDnsRecordSet -Name $("selector1._domainkey",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -Ttl 3600 -RecordType CNAME -DnsRecords (New-AzDnsRecordConfig -Cname $msolTenant.selector1Cname)
        }
}
Function selector2Cname{
    <#
        .SYNOPSIS
        
        .DESCRIPTION
        
        .PARAMETER
        
        .EXPAMPLE
    #>
    
    param (
        [parameter (Mandatory=$true,position=0)]
        $rgName,
        [parameter (Mandatory=$true,position=1)]
        $msolTenant,
        [parameter (Mandatory=$false,position=2)]
        [Switch]$Remove
    )
        #Check to see if recordset already exists
        #
        if (Get-AzDnsRecordSet -Name $("selector2._domainkey",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType CNAME) {
            $rs = Get-AzDnsRecordSet -Name $("selector2._domainkey",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType CNAME
            $rs.records[0].Cname = $msolTenant.selector2Cname
            if ($remove) {
                Remove-AzDnsRecordSet -RecordSet $rs
            #Otherswise update Recordset
            } else {
            Set-AzDnsRecordSet -RecordSet $rs
            }
        }
        #If RecordSet does not exist and we didn't call -Remove switch
        elseif (-not $remove) {
            New-AzDnsRecordSet -Name $("selector2._domainkey",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -Ttl 3600 -RecordType CNAME -DnsRecords (New-AzDnsRecordConfig -Cname $msolTenant.selector2Cname)
        }
}
Function spfTxt{
    <#
        .SYNOPSIS
        
        .DESCRIPTION
        
        .PARAMETER
        
        .EXPAMPLE
    #>
    
    param (
        [parameter (Mandatory=$true,position=0)]
        $rgName,
        [parameter (Mandatory=$true,position=1)]
        $msolTenant,
        [parameter (Mandatory=$false,position=2)]
        [Switch]$Remove
    )
        #Check to see if recordset already exists
        #
        if (Get-AzDnsRecordSet -Name $msolTenant.Name -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType TXT) {
            $rs = Get-AzDnsRecordSet -Name $msolTenant.Name -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType TXT
            $rs.records[0].Value = $msolTenant.spfValue
            if ($remove) {
                Remove-AzDnsRecordSet -RecordSet $rs
            #Otherswise update Recordset
            } else {
            Set-AzDnsRecordSet -RecordSet $rs
            }
        }
        #If RecordSet does not exist and we didn't call -Remove switch
        elseif (-not $remove) {
            New-AzDnsRecordSet -Name $msolTenant.Name -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -Ttl 3600 -RecordType TXT -DnsRecords (New-AzDnsRecordConfig -Value $msolTenant.spfValue)
        }
}
    Function dmarcTxt{
        <#
            .SYNOPSIS
            
            .DESCRIPTION
            
            .PARAMETER
            
            .EXPAMPLE
        #>
        
        param (
            [parameter (Mandatory=$true,position=0)]
            $rgName,
            [parameter (Mandatory=$true,position=1)]
            $msolTenant,
            [parameter (Mandatory=$false,position=2)]
            [Switch]$Remove
        )
            #Check to see if recordset already exists
            #
            if (Get-AzDnsRecordSet -Name ("_dmarc",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType TXT) {
                $rs = Get-AzDnsRecordSet -Name ("_dmarc",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType TXT
                $rs.records[0].Value = $msolTenant.dmarcTxt
                if ($remove) {
                    Remove-AzDnsRecordSet -RecordSet $rs
                #Otherswise update Recordset
                } else {
                Set-AzDnsRecordSet -RecordSet $rs
                }
            }
            #If RecordSet does not exist and we didn't call -Remove switch
            elseif (-not $remove) {
                New-AzDnsRecordSet -Name ("_dmarc",$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -Ttl 3600 -RecordType TXT -DnsRecords (New-AzDnsRecordConfig -Value $msolTenant.dmarcTxt)
            }
    }
    Function sipSrv{
        <#
            .SYNOPSIS
            
            .DESCRIPTION
            
            .PARAMETER
            
            .EXPAMPLE
        #>
        
        param (
            [parameter (Mandatory=$true,position=0)]
            $rgName,
            [parameter (Mandatory=$true,position=1)]
            $msolTenant,
            [parameter (Mandatory=$false,position=2)]
            [Switch]$Remove
        )
            #Check to see if recordset already exists
            #
            if (Get-AzDnsRecordSet -Name $("_sip",$msolTenant.sipSrvProtocol,$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType SRV) {
                $rs = Get-AzDnsRecordSet -Name $("_sip",$msolTenant.sipSrvProtocol,$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType SRV
                $rs.records[0].Priority = $msolTenant.sipSrvPriority
                $rs.records[0].Port = $msolTenant.sipSrvPort
                $rs.records[0].Weight = $msolTenant.sipSrvWeight
                $rs.records[0].Target = $msolTenant.sipSrv
                if ($remove) {
                    Remove-AzDnsRecordSet -RecordSet $rs
                #Otherswise update Recordset
                } else {
                Set-AzDnsRecordSet -RecordSet $rs
                }
            }
            #If RecordSet does not exist and we didn't call -Remove switch
            elseif (-not $remove) {
                New-AzDnsRecordSet -Name $("_sip",$msolTenant.sipSrvProtocol,$msolTenant.Name -join ".")  -RecordType SRV -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -Ttl 3600 -DnsRecords (New-AzDnsRecordConfig -Priority $msolTenant.sipSrvPriority -Weight $msolTenant.sipSrvWeight -Port $msolTenant.sipSrvPort -Target $msolTenant.sipSrv)

            }
    }
    Function sipfederationtlsSrv{
        <#
            .SYNOPSIS
            
            .DESCRIPTION
            
            .PARAMETER
            
            .EXPAMPLE
        #>
        
        param (
            [parameter (Mandatory=$true,position=0)]
            $rgName,
            [parameter (Mandatory=$true,position=1)]
            $msolTenant,
            [parameter (Mandatory=$false,position=2)]
            [Switch]$Remove
        )
            #Check to see if recordset already exists
            #
            if (Get-AzDnsRecordSet -Name $("_sipfederationtls",$msolTenant.sipfederationtlsSrvProtocol,$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType SRV) {
                $rs = Get-AzDnsRecordSet -Name $("_sipfederationtls",$msolTenant.sipfederationtlsSrvProtocol,$msolTenant.Name -join ".") -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -RecordType SRV
                $rs.records[0].Priority = $msolTenant.sipfederationtlsSrvPriority
                $rs.records[0].Port = $msolTenant.sipfederationtlsSrvPort
                $rs.records[0].Weight = $msolTenant.sipfederationtlsSrvWeight
                $rs.records[0].Target = $msolTenant.sipfederationtlsSrv
                if ($remove) {
                    Remove-AzDnsRecordSet -RecordSet $rs
                #Otherswise update Recordset
                } else {
                Set-AzDnsRecordSet -RecordSet $rs
                }
            }
            #If RecordSet does not exist and we didn't call -Remove switch
            elseif (-not $remove) {
                New-AzDnsRecordSet -Name $("_sipfederationtls",$msolTenant.sipfederationtlsSrvProtocol,$msolTenant.Name -join ".")  -RecordType SRV -ZoneName $msolTenant.labSuffix -ResourceGroupName $rgName -Ttl 3600 -DnsRecords (New-AzDnsRecordConfig -Priority $msolTenant.sipfederationtlsSrvPriority -Weight $msolTenant.sipfederationtlsSrvWeight -Port $msolTenant.sipfederationtlsSrvPort -Target $msolTenant.sipfederationtlsSrv)

            }
    }
#Main Script
    $tenant = [demoTenant]::new($tenantAdminUser,$tenantAdminPassword,$tenantVanityDomain)
    $tenant.addCustomDomain()
    addVerifyTxt -rgName $myDnsResourceGroupName -msolTenant $tenant
    Start-sleep -Seconds 5
    $result = Confirm-MsolDomain -DomainName $tenant.labName()
    #TEST $result.Availability -eq "AvailableImmediately"
    
    Set-MsolDomain -Name $tenant.labName() -IsDefault

    mxRecord -rgName $myDnsResourceGroupName -msolTenant $tenant
    autodiscoverCname -rgName $myDnsResourceGroupName -msolTenant $tenant
    sipCname -rgName $myDnsResourceGroupName -msolTenant $tenant
    lyncDiscoverCname -rgName $myDnsResourceGroupName -msolTenant $tenant
    spfTxt -rgName $myDnsResourceGroupName -msolTenant $tenant
    enterpriseregistrationCname -rgName $myDnsResourceGroupName -msolTenant $tenant
    enterpriseenrollmentCname -rgName $myDnsResourceGroupName -msolTenant $tenant
    sipSrv -rgName $myDnsResourceGroupName -msolTenant $tenant
    sipfederationtlsSrv -rgName $myDnsResourceGroupName -msolTenant $tenant
    selector1Cname -rgName $myDnsResourceGroupName -msolTenant $tenant
    selector2Cname -rgName $myDnsResourceGroupName -msolTenant $tenant
    dmarcTxt -rgName $myDnsResourceGroupName -msolTenant $tenant