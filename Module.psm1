<#
    Avanade.AzureAD.Graph
    Simple REST Wrappers for the Azure AD Graph
#>

<#
    .SYNOPSIS
        Wrapper method for paging OData REST calls
#>
Function GetAzureGraphODataResult
{
    [CmdletBinding(ConfirmImpact='None')]
    param
    (
        [Parameter(Mandatory=$true)]
        [System.Uri]
        $Uri,
        [Parameter(Mandatory=$true)]
        [hashtable]
        $Headers,
        [Parameter(Mandatory=$false)]
        [System.String]
        $ContentType='application/json',
        [Parameter(Mandatory=$false)]
        [System.Int32]
        $LimitResultPages,
        [Parameter(Mandatory=$false)]
        [System.String]
        $ValueProperty='value',
        [Parameter(Mandatory=$false)]
        [System.String]
        $NextLinkProperty='@odata.nextLink',
        [Parameter(Mandatory=$false)]
        [System.String]
        $ErrorProperty='error'
    )

    $ResultPages=0
    $TotalItems=0
    do
    {
        $ResultPages++
        try
        {
            $ArmResult=Invoke-RestMethod -Uri $Uri -Headers $Headers -ContentType $ContentType
            if ($ArmResult -ne $null)
            {
                if($ArmResult.PSobject.Properties.name -match $ErrorProperty)
                {
                    throw ($ArmResult|Select-Object -ExpandProperty $ErrorProperty)|ConvertTo-Json
                }
                elseif($ArmResult.PSobject.Properties.name -match $ValueProperty)
                {
                    $RequestValue=$ArmResult|Select-Object -ExpandProperty $ValueProperty
                }
                else
                {
                    $RequestValue=$null
                }
                $TotalItems+=$RequestValue.Count
                if ($LimitResultPages -gt 0)
                {
                    if ($ResultPages -lt $LimitResultPages)
                    {
                        if($ArmResult.PSobject.Properties.name -match $NextLinkProperty)
                        {
                            $Uri=$ArmResult|Select-Object -ExpandProperty $NextLinkProperty
                            Write-Verbose "[GetAzureGraphODataResult] Total Items:$TotalItems. More items available @ $Uri"
                        }
                        else
                        {
                            $Uri=$null
                        }
                    }
                    else
                    {
                        $Uri=$null
                        Write-Verbose "[GetAzureGraphODataResult] Stopped iterating at $ResultPages pages. Iterated Items:$TotalItems More data available?:$([string]::IsNullOrEmpty($ArmResult.value))"
                    }
                }
                else
                {
                    if($ArmResult.PSobject.Properties.name -match $NextLinkProperty)
                    {
                        $Uri=$ArmResult|Select-Object -ExpandProperty $NextLinkProperty
                        Write-Verbose "[GetAzureGraphODataResult] Total Items:$TotalItems. More items available @ $Uri"
                    }
                    else
                    {
                        $Uri=$null
                    }
                }
                Write-Output $RequestValue
            }
            else
            {
                $Uri=$null
            }
        }
        catch
        {
            Write-Warning "[GetAzureGraphODataResult]Error $Uri $_"
            $Uri=$null
        }
    } while ($Uri -ne $null)
}

<#
    .SYNOPSIS
        Retrieves the graph report metadata for the desired tenant(s)
    .PARAMETER TenantName
        The tenant name(s)
    .PARAMETER AccessToken
        The OAuth Bearer token
    .PARAMETER GraphApiEndpoint
        The Azure Graph API Uri
    .PARAMETER GraphApiVersion
        The Azure Graph API Version        
#>
Function Get-AzureADADGraphReportMetadata
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [String]
        $AccessToken,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [String[]]
        $TenantName, 
        [Parameter(Mandatory=$false)]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false)]
        [String]
        $GraphApiVersion='beta'               
    )
    BEGIN
    {
        $Headers=@{Authorization="Bearer $AccessToken";Accept='application/json'}
        $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
        $GraphUriBld.Query="api-version=$GraphApiVersion"
    }
    PROCESS
    {
        foreach ($item in $TenantName)
        {
            try
            {
                $GraphUriBld.Path="$item/reports/`$metadata"
                $GraphResult=Invoke-RestMethod -Uri $GraphUriBld.Uri -Headers $Headers -ContentType 'application/json'
                Write-Output $GraphResult
            }
            catch
            {
                Write-Warning "[Get-AzureADGraphReportMetadata] $item api-version=$GraphApiVersion $_"
            }
        } 
    }
    END
    {

    }
}

<#
    .SYNOPSIS
        Retrieves a list of audit events
    .PARAMETER TenantName
        The tenant name(s)
    .PARAMETER AccessToken
        The OAuth Bearer token
    .PARAMETER LimitResultPages
        Limit the number of paged results
    .PARAMETER Top
        Limits the result set
    .PARAMETER Filter
        OData filter clause
    .PARAMETER GraphApiEndpoint
        The Azure Graph API Uri
    .PARAMETER GraphApiVersion
        The Azure Graph API Version
#>
Function Get-AzureADGraphAuditEvent
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [String]
        $AccessToken,
        [ValidateRange(0,1000)]
        [Parameter(Mandatory=$false)]
        [int]
        $LimitResultPages,
        [Parameter(Mandatory=$false)]
        [ValidateRange(0,1000)]
        [int]
        $Top,        
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [String[]]
        $TenantName,
        [Parameter(Mandatory=$false)]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false)]
        [String]
        $GraphApiVersion='beta',
        [Parameter(Mandatory=$false)]
        [String]
        $Filter        
    )
    BEGIN
    {
        $Headers=@{Authorization="Bearer $AccessToken";Accept="application/json"}
        $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
        $GraphQuery="api-version=$GraphApiVersion"
    }
    PROCESS
    {
        foreach ($Tenant in $TenantName)
        {
            try
            {
                $GraphUriBld.Path="$Tenant/activities/audit"
                if ([String]::IsNullOrEmpty($Filter) -eq $false) {
                    $GraphQuery+="`$filter=$Filter"
                }    
                if ($Top -gt 0) {
                    $GraphQuery+="`$Top=$Top"
                }
                $GraphUriBld.Query=$GraphQuery
                $Result=GetAzureGraphODataResult -Uri $GraphUriBld.Uri -Headers $Headers `
                    -ContentType 'application/json' -LimitResultPages $LimitResultPages `
                    -ValueProperty 'value' -NextLinkProperty '@odata.nextLink'
                Write-Output $Result
            }
            catch
            {
                Write-Warning "[Get-AzureADGraphAuditEvent] $Tenant api-version=$GraphApiVersion $_"
            }              
        }  
    }
    END
    {

    }
}

<#
    .SYNOPSIS
        Retrieves the list of graph signin events
    .PARAMETER TenantName
        The tenant name(s)        
    .PARAMETER AccessToken
        The OAuth Bearer token
    .PARAMETER LimitResultPages
        Limit the number of paged results
    .PARAMETER Top
        Limits the result set
    .PARAMETER Filter
        OData filter clause
    .PARAMETER GraphApiEndpoint
        The Azure Graph API Uri
    .PARAMETER GraphApiVersion
        The Azure Graph API Version        
#>
Function Get-AzureADGraphSigninEvent
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [String[]]
        $TenantName,        
        [Parameter(Mandatory=$true)]
        [String]
        $AccessToken,
        [ValidateRange(0,1000)]
        [Parameter(Mandatory=$false)]
        [int]
        $LimitResultPages,
        [Parameter(Mandatory=$false)]
        [ValidateRange(0,1000)]
        [int]
        $Top,
        [Parameter(Mandatory=$false)]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false)]
        [String]
        $GraphApiVersion='beta',
        [Parameter(Mandatory=$false)]
        [String]
        $Filter        
    )
    BEGIN
    {
        $Headers=@{Authorization="Bearer $AccessToken";Accept="application/json"}
        $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
        $GraphQuery="api-version=$GraphApiVersion"
    }
    PROCESS
    {
        foreach ($Tenant in $TenantName)
        {
            try
            {
                $GraphUriBld.Path="$Tenant/activities/signinEvents"
                if ([String]::IsNullOrEmpty($Filter) -eq $false) {
                    $GraphQuery+="`$filter=$Filter"
                }    
                if ($Top -gt 0) {
                    $GraphQuery+="`$top=$Top"
                }
                $GraphUriBld.Query=$GraphQuery
                $Result=GetAzureGraphODataResult -Uri $GraphUriBld.Uri -Headers $Headers `
                    -ContentType 'application/json' -LimitResultPages $LimitResultPages `
                    -ValueProperty 'value' -NextLinkProperty '@odata.nextLink'
                Write-Output $Result                   
            }
            catch
            {
                Write-Warning "[Get-AzureADGraphSigninEvent] $Tenant api-version=$GraphApiVersion $_"    
            }
        }         
    }
    END
    {

    }
}

<#
    .SYNOPSIS
        Retrieves a report of the desired audit event elements
    .PARAMETER TenantName
        The tenant name(s)
    .PARAMETER Element
        The audit event element(s)
    .PARAMETER AccessToken
        The OAuth Bearer token
    .PARAMETER LimitResultPages
        Limit the number of paged results
    .PARAMETER Top
        Limits the result set
    .PARAMETER Filter
        OData filter clause
    .PARAMETER GraphApiEndpoint
        The Azure Graph API Uri
    .PARAMETER GraphApiVersion
        The Azure Graph API Version
#>
Function Get-AzureADGraphReport
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$false)]
        [ValidateSet(
            'b2cAuthenticationCountSummary','b2cMfaRequestCount','b2cMfaRequestEvent',
            'b2cAuthenticationEvent','b2cAuthenticationCount','b2cMfaRequestCountSummary',
            'tenantUserCount','applicationUsageDetailEvents','applicationUsageSummaryEvents',
            'b2cUserJourneySummaryEvents','b2cUserJourneyEvents','cloudAppDiscoveryEvents',
            'mimSsgmGroupActivityEvents','ssgmGroupActivityEvents','mimSsprActivityEvents',
            'ssprActivityEvents','mimSsprRegistrationActivityEvents','ssprRegistrationActivityEvents',
            'threatenedCredentials','weakCredentials','compromisedCredentials',
            'allUserSignInActivityEvents','auditEvents','accountProvisioningEvents',
            'signInsFromUnknownSourcesEvents','signInsFromIPAddressesWithSuspiciousActivityEvents',
            'signInsFromMultipleGeographiesEvents','signInsFromPossiblyInfectedDevicesEvents',
            'irregularSignInActivityEvents','allUsersWithAnomalousSignInActivityEvents',
            'signInsAfterMultipleFailuresEvents','applicationUsageSummary',
            'userActivitySummary','groupActivitySummary'
        )]
        [String[]]
        $Element='auditEvents',
        [Parameter(Mandatory=$true)]
        [String]
        $AccessToken,
        [ValidateRange(0,1000)]
        [Parameter(Mandatory=$false)]
        [int]
        $LimitResultPages,
        [Parameter(Mandatory=$false)]
        [ValidateRange(0,1000)]
        [int]
        $Top,        
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [String[]]
        $TenantName,
        [Parameter(Mandatory=$false)]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false)]
        [String]
        $GraphApiVersion='beta',
        [Parameter(Mandatory=$false)]
        [String]
        $Filter 
    )
    BEGIN
    {
        $Headers=@{Authorization="Bearer $AccessToken";Accept="application/json"}
        $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
        $GraphQuery="api-version=$GraphApiVersion"
    }
    PROCESS
    {
        foreach ($Tenant in $TenantName)
        {
            foreach ($item in $Element)
            {
                try
                {
                    $GraphUriBld.Path="$Tenant/reports/$Element"
                    if ([String]::IsNullOrEmpty($Filter) -eq $false) {
                        $GraphQuery+="&`$filter=$Filter"
                    }    
                    if ($Top -gt 0) {
                        $GraphQuery+="&`$top=$Top"
                    }
                    $GraphUriBld.Query=$GraphQuery
                    $Result=GetAzureGraphODataResult -Uri $GraphUriBld.Uri -Headers $Headers `
                            -ContentType 'application/json' -LimitResultPages $LimitResultPages `
                            -ValueProperty 'value' -NextLinkProperty '@odata.nextLink'
                    Write-Output $Result                    
                }
                catch
                {
                    Write-Warning "[Get-AzureADGraphReport] $Tenant $item api-version=$GraphApiVersion $_"
                }
            }
        }
    }
    END
    {

    }
}

<#
    .SYNOPSIS
        Retrieves a list of the OAuth permission grants for the
        current tenant scope
    .PARAMETER AccessToken
        The OAuth Bearer token
    .PARAMETER LimitResultPages
        Limit the number of paged results
    .PARAMETER Top
        Limits the result set
    .PARAMETER Filter
        OData filter clause
    .PARAMETER GraphApiEndpoint
        The Azure Graph API Uri
    .PARAMETER GraphApiVersion
        The Azure Graph API Version
#>
Function Get-AzureADGraphOauthPermissionGrant
{
    [CmdletBinding(ConfirmImpact='None')]
    param
    (
        [Parameter(Mandatory=$true)]
        [String]
        $AccessToken,
        [ValidateRange(0,1000)]
        [Parameter(Mandatory=$false)]
        [int]
        $LimitResultPages,
        [Parameter(Mandatory=$false)]
        [ValidateRange(0,1000)]
        [int]
        $Top,       
        [Parameter(Mandatory=$false)]
        [String]
        $Filter,        
        [Parameter(Mandatory=$false)]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false)]
        [String]
        $GraphApiVersion='beta',
        [Parameter(Mandatory=$false)]
        [String]
        $TenantName='myOrganization' 
    )
    $Headers=@{Authorization="Bearer $AccessToken";Accept="application/json"}
    $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
    $GraphUriBld.Path="$TenantName/oauth2PermissionGrants"
    $GraphQuery="api-version=$GraphApiVersion"
    if ([String]::IsNullOrEmpty($Filter) -eq $false) {
        $GraphQuery+="&`$filter=$Filter"
    }    
    if ($Top -gt 0) {
        $GraphQuery+="&`$top=$Top"
    }
    $GraphUriBld.Query=$GraphQuery
    #odata call
    $Result=GetAzureGraphODataResult -Uri $GraphUriBld.Uri -Headers $Headers `
        -ContentType 'application/json' -LimitResultPages $LimitResultPages `
        -ValueProperty 'value' -NextLinkProperty '@odata.nextLink'
    Write-Output $Result
}

<#
    .SYNOPSIS
        Retrieves a list of the domains
    .PARAMETER DomainName
        The domain name(s)
    .PARAMETER AccessToken
        The OAuth Bearer token
    .PARAMETER GraphApiEndpoint
        The Azure Graph API Uri
    .PARAMETER GraphApiVersion
        The Azure Graph API Version
#>
Function Get-AzureADGraphDomain
{
    [CmdletBinding(ConfirmImpact='None')]
    param
    (
        [Parameter(Mandatory=$false)]
        [String[]]
        $DomainName,
        [Parameter(Mandatory=$true)]
        [String]
        $AccessToken,      
        [Parameter(Mandatory=$false)]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false)]
        [String]
        $GraphApiVersion='beta',
        [Parameter(Mandatory=$false)]
        [String]
        $TenantName='myOrganization'
    )
    BEGIN
    {
        $Headers=@{Authorization="Bearer $AccessToken";Accept="application/json"}
        $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
        $GraphUriBld.Path="$TenantName/domains"
        $GraphUriBld.Query="api-version=$GraphApiVersion"
    }
    PROCESS
    {
        if ($DomainName -ne $null) 
        {
            foreach ($Domain in $DomainName)
            {
                try
                {
                    $GraphUriBld.Path="myOrganization/domains('$Domain')"
                    $Result=Invoke-RestMethod -Uri $GraphUriBld.Uri -Headers $Headers -ContentType 'application/json'
                    Write-Output $Result
                }
                catch {
                    Write-Warning "[Get-AzureADGraphDomain] $Domain api-version=$GraphApiVersion $_"
                }
            }    
        }
        else
        {
            $Result=GetAzureGraphODataResult -Uri $GraphUriBld.Uri -Headers $Headers `
                -ContentType 'application/json' -LimitResultPages $LimitResultPages `
                -ValueProperty 'value' -NextLinkProperty '@odata.nextLink'
            Write-Output $Result
        }
    }
    END
    {

    }
}

<#
    .SYNOPSIS
        Retrieves a list of the policies
    .PARAMETER PolicyId
        The policy id(s)
    .PARAMETER AccessToken
        The OAuth Bearer token
    .PARAMETER GraphApiEndpoint
        The Azure Graph API Uri
    .PARAMETER GraphApiVersion
        The Azure Graph API Version
#>
Function Get-AzureADGraphPolicy
{
    [CmdletBinding(ConfirmImpact='None')]
    param
    (
        [Parameter(Mandatory=$false)]
        [String[]]
        $PolicyId,
        [Parameter(Mandatory=$true)]
        [String]
        $AccessToken,      
        [Parameter(Mandatory=$false)]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false)]
        [String]
        $GraphApiVersion='1.6',
        [Parameter(Mandatory=$false)]
        [String]
        $TenantName='myOrganization' 
    )
    BEGIN
    {
        $Headers=@{Authorization="Bearer $AccessToken";Accept="application/json"}
        $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
        $GraphUriBld.Path="$TenantName/policies"
        $GraphUriBld.Query="api-version=$GraphApiVersion"
    }
    PROCESS
    {
        if ($PolicyId -ne $null) 
        {
            foreach ($Policy in $PolicyId)
            {
                try
                {
                    $GraphUriBld.Path="myorganization/policies/$Policy"
                    $Result=Invoke-RestMethod -Uri $GraphUriBld.Uri -Headers $Headers -ContentType 'application/json'
                    Write-Output $Result
                }
                catch {
                    Write-Warning "[Get-AzureADGraphPolicy] $Domain api-version=$GraphApiVersion $_"
                }
            }    
        }
        else
        {
            $Result=GetAzureGraphODataResult -Uri $GraphUriBld.Uri -Headers $Headers `
                -ContentType 'application/json' -LimitResultPages $LimitResultPages `
                -ValueProperty 'value' -NextLinkProperty '@odata.nextLink'
            Write-Output $Result
        }
    }
    END
    {

    }
}

<#
    .SYNOPSIS
        Retrieves a list of the roles
    .PARAMETER RoleId
        The policy id(s)
    .PARAMETER AccessToken
        The OAuth Bearer token
    .PARAMETER GraphApiEndpoint
        The Azure Graph API Uri
    .PARAMETER GraphApiVersion
        The Azure Graph API Version
#>
Function Get-AzureADGraphRole
{
    [CmdletBinding(ConfirmImpact='None')]
    param
    (
        [Parameter(Mandatory=$false)]
        [String[]]
        $RoleId,
        [Parameter(Mandatory=$true)]
        [String]
        $AccessToken,      
        [Parameter(Mandatory=$false)]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false)]
        [String]
        $GraphApiVersion='1.6',
        [Parameter(Mandatory=$false)]
        [String]
        $TenantName='myOrganization'
    )
    BEGIN
    {
        $Headers=@{Authorization="Bearer $AccessToken";Accept="application/json"}
        $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
        $GraphUriBld.Path="$TenantName/directoryRoles"
        $GraphUriBld.Query="api-version=$GraphApiVersion"
    }
    PROCESS
    {
        if ($RoleId -ne $null) 
        {
            foreach ($Role in $RoleId)
            {
                try
                {
                    $GraphUriBld.Path="myorganization/directoryRoles/$Role"
                    $Result=Invoke-RestMethod -Uri $GraphUriBld.Uri -Headers $Headers -ContentType 'application/json'
                    Write-Output $Result
                }
                catch {
                    Write-Warning "[Get-AzureADGraphRole] $Domain api-version=$GraphApiVersion $_"
                }
            }    
        }
        else
        {
            $Result=GetAzureGraphODataResult -Uri $GraphUriBld.Uri -Headers $Headers `
                -ContentType 'application/json' -LimitResultPages $LimitResultPages `
                -ValueProperty 'value' -NextLinkProperty '@odata.nextLink'
            Write-Output $Result
        }
    }
    END
    {

    }
}

<#
    .SYNOPSIS
        Retrieves a list of the roles
    .PARAMETER TemplateId
        The policy id(s)
    .PARAMETER AccessToken
        The OAuth Bearer token
    .PARAMETER GraphApiEndpoint
        The Azure Graph API Uri
    .PARAMETER GraphApiVersion
        The Azure Graph API Version
#>
Function Get-AzureADGraphRoleTemplate
{
    [CmdletBinding(ConfirmImpact='None')]
    param
    (
        [Parameter(Mandatory=$false)]
        [String[]]
        $TemplateId,
        [Parameter(Mandatory=$true)]
        [String]
        $AccessToken,      
        [Parameter(Mandatory=$false)]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false)]
        [String]
        $GraphApiVersion='1.6',
        [Parameter(Mandatory=$false)]
        [String]
        $TenantName='myOrganization'
    )
    BEGIN
    {
        $Headers=@{Authorization="Bearer $AccessToken";Accept="application/json"}
        $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
        $GraphUriBld.Path="$TenantName/directoryRoleTemplates"
        $GraphUriBld.Query="api-version=$GraphApiVersion"
    }
    PROCESS
    {
        if ($TemplateId -ne $null) 
        {
            foreach ($Template in $TemplateId)
            {
                try
                {
                    $GraphUriBld.Path="myorganization/directoryRoleTemplates/$Template"
                    $Result=Invoke-RestMethod -Uri $GraphUriBld.Uri -Headers $Headers -ContentType 'application/json'
                    Write-Output $Result
                }
                catch {
                    Write-Warning "[Get-AzureADGraphRoleTemplate] $Domain api-version=$GraphApiVersion $_"
                }
            }    
        }
        else
        {
            $Result=GetAzureGraphODataResult -Uri $GraphUriBld.Uri -Headers $Headers `
                -ContentType 'application/json' -LimitResultPages $LimitResultPages `
                -ValueProperty 'value' -NextLinkProperty '@odata.nextLink'
            Write-Output $Result
        }
    }
    END
    {

    }
}