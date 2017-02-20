<#
    Avanade.AzureAD.Graph
    Simple REST Wrappers for the Azure AD Graph
#>

#region Helper Methods

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
        $GraphApiRoot,
        [Parameter(Mandatory=$true)]
        [System.Uri]
        $Path,        
        [Parameter(Mandatory=$false)]
        [String]
        $Filter,
        [Parameter(Mandatory=$true)]
        [String]
        $GraphApiVersion,
        [Parameter(Mandatory=$true)]
        [hashtable]
        $Headers,
        [Parameter(Mandatory=$false)]
        [System.String]
        $ContentType='application/json',
        [Parameter(Mandatory=$false)]
        [System.Int32]
        $Top,        
        [Parameter(Mandatory=$false)]
        [System.Int32]
        $LimitResultPages,
        [Parameter(Mandatory=$false)]
        [System.String]
        $ValueProperty='value',
        [Parameter(Mandatory=$false)]
        [System.String]
        $NextLinkProperty='odata.nextLink',
        [Parameter(Mandatory=$false)]
        [System.String]
        $ErrorProperty='error'
    )

    $ResultPages=0
    $TotalItems=0
    $UriBld=New-Object System.UriBuilder($GraphApiRoot)
    $UriQuery="api-version=$GraphApiVersion"
    if ([String]::IsNullOrEmpty($Filter) -eq $false) {
        $UriQuery+="&`$filter=$Filter"
        if ($Filter -like '$top=') {
            $Top=0
        }
    }
    if ($Top -gt 0) {
        $UriQuery+="&`$top=$Top"
    }
    $UriBld.Path=$Path
    $UriBld.Query=$UriQuery
    do
    {
        $ResultPages++
        try
        {
            $GraphResult=Invoke-RestMethod -Uri $UriBld.Uri -Headers $Headers -ContentType $ContentType  -ErrorAction Stop
            if($GraphResult.PSobject.Properties.name -match $ValueProperty)
            {
                $RequestValue=@($GraphResult|Select-Object -ExpandProperty $ValueProperty)
                $TotalItems+=$RequestValue.Count
                if($GraphResult.PSobject.Properties.name -match $NextLinkProperty)
                {
                    if ($Top -gt 0 -or $LimitResultPages -gt 0) {
                        if ($TotalItems -ge $Top -or $ResultPages -ge $LimitResultPages) {
                            Write-Verbose "[GetAzureGraphODataResult] Stopped Iterating Page:$ResultPages Total Items:$TotalItems"
                            $UriBld=$null
                        }
                    }
                    else
                    {
                        $NextLinkValue=$GraphResult|Select-Object -ExpandProperty $NextLinkProperty
                        if ($NextLinkValue -like 'http*') {
                            $UriBld=New-Object System.UriBuilder([Uri]::UnescapeDataString($NextLinkValue))
                        }
                        else {
                            $UpdatedQuery=[Uri]::UnescapeDataString((($NextLinkValue.Split('?')|Select-Object -Last 1).Split('&')|Select-Object -Last 1))
                            $UriBld.Query="$UriQuery&$UpdatedQuery"
                        }
                        Write-Verbose "[GetAzureGraphODataResult] Page:$ResultPages Page Size:$($RequestValue.Count) More Results Available @ $($UriBld.Uri)"                     
                    }
                }
                else {
                    $UriBld=$null
                }
                Write-Verbose "[GetAzureGraphODataResult] Page:$ResultPages Total Items:$TotalItems"
                Write-Output $RequestValue
            }
            else
            {
                $UriBld=$null
            }
        }
        catch
        {
            Write-Warning "[GetAzureGraphODataResult] $($UriBld.Uri) $_"
            $UriBld=$null
        }
    } until ($UriBld -eq $null)
}

#endregion

#region Graph Functions

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
Function Get-AzureADGraphReportMetadata
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
                $Result=GetAzureGraphODataResult -Path $GraphUriBld.Path -Headers $Headers `
                    -ContentType 'application/json' `
                    -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' -Filter $Filter `
                    -Top $Top -LimitResultPages $LimitResultPages -GraphApiRoot $GraphApiEndpoint `
                    -GraphApiVersion $GraphApiVersion
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
                $Result=GetAzureGraphODataResult -Path $GraphUriBld.Path -Headers $Headers `
                    -ContentType 'application/json' `
                    -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' -Filter $Filter `
                    -Top $Top -LimitResultPages $LimitResultPages -GraphApiRoot $GraphApiEndpoint `
                    -GraphApiVersion $GraphApiVersion
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
                    $GraphUriBld.Path="$Tenant/reports/$item"
                    if ([String]::IsNullOrEmpty($Filter) -eq $false) {
                        $GraphQuery+="&`$filter=$Filter"
                    }    
                    if ($Top -gt 0) {
                        $GraphQuery+="&`$top=$Top"
                    }
                    $GraphUriBld.Query=$GraphQuery
                    $Result=GetAzureGraphODataResult -Path $GraphUriBld.Path -Headers $Headers `
                        -ContentType 'application/json' `
                        -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' -Filter $Filter `
                        -Top $Top -LimitResultPages $LimitResultPages -GraphApiRoot $GraphApiEndpoint `
                        -GraphApiVersion $GraphApiVersion
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
        $GraphApiVersion='1.6',
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
    $Result=GetAzureGraphODataResult -Path $GraphUriBld.Path -Headers $Headers `
        -ContentType 'application/json' `
        -ValueProperty 'value' -NextLinkProperty 'odata.nextLink' -Filter $Filter `
        -Top $Top -LimitResultPages $LimitResultPages -GraphApiRoot $GraphApiEndpoint `
        -GraphApiVersion $GraphApiVersion
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
        $GraphApiVersion='1.6',
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
                -ValueProperty 'value' -NextLinkProperty 'odata.nextLink'
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
        $TenantName='myOrganization',
        [ValidateRange(0,1000)]
        [Parameter(Mandatory=$false)]
        [int]
        $LimitResultPages,
        [Parameter(Mandatory=$false)]
        [ValidateRange(0,1000)]
        [int]
        $Top        
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
                    $GraphUriBld.Path="$TenantName/policies/$Policy"
                    $Result=Invoke-RestMethod -Uri $GraphUriBld.Uri -Headers $Headers -ContentType 'application/json'
                    Write-Output $Result
                }
                catch {
                    Write-Warning "[Get-AzureADGraphPolicy] $Policy api-version=$GraphApiVersion $_"
                }
            }    
        }
        else
        {
            $Result=GetAzureGraphODataResult -Path $GraphUriBld.Path -Headers $Headers `
                -ContentType 'application/json' `
                -ValueProperty 'value' -NextLinkProperty 'odata.nextLink' -Filter $Filter `
                -Top $Top -LimitResultPages $LimitResultPages -GraphApiRoot $GraphApiEndpoint `
                -GraphApiVersion $GraphApiVersion
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
    [CmdletBinding(ConfirmImpact='None',DefaultParameterSetName='noquery')]
    param
    (
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [String[]]
        $RoleId,
        [Parameter(Mandatory=$true,ParameterSetName='noquery')]
        [String]
        $AccessToken,
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [String]
        $GraphApiVersion='beta',
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [String]
        $TenantName='myOrganization'
    )
    BEGIN
    {
        $Headers=@{Authorization="Bearer $AccessToken";Accept="application/json"}
        $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
        $GraphUriBld.Path="$TenantName/directoryRoles"
        $GraphQuery="api-version=$GraphApiVersion"
        if ($PSCmdlet.ParameterSetName -eq 'query') {
            if ($Filter.Contains('top=') -eq $false -and $Top -gt 0) {
                $GraphQuery+="&`$top=$Top"
            }
        }
        else {
            if ($Top -gt 0) {
                $GraphQuery+="&`$top=$Top"
            }
        }
    }
    PROCESS
    {
        if ($PSCmdlet.ParameterSetName -eq 'noquery' -and $RoleId -ne $null) 
        {
            $GraphUriBld.Query=$GraphQuery
            foreach ($Role in $RoleId)
            {
                try
                {
                    $GraphUriBld.Path="$TenantName/directoryRoles/$Role"
                    $Result=Invoke-RestMethod -Uri $GraphUriBld.Uri -Headers $Headers -ContentType 'application/json'
                    Write-Output $Result
                }
                catch {
                    Write-Warning "[Get-AzureADGraphRole] $Role api-version=$GraphApiVersion $_"
                }
            }    
        }
        else
        {
            $Result=GetAzureGraphODataResult -Path $GraphUriBld.Path -Headers $Headers `
                -ContentType 'application/json' `
                -ValueProperty 'value' -NextLinkProperty 'odata.nextLink' `
                -GraphApiRoot $GraphApiEndpoint `
                -GraphApiVersion $GraphApiVersion
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
    [CmdletBinding(ConfirmImpact='None',DefaultParameterSetName='noquery')]
    param
    (
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [String[]]
        $TemplateId,
        [Parameter(Mandatory=$true,ParameterSetName='query')]
        [String]
        $Filter,
        [Parameter(Mandatory=$true,ParameterSetName='query')]
        [Parameter(Mandatory=$true,ParameterSetName='noquery')]
        [String]
        $AccessToken,
        [Parameter(Mandatory=$false,ParameterSetName='query')]
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ParameterSetName='query')]
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [String]
        $GraphApiVersion='1.6',
        [Parameter(Mandatory=$false,ParameterSetName='query')]
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [String]
        $TenantName='myOrganization',
        [ValidateRange(0,1000)]
        [Parameter(Mandatory=$false,ParameterSetName='query')]
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [int]
        $LimitResultPages,
        [Parameter(Mandatory=$false,ParameterSetName='query')]
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [ValidateRange(0,1000)]
        [int]
        $Top
    )
    BEGIN
    {
        $Headers=@{Authorization="Bearer $AccessToken";Accept="application/json"}
        $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
        $GraphUriBld.Path="$TenantName/directoryRoleTemplates"
        $GraphQuery="api-version=$GraphApiVersion"
        if ($PSCmdlet.ParameterSetName -eq 'query') {
            if ($Filter.Contains('top=') -eq $false -and $Top -gt 0) {
                $GraphQuery+="&`$top=$Top"
            }
        }
        else {
            if ($Top -gt 0) {
                $GraphQuery+="&`$top=$Top"
            }
        }
    }
    PROCESS
    {
        if ($PSCmdlet.ParameterSetName -eq 'noquery' -and $TemplateId -ne $null) 
        {
            $GraphUriBld.Query=$GraphQuery
            foreach ($Template in $TemplateId)
            {
                try
                {
                    $GraphUriBld.Path="$TenantName/directoryRoleTemplates/$Template"
                    $Result=Invoke-RestMethod -Uri $GraphUriBld.Uri -Headers $Headers -ContentType 'application/json'
                    Write-Output $Result
                }
                catch {
                    Write-Warning "[Get-AzureADGraphRoleTemplate] $Template api-version=$GraphApiVersion $_"
                }
            }    
        }
        else
        {
            $Result=GetAzureGraphODataResult -Path $GraphUriBld.Path -Headers $Headers `
                -ContentType 'application/json' `
                -ValueProperty 'value' -NextLinkProperty 'odata.nextLink' -Filter $Filter `
                -Top $Top -LimitResultPages $LimitResultPages -GraphApiRoot $GraphApiEndpoint `
                -GraphApiVersion $GraphApiVersion
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
    .PARAMETER UserId
        The user id(s)
    .PARAMETER AccessToken
        The OAuth Bearer token
    .PARAMETER GraphApiEndpoint
        The Azure Graph API Uri
    .PARAMETER GraphApiVersion
        The Azure Graph API Version
#>
Function Get-AzureADGraphUser
{
    [CmdletBinding(ConfirmImpact='None',DefaultParameterSetName='noquery')]
    param
    (
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [String[]]
        $UserId,
        [Parameter(Mandatory=$true,ParameterSetName='query')]
        [String]
        $Filter,
        [Parameter(Mandatory=$true,ParameterSetName='query')]
        [Parameter(Mandatory=$true,ParameterSetName='noquery')]
        [String]
        $AccessToken,
        [Parameter(Mandatory=$false,ParameterSetName='query')]
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ParameterSetName='query')]
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [String]
        $GraphApiVersion='beta',
        [Parameter(Mandatory=$false,ParameterSetName='query')]
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [String]
        $TenantName='myOrganization',
        [ValidateRange(0,1000)]
        [Parameter(Mandatory=$false,ParameterSetName='query')]
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [int]
        $LimitResultPages,
        [Parameter(Mandatory=$false,ParameterSetName='query')]
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [ValidateRange(0,1000)]
        [int]
        $Top
    )
    BEGIN
    {
        $Headers=@{Authorization="Bearer $AccessToken";Accept="application/json"}
        $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
        $GraphUriBld.Path="$TenantName/users"
        $GraphQuery="api-version=$GraphApiVersion"
        if ($PSCmdlet.ParameterSetName -eq 'query') {
            if ($Filter.Contains('top=') -eq $false -and $Top -gt 0) {
                $GraphQuery+="&`$top=$Top"
            }
        }
        else {
            if ($Top -gt 0) {
                $GraphQuery+="&`$top=$Top"
            }
        }
    }
    PROCESS
    {
        if ($PSCmdlet.ParameterSetName -eq 'noquery' -and $UserId -ne $null)
        {
            $GraphUriBld.Query=$GraphQuery
            foreach ($User in $UserId)
            {
                try
                {
                    $GraphUriBld.Path="$TenantName/users/$User"
                    $Result=Invoke-RestMethod -Uri $GraphUriBld.Uri -Headers $Headers -ContentType 'application/json'
                    Write-Output $Result
                }
                catch {
                    Write-Warning "[Get-AzureADGraphUser] $User api-version=$GraphApiVersion $_"
                }
            }            
        }
        else
        {
            $Result=GetAzureGraphODataResult -Path $GraphUriBld.Path -Headers $Headers `
                -ContentType 'application/json' `
                -ValueProperty 'value' -NextLinkProperty 'odata.nextLink' -Filter $Filter `
                -Top $Top -LimitResultPages $LimitResultPages -GraphApiRoot $GraphApiEndpoint `
                -GraphApiVersion $GraphApiVersion
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
    .PARAMETER UserId
        The user id(s)
    .PARAMETER AccessToken
        The OAuth Bearer token
    .PARAMETER GraphApiEndpoint
        The Azure Graph API Uri
    .PARAMETER GraphApiVersion
        The Azure Graph API Version
#>
Function Get-AzureADGraphGroup
{
    [CmdletBinding(ConfirmImpact='None',DefaultParameterSetName='noquery')]
    param
    (
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [String[]]
        $GroupId,
        [Parameter(Mandatory=$true,ParameterSetName='query')]
        [String]
        $Filter,
        [Parameter(Mandatory=$true,ParameterSetName='query')]
        [Parameter(Mandatory=$true,ParameterSetName='noquery')]
        [String]
        $AccessToken,
        [Parameter(Mandatory=$false,ParameterSetName='query')]
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ParameterSetName='query')]
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [String]
        $GraphApiVersion='beta',
        [Parameter(Mandatory=$false,ParameterSetName='query')]
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [String]
        $TenantName='myOrganization',
        [ValidateRange(0,1000)]
        [Parameter(Mandatory=$false,ParameterSetName='query')]
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [int]
        $LimitResultPages,
        [Parameter(Mandatory=$false,ParameterSetName='query')]
        [Parameter(Mandatory=$false,ParameterSetName='noquery')]
        [ValidateRange(0,1000)]
        [int]
        $Top
    )
    BEGIN
    {
        $Headers=@{Authorization="Bearer $AccessToken";Accept="application/json"}
        $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
        $GraphUriBld.Path="$TenantName/groups"
        $GraphQuery="api-version=$GraphApiVersion"
        if ($PSCmdlet.ParameterSetName -eq 'query') {
            if ($Filter.Contains('top=') -eq $false -and $Top -gt 0) {
                $GraphQuery+="&`$top=$Top"
            }
        }
        else {
            if ($Top -gt 0) {
                $GraphQuery+="&`$top=$Top"
            }
        }
    }
    PROCESS
    {
        if ($PSCmdlet.ParameterSetName -eq 'noquery' -and $GroupId -ne $null)
        {
            $GraphUriBld.Query=$GraphQuery
            foreach ($Group in $GroupId)
            {
                try
                {
                    $GraphUriBld.Path="$TenantName/groups/$Group"
                    $Result=Invoke-RestMethod -Uri $GraphUriBld.Uri -Headers $Headers -ContentType 'application/json'
                    Write-Output $Result
                }
                catch {
                    Write-Warning "[Get-AzureADGraphGroup] $User api-version=$GraphApiVersion $_"
                }
            }            
        }
        else
        {
            $Result=GetAzureGraphODataResult -Path $GraphUriBld.Path -Headers $Headers `
                -ContentType 'application/json' `
                -ValueProperty 'value' -NextLinkProperty 'odata.nextLink' -Filter $Filter `
                -Top $Top -LimitResultPages $LimitResultPages -GraphApiRoot $GraphApiEndpoint `
                -GraphApiVersion $GraphApiVersion
            Write-Output $Result
        }
    }
    END
    {

    }
}

#endregion