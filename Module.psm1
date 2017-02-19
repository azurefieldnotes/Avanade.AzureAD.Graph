<#
    Avanade.AzureAD.Graph
    Simple REST Wrappers for the Azure AD Graph
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
                            Write-Verbose "[GetArmODataResult] Total Items:$TotalItems. More items available @ $Uri"
                        }
                        else
                        {
                            $Uri=$null
                        }
                    }
                    else
                    {
                        $Uri=$null
                        Write-Verbose "[GetArmODataResult] Stopped iterating at $ResultPages pages. Iterated Items:$TotalItems More data available?:$([string]::IsNullOrEmpty($ArmResult.value))"
                    }
                }
                else
                {
                    if($ArmResult.PSobject.Properties.name -match $NextLinkProperty)
                    {
                        $Uri=$ArmResult|Select-Object -ExpandProperty $NextLinkProperty
                        Write-Verbose "[GetArmODataResult] Total Items:$TotalItems. More items available @ $Uri"
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
            Write-Warning "[GetArmODataResult]Error $Uri $_"
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
Function Get-AzureGraphReportMetadata
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
        foreach ($item in $TenantName) {
            $GraphUriBld.Path="$item/reports/`$metadata"
            $GraphResult=Invoke-RestMethod -Uri $GraphUriBld.Uri -Headers $Headers -ContentType 'application/json'
            Write-Output $GraphResult
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
Function Get-AzureGraphAuditEvent
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
        $GraphUriBld.Path="$TenantName/activities/audit"
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
        
        # $Result=GetAzureGraphResult -AccessToken $AccessToken `
        #     -LimitResultPages $LimitResultPages -Top $Top `
        #     -UriPath "$TenantName/activities/audit" `
        #     -Filter $Filter
        #     -GraphApiEndpoint $GraphApiEndpoint `
        #     -GraphApiVersion $GraphApiVersion
        Write-Output $Result            
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
Function Get-AzureGraphSigninEvent
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
        $GraphUriBld.Path="$TenantName/activities/signinEvents"
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

        # $Result=GetAzureGraphResult -AccessToken $AccessToken `
        #     -LimitResultPages $LimitResultPages -Top $Top `
        #     -UriPath "$TenantName/activities/signinEvents" `
        #     -Filter $Filter
        #     -GraphApiEndpoint $GraphApiEndpoint `
        #     -GraphApiVersion $GraphApiVersion
        Write-Output $Result            
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
Function Get-AzureGraphReport
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
Function Get-AzureGraphOauthPermissionGrant
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
        [Parameter(Mandatory=$false)]
        [String]
        $Filter,        
        [Parameter(Mandatory=$false)]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false)]
        [String]
        $GraphApiVersion='beta'    
    )
    $Headers=@{Authorization="Bearer $AccessToken";Accept="application/json"}
    $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
    $GraphUriBld.Path="myOrganization/oauth2PermissionGrants"
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