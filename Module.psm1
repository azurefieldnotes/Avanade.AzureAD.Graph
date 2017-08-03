<#
    Avanade.AzureAD.Graph
    Simple REST Wrappers for the Azure AD Graph
#>

#region Enums

$Global:Azure_ADGraph_Enums=@{
    'RoleAssignmentPrincipalTypes'=@{
        'ServicePrincipal'='ServicePrincipal';
        'User'='User';
        'Group'='Group';
    };
    'PermissionType'=@{
        'Application'='Application';
        'Delegated'='Delegated'
    };
}

$Global:Azure_ADGraph_WelknownIds=@{
    'Permissions' = [ordered]@{
        AccessDirectoryAsSignedInUser      = "a42657d6-7f20-40e3-b6f0-cee03008a62a"
        EnableSignOnAndReadUserProfiles    = "311a71cc-e848-46a1-bdf8-97ff7156d8e6"
        ReadAllGroups                      = "6234d376-f627-4f0f-90e0-dff25c5211a3"
        ReadAllUsersBasicProfile           = "cba73afc-7f69-4d86-8450-4978e04ecd1a"
        ReadAllUsersFullProfile            = "c582532d-9d9e-43bd-a97c-2667a28ce295"
        ReadDirectoryData                  = "5778995a-e1bf-45b8-affa-663a9f3f4d04"
        ManageAppsThatThisAppCreatesOrOwns = "824c81eb-e3f8-4ee6-8f6d-de7f50d565b7"
    };
    'PermissionScopes' = [ordered]@{
        AccessDirectoryAsSignedInUser      = "Directory.AccessAsUser.All"
        EnableSignOnAndReadUserProfiles    = "User.Read"
        ReadAllGroups                      = "Group.Read.All"
        ReadAllUsersBasicProfile           = "User.ReadBasic.All"
        ReadAllUsersFullProfile            = "User.Read.All"
        ReadDirectoryData                  = "Directory.Read.All"
        ManageAppsThatThisAppCreatesOrOwns = "Application.ReadWrite.OwnedBy"
    };
    'OAuthConsentTypes'=@{
        'AllPrincipals'='AllPrincipals'
        'Principal'='Principal'
    };
}

#endregion

#region Helper Methods

Function ConvertFromSecureString
{
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [SecureString[]]
        $InputObject
    )
    PROCESS
    {
        foreach ($item in $InputObject)
        {
            $ValuePtr=[System.IntPtr]::Zero
            try
            {
                $ValuePtr=[System.Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($item)
                Write-Output $([System.Runtime.InteropServices.Marshal]::PtrToStringUni($ValuePtr))
            }
            finally {
                [System.Runtime.InteropServices.Marshal]::ZeroFreeGlobalAllocUnicode($ValuePtr)
            }    
        }
    }
}


Function Invoke-AzureADGraphRequest
{
    [CmdletBinding(ConfirmImpact='None')]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [uri]
        $Uri,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [string]
        $ContentType='application/json',
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [System.Object]
        $Body,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [Microsoft.PowerShell.Commands.WebRequestMethod]
        $Method="GET",
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]
        $AccessToken,
        [ValidateNotNull()]
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [System.Collections.IDictionary]
        $AdditionalHeaders=@{Accept='application/json'},
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [string]
        $ValueProperty='value',
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [System.String]
        $NextLinkProperty='$odata.nextLink',
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [System.String]
        $ErrorProperty='odata.error',
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [Int32]
        $LimitResultPages,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [Int32]
        $RequestDelayMilliseconds=100
    )
    PROCESS
    {
        $ResultPages=0
        $TotalItems=0
        $RequestHeaders=$AdditionalHeaders
        $RequestHeaders['client-request-id']=[Guid]::NewGuid().ToString()
        $RequestHeaders['User-Agent']="PowerShell $($PSVersionTable.PSVersion.ToString())"
        $RequestHeaders['Authorization']="Bearer $AccessToken"
        $BaseUri="$($Uri.Scheme)://$($Uri.Host)"
        $RequestParams=@{
            Headers=$RequestHeaders;
            Uri=$Uri;
            ContentType=$ContentType;
            Method=$Method;
        }
        if ($Body -ne $null)
        {
            $RequestParams['Body']=$Body|ConvertTo-Json -Depth 10
        }
        $RequestResult=$null
        try
        {
            $Response=Invoke-WebRequest @RequestParams -UseBasicParsing -ErrorAction Stop
            Write-Verbose "[Invoke-AzureADGraphRequest]$Method $Uri Response:$($Response.StatusCode)-$($Response.StatusDescription) Content-Length:$($Response.RawContentLength)"
            $RequestResult=$Response.Content|ConvertFrom-Json
        }
        catch
        {
            #See if we can unwind an exception from a response
            if($_.Exception.Response -ne $null)
            {
                $ExceptionResponse=$_.Exception.Response
                $ErrorStream=$ExceptionResponse.GetResponseStream()
                $ErrorStream.Position=0
                $StreamReader=New-Object System.IO.StreamReader($ErrorStream)
                try
                {
                    $ErrorContent=$StreamReader.ReadToEnd()
                    $StreamReader.Close()
                    if(-not [String]::IsNullOrEmpty($ErrorContent))
                    {
                        $ErrorObject=$ErrorContent|ConvertFrom-Json
                        if (-not [String]::IsNullOrEmpty($ErrorProperty) -and  $ErrorObject.PSobject.Properties.name -match $ErrorProperty)
                        {
                            $ErrorContent=($ErrorObject|Select-Object -ExpandProperty $ErrorProperty)|ConvertTo-Json
                        }
                    }
                }
                catch
                {
                }
                finally
                {
                    $StreamReader.Close()
                }
                $ErrorMessage="Error: $($ExceptionResponse.Method) $($ExceptionResponse.ResponseUri) Returned $($ExceptionResponse.StatusCode) $ErrorContent"
            }
            else
            {
                $ErrorMessage="An error occurred $_"
            }
            Write-Verbose "[Invoke-AzureADGraphRequest] $ErrorMessage"
            throw $ErrorMessage        
        }
        if ($RequestResult -ne $null)
        {
            if ($RequestResult.PSobject.Properties.name -match $ValueProperty)
            {
                $Result=$RequestResult|Select-Object -ExpandProperty $ValueProperty
                $TotalItems+=$Result.Count
                Write-Output $Result
            }
            else
            {
                Write-Output $RequestResult
                $TotalItems++ #not sure why I am incrementing..
            }
            #Loop to aggregate OData continutation tokens
            while ($RequestResult.PSobject.Properties.name -match $NextLinkProperty)
            {
                #Throttle the requests a bit..
                Start-Sleep -Milliseconds $RequestDelayMilliseconds
                $ResultPages++
                $UriBld=New-Object System.UriBuilder($BaseUri)
                $NextUri=$RequestResult|Select-Object -ExpandProperty $NextLinkProperty
                if($LimitResultPages -gt 0 -and $ResultPages -eq $LimitResultPages -or [String]::IsNullOrEmpty($NextUri))
                {
                    break
                }
                Write-Verbose "[Invoke-AzureADGraphRequest] Item Count:$TotalItems Page:$ResultPages More Items available @ $NextUri"
                #Is this an absolute or relative uri?
                if($NextUri -match "$BaseUri*")
                {
                    $UriBld=New-Object System.UriBuilder($NextUri)
                }
                else
                {
                    $Path=$NextUri.Split('?')|Select-Object -First 1
                    $NextQuery=[Uri]::UnescapeDataString(($NextUri.Split('?')|Select-Object -Last 1))
                    $UriBld.Path=$Path
                    $UriBld.Query=$NextQuery
                }
                try
                {
                    $RequestParams['Uri']=$UriBld.Uri
                    $Response=Invoke-WebRequest @RequestParams -UseBasicParsing -ErrorAction Stop
                    Write-Verbose "[Invoke-AzureADGraphRequest]$Method $Uri Response:$($Response.StatusCode)-$($Response.StatusDescription) Content-Length:$($Response.RawContentLength)"
                    $RequestResult=$Response.Content|ConvertFrom-Json
                    if ($RequestResult.PSobject.Properties.name -match $ValueProperty)
                    {
                        $Result=$RequestResult|Select-Object -ExpandProperty $ValueProperty
                        $TotalItems+=$Result.Count
                        Write-Output $Result
                    }
                    else
                    {
                        Write-Output $RequestResult
                        $TotalItems++ #not sure why I am incrementing..
                    }
                }
                catch
                {
                    #See if we can unwind an exception from a response
                    if($_.Exception.Response -ne $null)
                    {
                        $ExceptionResponse=$_.Exception.Response
                        $ErrorStream=$ExceptionResponse.GetResponseStream()
                        $ErrorStream.Position=0
                        $StreamReader = New-Object System.IO.StreamReader($ErrorStream)
                        try
                        {
                            $ErrorContent=$StreamReader.ReadToEnd()
                            $StreamReader.Close()
                            if(-not [String]::IsNullOrEmpty($ErrorContent))
                            {
                                $ErrorObject=$ErrorContent|ConvertFrom-Json
                                if (-not [String]::IsNullOrEmpty($ErrorProperty) -and  $ErrorObject.PSobject.Properties.name -match $ErrorProperty)
                                {
                                    $ErrorContent=($ErrorObject|Select-Object -ExpandProperty $ErrorProperty)|ConvertTo-Json
                                }
                            }
                        }
                        catch
                        {
                        }
                        finally
                        {
                            $StreamReader.Close()
                        }
                        $ErrorMessage="Error: $($ExceptionResponse.Method) $($ExceptionResponse.ResponseUri) Returned $($ExceptionResponse.StatusCode) $ErrorContent"
                    }
                    else
                    {
                        $ErrorMessage="An error occurred $_"
                    }
                    Write-Verbose "[Invoke-AzureADGraphRequest] $ErrorMessage"
                    throw $ErrorMessage
                }
            }        
        }        
    }
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
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [String]
        $AccessToken,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [String[]]
        $TenantName, 
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
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
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [String]
        $AccessToken,
        [ValidateRange(0,1000)]
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [int]
        $LimitResultPages,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [ValidateRange(0,1000)]
        [int]
        $Top,        
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [String[]]
        $TenantName,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [String]
        $GraphApiVersion='beta',
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [String]
        $Filter        
    )
    BEGIN
    {
        $Headers=@{Accept="application/json"}
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
                    $GraphQuery+="&`$filter=$Filter"
                }    
                if ($Top -gt 0) {
                    $GraphQuery+="&`$Top=$Top"
                }
                $GraphUriBld.Query=$GraphQuery
                $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                    -ContentType 'application/json' `
                    -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' -ErrorProperty 'odata.error' `
                    -LimitResultPages $LimitResultPages -AccessToken $AccessToken
                Write-Output $Result
            }
            catch
            {
                Write-Warning "[Get-AzureADGraphAuditEvent] $Tenant api-version=$GraphApiVersion $_"
            }              
        }  
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
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [String[]]
        $TenantName,        
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [String]
        $AccessToken,
        [ValidateRange(0,1000)]
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [int]
        $LimitResultPages,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [ValidateRange(0,1000)]
        [int]
        $Top,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [String]
        $GraphApiVersion='beta',
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [String]
        $Filter        
    )
    BEGIN
    {
        $Headers=@{Accept="application/json"}
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
                    $GraphQuery+="&`$filter=$Filter"
                }    
                if ($Top -gt 0) {
                    $GraphQuery+="&`$top=$Top"
                }
                $GraphUriBld.Query=$GraphQuery
                $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                    -ContentType 'application/json' `
                    -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' -ErrorProperty 'odata.error' `
                    -LimitResultPages $LimitResultPages -AccessToken $AccessToken
                Write-Output $Result               
            }
            catch
            {
                Write-Warning "[Get-AzureADGraphSigninEvent] $Tenant api-version=$GraphApiVersion $_"    
            }
        }         
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
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
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
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [String]
        $AccessToken,
        [ValidateRange(0,1000)]
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [int]
        $LimitResultPages,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [ValidateRange(0,1000)]
        [int]
        $Top,        
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [String[]]
        $TenantName,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [String]
        $GraphApiVersion='beta',
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [String]
        $Filter 
    )
    BEGIN
    {
        $Headers=@{Accept="application/json"}
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
                    $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                        -ContentType 'application/json' `
                        -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' -ErrorProperty 'odata.error' `
                        -LimitResultPages $LimitResultPages -AccessToken $AccessToken
                    if ($Result -ne $null) {
                        Write-Output $Result
                    }
                }
                catch
                {
                    Write-Warning "[Get-AzureADGraphReport] $Tenant $item api-version=$GraphApiVersion $_"
                }
            }
        }
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
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [String]
        $AccessToken,
        [ValidateRange(0,1000)]
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [int]
        $LimitResultPages,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [int]
        $Top,       
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [String]
        $Filter,        
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [String]
        $GraphApiVersion='1.6',
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [String]
        $TenantName='myOrganization' 
    )
    $Headers=@{Accept="application/json"}
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
    $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
        -ContentType 'application/json' `
        -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' -ErrorProperty 'odata.error' `
        -LimitResultPages $LimitResultPages -AccessToken $AccessToken
    if ($Result -ne $null) {
        Write-Output $Result
    }
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
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [String[]]
        $DomainName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [String]
        $AccessToken,      
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [String]
        $GraphApiVersion='1.6',
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
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
                    $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                        -ContentType 'application/json' `
                        -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' -ErrorProperty 'odata.error' `
                        -LimitResultPages $LimitResultPages -AccessToken $AccessToken
                    if ($Result -ne $null) {
                        Write-Output $Result
                    }
                }
                catch {
                    Write-Warning "[Get-AzureADGraphDomain] $Domain api-version=$GraphApiVersion $_"
                }
            }    
        }
        else
        {
            $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                -ContentType 'application/json' `
                -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' -ErrorProperty 'odata.error' `
                -LimitResultPages $LimitResultPages -AccessToken $AccessToken
            if ($Result -ne $null) {
                Write-Output $Result
            }
        }
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
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [String[]]
        $PolicyId,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [String]
        $AccessToken,      
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [String]
        $GraphApiVersion='1.6',
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [String]
        $TenantName='myOrganization',
        [ValidateRange(0,1000)]
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [int]
        $LimitResultPages,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [ValidateRange(0,1000)]
        [int]
        $Top        
    )
    BEGIN
    {
        $Headers=@{Accept="application/json"}
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
                    $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                        -ContentType 'application/json' `
                        -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' -ErrorProperty 'odata.error' `
                        -LimitResultPages $LimitResultPages -AccessToken $AccessToken
                    if ($Result -ne $null) {
                        Write-Output $Result
                    }
                }
                catch
                {
                    Write-Warning "[Get-AzureADGraphPolicy] $Policy api-version=$GraphApiVersion $_"
                }
            }    
        }
        else
        {
            $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                -ContentType 'application/json' `
                -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' -ErrorProperty 'odata.error' `
                -LimitResultPages $LimitResultPages -AccessToken $AccessToken
            if ($Result -ne $null) {
                Write-Output $Result
            }
        }
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
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [String[]]
        $RoleId,
        [Parameter(Mandatory=$true,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [String]
        $AccessToken,
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [String]
        $GraphApiVersion='beta',
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [String]
        $TenantName='myOrganization'
    )
    BEGIN
    {
        $Headers=@{Accept="application/json"}
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
        $GraphUriBld.Query=$GraphQuery
    }
    PROCESS
    {
        if ($PSCmdlet.ParameterSetName -eq 'noquery' -and $RoleId -ne $null) 
        {
            foreach ($Role in $RoleId)
            {
                try
                {
                    $GraphUriBld.Path="$TenantName/directoryRoles/$Role"
                    $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                        -ContentType 'application/json' `
                        -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' -ErrorProperty 'odata.error' `
                        -LimitResultPages $LimitResultPages -AccessToken $AccessToken
                    if ($Result -ne $null) {
                        Write-Output $Result
                    }
                }
                catch {
                    Write-Warning "[Get-AzureADGraphRole] $Role api-version=$GraphApiVersion $_"
                }
            }    
        }
        else
        {
            $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                -ContentType 'application/json' `
                -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' -ErrorProperty 'odata.error' `
                -LimitResultPages $LimitResultPages -AccessToken $AccessToken
            if ($Result -ne $null) {
                Write-Output $Result
            }
        }
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
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [String[]]
        $TemplateId,
        [Parameter(Mandatory=$true,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [String]
        $Filter,
        [Parameter(Mandatory=$true,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [String]
        $AccessToken,
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [String]
        $GraphApiVersion='1.6',
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [String]
        $TenantName='myOrganization',
        [ValidateRange(0,1000)]
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [int]
        $LimitResultPages,
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [ValidateRange(0,1000)]
        [int]
        $Top
    )
    BEGIN
    {
        $Headers=@{Accept="application/json"}
        $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
        $GraphUriBld.Path="$TenantName/directoryRoleTemplates"
        $GraphQuery="api-version=$GraphApiVersion"
        if ($PSCmdlet.ParameterSetName -eq 'query')
        {
            if ($Filter.Contains('top=') -eq $false -and $Top -gt 0)
            {
                $GraphQuery+="&`$top=$Top"
            }
        }
        else
        {
            if ($Top -gt 0)
            {
                $GraphQuery+="&`$top=$Top"
            }
        }
        $GraphUriBld.Query=$GraphQuery
    }
    PROCESS
    {
        if ($PSCmdlet.ParameterSetName -eq 'noquery' -and $TemplateId -ne $null) 
        {
            foreach ($Template in $TemplateId)
            {
                try
                {
                    $GraphUriBld.Path="$TenantName/directoryRoleTemplates/$Template"
                    $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                        -ContentType 'application/json' `
                        -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' -ErrorProperty 'odata.error' `
                        -LimitResultPages $LimitResultPages -AccessToken $AccessToken
                    if ($Result -ne $null) {
                        Write-Output $Result
                    }
                }
                catch {
                    Write-Warning "[Get-AzureADGraphRoleTemplate] $Template api-version=$GraphApiVersion $_"
                }
            }    
        }
        else
        {
            $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                -ContentType 'application/json' `
                -ValueProperty 'value' -NextLinkProperty '@odata.nextLink'  -ErrorProperty 'odata.error' `
                -LimitResultPages $LimitResultPages -AccessToken $AccessToken
            if ($Result -ne $null) {
                Write-Output $Result
            }
        }
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
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [String[]]$UserId,
        [Parameter(Mandatory=$true,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [String]$Filter,
        [Parameter(Mandatory=$true,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [String]$AccessToken,
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [System.Uri]$GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [String]$GraphApiVersion='beta',
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [String]$TenantName='myOrganization',
        [ValidateRange(0,1000)]
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [int]$LimitResultPages,
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [ValidateRange(0,1000)]
        [int]$Top
    )
    BEGIN
    {
        $Headers=@{Accept="application/json"}
        $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
        $GraphUriBld.Path="$TenantName/users"
        $GraphQuery="api-version=$GraphApiVersion"
        if ($PSCmdlet.ParameterSetName -eq 'query')
        {
            if ($Filter.Contains('top=') -eq $false -and $Top -gt 0)
            {
                $GraphQuery+="&`$top=$Top"
            }
        }
        else
        {
            if ($Top -gt 0)
            {
                $GraphQuery+="&`$top=$Top"
            }
        }
        $GraphUriBld.Query=$GraphQuery
    }
    PROCESS
    {
        if ($PSCmdlet.ParameterSetName -eq 'noquery' -and $UserId -ne $null)
        {
            foreach ($User in $UserId)
            {
                try
                {
                    $GraphUriBld.Path="$TenantName/users/$User"
                    $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                        -ContentType 'application/json' `
                        -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' -ErrorProperty 'odata.error' `
                        -LimitResultPages $LimitResultPages -AccessToken $AccessToken
                    if ($Result -ne $null) {
                        Write-Output $Result
                    }
                }
                catch {
                    Write-Warning "[Get-AzureADGraphUser] $User api-version=$GraphApiVersion $_"
                }
            }            
        }
        else
        {
            $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                -ContentType 'application/json' `
                -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' -ErrorProperty 'odata.error' `
                -LimitResultPages $LimitResultPages -AccessToken $AccessToken
            if ($Result -ne $null) {
                Write-Output $Result
            }
        }
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
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [String[]]$GroupId,
        [Parameter(Mandatory=$true,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [String]$Filter,
        [Parameter(Mandatory=$true,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [String]$AccessToken,
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [System.Uri]$GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [String]$GraphApiVersion='beta',
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [String]$TenantName='myOrganization',
        [ValidateRange(0,1000)]
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [int]$LimitResultPages,
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [ValidateRange(0,1000)]
        [int]$Top
    )
    BEGIN
    {
        $Headers=@{Accept="application/json"}
        $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
        $GraphUriBld.Path="$TenantName/groups"
        $GraphQuery="api-version=$GraphApiVersion"
        if ($PSCmdlet.ParameterSetName -eq 'query')
        {
            if ($Filter.Contains('top=') -eq $false -and $Top -gt 0)
            {
                $GraphQuery+="&`$top=$Top"
            }
        }
        else
        {
            if ($Top -gt 0)
            {
                $GraphQuery+="&`$top=$Top"
            }
        }
        $GraphUriBld.Query=$GraphQuery
    }
    PROCESS
    {
        if ($PSCmdlet.ParameterSetName -eq 'noquery' -and $GroupId -ne $null)
        {
            foreach ($Group in $GroupId)
            {
                try
                {
                    $GraphUriBld.Path="$TenantName/groups/$Group"
                    $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                        -ContentType 'application/json' `
                        -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' -ErrorProperty 'odata.error' `
                        -LimitResultPages $LimitResultPages -AccessToken $AccessToken
                    if ($Result -ne $null) {
                        Write-Output $Result
                    }
                }
                catch
                {
                    Write-Warning "[Get-AzureADGraphGroup] $User api-version=$GraphApiVersion $_"
                }
            }            
        }
        else
        {
            $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                -ContentType 'application/json' `
                -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' `
                -LimitResultPages $LimitResultPages -AccessToken $AccessToken
            if ($Result -ne $null) {
                Write-Output $Result
            }
        }
    }
}

<#
    .SYNOPSIS
        Retrieves details for the tenant
    .PARAMETER AccessToken
        The OAuth Bearer token
    .PARAMETER GraphApiEndpoint
        The Azure Graph API Uri
    .PARAMETER GraphApiVersion
        The Azure Graph API version
    .PARAMETER TenantName
        The directory tenant name
#>
Function Get-AzureADTenantDetails
{
    [CmdletBinding(ConfirmImpact='None')]
    param
    (
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [String[]]$TenantName='myOrganization',        
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [String]$AccessToken,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [System.Uri]$GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [String]$GraphApiVersion='beta',
        [ValidateRange(0,1000)]
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [int]$LimitResultPages,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [ValidateRange(0,1000)]
        [int]$Top
    )
    BEGIN
    {
        $Headers=@{Accept="application/json"}
        $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
        $GraphUriBld.Query="api-version=$GraphApiVersion"
    }
    PROCESS
    {
        foreach ($item in $TenantName)
        {
            try
            {
                $GraphUriBld.Path="$item/tenantDetails"
                $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                    -ContentType 'application/json' `
                    -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' -ErrorProperty 'odata.error' `
                    -LimitResultPages $LimitResultPages -AccessToken $AccessToken
                if ($Result -ne $null) {
                    Write-Output $Result
                }              
            }
            catch {
                Write-Verbose "[Get-AzureADTenantDetails] Error retrieving details for $item $GraphApiVersion $_"   
            } 
        }
    }
}

<#
    .SYNOPSIS
        Retrieves an application from the directory
    .PARAMETER ApplicationId
        The application id(s)
    .PARAMETER ApplicationUri
        The application identifier uri(s)
    .PARAMETER DisplayName
        The application display name
    .PARAMETER Filter
        The OData filter to be applied
    .PARAMETER AccessToken
        The OAuth Bearer token
    .PARAMETER GraphApiEndpoint
        The Azure Graph API Uri
    .PARAMETER GraphApiVersion
        The Azure Graph API version
    .PARAMETER TenantName
        The directory tenant name    
#>
Function Get-AzureADGraphApplication
{
    [CmdletBinding(ConfirmImpact='None',DefaultParameterSetName='noquery')]
    param
    (
        [Parameter(Mandatory=$true,ParameterSetName='byAppId',ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [Guid[]]$ApplicationId,
        [Parameter(Mandatory=$true,ParameterSetName='byAppUri',ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [System.Uri[]]$ApplicationUri,
        [Parameter(Mandatory=$true,ParameterSetName='displayName',ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [System.String[]]$DisplayName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [String]$Filter,
        [Parameter(Mandatory=$true,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true,ParameterSetName='byAppUri',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true,ParameterSetName='byAppId',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true,ParameterSetName='displayName',ValueFromPipelineByPropertyName=$true)]
        [String]$AccessToken,
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='byAppUri',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='byAppId',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='displayName',ValueFromPipelineByPropertyName=$true)]
        [System.Uri]$GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='byAppUri',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='byAppId',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='displayName',ValueFromPipelineByPropertyName=$true)]
        [String]$GraphApiVersion='beta',
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='byAppUri',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='byAppId',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='displayName',ValueFromPipelineByPropertyName=$true)]
        [String]$TenantName='myOrganization',
        [ValidateRange(0,1000)]
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='displayName',ValueFromPipelineByPropertyName=$true)]
        [int]$LimitResultPages,
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='displayName',ValueFromPipelineByPropertyName=$true)]
        [ValidateRange(0,1000)]
        [int]$Top
    )
    BEGIN
    {
        $Headers=@{Accept="application/json"}
        $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
        $GraphUriBld.Path="$TenantName/applications()"
        $GraphQuery="api-version=$GraphApiVersion"
    }
    PROCESS
    {
        if ($PSCmdlet.ParameterSetName -in 'query','noquery')
        {
            if ([String]::IsNullOrEmpty($Filter) -eq $false)
            {
                $GraphQuery+="&`$filter=$Filter"
            }
            $GraphUriBld.Query=$GraphQuery
            $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                -ContentType 'application/json' `
                -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' `
                -LimitResultPages $LimitResultPages -AccessToken $AccessToken
            if ($Result -ne $null) {
                Write-Output $Result
            }            
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'byAppUri')
        {
            foreach ($Uri in $ApplicationUri)
            {
                try
                {
                    $GraphQuery+="&`$filter=identifierUris/any(i:i eq '$($Uri.OriginalString)')"
                    $GraphUriBld.Query=$GraphQuery
                    $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                        -ContentType 'application/json' `
                        -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' `
                        -LimitResultPages $LimitResultPages -AccessToken $AccessToken
                    if ($Result -ne $null) {
                        Write-Output $Result
                    }  
                }
                catch
                {
                    Write-Warning "[Get-AzureADGraphApplication] $GraphApiVersion $_"
                }                
            }
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'byAppId')
        {
            foreach ($Id in $ApplicationId)
            {
                try
                {
                    $GraphQuery+="&`$filter=appId eq '$Id'"
                    $GraphUriBld.Query=$GraphQuery
                    $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                        -ContentType 'application/json' `
                        -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' `
                        -LimitResultPages $LimitResultPages -AccessToken $AccessToken
                    if ($Result -ne $null) {
                        Write-Output $Result
                    }                   
                }
                catch
                {
                    Write-Warning "[Get-AzureADGraphApplication] $GraphApiVersion $_" 
                }               
            }
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'displayName')
        {
            foreach ($Name in $DisplayName)
            {
                try
                {
                    $GraphQuery+="&`$filter=displayName eq '$([Uri]::EscapeDataString($Name))'"
                    $GraphUriBld.Query=$GraphQuery
                    $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                        -ContentType 'application/json' `
                        -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' `
                        -LimitResultPages $LimitResultPages -AccessToken $AccessToken
                    if ($Result -ne $null) {
                        Write-Output $Result
                    }                     
                }
                catch
                {
                    Write-Warning "[Get-AzureADGraphApplication] $GraphApiVersion $_" 
                }               
            }
        }
    }
}

<#
    .SYNOPSIS
        Retrieves a service principal from the directory
    .PARAMETER ApplicationId
        The application id(s)
    .PARAMETER DisplayName
        The application display name
    .PARAMETER Filter
        The OData filter to be applied
    .PARAMETER AccessToken
        The OAuth Bearer token
    .PARAMETER GraphApiEndpoint
        The Azure Graph API Uri
    .PARAMETER GraphApiVersion
        The Azure Graph API version
    .PARAMETER TenantName
        The directory tenant name    
#>
Function Get-AzureADGraphServicePrincipal
{
    [CmdletBinding(ConfirmImpact='None',DefaultParameterSetName='noquery')]
    param
    (
        [Parameter(Mandatory=$true,ParameterSetName='byAppId',ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [Guid[]]
        $ApplicationId,
        [Parameter(Mandatory=$true,ParameterSetName='displayName',ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [System.String[]]
        $DisplayName,
        [Parameter(Mandatory=$true,ParameterSetName='appDisplayName',ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [System.String[]]
        $AppDisplayName,                
        [Parameter(Mandatory=$true,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [String]
        $Filter,
        [Parameter(Mandatory=$true,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true,ParameterSetName='byAppId',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true,ParameterSetName='displayName',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true,ParameterSetName='appDisplayName',ValueFromPipelineByPropertyName=$true)]
        [String]
        $AccessToken,
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='byAppId',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='displayName',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true,ParameterSetName='appDisplayName',ValueFromPipelineByPropertyName=$true)]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='byAppId',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='displayName',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true,ParameterSetName='appDisplayName',ValueFromPipelineByPropertyName=$true)]
        [String]
        $GraphApiVersion='1.6',
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='byAppId',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='displayName',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true,ParameterSetName='appDisplayName',ValueFromPipelineByPropertyName=$true)]
        [String]
        $TenantName='myOrganization',
        [ValidateRange(0,1000)]
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='displayName',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true,ParameterSetName='appDisplayName',ValueFromPipelineByPropertyName=$true)]
        [int]
        $LimitResultPages,
        [Parameter(Mandatory=$false,ParameterSetName='query',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='noquery',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='displayName',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true,ParameterSetName='appDisplayName',ValueFromPipelineByPropertyName=$true)]
        [ValidateRange(0,1000)]
        [int]
        $Top
    )
    BEGIN
    {
        $Headers=@{Accept="application/json"}
        $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
        $GraphUriBld.Path="$TenantName/servicePrincipals()"
        $GraphQuery="api-version=$GraphApiVersion"
    }
    PROCESS
    {
        if ($PSCmdlet.ParameterSetName -in 'query','noquery')
        {
            if ([String]::IsNullOrEmpty($Filter) -eq $false)
            {
                $GraphQuery+="&`$filter=$Filter"
            }
            $GraphUriBld.Query=$GraphQuery
            $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                -ContentType 'application/json' `
                -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' `
                -LimitResultPages $LimitResultPages -AccessToken $AccessToken
            if ($Result -ne $null) {
                Write-Output $Result
            }            
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'byAppId')
        {
            foreach ($Id in $ApplicationId)
            {
                try
                {
                    $GraphQuery+="&`$filter=appId eq '$Id'"
                    $GraphUriBld.Query=$GraphQuery
                    $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                        -ContentType 'application/json' `
                        -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' `
                        -LimitResultPages $LimitResultPages -AccessToken $AccessToken
                    if ($Result -ne $null) {
                        Write-Output $Result
                    }                   
                }
                catch
                {
                    Write-Warning "[Get-AzureADGraphApplication] $GraphApiVersion $_" 
                }               
            }
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'displayName')
        {
            foreach ($Name in $DisplayName)
            {
                try
                {
                    $GraphQuery+="&`$filter=displayName eq '$([Uri]::EscapeDataString($Name))'"
                    $GraphUriBld.Query=$GraphQuery
                    $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                        -ContentType 'application/json' `
                        -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' `
                        -LimitResultPages $LimitResultPages -AccessToken $AccessToken
                    if ($Result -ne $null) {
                        Write-Output $Result
                    }                     
                }
                catch
                {
                    Write-Warning "[Get-AzureADGraphApplication] $GraphApiVersion $_" 
                }               
            }
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'appDisplayName')
        {
            foreach ($Name in $AppDisplayName)
            {
                try
                {
                    $GraphQuery+="&`$filter=appDisplayName eq '$([Uri]::EscapeDataString($Name))'"
                    $GraphUriBld.Query=$GraphQuery
                    $Result=Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -AdditionalHeaders $Headers `
                        -ContentType 'application/json' `
                        -ValueProperty 'value' -NextLinkProperty '@odata.nextLink' `
                        -LimitResultPages $LimitResultPages -AccessToken $AccessToken
                    if ($Result -ne $null) {
                        Write-Output $Result
                    }                     
                }
                catch
                {
                    Write-Warning "[Get-AzureADGraphApplication] $GraphApiVersion $_" 
                }               
            }
        }        
    }
}

<#
    .SYNOPSIS
        Removes an item from the directory
    .PARAMETER Object
        The azure active directory object
    .PARAMETER ObjectId
        The azure active directory object id
    .PARAMETER AccessToken
        The OAuth Bearer token
    .PARAMETER GraphApiEndpoint
        The Azure Graph API Uri
    .PARAMETER GraphApiVersion
        The Azure Graph API version
    .PARAMETER TenantName
        The directory tenant name
#>
Function Remove-AzureADGraphObject
{
    [CmdletBinding(ConfirmImpact='None',DefaultParameterSetName='object')]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [Object[]]$Object,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ParameterSetName='id',ValueFromPipelineByPropertyName=$true)]
        [string[]]$ObjectId,
        [Parameter(Mandatory=$true,ParameterSetName='id',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [string]$AccessToken,
        [Parameter(Mandatory=$false,ParameterSetName='id',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [System.Uri]$GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ParameterSetName='id',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [String]$GraphApiVersion='beta',
        [Parameter(Mandatory=$false,ParameterSetName='id',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [String]$TenantName='myOrganization'                        
    )
    BEGIN
    {
        $Headers=@{Accept='application/json'}
        $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
        $GraphUriBld.Query="api-version=$GraphApiVersion"
    }
    PROCESS
    {
        if ($PSCmdlet.ParameterSetName -eq "object") {
            $ObjectId=$Object|Select-Object -ExpandProperty 'objectId'
        }
        foreach ($Id in $ObjectId)
        {
            try
            {
                $GraphUriBld.Path="$TenantName/directoryObjects/$Id"
                Invoke-AzureADGraphRequest -Uri $GraphUriBld.Uri -Method Delete -AdditionalHeaders $Headers `
                    -ContentType 'application/json' -AccessToken $AccessToken | Out-Null
            }
            catch
            {
                Write-Warning "[Remove-AzureADGraphObject] Error removing $Id $GraphApiVersion $_"
            }
        }
    }
}

Function New-AzureADGraphUser
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true,ParameterSetName='object',ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [object[]]$User,
        [Parameter(Mandatory=$true,ParameterSetName='explicit',ValueFromPipelineByPropertyName=$true)]
        [bool]$AccountEnabled,        
        [Parameter(Mandatory=$true,ParameterSetName='explicit',ValueFromPipelineByPropertyName=$true)]
        [string]$DisplayName,
        [Parameter(Mandatory=$true,ParameterSetName='explicit',ValueFromPipelineByPropertyName=$true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory=$true,ParameterSetName='explicit',ValueFromPipelineByPropertyName=$true)]
        [securestring]$Password, 
        [Parameter(Mandatory=$false,ParameterSetName='explicit',ValueFromPipelineByPropertyName=$true)]
        [bool]$ForcePasswordChange,          
        [Parameter(Mandatory=$true,ParameterSetName='explicit',ValueFromPipelineByPropertyName=$true)]
        [string]$MailNickName,
        [Parameter(Mandatory=$true,ParameterSetName='explicit',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [string]$AccessToken,
        [Parameter(Mandatory=$false,ParameterSetName='explicit',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [System.Uri]$GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ParameterSetName='explicit',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [String]$GraphApiVersion='beta',
        [Parameter(Mandatory=$false,ParameterSetName='explicit',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [String]$TenantName='myOrganization'
    )
    BEGIN
    {
            $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
            $GraphUriBld.Path="$TenantName/users"
            $GraphUriBld.Query="api-version=$GraphApiVersion"
    }
    PROCESS
    {
        if ($PSCmdlet.ParameterSetName -eq 'explicit')
        {
            $PasswordProfile=[ordered]@{
                'password'=$($Password|ConvertFromSecureString);
                'forceChangePasswordNextLogin'=$ForcePasswordChange
            }
            $UserProperties=[ordered]@{
                'accountEnabled'=$AccountEnabled;
                'displayName'=$DisplayName;
                'mailNickname'=$MailNickName;
                'passwordProfile'=$(New-Object psobject -Property $PasswordProfile)
                'userPrincipalName'=$UserPrincipalName;
            }
            $User=@($(New-Object psobject -Property $UserProperties))
        }
        foreach ($item in $User)
        {
            $GraphRequestParams=@{
                Uri=$GraphUriBld.Uri;
                ContentType='application/json';
                Method='POST';
                AccessToken=$AccessToken;
                Body=$item;
            }
            $NewUser=Invoke-AzureADGraphRequest @GraphRequestParams
            Write-Output $NewUser
        }
    }
}

Function New-AzureADGraphGroup
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true,ParameterSetName='object',ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [object[]]$Group,      
        [Parameter(Mandatory=$true,ParameterSetName='explicit',ValueFromPipelineByPropertyName=$true)]
        [string]$DisplayName,
        [Parameter(Mandatory=$false,ParameterSetName='explicit',ValueFromPipelineByPropertyName=$true)]
        [bool]$MailEnabled,
        [Parameter(Mandatory=$true,ParameterSetName='explicit',ValueFromPipelineByPropertyName=$true)]
        [string]$MailNickName,
        [Parameter(Mandatory=$true,ParameterSetName='explicit',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [string]$AccessToken,
        [Parameter(Mandatory=$false,ParameterSetName='explicit',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [System.Uri]$GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ParameterSetName='explicit',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [String]$GraphApiVersion='beta',
        [Parameter(Mandatory=$false,ParameterSetName='explicit',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [String]$TenantName='myOrganization'
    )
    BEGIN
    {
            $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
            $GraphUriBld.Path="$TenantName/groups"
            $GraphUriBld.Query="api-version=$GraphApiVersion"
    }
    PROCESS
    {
        if ($PSCmdlet.ParameterSetName -eq 'explicit')
        {
            $GroupProperties=[ordered]@{
                'displayName'=$DisplayName;
                'mailNickname'=$MailNickName;
                'mailEnabled'=$MailEnabled;
                'securityEnabled'=$true
            }
            $Group=@($(New-Object psobject -Property $GroupProperties))
        }
        foreach ($item in $User)
        {
            $GraphRequestParams=@{
                Uri=$GraphUriBld.Uri;
                ContentType='application/json';
                Method='POST';
                AccessToken=$AccessToken;
                Body=$item;
            }
            $NewGroup=Invoke-AzureADGraphRequest @GraphRequestParams
            Write-Output $NewGroup
        }
    }
}

Function New-AzureADGraphServicePrincipal
{
    [CmdletBinding()]
    param
    (

        [Parameter(Mandatory=$true,ParameterSetName='id',ValueFromPipelineByPropertyName=$true)]
        [object[]]$ServicePrincipal,
        [Parameter(Mandatory=$true,ParameterSetName='id',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [string]$AccessToken,
        [Parameter(Mandatory=$false,ParameterSetName='id',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [System.Uri]$GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ParameterSetName='id',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [String]$GraphApiVersion='beta',
        [Parameter(Mandatory=$false,ParameterSetName='id',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$false,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [String]$TenantName='myOrganization'
    )
    BEGIN
    {
        $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
        $GraphUriBld.Path="$TenantName/servicePrincipals"
        $GraphUriBld.Query="api-version=$GraphApiVersion"
    }
    PROCESS
    {
        foreach ($item in $ServicePrincipal)
        {
            $RequestParams=@{
                Uri=$GraphUriBld.Uri;
                AccessToken=$AccessToken;
                Method='POST'
                ContentType='application/json';
                Body=$item;
            }
            $NewApp=Invoke-AzureADGraphRequest @RequestParams
            Write-Output $NewApp 
        }
    }
}

Function New-AzureADGraphApplication
{
    [CmdletBinding()]
    param
    (

        [Parameter(Mandatory=$true,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [object[]]
        $Application,
        [Parameter(Mandatory=$true,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [string]
        $AccessToken,
        [Parameter(Mandatory=$false,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [System.Uri]
        $GraphApiEndpoint='https://graph.windows.net',
        [Parameter(Mandatory=$false,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [String]
        $GraphApiVersion='beta',
        [Parameter(Mandatory=$false,ParameterSetName='object',ValueFromPipelineByPropertyName=$true)]
        [String]
        $TenantName='myOrganization'
    )
    BEGIN
    {
            $GraphUriBld=New-Object System.UriBuilder($GraphApiEndpoint)
            $GraphUriBld.Path="$TenantName/applications"
            $GraphUriBld.Query="api-version=$GraphApiVersion"
    }
    PROCESS
    {s
        foreach ($item in $Application)
        {
            $RequestParams=@{
                Uri=$GraphUriBld.Uri;
                AccessToken=$AccessToken;
                Method='POST'
                ContentType='application/json';
                Body=$item;
            }
            $NewApp=Invoke-AzureADGraphRequest @RequestParams
            Write-Output $NewApp
        }
    }
}

function New-AzureADGraphOauthPermissionGrant
{
    [CmdletBinding()]
    param
    (
        
    )
    
    begin
    {
    }
    
    process
    {
    }
    
    end
    {
    }
}

#endregion

#region Request Body Helpers
function New-AzureADGraphUserParameter
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [bool]$AccountEnabled,        
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$DisplayName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [securestring]$Password, 
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [bool]$ForcePasswordChange,          
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$MailNickName        
    )
    PROCESS
    {
        $PasswordProfile=[ordered]@{
            'password'=$($Password|ConvertFromSecureString);
            'forceChangePasswordNextLogin'=$ForcePasswordChange
        }
        $UserProperties=[ordered]@{
            'accountEnabled'=$AccountEnabled;
            'displayName'=$DisplayName;
            'mailNickname'=$MailNickName;
            'passwordProfile'=$(New-Object psobject -Property $PasswordProfile)
            'userPrincipalName'=$UserPrincipalName;
        }
        $User=@($(New-Object psobject -Property $UserProperties))
        Write-Output $User        
    }
}

function New-AzureADGraphGroupParameter
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [bool]$AccountEnabled,        
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$DisplayName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [securestring]$Password, 
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [bool]$ForcePasswordChange,          
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$MailNickName,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [bool]$SecurityEnabled=$true              
    )
    process
    {
        $GroupProperties=[ordered]@{
            'displayName'=$DisplayName;
            'mailNickname'=$MailNickName;
            'mailEnabled'=$MailEnabled;
            'securityEnabled'=$SecurityEnabled
        }
        $Group=New-Object psobject -Property $GroupProperties
        Write-Output $Group
    }
}

function New-AzureADGraphServicePrincipalParameter
{
    [CmdletBinding()]
    param
    (
        
    )
    
    begin
    {
    }
    
    process
    {
    }
    
    end
    {
    }
}

function New-AzureADGraphApplicationParameter
{
    [CmdletBinding()]
    param
    (
        
    )
    
    begin
    {
    }
    
    process
    {
    }
    
    end
    {
    }
}

function New-AzureADGraphOauthPermissionGrantParameter
{
    [CmdletBinding()]
    param
    (
        
    )
    
    begin
    {
    }
    
    process
    {
    }
    
    end
    {
    }
}

#endregion