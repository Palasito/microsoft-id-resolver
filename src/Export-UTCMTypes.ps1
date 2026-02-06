<#
.SYNOPSIS
    Crawls Microsoft Learn pages to extract resource types and add appropriate prefixes for snapshot creation.

.DESCRIPTION
    This script periodically fetches resource types from Microsoft Learn documentation pages
    and adds the appropriate 'microsoft.<type>.' prefix to each resource type.
    The prefix type is dynamically determined from the URL pattern.

.NOTES
    Version: 1.0
    Author: Pipeline Automation
    Last Updated: February 2026
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\_info",

    [Parameter(Mandatory = $false)]
    [string]$FileName = "utcm-resource-types.json",

    [Parameter(Mandatory = $false)]
    [switch]$ExportToJson,

    [Parameter(Mandatory = $false)]
    [switch]$ExportToCsv
)

# Static list of URLs to crawl - Update this list manually as needed
$ResourceUrls = @(
    @{
        Url  = "https://learn.microsoft.com/en-us/graph/utcm-intune-resources"
        Type = "intune"
    },
    @{
        Url  = "https://learn.microsoft.com/en-us/graph/utcm-exchange-resources"
        Type = "exchange"
    },
    @{
        Url  = "https://learn.microsoft.com/en-us/graph/utcm-entra-resources"
        Type = "entra"
    },
    @{
        Url  = "https://learn.microsoft.com/en-us/graph/utcm-securityandcompliance-resources"
        Type = "securityandcompliance"
    },
    @{
        Url  = "https://learn.microsoft.com/en-us/graph/utcm-teams-resources"
        Type = "teams"
    }
    # Add more resource types here as needed:
    # @{
    #     Url = "https://learn.microsoft.com/en-us/graph/utcm-<type>-resources"
    #     Type = "<type>"
    # }
)

# Function to extract resource types from Microsoft Learn page
function Get-ResourceTypesFromPage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url,
        
        [Parameter(Mandatory = $true)]
        [string]$ResourceType
    )
    
    try {
        Write-Host "Fetching resource types from: $Url" -ForegroundColor Cyan
        
        # Fetch the page content
        $response = Invoke-WebRequest -Uri $Url -UseBasicParsing -ErrorAction Stop
        
        # Parse the HTML content
        $content = $response.Content
        
        # Extract resource types from the page
        $resourceTypes = @()
        
        # Primary Pattern: Look for anchor links with pattern #<resourcetype>-resource-type
        # Example: href="#deviceconfigurationpolicyios-resource-type"
        $anchorMatches = [regex]::Matches($content, 'href="#([a-zA-Z0-9_]+)-resource-type"')
        
        Write-Host "Found $($anchorMatches.Count) anchor links with resource type pattern" -ForegroundColor Gray
        
        foreach ($match in $anchorMatches) {
            $resourceName = $match.Groups[1].Value
            
            if ($resourceName -and $resourceName -notmatch '^(http|https|www|learn|microsoft|graph|resource|type)$' -and $resourceName.Length -gt 2) {
                $resourceTypes += $resourceName
            }
        }
        
        # Secondary Pattern: Look for ID attributes with resource type pattern
        # Example: id="deviceconfigurationpolicyios-resource-type"
        $idMatches = [regex]::Matches($content, 'id="([a-zA-Z0-9_]+)-resource-type"')
        
        Write-Host "Found $($idMatches.Count) ID elements with resource type pattern" -ForegroundColor Gray
        
        foreach ($match in $idMatches) {
            $resourceName = $match.Groups[1].Value
            
            if ($resourceName -and $resourceName -notmatch '^(http|https|www|learn|microsoft|graph|resource|type)$' -and $resourceName.Length -gt 2) {
                $resourceTypes += $resourceName
            }
        }

        # Tertiary Pattern: Look for heading elements containing "resource type"
        # Example: <h2>deviceConfigurationPolicyiOS resource type</h2>
        $headingMatches = [regex]::Matches($content, '<h[2-4][^>]*>([a-zA-Z0-9_]+)\s+resource\s+type</h[2-4]>', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        
        Write-Host "Found $($headingMatches.Count) heading elements with resource type pattern" -ForegroundColor Gray
        
        foreach ($match in $headingMatches) {
            $resourceName = $match.Groups[1].Value
            
            if ($resourceName -and $resourceName -notmatch '^(http|https|www|learn|microsoft|graph|resource|type)$' -and $resourceName.Length -gt 2) {
                $resourceTypes += $resourceName
            }
        }
        
        # Remove duplicates (case-insensitive) and sort
        $uniqueResourceTypes = $resourceTypes | Sort-Object -Unique -Property @{Expression={$_.ToLower()}}
        
        Write-Host "Found $($uniqueResourceTypes.Count) unique resource types for $ResourceType" -ForegroundColor Green
        
        # Add prefix to each resource type
        $prefixedResources = @()
        foreach ($resource in $uniqueResourceTypes) {
            $prefixedResource = "microsoft.$ResourceType.$resource"
            
            # Try to extract permissions for this resource type
            $permissionsData = Get-PermissionsForResourceType -Content $content -ResourceName $resource
            
            $resourceObject = @{
                OriginalName           = $resource
                PrefixedName           = $prefixedResource
                ResourceType           = $ResourceType
                AnchorUrl              = "$Url#$($resource.ToLower())-resource-type"
                LastUpdated            = (Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ")
            }
            
            # Add permissions data if available
            if ($permissionsData -and $permissionsData.AllPermissions) {
                $resourceObject.ApplicationPermissions = $permissionsData.AllPermissions
                $resourceObject.OperationPermissions = $permissionsData.OperationPermissions
            } else {
                $resourceObject.ApplicationPermissions = @()
                $resourceObject.OperationPermissions = @()
            }
            
            # Add Exchange.ManageAsApp permission for all Exchange resource types
            if ($ResourceType -eq "exchange") {
                if ($resourceObject.ApplicationPermissions -notcontains "Exchange.ManageAsApp") {
                    $resourceObject.ApplicationPermissions = @($resourceObject.ApplicationPermissions) + @("Exchange.ManageAsApp")
                    Write-Verbose "Added Exchange.ManageAsApp permission to $resource"
                }
            }
            
            $prefixedResources += $resourceObject
        }
        
        return $prefixedResources
    }
    catch {
        Write-Error "Failed to fetch resource types from $Url : $_"
        return @()
    }
}

# Function to extract permissions for a specific resource type
function Get-PermissionsForResourceType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Content,
        
        [Parameter(Mandatory = $true)]
        [string]$ResourceName
    )
    
    try {
        # Look for the section containing this resource type
        # Pattern: Find the section starting with the resource type anchor/heading
        # Search up to 30000 characters or until next h2 heading ONLY (h3 is subsection)
        Write-Verbose "`n=== DEBUG: Searching for permissions: $ResourceName ==="
        $sectionPattern = "id=`"$ResourceName-resource-type`"[\s\S]{0,30000}?(?=<h2\s|$)"
        $sectionMatch = [regex]::Match($Content, $sectionPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        
        if (-not $sectionMatch.Success) {
            # Try alternate pattern with the anchor link
            Write-Verbose "DEBUG: ID pattern failed, trying href..."
            $sectionPattern = "href=`"#$ResourceName-resource-type`"[\s\S]{0,30000}?(?=<h2\s|$)"
            $sectionMatch = [regex]::Match($Content, $sectionPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            
            if (-not $sectionMatch.Success) {
                Write-Verbose "DEBUG: ❌ Could not find section for resource: $ResourceName"
                return @()
            }
        }
        
        Write-Verbose "DEBUG: ✓ Section found (length: $($sectionMatch.Value.Length) chars)"
        $sectionContent = $sectionMatch.Value
        
        # Show preview of section
        $preview = $sectionContent.Substring(0, [Math]::Min(300, $sectionContent.Length)) -replace '\s+', ' '
        Write-Verbose "DEBUG: Section preview: $preview..."
        
        # Try to find "Application permissions" or "Microsoft Graph permissions" heading followed by table content
        # Support multiple HTML structures: standard tables, divs with tables, markdown-rendered tables
        Write-Verbose "DEBUG: Searching for permission tables..."
        $permissionsPatterns = @(
            # Pattern 1: Standard table after "Application permissions"
            'Application\s+permissions[\s\S]{0,3000}?<table[\s\S]{0,5000}?</table>',
            # Pattern 2: Microsoft Graph + Application permissions + table
            'Microsoft\s+Graph[\s\S]{0,500}?Application\s+permissions[\s\S]{0,3000}?<table[\s\S]{0,5000}?</table>',
            # Pattern 3: Any heading with "permissions" + table
            '<h[3-5][^>]*>.*?[Pp]ermissions.*?</h[3-5]>[\s\S]{0,3000}?<table[\s\S]{0,5000}?</table>',
            # Pattern 4: Permissions in div/section wrapper
            'Application\s+permissions[\s\S]{0,500}?<(?:div|section)[^>]*>[\s\S]{0,5000}?<table[\s\S]{0,5000}?</table>',
            # Pattern 5: Markdown-style table (pipe-delimited)
            'Application\s+permissions[\s\S]{0,500}?\|[^\n]*\|[\s\S]{0,3000}?(?=\n\n|<h|$)'
        )
        
        $tableContent = $null
        $patternIndex = 0
        foreach ($pattern in $permissionsPatterns) {
            $patternIndex++
            Write-Verbose "DEBUG: Trying pattern $patternIndex..."
            $permissionsMatch = [regex]::Match($sectionContent, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if ($permissionsMatch.Success) {
                $tableContent = $permissionsMatch.Value
                Write-Verbose "DEBUG: ✓ Pattern $patternIndex matched! Table length: $($tableContent.Length)"
                break
            }
        }
        
        if (-not $tableContent) {
            Write-Verbose "DEBUG: ⚠ No permission table found with standard patterns (tried $patternIndex patterns)"
            Write-Verbose "DEBUG: Attempting fallback: extracting permission strings from section..."
            
            # Fallback: Search for "Application permissions" section and extract any Graph permission strings
            $permSectionPattern = 'Application\s+permissions[\s\S]{0,5000}?(?=<h[3-5]|$)'
            $permSectionMatch = [regex]::Match($sectionContent, $permSectionPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            
            if ($permSectionMatch.Success) {
                Write-Verbose "DEBUG: ✓ Found permissions section (fallback)"
                $permSectionText = $permSectionMatch.Value
                
                # Extract all Graph permission strings directly (Capital.Capital.Capital pattern)
                $permissionPattern = '\b([A-Z][A-Za-z0-9]*\.[A-Z][A-Za-z0-9]*\.[A-Z][A-Za-z0-9]*)\b'
                $permMatches = [regex]::Matches($permSectionText, $permissionPattern)
                
                if ($permMatches.Count -gt 0) {
                    Write-Verbose "DEBUG: ✓ Extracted $($permMatches.Count) permissions using fallback pattern"
                    $fallbackPerms = @()
                    foreach ($match in $permMatches) {
                        $perm = $match.Groups[1].Value
                        if ($perm -and $perm -notmatch '^(Operation|Supported|Permissions|Microsoft|Graph)$') {
                            $fallbackPerms += $perm
                        }
                    }
                    
                    if ($fallbackPerms.Count -gt 0) {
                        $uniquePerms = $fallbackPerms | Select-Object -Unique | Sort-Object
                        Write-Verbose "DEBUG: ✓ Returning $($uniquePerms.Count) unique permissions (fallback) for: $ResourceName"
                        Write-Host "  Found $($uniquePerms.Count) permissions for $ResourceName (fallback)" -ForegroundColor DarkYellow
                        return @{
                            AllPermissions = $uniquePerms
                            OperationPermissions = @()  # Can't determine operations without table structure
                        }
                    }
                }
            }
            
            Write-Verbose "DEBUG: ❌ No permissions found for: $ResourceName (exhausted all methods)"
            return @()
        }
        
        # Extract all table rows (excluding header)
        $rowMatches = [regex]::Matches($tableContent, '<tr[^>]*>(.*?)</tr>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        
        Write-Verbose "DEBUG: Found $($rowMatches.Count) table rows"
        
        if ($rowMatches.Count -eq 0) {
            Write-Verbose "DEBUG: ❌ No table rows found for: $ResourceName"
            return @()
        }
        
        $permissions = @()
        $operationPermissions = @()
        $rowsProcessed = 0
        
        foreach ($rowMatch in $rowMatches) {
            $rowContent = $rowMatch.Groups[1].Value
            
            # Skip header rows
            if ($rowContent -match '<th') {
                continue
            }
            
            $rowsProcessed++
            
            # Extract all table cells
            $cellMatches = [regex]::Matches($rowContent, '<td[^>]*>(.*?)</td>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
            
            if ($cellMatches.Count -ge 2) {
                $operation = $null
                
                # First cell is typically the Operation
                if ($cellMatches.Count -gt 0) {
                    $operationCell = $cellMatches[0].Groups[1].Value
                    $operation = $operationCell -replace '<[^>]+>', '' -replace '&nbsp;', ' ' -replace '&amp;', '&' -replace '\s+', ' '
                    $operation = $operation.Trim()
                }
                
                # Second cell is typically the Supported permissions
                if ($cellMatches.Count -gt 1) {
                    $permissionCell = $cellMatches[1].Groups[1].Value
                    
                    # Remove HTML tags but keep the text
                    $cleanCell = $permissionCell -replace '<[^>]+>', ' '
                    
                    # Extract all permission patterns (e.g., Group.Read.All, DeviceManagementConfiguration.Read.All)
                    # Pattern: word.word.word (each part starts with capital letter)
                    $permissionMatches = [regex]::Matches($cleanCell, '\b([A-Z][A-Za-z0-9]*\.[A-Z][A-Za-z0-9]*\.[A-Z][A-Za-z0-9]*)\b')
                    
                    $cellPermissions = @()
                    foreach ($permMatch in $permissionMatches) {
                        $perm = $permMatch.Groups[1].Value
                        if ($perm -and $perm -notmatch '^(Operation|Supported|Permissions|Microsoft|Graph)$') {
                            $cellPermissions += $perm
                            $permissions += $perm
                        }
                    }
                    
                    if ($operation -and $cellPermissions.Count -gt 0) {
                        Write-Verbose "DEBUG: Found permissions for operation '$operation': $($cellPermissions -join ', ')"
                        $operationPermissions += @{
                            Operation = $operation
                            Permissions = $cellPermissions
                        }
                    }
                }
            }
        }
        
        Write-Verbose "DEBUG: Processed $rowsProcessed data rows, found $($permissions.Count) permission mentions"
        
        # Return unique permissions with operation context
        $uniquePermissions = $permissions | Select-Object -Unique | Sort-Object
        
        if ($uniquePermissions.Count -gt 0) {
            Write-Verbose "DEBUG: ✓ Returning $($uniquePermissions.Count) unique permissions for: $ResourceName"
            Write-Host "  Found $($uniquePermissions.Count) permissions for $ResourceName" -ForegroundColor DarkGreen
            return @{
                AllPermissions = $uniquePermissions
                OperationPermissions = $operationPermissions
            }
        }
        
        Write-Verbose "DEBUG: ⚠ No permissions found for: $ResourceName"
        return @()
    }
    catch {
        Write-Verbose "DEBUG: ❌ ERROR extracting permissions for $ResourceName : $_"
        Write-Host "  ERROR: $_" -ForegroundColor Red
        return @()
    }
}

# Function to dynamically determine resource type from URL if not provided
function Get-ResourceTypeFromUrl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url
    )
    
    # Extract resource type from URL pattern: utcm-<type>-resources
    if ($Url -match 'utcm-([a-zA-Z0-9]+)-resources') {
        return $matches[1]
    }
    
    # Fallback: try to extract from path
    if ($Url -match '/([a-zA-Z0-9]+)-resources') {
        return $matches[1]
    }
    
    return "unknown"
}

# Main execution
Write-Host "==================================================" -ForegroundColor Yellow
Write-Host "Resource Type Crawler - Starting" -ForegroundColor Yellow
Write-Host "==================================================" -ForegroundColor Yellow
Write-Host ""

$allResources = @()
$summary = @()

foreach ($resourceConfig in $ResourceUrls) {
    $url = $resourceConfig.Url
    $type = $resourceConfig.Type
    
    # If type is not provided, try to determine it from URL
    if ([string]::IsNullOrWhiteSpace($type)) {
        $type = Get-ResourceTypeFromUrl -Url $url
        Write-Host "Dynamically determined resource type: $type" -ForegroundColor Magenta
    }
    
    # Fetch and process resources
    $resources = Get-ResourceTypesFromPage -Url $url -ResourceType $type
    
    if ($resources.Count -gt 0) {
        $allResources += $resources
        $summary += @{
            ResourceType = $type
            Url          = $url
            Count        = $resources.Count
        }
    }
    
    Write-Host ""
}

# Display summary
Write-Host "==================================================" -ForegroundColor Yellow
Write-Host "Summary" -ForegroundColor Yellow
Write-Host "==================================================" -ForegroundColor Yellow

foreach ($item in $summary) {
    Write-Host "Resource Type: $($item.ResourceType)" -ForegroundColor Cyan
    Write-Host "  URL: $($item.Url)" -ForegroundColor Gray
    Write-Host "  Count: $($item.Count)" -ForegroundColor Green
    Write-Host ""
}

Write-Host "Total prefixed resources: $($allResources.Count)" -ForegroundColor Green
Write-Host ""

# Export results
$jsonOutput = @{
    GeneratedDate  = (Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ")
    TotalResources = $allResources.Count
    ResourceTypes  = $summary
    Resources      = $allResources
} | ConvertTo-Json -Depth 10

$OutputPath = Join-Path $OutputPath $FileName

$jsonOutput | Out-File -FilePath $OutputPath -Encoding UTF8 -Force
Write-Host "Results exported to: $OutputPath" -ForegroundColor Green


$csvPath = $OutputPath -replace '\.json$', '.csv'
$allResources | ForEach-Object {
    [PSCustomObject]@{
        AnchorUrl              = $_.AnchorUrl
        ApplicationPermissions = if ($_.ApplicationPermissions) { $_.ApplicationPermissions -join ', ' } else { '' }
        OperationPermissions   = if ($_.OperationPermissions) { 
            ($_.OperationPermissions | ForEach-Object { "$($_.Operation): $($_.Permissions -join ', ')" }) -join ' | ' 
        } else { '' }
        ResourceType           = $_.ResourceType
        OriginalName           = $_.OriginalName
        PrefixedName           = $_.PrefixedName
        LastUpdated            = $_.LastUpdated
    }
} | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
Write-Host "Results exported to CSV: $csvPath" -ForegroundColor Green


# Output sample of prefixed resources for verification
Write-Host "==================================================" -ForegroundColor Yellow
Write-Host "Sample Prefixed Resources (first 5)" -ForegroundColor Yellow
Write-Host "==================================================" -ForegroundColor Yellow

$allResources | Select-Object -First 5 | ForEach-Object {
    Write-Host ""
    Write-Host "$($_.PrefixedName)" -ForegroundColor White
    Write-Host "  Anchor: $($_.AnchorUrl)" -ForegroundColor DarkGray
    
    if ($_.ApplicationPermissions -and $_.ApplicationPermissions.Count -gt 0) {
        Write-Host "  All Permissions: $($_.ApplicationPermissions -join ', ')" -ForegroundColor Cyan
        
        if ($_.OperationPermissions -and $_.OperationPermissions.Count -gt 0) {
            Write-Host "  Permission Details:" -ForegroundColor Gray
            foreach ($opPerm in $_.OperationPermissions) {
                Write-Host "    - $($opPerm.Operation): $($opPerm.Permissions -join ', ')" -ForegroundColor DarkCyan
            }
        }
    } else {
        Write-Host "  No permissions found" -ForegroundColor DarkYellow
    }
}

Write-Host ""
Write-Host "==================================================" -ForegroundColor Yellow

# Count resources with permissions
$resourcesWithPermissions = ($allResources | Where-Object { $_.ApplicationPermissions -and $_.ApplicationPermissions.Count -gt 0 }).Count
Write-Host "Resources with permissions: $resourcesWithPermissions / $($allResources.Count)" -ForegroundColor Green

Write-Host ""
Write-Host "Script completed successfully!" -ForegroundColor Green