<#
.SYNOPSIS
    Extracts UTCM resource types from the official JSON schema at schemastore.org.

.DESCRIPTION
    This script fetches the authoritative UTCM JSON schema from schemastore.org and extracts
    all resource types with their metadata. The schema is always up-to-date and contains
    all resource types in microsoft.<type>.<resourcename> format.

.NOTES
    Version: 2.0
    Author: Pipeline Automation
    Last Updated: February 2026
    Schema Source: https://www.schemastore.org/utcm-monitor.json
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\_info",

    [Parameter(Mandatory = $false)]
    [string]$JsonFileName = "utcm-resource-types.json",

    [Parameter(Mandatory = $false)]
    [string]$CsvFileName = "utcm-resource-types.csv",

    [Parameter(Mandatory = $false)]
    [string]$SchemaUrl = "https://www.schemastore.org/utcm-monitor.json"
)

# Function to convert camelCase/PascalCase to friendly display name
function ConvertTo-FriendlyName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )
    
    # Check if the string is all lowercase (no camelCase)
    $hasUpperCase = $Name -cmatch '[A-Z]'
    
    if (-not $hasUpperCase) {
        # All lowercase - use intelligent word boundary detection
        # Strategy: comprehensive English word library to split concatenated words
        
        $friendlyName = $Name
        
        # Build comprehensive word library (sorted by length descending for longest-match-first)
        $words = @(
            # 5+ letters (most specific)
            'authentication', 'authorization', 'administrative', 'administrator', 'availability', 
            'configuration', 'classification', 'synchronization', 'notification', 'registration',
            'certificate', 'assignment', 'encryption', 'federation', 'deployment', 'integration',
            'distribution', 'subscription', 'organization', 'optimization', 'restriction',
            'template', 'delivery', 'detection', 'interface', 'monitoring', 'onboarding',
            'reduction', 'enrollment', 'membership', 'ownership', 'migration', 'retention',
            'expiration', 'activation', 'deactivation', 'validation', 'verification',
            'settings', 'category', 'firmware', 'software', 'hardware', 'platform',
            'defender', 'endpoint', 'antivirus', 'boundary', 'baseline', 'standard',
            'advanced', 'custom', 'attribute', 'context', 'reference', 'method',
            'policy', 'profile', 'device', 'windows', 'android', 'account',
            'control', 'surface', 'catalog', 'setting', 'domain', 'email',
            'health', 'identity', 'imported', 'kiosk', 'network', 'trusted',
            'wired', 'wireless', 'status', 'response', 'exploit', 'cleanup',
            'local', 'group', 'class', 'claim', 'unit', 'application',
            'protection', 'management', 'compliance', 'information', 'permission',
            'principal', 'processing', 'provider', 'relationship', 'schedule',
            'security', 'service', 'tenant', 'transport', 'mailbox',
            'accepted', 'access', 'address', 'allowed', 'attachment',
            'blocked', 'calendar', 'client', 'collection', 'connection',
            'connector', 'content', 'delivery', 'filter', 'hosted',
            'inbound', 'instance', 'internal', 'journal', 'location',
            'manager', 'mapping', 'member', 'message', 'named',
            'outbound', 'password', 'request', 'rule', 'user',
            'strength', 'conditional', 'crosstenantaccess', 'entitlement',
            'package', 'catalog', 'resource', 'connected', 'lifecycle',
            'naming', 'external', 'definition', 'eligibility', 'temporary',
            'voice', 'authenticator', 'fido', 'software', 'authorization',
            'tenant', 'default', 'partner', 'malware', 'spam', 'connection',
            'safe', 'links', 'broadcast', 'emergency', 'routing', 'enhanced',
            'dialin', 'conferencing', 'recording', 'cortana', 'translation',
            'unassigned', 'treatment', 'upgrade', 'voicemail', 'calling',
            'mobility', 'roaming', 'feedback', 'federation', 'guest',
            'messaging', 'meeting', 'online', 'orgwide', 'workload',
            'sensitivity', 'supervisory', 'review', 'fileplan', 'authority',
            'citation', 'department', 'subcategory', 'compliance', 'retention',
            'entitlement', 'autopilot', 'hybrid', 'joined', 'limit',
            'restriction', 'enrollment', 'android', 'administrator',
            'opensource', 'enterprise', 'autoreply', 'permission', 'tips',
            'intraorganization', 'offlineaddress', 'book', 'perimeter',
            'policytip', 'config', 'recipient', 'submission', 'roleassignment',
            'shared',  'sweep', 'accountprotection', 'usergroupmembership',
            'appconfiguration', 'appprotection', 'deviceandappmanagement',
            'assignmentfilter', 'cleanup', 'openvpn', 'windows', 'macos',
            
            # 4-letter words
            'role', 'mail', 'fido', 'wifi', 'rule', 'user', 'page',
            'site', 'team', 'hold', 'park', 'call', 'chat', 'file',
            'auto', 'list', 'plan', 'irm', 'owa', 'dlp', 'case',
            'book', 'tips', 'pstn', 'sync', 'task', 'form', 'work',
            'owner', 'join', 'data', 'zone', 'note', 'link', 'path',
            
            # 3-letter words
            'cas', 'atp', 'ume', 'eop', 'ome', 'vdi', 'app',
            'api', 'url', 'sms', 'vpn', 'tab', 'set', 'log',
            'tag', 'tip', 'org', 'key', 'map', 'for', 'and',
            'the', 'ip', 'ad',
            
            # 2-letter words (rarely needed)
            'um'
        )
        
        # Sort by length descending to match longest first
        $words = $words | Sort-Object -Property Length -Descending | Select-Object -Unique
        
        # Iterative word extraction
        $result = @()
        $remaining = $friendlyName.ToLower()
        
        while ($remaining.Length -gt 0) {
            $matched = $false
            foreach ($word in $words) {
                if ($remaining.StartsWith($word)) {
                    $result += $word
                    $remaining = $remaining.Substring($word.Length)
                    $matched = $true
                    break
                }
            }
            
            if (-not $matched) {
                # No match found - take first character and continue
                $result += $remaining[0].ToString()
                $remaining = $remaining.Substring(1)
            }
        }
        
        # Join with spaces and convert to Title Case
        $friendlyName = ($result -join ' ').Trim()
        $friendlyName = $friendlyName -replace '\s+', ' '  # Collapse multiple spaces
        $friendlyName = (Get-Culture).TextInfo.ToTitleCase($friendlyName)
    } else {
        # For camelCase names, protect iOS, macOS, AD, and IP from being split
        $tempName = $Name
        $tempName = $tempName -creplace 'iOS', '~ios~'      # Case-sensitive: only 'iOS'
        $tempName = $tempName -creplace 'macOS', '~macos~'  # Case-sensitive: only 'macOS'
        $tempName = $tempName -creplace '(?<=[a-z])AD(?=[A-Z]|$)', '~ad~'    # AD between lowercase and uppercase/end
        $tempName = $tempName -creplace '(?<=[a-z])IP(?=[A-Z]|$)', '~ip~'    # IP between lowercase and uppercase/end
        $tempName = $tempName -creplace '^AD(?=[A-Z])', '~ad~'    # AD at start before uppercase
        $tempName = $tempName -creplace '^IP(?=[A-Z])', '~ip~'    # IP at start before uppercase
        
        # Insert spaces before capital letters
        $friendlyName = $tempName -creplace '([A-Z])', ' $1'
        $friendlyName = $friendlyName.Trim() -replace '\s+', ' '
        
        # Restore protected OS names and abbreviations
        $friendlyName = $friendlyName -replace '~ios~', ' iOS'
        $friendlyName = $friendlyName -replace '~macos~', ' macOS'
        $friendlyName = $friendlyName -replace '~ad~', ' AD'
        $friendlyName = $friendlyName -replace '~ip~', ' IP'
        
        # Capitalize first letter
        if ($friendlyName.Length -gt 0) {
            $friendlyName = $friendlyName.Substring(0, 1).ToUpper() + $friendlyName.Substring(1)
        }
    }
    
    # Handle common abbreviations
    $friendlyName = $friendlyName -replace '\bIos\b', 'iOS'
    $friendlyName = $friendlyName -replace '\bMac\s*Os\b', 'macOS'
    $friendlyName = $friendlyName -replace '\bApi\b', 'API'
    $friendlyName = $friendlyName -replace '\bId\b', 'ID'
    $friendlyName = $friendlyName -replace '\bVpp\b', 'VPP'
    $friendlyName = $friendlyName -replace '\bMdm\b', 'MDM'
    $friendlyName = $friendlyName -replace '\bMam\b', 'MAM'
    $friendlyName = $friendlyName -replace '\bUrl\b', 'URL'
    $friendlyName = $friendlyName -replace '\bVpn\b', 'VPN'
    $friendlyName = $friendlyName -replace '\bWifi\b', 'WiFi'
    $friendlyName = $friendlyName -replace '\bScep\b', 'SCEP'
    $friendlyName = $friendlyName -replace '\bPkcs\b', 'PKCS'
    $friendlyName = $friendlyName -replace '\bPfx\b', 'PFX'
    $friendlyName = $friendlyName -replace '\bAd\b', 'AD'
    $friendlyName = $friendlyName -replace '\bIp\b', 'IP'
    $friendlyName = $friendlyName -replace '\bAsr\b', 'ASR'
    $friendlyName = $friendlyName -replace 'Windows10', 'Windows 10'
    $friendlyName = $friendlyName -replace 'O365', 'Office 365'
    
    return $friendlyName
}

# Function to extract h2 headers from UTCM documentation pages
function Get-UTCMResourceNamesFromDocs {
    [CmdletBinding()]
    param()
    
    $utcmPages = @(
        'https://learn.microsoft.com/en-us/graph/utcm-entra-resources'
        'https://learn.microsoft.com/en-us/graph/utcm-exchange-resources'
        'https://learn.microsoft.com/en-us/graph/utcm-intune-resources'
        'https://learn.microsoft.com/en-us/graph/utcm-securityandcompliance-resources'
        'https://learn.microsoft.com/en-us/graph/utcm-teams-resources'
    )
    
    $resourceNameMap = @{}
    
    foreach ($pageUrl in $utcmPages) {
        Write-Verbose "Fetching: $pageUrl"
        try {
            $response = Invoke-WebRequest -Uri $pageUrl -UseBasicParsing -ErrorAction Stop
            $htmlContent = $response.Content
            
            # Extract h2 headers using regex
            # Actual format: <h2 id="administrativeunit-resource-type">administrativeUnit resource type</h2>
            # We need to capture the camelCase name from the text content
            $h2Pattern = '<h2[^>]*id="([^"]+)-resource-type"[^>]*>(\w+)\s+resource\s+type</h2>'
            $activeMatches = [regex]::Matches($htmlContent, $h2Pattern)
            
            foreach ($match in $activeMatches) {
                $lowercaseKey = $match.Groups[1].Value    # from id (e.g., "administrativeunit")  
                $camelCaseName = $match.Groups[2].Value  # from text (e.g., "administrativeUnit")
                
                # Store mapping: lowercase -> camelCase
                if (-not $resourceNameMap.ContainsKey($lowercaseKey)) {
                    $resourceNameMap[$lowercaseKey] = $camelCaseName
                    Write-Verbose "  Found: $lowercaseKey -> $camelCaseName"
                }
            }
            
            Write-Verbose "  Extracted $($activeMatches.Count) resource types from page"
        } catch {
            Write-Warning "Failed to fetch $pageUrl : $_"
        }
    }
    
    Write-Host "✓ Extracted $($resourceNameMap.Count) camelCase resource names from documentation" -ForegroundColor Green
    return $resourceNameMap
}

# Main execution
Write-Host "==================================================" -ForegroundColor Yellow
Write-Host "UTCM Resource Type Extractor - Starting" -ForegroundColor Yellow
Write-Host "==================================================" -ForegroundColor Yellow
Write-Host ""

# Step 1: Fetch camelCase names from documentation pages
Write-Host "Step 1: Fetching resource names from Microsoft Learn documentation pages..." -ForegroundColor Cyan
$camelCaseMap = Get-UTCMResourceNamesFromDocs
Write-Host ""

# Step 1.5: Manual overrides for resources not found in docs or with poor word splitting
$manualOverrides = @{
    'globaladdresslist' = 'Global Address List'
    'managementroleentry' = 'Management Role Entry'
    'addressbookpolicy' = 'Address Book Policy'
    'addresslist' = 'Address List'
    'clientaccessrule' = 'Client Access Rule'
    'eopprotectionpolicyrule' = 'EOP Protection Policy Rule'
    'externalinoutlook' = 'External In Outlook'
    'offlineaddressbook' = 'Offline Address Book'
    'sweeprule' = 'Sweep Rule'
    'groupsnamingpolicy' = 'Groups Naming Policy'
    'groupssettings' = 'Groups Settings'
    'channel' = 'Channel'
    'channeltab' = 'Channel Tab'
    'orgwideappsettings' = 'Org Wide App Settings'
    'accountprotectionlocaladministratorpasswordsolutionpolicy' = 'Account Protection Local Administrator Password Solution Policy'
    'asrrulespolicywindows10' = 'ASR Rules Policy Windows 10'
    'auditconfigurationpolicy' = 'Audit Configuration Policy'
    'autosensitivitylabelrule' = 'Auto Sensitivity Label Rule'
    'dlpcompliancerule' = 'DLP Compliance Rule'
    'sensitivitylabel' = 'Sensitivity Label'
    'place' = 'Place'
}

# Step 2: Fetch the JSON schema
Write-Host "Step 2: Fetching JSON schema from: $SchemaUrl" -ForegroundColor Cyan

try {
    # Fetch the JSON schema
    $schema = Invoke-RestMethod -Uri $SchemaUrl -ErrorAction Stop
    Write-Host "✓ Schema fetched successfully" -ForegroundColor Green
    
    # Ensure output directory exists before saving schema
    if (-not (Test-Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    }
    
    # Export the raw schema to file for reference
    $schemaOutputPath = Join-Path $OutputPath "utcm-monitor-schema.json"
    $schema | ConvertTo-Json -Depth 100 | Set-Content -Path $schemaOutputPath -Encoding UTF8
    Write-Host "✓ Raw schema exported to: $schemaOutputPath" -ForegroundColor Green
    
    # Extract resource types from $defs
    $resourceDefinitions = $schema.'$defs'
    
    if (-not $resourceDefinitions) {
        throw "Schema does not contain `$defs section"
    }
    
    $resourceTypeNames = $resourceDefinitions.PSObject.Properties.Name
    Write-Host "✓ Found $($resourceTypeNames.Count) resource type definitions" -ForegroundColor Green
    Write-Host ""
    
    # Process each resource type
    $allResources = @()
    $summary = @{}
    
    foreach ($prefixedName in $resourceTypeNames) {
        # Parse the prefixed name: microsoft.<type>.<resourcename>
        $parts = $prefixedName -split '\.'
        
        if ($parts.Count -ne 3 -or $parts[0] -ne 'microsoft') {
            Write-Warning "Skipping invalid resource name format: $prefixedName"
            continue
        }
        
        $resourceType = $parts[1]
        $originalName = $parts[2]
        
        # Determine friendly name with priority: manual override > camelCase from docs > word splitting
        $lowercaseKey = $originalName.ToLower()
        $friendlyName = $null
        
        # Priority 1: Check manual overrides first
        if ($manualOverrides.ContainsKey($lowercaseKey)) {
            $friendlyName = $manualOverrides[$lowercaseKey]
            Write-Verbose "Using manual override: $lowercaseKey → $friendlyName"
        }
        # Priority 2: Use camelCase from documentation
        elseif ($camelCaseMap.ContainsKey($lowercaseKey)) {
            $camelCaseName = $camelCaseMap[$lowercaseKey]
            $friendlyName = ConvertTo-FriendlyName -Name $camelCaseName
            Write-Verbose "Using camelCase from docs: $lowercaseKey → $camelCaseName → $friendlyName"
        }
        # Priority 3: Fall back to word splitting on original name
        else {
            $friendlyName = ConvertTo-FriendlyName -Name $originalName
            Write-Verbose "Using word splitting: $originalName → $friendlyName"
        }
        
        # Get the resource definition for description
        $resourceDef = $resourceDefinitions.$prefixedName
        $description = $resourceDef.description
        
        # Create resource object
        $resourceObject = [PSCustomObject]@{
            PrefixedName = $prefixedName
            ResourceType = $resourceType
            OriginalName = $originalName
            FriendlyName = $friendlyName
            Description  = $description
            LastUpdated  = (Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ")
        }
        
        $allResources += $resourceObject
        
        # Update summary
        if (-not $summary.ContainsKey($resourceType)) {
            $summary[$resourceType] = 0
        }
        $summary[$resourceType]++
    }
    
    # Sort resources
    $allResources = $allResources | Sort-Object PrefixedName
    
    Write-Host "==================================================" -ForegroundColor Yellow
    Write-Host "Summary by Resource Type" -ForegroundColor Yellow
    Write-Host "==================================================" -ForegroundColor Yellow
    
    foreach ($type in ($summary.Keys | Sort-Object)) {
        Write-Host "$type : $($summary[$type]) resources" -ForegroundColor Cyan
    }
    
    Write-Host ""
    Write-Host "Total resources: $($allResources.Count)" -ForegroundColor Green
    Write-Host ""
    
    # Ensure output directory exists
    if (-not (Test-Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        Write-Host "Created output directory: $OutputPath" -ForegroundColor Gray
    }
    
    # Export to JSON
    $jsonData = @{
        GeneratedDate  = (Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ")
        SchemaSource   = $SchemaUrl
        TotalResources = $allResources.Count
        Summary        = $summary
        Resources      = $allResources
    }
    
    $jsonPath = Join-Path $OutputPath $JsonFileName
    $jsonData | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath -Encoding UTF8 -Force
    Write-Host "✓ JSON exported to: $jsonPath" -ForegroundColor Green
    
    # Export to CSV
    $csvPath = Join-Path $OutputPath $CsvFileName
    $allResources | Select-Object PrefixedName, ResourceType, OriginalName, FriendlyName, Description, LastUpdated |
        Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
    Write-Host "✓ CSV exported to: $csvPath" -ForegroundColor Green
    
    Write-Host ""
    Write-Host "==================================================" -ForegroundColor Yellow
    Write-Host "Sample Resources (first 10)" -ForegroundColor Yellow
    Write-Host "==================================================" -ForegroundColor Yellow
    
    $allResources | Select-Object -First 10 | ForEach-Object {
        Write-Host ""
        Write-Host "$($_.PrefixedName)" -ForegroundColor White
        Write-Host "  Original Name: $($_.OriginalName)" -ForegroundColor Gray
        Write-Host "  Friendly Name: $($_.FriendlyName)" -ForegroundColor Cyan
        Write-Host "  Type: $($_.ResourceType)" -ForegroundColor Gray
        if ($_.Description) {
            $desc = if ($_.Description.Length -gt 100) { 
                $_.Description.Substring(0, 97) + "..." 
            } else { 
                $_.Description 
            }
            Write-Host "  Description: $desc" -ForegroundColor DarkGray
        }
    }
    
    Write-Host ""
    Write-Host "==================================================" -ForegroundColor Yellow
    Write-Host "Script completed successfully!" -ForegroundColor Green
    Write-Host "==================================================" -ForegroundColor Yellow
}
catch {
    Write-Error "Failed to process schema: $_"
    Write-Error $_.Exception.Message
    exit 1
}
