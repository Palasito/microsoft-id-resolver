#Requires -Version 7.0
<#
.SYNOPSIS
    Exports all Microsoft license SKUs with their human readable names, service plans and retail prices.

.DESCRIPTION
    SKU identifiers, display names and service plans are read from the official Microsoft licensing reference
    ("Product names and service plan identifiers for licensing"), which lists every commercial SKU.

    Prices are read from the Partner Center license-based price sheet for the requested market. Only the
    Estimated Retail Price (ERP) is exported. The partner UnitPrice is never written to the output or logs
    because this repository is public.

    Partner Center offers are matched to SKU GUIDs by normalized product name. Mismatches and SKUs that
    Partner Center does not price can be fixed in ./customdata/LicensePriceOverrides.csv.

    When neither a Partner Center access token nor a price sheet file is supplied, prices from the previous
    export are carried forward so an expired token never wipes committed prices.

.PARAMETER PartnerCenterAccessToken
    Delegated (App+User) access token for https://api.partner.microsoft.com, issued to a user with the
    Admin Agent or Sales Agent role.

.PARAMETER PriceSheetPath
    Path to a license-based price sheet CSV that was already downloaded from Partner Center.
    Takes precedence over PartnerCenterAccessToken.

.EXAMPLE
    ./src/Export-LicenseSkus.ps1 -PartnerCenterAccessToken $token

    Exports SKUs, service plans and Greek EUR retail prices to .\_info.

.EXAMPLE
    ./src/Export-LicenseSkus.ps1 -PriceSheetPath ./pricesheet.csv -Market GR -Currency EUR

    Uses a price sheet that was manually downloaded from Partner Center.

.EXAMPLE
    ./src/Export-LicenseSkus.ps1

    Refreshes SKU names and service plans while keeping the previously exported prices.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Path to output the csv and json files")]
    [string]$OutputPath = ".\_info",

    [Parameter(Mandatory = $false, HelpMessage = "Path to csv file with price and offer mapping overrides")]
    [string]$OverridesPath = "./customdata/LicensePriceOverrides.csv",

    [Parameter(Mandatory = $false)]
    [string]$LicensingCsvUrl = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv",

    [Parameter(Mandatory = $false, HelpMessage = "Two letter Partner Center market code")]
    [string]$Market = "GR",

    [Parameter(Mandatory = $false)]
    [string]$Currency = "EUR",

    [Parameter(Mandatory = $false)]
    [securestring]$PartnerCenterAccessToken,

    [Parameter(Mandatory = $false)]
    [string]$PriceSheetPath
)

$invariantCulture = [System.Globalization.CultureInfo]::InvariantCulture

function ConvertTo-NormalizedProductName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Name
    )

    $normalized = $Name.ToLowerInvariant() -replace '[\u2122\u00AE\u00A9]', '' -replace '[^a-z0-9]+', ' '
    return $normalized.Trim()
}

function ConvertTo-Price {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Value
    )

    $price = [decimal]0
    if ([decimal]::TryParse($Value, [System.Globalization.NumberStyles]::Number, $invariantCulture, [ref]$price)) {
        return $price
    }
    return $null
}

function Get-LicensingReference {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri
    )

    $tempFile = New-TemporaryFile
    try {
        Invoke-WebRequest -Uri $Uri -OutFile $tempFile -UseBasicParsing -ErrorAction Stop
        $rows = Import-Csv -Path $tempFile -Encoding UTF8
    }
    finally {
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
    }

    if (-not $rows -or -not ($rows[0].PSObject.Properties.Name -contains 'GUID')) {
        throw "Licensing reference at $Uri did not contain the expected columns"
    }

    $skus = $rows | Group-Object { $_.GUID.Trim().ToLowerInvariant() } | ForEach-Object {
        $first = $_.Group[0]
        [PSCustomObject]@{
            SkuId         = $_.Name
            SkuPartNumber = $first.String_Id.Trim()
            DisplayName   = $first.Product_Display_Name.Trim()
        }
    }

    $servicePlans = $rows | ForEach-Object {
        [PSCustomObject]@{
            SkuId                  = $_.GUID.Trim().ToLowerInvariant()
            SkuPartNumber          = $_.String_Id.Trim()
            ServicePlanId          = $_.Service_Plan_Id.Trim().ToLowerInvariant()
            ServicePlanName        = $_.Service_Plan_Name.Trim()
            ServicePlanDisplayName = $_.Service_Plans_Included_Friendly_Names.Trim()
        }
    } | Sort-Object SkuId, ServicePlanId -Unique

    return [PSCustomObject]@{
        Skus         = @($skus)
        ServicePlans = @($servicePlans)
    }
}

function Save-PartnerCenterPriceSheet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [securestring]$AccessToken,

        [Parameter(Mandatory = $true)]
        [string]$Market,

        [Parameter(Mandatory = $true)]
        [string]$Destination
    )

    $uri = "https://api.partner.microsoft.com/v1.0/sales/pricesheets(Market='$Market',PricesheetView='updatedlicensebased')/`$value"
    Write-Verbose "Downloading price sheet: $uri"
    Invoke-WebRequest -Uri $uri -Authentication Bearer -Token $AccessToken -OutFile $Destination -ErrorAction Stop
}

# Reduces the price sheet to current ERP prices per offer, keyed by "ProductId:SkuId".
# UnitPrice (partner cost) is intentionally never read.
function Get-RetailPricesFromPriceSheet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Market,

        [Parameter(Mandatory = $true)]
        [string]$Currency
    )

    $now = [datetime]::UtcNow
    $offers = @{}

    Import-Csv -Path $Path -Encoding UTF8 | Where-Object {
        $_.Market -eq $Market -and
        $_.Currency -eq $Currency -and
        $_.Tags -match 'License' -and
        $_.Tags -notmatch 'Trial' -and
        (-not $_.PSObject.Properties['Segment'] -or -not $_.Segment -or $_.Segment -eq 'Commercial')
    } | ForEach-Object {
        $start = [datetime]::MinValue
        $end = [datetime]::MaxValue
        if ($_.EffectiveStartDate) { [void][datetime]::TryParse($_.EffectiveStartDate, $invariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal, [ref]$start) }
        if ($_.EffectiveEndDate) { [void][datetime]::TryParse($_.EffectiveEndDate, $invariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal, [ref]$end) }
        if ($start -gt $now -or $end -le $now) { return }

        $erp = ConvertTo-Price -Value $_.'ERP Price'
        if ($null -eq $erp) { return }

        $offerId = "$($_.ProductId):$($_.SkuId)"
        if (-not $offers.ContainsKey($offerId)) {
            $offers[$offerId] = [PSCustomObject]@{
                OfferId                      = $offerId
                ProductTitle                 = $_.ProductTitle
                SkuTitle                     = $_.SkuTitle
                MonthToMonthPrice            = $null
                MonthToMonthStart            = [datetime]::MinValue
                AnnualCommitmentMonthlyPrice = $null
                AnnualCommitmentStart        = [datetime]::MinValue
            }
        }
        $offer = $offers[$offerId]

        # Latest effective row wins when a price changes mid-month.
        if ($_.TermDuration -eq 'P1M' -and $_.BillingPlan -eq 'Monthly' -and $start -ge $offer.MonthToMonthStart) {
            $offer.MonthToMonthPrice = $erp
            $offer.MonthToMonthStart = $start
        }
        elseif ($_.TermDuration -eq 'P1Y' -and $_.BillingPlan -eq 'Annual' -and $start -ge $offer.AnnualCommitmentStart) {
            $offer.AnnualCommitmentMonthlyPrice = [math]::Round($erp / 12, 2)
            $offer.AnnualCommitmentStart = $start
        }
    }

    return $offers
}

function Get-PreviousPrices {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $previous = @{}
    if (Test-Path $Path) {
        Get-Content -Path $Path -Raw -Encoding UTF8 | ConvertFrom-Json | ForEach-Object {
            $previous[$_.SkuId] = $_
        }
    }
    return $previous
}

Write-Host "==================================================" -ForegroundColor Yellow
Write-Host "License SKU Exporter - Starting" -ForegroundColor Yellow
Write-Host "==================================================" -ForegroundColor Yellow

$skuJsonPath = Join-Path $OutputPath "LicenseSkus.json"
$skuCsvPath = Join-Path $OutputPath "LicenseSkus.csv"
$planJsonPath = Join-Path $OutputPath "LicenseServicePlans.json"
$planCsvPath = Join-Path $OutputPath "LicenseServicePlans.csv"
$downloadedPriceSheet = $null

try {
    Write-Host "Step 1: Fetching licensing reference from Microsoft" -ForegroundColor Cyan
    $reference = Get-LicensingReference -Uri $LicensingCsvUrl
    Write-Host "[OK] Found $($reference.Skus.Count) SKUs and $($reference.ServicePlans.Count) service plan assignments" -ForegroundColor Green

    Write-Host "Step 2: Resolving retail prices ($Market / $Currency)" -ForegroundColor Cyan
    $offers = $null
    if ($PriceSheetPath) {
        $offers = Get-RetailPricesFromPriceSheet -Path $PriceSheetPath -Market $Market -Currency $Currency
    }
    elseif ($PartnerCenterAccessToken) {
        $downloadedPriceSheet = New-TemporaryFile
        Save-PartnerCenterPriceSheet -AccessToken $PartnerCenterAccessToken -Market $Market -Destination $downloadedPriceSheet
        $offers = Get-RetailPricesFromPriceSheet -Path $downloadedPriceSheet -Market $Market -Currency $Currency
    }

    if ($null -ne $offers) {
        Write-Host "[OK] Found $($offers.Count) current $Market offers with retail prices" -ForegroundColor Green
    }
    else {
        Write-Warning "No Partner Center price sheet available, carrying forward previously exported prices"
    }

    $overrides = @{}
    if (Test-Path $OverridesPath) {
        Import-Csv -Path $OverridesPath | ForEach-Object { $overrides[$_.SkuId.Trim().ToLowerInvariant()] = $_ }
    }

    # SKU titles are specific, product titles can be shared by several SKUs, so they are only a fallback.
    $offersBySkuTitle = @{}
    $offersByProductTitle = @{}
    if ($null -ne $offers) {
        foreach ($offer in $offers.Values) {
            foreach ($lookup in @(@{ Index = $offersBySkuTitle; Title = $offer.SkuTitle }, @{ Index = $offersByProductTitle; Title = $offer.ProductTitle })) {
                if (-not $lookup.Title) { continue }
                $key = ConvertTo-NormalizedProductName -Name $lookup.Title
                if (-not $lookup.Index.ContainsKey($key)) { $lookup.Index[$key] = [System.Collections.Generic.List[object]]::new() }
                $lookup.Index[$key].Add($offer)
            }
        }
    }

    $previousPrices = Get-PreviousPrices -Path $skuJsonPath
    $matchedOfferIds = [System.Collections.Generic.HashSet[string]]::new()
    $ambiguous = @()

    $licenseSkus = foreach ($sku in $reference.Skus) {
        $override = $overrides[$sku.SkuId]
        $offer = $null
        $priceSource = $null
        $monthToMonth = $null
        $annualMonthly = $null

        if ($null -ne $offers) {
            if ($override -and $override.PartnerCenterOfferId) {
                $offer = $offers[$override.PartnerCenterOfferId.Trim()]
                if (-not $offer) { Write-Warning "Override offer $($override.PartnerCenterOfferId) for $($sku.SkuPartNumber) not found in price sheet" }
            }
            else {
                $nameKey = ConvertTo-NormalizedProductName -Name $sku.DisplayName
                $candidates = $offersBySkuTitle[$nameKey]
                if ($candidates.Count -eq 0) { $candidates = $offersByProductTitle[$nameKey] }
                if ($candidates.Count -eq 1) {
                    $offer = $candidates[0]
                }
                elseif ($candidates.Count -gt 1) {
                    $ambiguous += "$($sku.SkuPartNumber) -> $(($candidates.OfferId) -join ', ')"
                }
            }

            if ($offer) {
                [void]$matchedOfferIds.Add($offer.OfferId)
                $monthToMonth = $offer.MonthToMonthPrice
                $annualMonthly = $offer.AnnualCommitmentMonthlyPrice
                $priceSource = "PartnerCenterERP"
            }
        }
        elseif ($previousPrices.ContainsKey($sku.SkuId)) {
            $previous = $previousPrices[$sku.SkuId]
            $monthToMonth = $previous.MonthToMonthPrice
            $annualMonthly = $previous.AnnualCommitmentMonthlyPrice
            $priceSource = $previous.PriceSource
            if ($previous.PartnerCenterOfferId) { $offer = [PSCustomObject]@{ OfferId = $previous.PartnerCenterOfferId } }
        }

        if ($override -and ($override.MonthToMonthPrice -or $override.AnnualCommitmentMonthlyPrice)) {
            if ($override.MonthToMonthPrice) { $monthToMonth = ConvertTo-Price -Value $override.MonthToMonthPrice }
            if ($override.AnnualCommitmentMonthlyPrice) { $annualMonthly = ConvertTo-Price -Value $override.AnnualCommitmentMonthlyPrice }
            $priceSource = "Override"
        }

        $hasPrice = $null -ne $monthToMonth -or $null -ne $annualMonthly
        [PSCustomObject][ordered]@{
            SkuId                        = $sku.SkuId
            SkuPartNumber                = $sku.SkuPartNumber
            DisplayName                  = $sku.DisplayName
            MonthToMonthPrice            = $monthToMonth
            AnnualCommitmentMonthlyPrice = $annualMonthly
            Currency                     = if ($hasPrice) { $Currency } else { $null }
            Market                       = if ($hasPrice) { $Market } else { $null }
            PriceSource                  = if ($hasPrice) { $priceSource } else { $null }
            PartnerCenterOfferId         = if ($offer) { $offer.OfferId } else { $null }
        }
    }

    $licenseSkus = @($licenseSkus | Sort-Object SkuPartNumber, SkuId)
    $pricedCount = @($licenseSkus | Where-Object { $_.PriceSource }).Count
    Write-Host "[OK] Priced $pricedCount of $($licenseSkus.Count) SKUs" -ForegroundColor Green

    if ($null -ne $offers) {
        $unmatchedOffers = @($offers.Values | Where-Object { -not $matchedOfferIds.Contains($_.OfferId) } | Sort-Object SkuTitle)
        Write-Host "  Partner Center offers without a SKU match: $($unmatchedOffers.Count) (map them in $OverridesPath)" -ForegroundColor Gray
        $unmatchedOffers | ForEach-Object { Write-Verbose "  Unmatched offer: $($_.OfferId) $($_.SkuTitle)" }
        if ($ambiguous.Count -gt 0) {
            Write-Warning "$($ambiguous.Count) SKUs match multiple offers and were left unpriced (map them in $OverridesPath)"
            $ambiguous | ForEach-Object { Write-Verbose "  Ambiguous: $_" }
        }
    }

    Write-Host "Step 3: Exporting to csv and json" -ForegroundColor Cyan
    New-Item -ItemType Directory -Force -Path $OutputPath | Out-Null

    $servicePlans = @($reference.ServicePlans | Sort-Object SkuPartNumber, ServicePlanName, ServicePlanId)

    ConvertTo-Json -InputObject $licenseSkus -Depth 3 | Out-File -FilePath $skuJsonPath -Encoding UTF8 -Force
    # Export-Csv formats decimals with the current culture, so prices are written invariantly for consumers.
    $licenseSkus | Select-Object -Property SkuId, SkuPartNumber, DisplayName,
        @{ Name = 'MonthToMonthPrice'; Expression = { if ($null -ne $_.MonthToMonthPrice) { ([decimal]$_.MonthToMonthPrice).ToString('0.00', $invariantCulture) } } },
        @{ Name = 'AnnualCommitmentMonthlyPrice'; Expression = { if ($null -ne $_.AnnualCommitmentMonthlyPrice) { ([decimal]$_.AnnualCommitmentMonthlyPrice).ToString('0.00', $invariantCulture) } } },
        Currency, Market, PriceSource, PartnerCenterOfferId |
        Export-Csv -Path $skuCsvPath -NoTypeInformation -Encoding UTF8 -Force
    ConvertTo-Json -InputObject $servicePlans -Depth 3 | Out-File -FilePath $planJsonPath -Encoding UTF8 -Force
    $servicePlans | Export-Csv -Path $planCsvPath -NoTypeInformation -Encoding UTF8 -Force

    Write-Host "[OK] Exported $skuCsvPath, $skuJsonPath" -ForegroundColor Green
    Write-Host "[OK] Exported $planCsvPath, $planJsonPath" -ForegroundColor Green
}
catch {
    Write-Error "Failed to export license SKUs: $_"
    exit 1
}
finally {
    if ($downloadedPriceSheet) { Remove-Item $downloadedPriceSheet -Force -ErrorAction SilentlyContinue }
}
