# microsoft-id-resolver

Federate GitHub to az cli
https://github.com/marketplace/actions/azure-ad-workload-identity-federation

This repo is a standalone fork of merill's microsoft-info with some minor tweaks.

## License SKUs and prices

`src/Export-LicenseSkus.ps1` exports every Microsoft license SKU with its display name, service plans and retail price:

| File | Contents |
|---|---|
| `_info/LicenseSkus.csv` / `.json` | SkuId, SkuPartNumber, DisplayName, MonthToMonthPrice, AnnualCommitmentMonthlyPrice, Currency, Market, PriceSource, PartnerCenterOfferId |
| `_info/LicenseServicePlans.csv` / `.json` | SkuId, SkuPartNumber, ServicePlanId, ServicePlanName, ServicePlanDisplayName |

- Names and service plans come from Microsoft's [Product names and service plan identifiers for licensing](https://learn.microsoft.com/entra/identity/users/licensing-service-plan-reference) reference.
- Prices are the Estimated Retail Price (ERP) from the Partner Center license-based price sheet, Greece (`GR`) market in `EUR`, excluding VAT. Partner unit prices are never exported.
- `MonthToMonthPrice` is the monthly term price (`P1M`). `AnnualCommitmentMonthlyPrice` is the annual term price (`P1Y`) divided by 12.
- Partner Center offers are matched to SKUs by name. Unmatched or ambiguous SKUs (run with `-Verbose` to list them) and prices not available in Partner Center can be set in `customdata/LicensePriceOverrides.csv`.
- If Partner Center is unavailable, the previously exported prices are kept.

Run locally with a downloaded price sheet (requires PowerShell 7+, run from the repository root):

```powershell
./src/Export-LicenseSkus.ps1 -PriceSheetPath ./pricesheet.csv
```

### Partner Center setup

The price sheet API only supports delegated (App+User) tokens from a user with the Admin Agent or Sales Agent role.

1. Register an app in the CSP partner tenant with the delegated `Microsoft Partner Center` → `user_impersonation` permission, a client secret and a **Web** redirect URI `http://localhost:8400/`. Grant admin consent.
2. Get the refresh token once, signing in as an Admin Agent or Sales Agent user (PowerShell 7+):

   ```powershell
   # Copies the refresh token to the clipboard
   ./src/New-PartnerCenterRefreshToken.ps1 -TenantId <tenant-id> -ClientId <client-id>

   # Or writes it straight to the repository secret with the GitHub CLI
   ./src/New-PartnerCenterRefreshToken.ps1 -TenantId <tenant-id> -ClientId <client-id> -GitHubRepository <owner>/<repo>
   ```

3. Configure the repository:
   - Variables: `PARTNER_CENTER_TENANT_ID`, `PARTNER_CENTER_CLIENT_ID`
   - Secrets: `PARTNER_CENTER_CLIENT_SECRET`, `PARTNER_CENTER_REFRESH_TOKEN`

The refresh token is renewed manually by running `New-PartnerCenterRefreshToken.ps1` again. When it expires the workflow shows a warning and keeps the previous prices until `PARTNER_CENTER_REFRESH_TOKEN` is replaced.
