###################################################################
#                                                                 #
# Made by SiebeScript™                                            #
#                                                                 #
# This script makes an Excel export of all M365 licenses per user #                                                                 
# Version 3 - 10/08/2023                                          #
#                                                                 #
###################################################################

# Check installed modules
$modules = "MSOnline", "ImportExcel"
$installed = @((Get-Module $modules -ListAvailable).Name | Select-Object -Unique)
$notInstalled = Compare-Object $modules $installed -PassThru


if ($notInstalled) {
    Write-Host "The following modules aren't installed: `n
        $notInstalled `n
        Install them from an elevated Powershell window first!"  -ForegroundColor Red
    Read-Host "Press Enter to quit..."
    Exit
}


# Connect M365 and write licenses to CSV
Connect-MsolService

$SKUFriendlyNames = @{
    "FLOW_PER_USER" = "Power Automate per user plan"
    "POWER_BI_PRO" = "Power BI Pro"
    "DYN365_TEAM_MEMBERS" = "Dynamics 365 Team Members"
    "DYN365_REGULATORY_SERVICE" = "Dynamics 365 Regulatory Service - Enterprise Edition Trial"
    "WINDOWS_STORE" = "Windows Store for Business"
    "Dynamics_365_for_Operations_Sandbox_Tier2_SKU" = "Dynamics 365 Operations - Sandbox Tier 2:Standard Acceptance Testing"
    "PROJECTPREMIUM" = "Project Online Premium"
    "DYN365_PROJECT_OPERATIONS_ATTACH" = "Dynamics 365 Operations - Attach"
    "ENTERPRISEPACK" = "Office 365 E3"
    "FLOW_FREE" = "Microsoft Power Automate Free"
    "CCIBOTS_PRIVPREV_VIRAL" = "Power Virtual Agents Viral Trial"
    "SPB" = "Microsoft 365 Business Premium"
    "POWERAPPS_VIRAL" = "Power Apps Plan 2 Trial"
    "EXCHANGESTANDARD" = "Exchange Online (Plan 1)"
    "DYN365_SCM" = "Dynamics 365 for Supply Chain Management"
    "O365_BUSINESS_PREMIUM" = "Microsoft 365 Business Standard"
    "POWER_BI_STANDARD" = "Power BI (free)"
    "DYN365_FINANCE_ATTACH" = "Dynamics 365 Finance - Attach"
    "Dynamics_365_for_Operations_Devices" = "Dynamics 365 Operations - Device"
    "Dyn365_Operations_Activity" = "Dynamics 365 Operations - Device"
    "Power_Pages_vTrial_for_Makers" = "Power Pages vTrial for Makers"
    "TEAMS_EXPLORATORY" = "Microsoft Teams Exploratory"
    "MCOMEETADV" = "Microsoft 365 Audio Conferencing"
    "M365_F1_COMM" = "Microsoft 365 F1"
    "AAD_PREMIUM" = "Azure AD Premium P1"
    "ATP_ENTERPRISE" = "Microsoft 365 Defender (Plan 1)"
    "SPE_E3" = "Microsoft 365 E3"
    "PROJECTPROFESSIONAL" = "Project Plan 3"
    "POWERAPPS_DEV" = "Microsoft PowerApps for Developer"
}

$AllLicensePlans = Get-MsolAccountSku | Where-Object ActiveUnits -gt 0 | Select-Object SkuPartNumber, ActiveUnits, ConsumedUnits | Sort-Object ConsumedUnits | ForEach-Object { $_.SkuPartNumber = $SKUFriendlyNames.Item($_.SkuPartNumber)
$_ } | Export-Csv -Path ".\Overview.csv" -Encoding UTF8 -NoTypeInformation
$ActiveLicensePlans = Get-MsolAccountSku | Where-Object ConsumedUnits -gt 0

foreach ($LicensePlan in $ActiveLicensePlans) {
    $CSVPath = $SKUFriendlyNames.Item($LicensePlan.SkuPartNumber) + ".csv"    
    Write-Host "Exporting $CSVPath" -ForegroundColor Green
    Get-MsolUser -All | Where-Object {($_.licenses).AccountSkuId -match $LicensePlan.AccountSkuId } | Select-Object DisplayName, UserPrincipalName | Sort-Object DisplayName | Export-Csv -Path $CSVPath -Encoding UTF8 -NoTypeInformation
}


# Import CSV's to Excel
$ExportExcelPath = "LicenseOverview.xlsx"
if (Test-Path $ExportExcelPath) { Remove-Item $ExportExcelPath -Force }

$CSVs = Get-ChildItem *.csv
foreach ($csv in $CSVs) {
    Write-Host "Importing $csv" -ForegroundColor Green
    Import-Csv -Path $csv.FullName | Export-Excel -Path $ExportExcelPath -WorksheetName $csv.BaseName
    Remove-Item $csv
} 

Write-Host "Overview exported to LicenseOverview.xlsx" -ForegroundColor Green
Read-Host "Press Enter to quit..."
