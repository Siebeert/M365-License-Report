###################################################################
#                                                                 #
# Made by SiebeScript™                                            #
#                                                                 #
# This script creates an overview of all users with their         #
# licenses, job title, manager and location.                      # 
#                                                                 #                                                                
# Version 5 - 31/10/2023                                          #
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


# Connect M365 and get licenses
Connect-MsolService
Connect-ExchangeOnline

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

$AllLicensePlans = Get-MsolAccountSku | Where-Object ActiveUnits -gt 0 | Select-Object SkuPartNumber, ActiveUnits, ConsumedUnits | Sort-Object ConsumedUnits -Descending | ForEach-Object { 
    $_.SkuPartNumber = $SKUFriendlyNames.Item($_.SkuPartNumber)
    $_ 
}
$ActiveLicensePlans = Get-MsolAccountSku | Where-Object ConsumedUnits -gt 0


# Get users and write license information to CSV
[System.Collections.Generic.List[System.Object]]$Users = Get-User | Select-Object DisplayName, UserPrincipalName, Title, Manager, CountryOrRegion | Sort-Object DisplayName
$MSOLUsers = Get-MsolUser -All | Select-Object UserPrincipalName, Licenses
$usersToRemove = New-Object System.Collections.Generic.List[System.Object]

Write-Host "Filtering out users without licenses" -ForeGroundColor Yellow
foreach ($user in $users) {
    $AssignedLicenses = $MSOLUsers | Where-Object UserPrincipalName -EQ $user.UserPrincipalName | Select-Object Licenses
    if ($AssignedLicenses.Licenses.Count -eq 0) {
        $usersToRemove.Add($user)
    }
}

foreach ($user in $usersToRemove) {
    $users.Remove($user) | Out-Null
}

Write-Host "Adding LastLogonTime to users" -ForeGroundColor Yellow
foreach ($User in $Users) {
    try {
        $lastLogonTime = $user | Select-Object -ExpandProperty userprincipalname | Get-MailboxStatistics -ErrorAction Stop | Select-Object LastLogonTime
        $user | Add-Member -MemberType NoteProperty -Name LastLogonTime -Value $lastLogonTime.LastLogonTime
    } catch {
        $user | Add-Member -MemberType NoteProperty -Name LastLogonTime -Value ""
    }
}

foreach ($LicensePlan in $ActiveLicensePlans) {
    foreach ($user in $Users) {
        try {
            $AssignedLicenses = $MSOLUsers | Where-Object UserPrincipalName -EQ $user.UserPrincipalName | Select-Object Licenses
            if ($LicensePlan.AccountSkuId -in $AssignedLicenses.Licenses.AccountSkuId) {
                $value = "Yes"
            } else {
                $value = "No"
            }
            
            $userName = $user.DisplayName
            $licenseName = $SKUFriendlyNames.Item($LicensePlan.SkuPartNumber)
            $user | Add-Member -MemberType NoteProperty -Name $licenseName -Value $value
            
            Write-Host "License information for $licenseName added to user $userName" -ForeGroundColor Green
        } catch {
            Write-Host "Error: user not found" -ForeGroundColor Red
        }
    }
}

$users | Export-Csv -Path ".\LicenseOverview.csv" -Encoding UTF8 -NoTypeInformation


# Convert CSV to Excel for User Overview
$ExportExcelPath = "LicenseOverview.xlsx"
if (Test-Path $ExportExcelPath) { Remove-Item $ExportExcelPath -Force }

Import-Csv -Path LicenseOverview.csv | Export-Excel -Path $ExportExcelPath -WorksheetName "User Overview"
Remove-Item "LicenseOverview.csv"


Write-Host "Overview exported to UserOverview.xlsx" -ForegroundColor Green
Read-Host "Press Enter to quit..."