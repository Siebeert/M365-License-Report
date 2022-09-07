###################################################################
#                                                                 #
# Made by SiebeScript™                                            #
#                                                                 #
# This script makes an Excel export of all M365 licenses per user #
#                                                                 #
###################################################################

Connect-MsolService

$SKUFriendlyNames = @{
    "FLOW_PER_USER" = "Power Automate per user plan"
    "POWER_BI_PRO" = "Power BI Pro"
    "WINDOWS_STORE" = "WINDOWS_STORE"
    "ENTERPRISEPACK" = "Office 365 E3"
    "FLOW_FREE" = "Microsoft Power Automate Free"
    "SPB" = "M365 Business Premium"
    "POWERAPPS_VIRAL" = "Power Apps Plan 2 Trial"
    "EXCHANGESTANDARD" = "Exchange Online (Plan 1)"
    "POWER_BI_STANDARD" = "Power BI (free)"
    "TEAMS_EXPLORATORY" = "Microsoft Teams Exploratory"
    "MCOMEETADV" = "M365 Audio Conferencing"
    "M365_F1_COMM" = "M365 F1"
    "AAD_PREMIUM" = "Azure AD Premium P1"
    "ATP_ENTERPRISE" = "M365 Defender (Plan 1)"
}

$AllLicensePlans = Get-MsolAccountSku | Where-Object ActiveUnits -gt 0 | Select-Object SkuPartNumber, ActiveUnits, ConsumedUnits | Sort-Object ConsumedUnits | ForEach-Object { $_.SkuPartNumber = $SKUFriendlyNames.Item($_.SkuPartNumber)
$_ } | Export-Csv -Path ".\Overview.csv" -Encoding UTF8 -NoTypeInformation
$ActiveLicensePlans = Get-MsolAccountSku | Where-Object ConsumedUnits -gt 0

foreach ($LicensePlan in $ActiveLicensePlans) {
    $CSVPath = $SKUFriendlyNames.Item($LicensePlan.SkuPartNumber) + ".csv"
    Get-MsolUser | Where-Object {($_.licenses).AccountSkuId -match $LicensePlan.AccountSkuId } | Select-Object DisplayName, UserPrincipalName | Sort-Object DisplayName | Export-Csv -Path $CSVPath -Encoding UTF8 -NoTypeInformation
}

$ExportXLSXPath = "C:\Temp\LicenseOverview.xlsx"
$ExportXLSX = New-Object -ComObject Excel.Application
$ExportXLSX.Visible = $true 
$ExportXLSX.DisplayAlerts = $false
$WB = $ExportXLSX.Workbooks.Add()

Get-ChildItem *.csv |
    ForEach-Object{
        Try{
            Write-Host "Moving $_" -ForegroundColor green
            $Sheet = $WB.Sheets.Add()
            $Sheet.Name = $_.BaseName
            $Data = Get-Content $_ -Raw
            Set-Clipboard $Data
            $Sheet.UsedRange.PasteSpecial(
                [Microsoft.Office.Interop.Excel.XlPasteType]::xlPasteAll,
                [Microsoft.Office.Interop.Excel.XlPasteSpecialOperation]::xlPasteSpecialOperationAdd
            )
            $Sheet.UsedRange.TextToColumns(
                $Sheet.UsedRange,
                [Microsoft.Office.Interop.Excel.XlTextParsingType]::xlDelimited,
                [Microsoft.Office.Interop.Excel.XlTextQualifier]::xlTextQualifierDoubleQuote,
                $false, $false, $false, $true
            )
        }
        Catch{
            
        } Start-Sleep -s 1
    } 

Get-ChildItem *.csv | ForEach-Object { Remove-Item $_ }

$WB.Sheets.Item('sheet1').Delete()
$WB.SaveAs($ExportXLSXPath)
$WB.Close()
$ExportXLSX.Quit()

Write-Host "Overview exported to C:\Temp\LicenseOverview.xlsx"
