# M365-License-Report
This script exports all Microsoft 365 licenses of a customer to an Excel file.

# How to
- Execute the script with Powershell
- Log in with the/your Azure administrator for that customer (works for every customer)
- The Excel is generated and saved to LicenseOverview.xlsx
- You will have to 'Format as table' manually for each sheet

# Opmerkingen
Make sure the following 2 Powershell modules are installed:
- ImportExcel
- MSOnline
