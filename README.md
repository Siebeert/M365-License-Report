# M365-License-Report
Dit script exporteert alle Microsoft 365 licenties van een klant naar Excel.

# How to:
- Voer het script uit met Powershell
- Meld aan met de/jouw Azure administrator van de klant
- De Excel wordt gegenereerd en is terug te vinden onder C:\Temp\LicenseOverview.xlsx
- Elk tabblad opmaken als tabel moet je zelf nog doen

# Let op:
- Werkt enkel op computers waar Excel geïnstalleerd is

# Technische details
- Er wordt een aparte CSV gegenereerd met een overzicht van alle licentieplannen onder de tenant en voor alle gebruikers per licentieplan (weggeschreven naar C:\Temp)
- Er wordt een Excel gegenereerd en de CSV's worden elk in een nieuw tabblad geïmporteerd
- Alle CSV's worden nadien terug verwijderd, zodat enkel de Excel overblijft
