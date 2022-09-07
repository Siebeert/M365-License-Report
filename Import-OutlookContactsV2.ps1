###################################################################
#                                                                 #
# Made by SiebeScript™                                            #
#                                                                 #
# This script imports all entries in the organization's GAL into  #
# the default Outlook contacts.                                   #
#                                                                 #
# This script works directly in Outlook under the user's context. #
#                                                                 #
###################################################################

#region Variables

$deleteContacts = ""                    # Contacts to delete, full name separated by comma (e.g. "John Smith","Jane Doe")
$LogFile = "C:\Temp\$env:Username.log"  # The script will write logging to this location, preferably a network location
$mainPhoneNumer = ""                    # For contacts that have the main phone number set, the phone number will be set to empty (to avoid multiple people having the same phone number)
$receptionJobTitle = ""                 # The jobTitle attribute of those that should be excluded from above

#endregion Variables

#region Functions

function Add-Contact 
{
    param($contact)

    $newContact = $contactList.Items.Add()

    $newContact.FirstName = $contact.FirstName
    $newContact.LastName = $contact.LastName
    $newContact.Email1Address = $contact.PrimarySmtpAddress
    if ($contact.MobileTelephoneNumber -ne "") { $newContact.MobileTelephoneNumber = $contact.MobileTelephoneNumber }
    $newContact.JobTitle = $contact.JobTitle
    $newContact.CompanyName = $contact.CompanyName
    $newContact.Department = $contact.Department
    if (($contact.BusinessTelephoneNumber -eq $mainPhoneNumber) -and ($contact.JobTitle -ne $receptionJobTitle)) {
        $newContact.BusinessTelephoneNumber = ""
    } else {
        $newContact.BusinessTelephoneNumber = $contact.BusinessTelephoneNumber
    }
    $newContact.Business2TelephoneNumber = $contact.Business2TelephoneNumber

    $newContact.Save()

    Add-Content -Path $LogFile -Value ((Get-Date).ToString() + " | Added new contact " +$contact.Name)
}

function Update-Contact
{
    param ($contact)

    $fullName = $contact.FirstName + " " + $contact.LastName
    $searchFilter = "[FullName] = " + $fullName
    $existingContact = $contactList.Items.Find($searchFilter)

    $existingContact.FirstName = $contact.FirstName
    $existingContact.LastName = $contact.LastName
    $existingContact.Email1Address = $contact.PrimarySmtpAddress
    if ($contact.MobileTelephoneNumber -ne "") {$existingContact.MobileTelephoneNumber = $contact.MobileTelephoneNumber }
    $existingContact.JobTitle = $contact.JobTitle
    $existingContact.CompanyName = $contact.CompanyName
    $existingContact.Department = $contact.Department
    if (($contact.BusinessTelephoneNumber -eq $mainPhoneNumber) -and ($contact.JobTitle -ne $receptionJobTitle)) {
        $existingContact.BusinessTelephoneNumber = ""
    } else {
        $existingContact.BusinessTelephoneNumber = $contact.BusinessTelephoneNumber
    }
    $existingContact.Business2TelephoneNumber = $contact.Business2TelephoneNumber

    $existingContact.Save()
    Add-Content -Path $LogFile -Value ((Get-Date).ToString() + " | Updated contact $fullName")
}

function Remove-Contact 
{
    param ($fullName)
    
    $searchFilter = "[FullName] = $fullName"
    $contactList.Items.Find($searchFilter).Delete()
}

function Contact-Exists
{
    param ($contact)
    
    $fullName = $contact.FirstName + " " + $contact.LastName
    $searchFilter = "[FullName] = " + $fullName
    return $contactList.Items.Find($searchFilter)
}

#endregion Functions

# Open Outlook and get default Contacts List
$olApp = new-object -comobject outlook.application
$namespace = $olApp.GetNamespace("MAPI")
$contactList = $namespace.GetDefaultFolder(10)

$GAL = $namespace.AddressLists("Global Address List").AddressEntries
$GALcontacts =  New-Object System.Collections.ArrayList($null)
foreach ($contact in $GAL) { $GALcontacts.Add($contact.GetExchangeUser()) | Out-Null }

Add-Content -Path $LogFile -Value ((Get-Date).ToString() + " | Importing contacts to " + $namespace.GetDefaultFolder(10).FolderPath)
   
# Add contacts
$progressCounter = 0
foreach ($contact in $GALcontacts) 
{ 
    $progressCounter++
    #Write-Progress -Activity "Adding contacts to outlook" -Status "$progressCounter of $($GALcontacts.Count) complete" -PercentComplete ($progressCounter / $GALcontacts.Items.Count * 100)    
    if ($contact -ne $null -and ($contact.FirstName -ne "" -and $contact.LastName -ne "")) {
        if (Contact-Exists $contact) {
            Update-Contact $contact | Out-Null
        } else {
            Add-Contact $contact | Out-Null
        }
    }
}

# Remove contacts
foreach ($fullName in $deleteContacts) 
{
    Remove-Contact $fullName
}

# Clean up sessions
$olApp.Quit | Out-Null
[GC]::Collect()