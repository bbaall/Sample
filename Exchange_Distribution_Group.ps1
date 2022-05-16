## Written by Brent Lasater (2022)

##### 1. Connect #####

Import-Module ExchangeOnlineManagement

$upn = read-host "Enter an account with privileges necessary to manage Exchange Online"

Connect-ExchangeOnline -UserPrincipalName $upn



##### 2. Import data #####

Add-Type -AssemblyName System.Windows.Forms
$csv = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('MyDocuments') }
$csv.ShowDialog()

$exchange_data = import-csv $csv.FileName




##### 3. Check for new groups and create group(s) #####

$ErrorActionPreference = "Stop"

$exchange_data | forEach {
	try {
		Get-DistributionGroup -Identity $_.caseName
	}
	
	catch [Exchange.Specific.Error] { ## *******************NOTE - I do not have access to a Microsoft tenant with Exchange Online capabilities (personal accounts cannot access Exchange Online) so I cannot get this specific error - THIS IS A PLACEHOLDER
		$filtered_case_name = $_.caseName -replace '\W','' ## Parse caseName, removing all spaces and special characters.
		New-DistributionGroup -Identity $_.caseName -Alias $filtered_case_name -RequireSenderAuthenticationEnabled $false
	}
	
	catch {
		"An error other than object not found/does not exist has occurred."
	}
}



##### 4. Update group membership #####

$ErrorActionPreference = "Continue"

$exchange_data | forEach {
	Update-DistributionGroupMember -Identity $_.caseName -Members $_.smtpAddress
}



##### Notes regarding automation #####
<#

1.A)
In my situation this script was typically ran every ~2 months,
therefore I did not set this script up as a scheduled task
as I would manually run this script after another user updated 
the data (.xlsx document) and then informed me that the data
was updated. Therefore, I would manually enter the account password
when running Connect-ExchangeOnline.

Automate this process by using Get-Secret (Microsoft.Powershell.SecretManagement)
using something like an Azure KeyVault, KeyPass, etc.

2.A)
This script assumes that the data is contained within a CSV file.
When working with the client data was contained within a XSLX
excel spreadsheet. Automating data conversion can be done by using
an Excel COM object (New-Object -COMObject excel.application and
saving the file as an CSV. However, if running on a server it is not
recommended to have Microsft Office installed. I have not used these
but it looks like there are some other PowerShell modules that can
import/export Excel spreadsheets such as the module ImportExcel.
Using this module does not require an installation of Excel on the
device running the PowerShell script.

2.B)
This step can be automated by using a specific folder where the
data would be stored. Use Get-ChildItem on this location to get
the data and then import that data using import-csv. Constraints
would be knowing the name of the file or only allowing one data file
in the folder.

Get-ChildItem -path "C:\Name\Of\Folder

or

import-csv -path "C:\Name\Of\Folder\data.csv"

3.A)
This step can be dealt with a number of ways. My first intuition
would be to maintain a separate spreadsheet/CSV that only includes
NEW distribution groups - these would later be merged into the
master document. This step could then create new distribution
groups and would not require any data processing if only NEW 
groups were held in a separate spreadsheet/CSV

If using only one master sheet that contains all data, then some 
processing would need to occur comparing the CURRENT distribution
groups using Get-DistributionGroup and any NEW groups from the
master sheet. This could be done with a FOR loop searching the
current distribution groups to see if groups found in the master
sheet exist and if not then create the new group.

I chose to use a Try-Catch because when running Get-DistributionGroup
because when the master sheet is updated, we would expect that some groups
do not exist. Therefore, an error would be thrown that the group does not
exist or that the object does not exist. When this error occurs we can
run New-DistributionGroup and create that group.


#>


<# Sample Data 
caseName			caseMembers			smtpAddress
-----------------------------------------------------------------------------------------
Valley v. Fields	Mark, Kevin, James	mark@domain.com,kevin@domain.com,james@domain.com
Water v. Fire		Mark, Kevin			mark@domain.com,kevin@domain.com
Light v. Dark		James				james@domain.com
Weak v. Strong		Mark, Kevin			mark@domain.com,kevin@domain.com

#>
