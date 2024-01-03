<#
 Microsoft provides programming examples for illustration only, without warranty either expressed or
 implied, including, but not limited to, the implied warranties of merchantability and/or fitness
 for a particular purpose.
 
 This sample assumes that you are familiar with the programming language being demonstrated and the
 tools used to create and debug procedures. Microsoft support professionals can help explain the
 functionality of a particular procedure, but they will not modify these examples to provide added
 functionality or construct procedures to meet your specific needs. if you have limited programming
 experience, you may want to contact a Microsoft Certified Partner or the Microsoft fee-based consulting
 line at (800) 936-5200.
 ---------------------------------------------------------------------------------------------------------------------------
 History
 ---------------------------------------------------------------------------------------------------------------------------
 v1.0 - 01/02/2024 - MARKKOR Created
 ============================================================================================================================
 
 This script is provided as is and without warranty. Please review it closely
 and fully test it in a non-production environment before implementing it
 in a production M365 tenant.
#>
$defaultLocation = "C:\Users\markkor\OneDrive\Code\PowerShell\PowerApps"

#Initialize variable for Source environment
$sourceEnv = "https://org5b2d9cff.crm.dynamics.com"

<############    YOU SHOULD NOT HAVE TO MODIFY ANYTHING BELOW THIS POINT    ############>

# ensure the destination directory exists
New-Item -ItemType Directory -Force -Path $defaultLocation | Out-Null

# set the default location
Set-Location -Path $defaultLocation

# determin the filename for the CSV file to be exported
$exportCSV = [System.IO.Path]::Combine($defaultLocation, "Exports.csv")

# Get user credentials for source environment
# $userName = Read-Host "Please Enter user name"
# $password = ConvertFrom-SecureString(Read-Host -AsSecureString "Please enter password") -AsPlainText

#Connect to Source Environment
Write-Host "Authenticating... "
pac auth create --name Source --username $username --password $password --url $sourceEnv

# get a list of unmanaged solutions e
Write-Host "Listing all solutions... "
$allSolutions = pac solution list

# get a reference to the header and use this to determine where to look for our values in the rows below
$header = $allSolutions | Select-Object -Skip 4 -First 1
$uniqueNameField   = 0
$FriendlyNameField = $header.IndexOf("Friendly Name")
$VersionField      = $header.IndexOf("Version")
$ManagedField      = $header.IndexOf("Managed")

# get all unmanaged solutions except the default and the "Common Data Services Default Solution"
$unmanagedSolutions = $allSolutions | Select-Object -skip 5 | `
  Where-Object { `
    $_.length -gt $ManagedField -and `
    $_.Substring($ManagedField) -eq "False" -and `
    $_.Substring($FriendlyNameField, $VersionField - $FriendlyNameField).Trim() -ne "Common Data Services Default Solution" -and `
    $_.Substring($uniqueNameField, $FriendlyNameField - $uniqueNameField).Trim() -ne "Default"} | `
    Select-Object `
      @{n="UniqueName";   ex={$_.Substring($uniqueNameField, $FriendlyNameField-$uniqueNameField).Trim()}},
      @{n="FriendlyName"; ex={$_.Substring($FriendlyNameField, $VersionField - $FriendlyNameField).Trim()}},
      @{n="Version";      ex={$_.Substring($VersionField, $ManagedField - $VersionField).Trim()}},
      @{n="Managed";      ex={$_.Substring($ManagedField).Trim()}}

# progress bar variables
$i = 0
$activity = "Exporting solutions..."
# Loop Through Solutions, export as zipped file and create Deployment Settings File
foreach ($solution in $unmanagedSolutions)
{
  $sol = $solution.UniqueName

  # progress bar
  $i++
  Write-Progress -Activity $activity -Status "($i/$($sol.Count)) - Exporting $sol" -PercentComplete ($i/$unmanagedSolutions.Count*100)

  Write-Host "Exporting Solution [$sol]"
  pac solution export --name $sol --path .\ --managed false
  pac solution create-settings --solution-zip .\$sol.zip --settings-file .\$sol.json

  # add a row to the CSV output file for this solution
  [PSCustomObject]@{
    UniqueName             = $sol
    FriendlyName           = $solution.FriendlyName
    Version                = $solution.Version
    Managed                = $solution.Managed
    ZipSavedLocation       = [System.IO.Path]::Combine($defaultLocation, "$sol.zip")
    JSONSavedLocation      = [System.IO.Path]::Combine($defaultLocation, "$sol.json")
    DestinationEnvironment = ""
  } | Export-Csv -Path $exportCSV -NoTypeInformation -Append
}
Write-Progress -Activity $activity -Completed -PercentComplete 100

Write-Host "Script completed!" -ForegroundColor Cyan