
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

# this should be the same location as was used in the export script
$defaultLocation = "C:\users\ryanstark\Documents\EnvExports"

<############    YOU SHOULD NOT HAVE TO MODIFY ANYTHING BELOW THIS POINT    ############>

# indicate if we have exceptions during import
$blnException = $false

#Get user credentials for target environment
$userName = Read-Host "Please Enter user name"
$password = ConvertFrom-SecureString(Read-Host -AsSecureString "Please enter password") -AsPlainText

# import the CSV contents from the default location
$csv = [System.IO.Path]::Combine($defaultLocation, "Exports.csv")
$imports = Import-Csv -Path $csv

# create an exception filename in case one is needed
$exceptionFile = [System.IO.Path]::Combine($defaultLocation, "exceptions.csv")

# get a unique list of destination environments
$destinationEnvironemnts = $imports | Select-Object -ExpandProperty DestinationEnvironment -Unique

# progress bar values
$i = 0
$activity = "Importing solutions..."

Write-Host "NOTE: Any import solution versions which match the versions in the environment will not be updated." -ForegroundColor Yellow

# loop through the destination environments
foreach ($destinationEnvironemnt in $destinationEnvironemnts)
{
    Write-Host "Working with destination environment $destinationEnvironemnt"

    # Connect to Target Environment
    pac auth create --name Destination --username $username --password $password --url $destinationEnvironemnt

    # get a list of solutions to import into that destination environment
    $solutions = $imports | Where-Object DestinationEnvironment -eq $destinationEnvironemnt

    # loop through all of the solutions for the destination environment
    foreach ($solution in $solutions) 
    {
        # get values from the current solution
        $sol = $solution.UnmanagedSolutionName
        $zip = $solution.ZipSavedLocation
        $json = $solution.JSONSavedLocation

        # update progress bar
        $i++
        Write-Progress -Activity $activity -Status "($i/$($destinationEnvironemnts.Count)) - Importing solution $sol" -PercentComplete ($i/$imports.Count*100)
        
        Write-Host "Importing solution: $sol"

        # verify that both the ZIP and JSON file exist
        if ((Test-Path $zip) -and (Test-Path $json))
        {
            # import the solution with settings
            pac solution import --path $zip --settings-file $json
        }
        else 
        {
            # one or both of the files needed for import were not found
            # write to the exception file
            Write-Host "Either ZIP or JSON file is missing for solution $sol" -ForegroundColor Red
            [PSCustomObject]@{
                SolutionName = $sol
                ZipLocation  = $zip
                JsonLocation = $json
                Exception    = "File not found"
            } | Export-Csv -Path $exceptionFile -NoTypeInformation -Append
            $blnException = $true
        }
    }
}

# check for exception
if ($blnException)
{
    Write-Host "Exceptions encountered during import. Review exception file below:" -ForegroundColor Red
    Write-Host "$exceptionFile`n" -ForegroundColor Red
}

Write-Host "Script Complete!" -ForegroundColor Cyan