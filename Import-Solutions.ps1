#Requires -Version 7.0
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
 v1.0 - 01/04/2024 - MARKKOR Created
 ============================================================================================================================
 
 This script is provided as is and without warranty. Please review it closely
 and fully test it in a non-production environment before implementing it
 in a production M365 tenant.

 Because of the -AsPlainText parameter on ConvertFrom-SecureString, PowerShell 7 is required
#>

# use the same $defaultLocation variable in both the import and export script
$defaultLocation = "[Path to store export/import files]"

<############    YOU SHOULD NOT HAVE TO MODIFY ANYTHING BELOW THIS POINT    ############>

# boolean variable to indicate if we have exceptions during import
$blnException = $false

# Get user credentials for target environment
$userName = Read-Host "Please enter user name"
$password = ConvertFrom-SecureString(Read-Host -AsSecureString "Please enter password") -AsPlainText

# import the CSV contents from the default location
$csv = [System.IO.Path]::Combine($defaultLocation, "Exports.csv")
$imports = Import-Csv -Path $csv

# create an exception filename in case one is needed
$exceptionFile = [System.IO.Path]::Combine($defaultLocation, "exceptions.csv")

# get a unique list of destination environments
$destinationEnvironments = $imports | Select-Object -ExpandProperty DestinationEnvironment -Unique

# progress bar values
$i = 0
$activity = "Importing solutions..."

Write-Host "NOTE: Any solution to be imported which has a version matching the version of a solution`n      by the same name that is already in the target environment will not be updated." -ForegroundColor Yellow

# loop through the destination environments
foreach ($destinationEnvironment in $destinationEnvironments)
{
    Write-Host "`nWorking with destination environment: " -NoNewline
    Write-Host "$destinationEnvironment" -ForegroundColor Yellow

    # Connect to Target Environment
    Write-Host "Authenticating..."
    pac auth create --name Destination --username $username --password $password --url $destinationEnvironment

    # get a list of solutions to import into that destination environment
    $solutions = $imports | Where-Object DestinationEnvironment -eq $destinationEnvironment

    # loop through all of the solutions for the destination environment
    foreach ($solution in $solutions) 
    {
        # get values from the current solution
        $solutionName  = $solution.UniqueName
        $friendlyName  = $solution.FriendlyName
        $sourceVersion = $solution.Version
        $managed       = $solution.Managed
        $zipLocation   = $solution.ZipSavedLocation
        $jsonLocation  = $solution.JSONSavedLocation

        # update progress bar
        $i++
        Write-Progress -Activity $activity -Status "($i/$($destinationEnvironments.Count)) - Importing solution $solutionName" -PercentComplete ($i/$imports.Count*100)
        
        Write-Host "Importing solution: " -NoNewline
        Write-Host "$solutionName" -ForegroundColor Yellow

        # verify that both the ZIP and JSON file exist
        if ((Test-Path $zipLocation) -and (Test-Path $jsonLocation))
        {
            # import the solution with settings
            $response     = pac solution import --path $zipLocation --settings-file $jsonLocation 

            #initilize error message variables
            $errorFound   = $false
            $errorMessage = ""

            # Look for an error message in the response
            foreach ($responseline in $response)
            {
                if ($responseLine.Contains("error") -or $errorFound)
                {
                    $errorFound   = $true
                    $errorMessage += $responseline
                }
            }

            # if an error was found, write it to the exception file
            if ($errorFound)
            {
                Write-Host "Exception during import:" -ForegroundColor Red
                Write-Host $errorMessage -ForegroundColor Red
                Write-Host ""
                [PSCustomObject]@{
                    SolutionName = $solutionName
                    FriendlyName = $friendlyName
                    Version      = $sourceVersion
                    Managed      = $managed
                    ZipLocation  = $zipLocation
                    JsonLocation = $json
                    Exception    = $errorMessage
                } | Export-Csv -Path $exceptionFile -NoTypeInformation -Append
            }
            Write-Host "Upload completed for solution " -NoNewline
            Write-Host "$solutionName" -ForegroundColor Yellow
        }
        else 
        {
            # one or both of the files needed for import were not found
            # write to the exception file
            Write-Host "Either ZIP or JSON file is missing for solution $solutionName" -ForegroundColor Red
            [PSCustomObject]@{
                SolutionName = $solutionName
                FriendlyName = $friendlyName
                Version      = $sourceVersion
                Managed      = $managed
                ZipLocation  = $zipLocation
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