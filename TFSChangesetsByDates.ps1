########################################################################################################################
# PARAMETERS
########################################################################################################################

param
(
    [Parameter(Mandatory=$true)]
    [string]$Path,
    [Parameter(Mandatory=$true)]
    [string]$CollectionUrl,
    [Parameter(Mandatory=$true)]
    [string]$Dates,
    [Parameter(Mandatory=$false)]
    [string]$TFPath = "C:\Program Files (x86)\Microsoft Visual Studio 15.0\Common7\IDE",
    [Parameter(Mandatory=$false)]
    [string]$File = [Environment]::GetFolderPath("Desktop") + "\ChangesetsByDates_" + (Get-Date).ToString('ddMMyyhhmmss') + ".txt"
)

########################################################################################################################
# INITIALIZATION
########################################################################################################################


#Get the current location (to come back later)

$currentLocation = (Get-Item -Path ".\" -Verbose).FullName



# Check if TF.exe exists in $TFPath

if (-not(Test-Path $TFPath))
{
    # Try using visual studio 2015
    
    $TFPath = "C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE"

    if (-not(Test-Path $TFPath))
    {
        $TFPath = "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE"
    }

    # Try using visual studio 2013
    
    if (-not(Test-Path $TFPath))
    {
        Write-Host "Error, TF.exe not found (see https://www.visualstudio.com/en-us/docs/tfvc/use-team-foundation-version-control-commands)"
        Exit 101
    }
}


# Validate Dates Parameter

$Dates.Split(";") | ForEach-Object { 

    #Separate multiple dates
    if($_.Contains("-"))
    {
        #Separate intial and final dates
        $_.Split("-") | ForEach-Object {
            
            #Check if the date meet the format [day/month/year]
            if($_.Split("/").Count -ne 3)
            {
                Write-Host "Error, invalid Dates (invalid format) [format: day/month/year-day/month/year;<...>]"
                Exit 103
            }

            #Check if the month digit is valid
            if(("1","2","3","4","5","6","7","8","9","10","11","12") -notcontains $_.Split("/")[1])
            {
                Write-Host "Error, invalid Dates (invalid month) [format: day/month/year-day/month/year;<...>]"
                Exit 104
            }

            #Check if the year digit is valid (have 4 digits)
            if($_.Split("/")[2].Length -ne 4)
            {
                Write-Host "Error, invalid Dates (invalid year) [format: day/month/year-day/month/year;<...>]"
                Exit 105
            } 
        }
    }
    else
    {
        Write-Host "Error, invalid Dates ('-' is missing) [format: day/month/year-day/month/year;<...>]"
        Exit 102
    }
}


# Set tf.exe command arguments

$Command = "history"
$NoPrompt = "/noprompt"
$Recursive = "/recursive"
$CollectionUrl = "/collection:" + $CollectionUrl


# Move to TF.exe folder

cd $TFPath


########################################################################################################################
# BUILD THE REPORT
########################################################################################################################

Clear-Host
Write-Host "Starting..."

$TempFile = $File.Substring(0,$File.Length-4)+"_temp.txt"
$totalDates = $Dates.Split(";").Count
$currentDateIndex = 1
$TotalChangeSets = 0


# Arrays for save summary details summary

[string[]]$allDates = @()
[string[]]$allTotals = @()


# CREATE THE DETAILS SECTION IN TEMP FILE

#For each date
$Dates.Split(";") | ForEach-Object { 
    $initialDate = $_.Split("-")[0]
    $finalDate = $_.Split("-")[1]
    $Period = "/v:D" + $initialDate + "~D" + $finalDate

    # Run TF.exe command to retrieve changesets
    $Results = .\tf.exe $Command $Path $NoPrompt $Recursive $Period $CollectionUrl
    $TotalResults = $Results.Count

    # Exit if the tf.exe tool failed
    if($TotalResults -eq 0)
    {
        Exit 106
    }
    
    # Set $TotalResults to zero if there are not
    if($Results -eq "No history entries were found for the item and version combination specified.")
    {
        $TotalResults = 0
    }


    # Calculate total changesets
    $TotalChangeSets = $TotalChangeSets + $TotalResults


    # Fill summary arrays
    $allTotals += ,$TotalResults
    $allDates += ,$_


    # Fill details section in temp file
    Add-Content $TempFile ""
    Add-Content $TempFile "From: $initialDate"
    Add-Content $TempFile "To: $finalDate"
    Add-Content $TempFile ""
    Add-Content $TempFile $Results
    Add-Content $TempFile ""
    Add-Content $TempFile "Total: $TotalResults"
    Add-Content $TempFile ""
    Add-Content $TempFile "======================================================================================"

    Write-Host "Date set $currentDateIndex of $totalDates Finished"
    $currentDateIndex++
}


# CREATE THE REPORT FILE

Write-Host "Creating the report..."
$TempFileContent = Get-Content $TempFile

Add-Content $File "======================================================================================"
Add-Content $File "                                                                       by Leon Jalfon "
Add-Content $File "    ___ _                               _        ___                                  "
Add-Content $File "   / __| |_  __ _ _ _  __ _ ___ ___ ___| |_ ___ / __|_  _ _ __  _ __  __ _ _ _ _  _   "
Add-Content $File "  | (__| ' \/ _' | ' \/ _' / -_|_-</ -_)  _(_-< \__ \ || | '  \| '  \/ _' | '_| || |  "
Add-Content $File "   \___|_||_\__,_|_||_\__, \___/__/\___|\__/__/ |___/\_,_|_|_|_|_|_|_\__,_|_|  \_, |  "
Add-Content $File "                      |___/                                                    |__/   "
Add-Content $File "                                                                                      "
Add-Content $File "======================================================================================"
Add-Content $File "                                       SUMMARY                                        "
Add-Content $File "======================================================================================"
Add-Content $File ""

$dateLine = "Date: " + (Get-Date).ToString('dd/MM/yyyy hh:mm:ss tt')

Add-Content $File $dateLine
Add-Content $File "Path: $Path"
Add-Content $File ""


for($i=0; $i -lt $allDates.Length; $i++){
    $SummaryLine = "(" + $allDates[$i] + ") -> " + $allTotals[$i]
    Add-Content $File $SummaryLine
    }

Add-Content $File ""
Add-Content $File "Total: $TotalChangeSets"
Add-Content $File ""
Add-Content $File "======================================================================================"
Add-Content $File "                                       DETAILS                                        "
Add-Content $File "======================================================================================"


# Add contents from temp file and remove it

Get-Content $TempFile | Add-Content $File
Remove-Item -Path $TempFile -Force


########################################################################################################################

# Final Message
Write-Host "Changesets Summary Available in {$File}"


# Return to the start location
cd $currentLocation


########################################################################################################################



