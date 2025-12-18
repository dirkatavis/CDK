# This script parses a log file for either RO or MVA numbers.

param (
    [string]$dataType
)

# Define the path to the input file
$inputFile = "C:\Temp\Code\Scripts\VBScript\CDK\CreateNew_ROs\VehicleData.log"

# Validate the input
if ($dataType -ne "RO" -and $dataType -ne "MVA") {
    Write-Host "Invalid data type specified. Please use 'RO' or 'MVA'."
    exit
}

# Set the regex and output file based on the data type
if ($dataType -eq "RO") {
    $regex = "RO: (\d{6})"
    $outputFile = "C:\Temp\Code\Scripts\VBScript\CDK\CreateNew_ROs\Parse_Data_" + $dataType + ".txt"
} else { # MVA
    $regex = "MVA: (\d{8})"
    $outputFile = "C:\Temp\Code\Scripts\VBScript\CDK\CreateNew_ROs\Parse_Data_" + $dataType + ".txt"
}

# Get the content of the input file and find all matches
$matches = Get-Content $inputFile | Select-String -Pattern $regex -AllMatches

if ($matches) {
    # Extract the numbers from the matches
    $numbers = $matches | ForEach-Object { $_.Matches.Groups[1].Value }

    # Write the extracted numbers to the output file, overwriting it if it exists
    $numbers | Out-File -FilePath $outputFile
} else {
    Write-Host "No $dataType numbers found in the log file."
}
