<#
.SYNOPSIS
    The script interacts with the user to process either 
    a) a large EWIP report or 
    b) compare PC names from Excel files, converting them to uppercase for matching. 

    It then exports the matched and unmatched PC data to separate CSV files.

.DESCRIPTION
    The main code section prompts the user for input based on the selected operation (processing a report or comparing PC names), 
    reads file paths, processes the data accordingly, and exports the results to specified locations.
.PARAMETER FilePath
    Specifies the path of the Excel file to be parsed.
.PARAMETER SheetName
    Specifies the name of the sheet within the Excel file to be parsed (optional).
.PARAMETER lower_case_PC_list
    Specifies a list of PC names in lowercase to be converted to uppercase.
.PARAMETER Array1
    Specifies the first array to be compared.
.PARAMETER Array2
    Specifies the second array to be compared.
.PARAMETER excelPath
    Specifies the path of the Excel file to be processed in the Filter-Large-Report function.
.PARAMETER excelPathExport
    Specifies the path where the processed Excel file will be exported.
.PARAMETER user_file
    Specifies the path of the Excel file containing a list of PC names.
.PARAMETER EWIP_file
    Specifies the path of the Excel file containing all PC information for comparison.
.PARAMETER path_matches
    Specifies the path where the found PC information will be exported.
.PARAMETER path_no_matches
    Specifies the path where the not-found PC information will be exported.
.NOTES
    Author: Michael Morra (3/21/2024)
    This script requires Excel to be installed on the system.
.EXAMPLE

    # Program prompts you to enter 1 or 2 (1 is for large report parsing, 2 is for comparing pc list to ewip report)

    1 

    # Program prompts you to enter file path to excel file to parse and another prompt to an export location


    C:\Path\To\Your\EWIP\File.xlsx
    C:\Path\To\Your\Export\File.xlsx
    
    Can easily change code to include sheet name
    ex: Import-Excel -FilePath $excelPath -SheetName "Sheet1"
#>


# FUNCTIONS

<# 
Import-Excel: Imports an Excel file and converts it to CSV format.

FindSheet: Finds a specific sheet within a workbook by name.

SetActiveSheet: Activates a specified sheet within a workbook.

Convert-To-Upper: Converts strings in an array to uppercase.

Compare-Arrays: Compares arrays based on matching attributes.

Name-Not-Found: Checks for names not found in a list of PC names.

Welcome-Message: Displays a welcome message and prompts the user for input.

Filter-Large-Report: Processes a large EWIP report, sorts data, and exports it to a new Excel file.

#>


# Function to import an Excel file and convert it to CSV
function Import-Excel([string]$FilePath, [string]$SheetName = "")
{
    $csvFile = Join-Path $env:temp ("{0}.csv" -f (Get-Item -path $FilePath).BaseName)
    if (Test-Path -path $csvFile) { Remove-Item -path $csvFile }

    # Convert Excel file to CSV file
    $xlCSVType = 6 
    $excelObject = New-Object -ComObject Excel.Application  
    $excelObject.Visible = $false 
    $workbookObject = $excelObject.Workbooks.Open($FilePath)
    SetActiveSheet $workbookObject $SheetName | Out-Null
    $workbookObject.SaveAs($csvFile,$xlCSVType) 
    $workbookObject.Saved = $true
    $workbookObject.Close()

    # Cleanup
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookObject) | Out-Null
    $excelObject.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelObject) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    # Import and return the data 
    Import-Csv -path $csvFile
}

# Function to find a sheet by name within a workbook
function FindSheet([Object]$workbook, [string]$name)
{
    $sheetNumber = 0
    for ($i=1; $i -le $workbook.Sheets.Count; $i++) {
        if ($name -eq $workbook.Sheets.Item($i).Name) { $sheetNumber = $i; break }
    }
    return $sheetNumber
}

# Function to activate a specific sheet within a workbook by name
function SetActiveSheet([Object]$workbook, [string]$name)
{
    if (!$name) { return }
    $sheetNumber = FindSheet $workbook $name
    if ($sheetNumber -gt 0) { $workbook.Worksheets.Item($sheetNumber).Activate() }
    return ($sheetNumber -gt 0)
}

# Function to convert strings in an array to uppercase
function Convert-To-Upper{
    param(
        [array]$lower_case_PC_list # list with lowercase PC names
    )

    $newArray = @()
    
    foreach ($obj in $lower_case_PC_list){
        $obj.'PC name' = $obj.'PC name'.toUpper() 
        $newArray += $obj
    }

    return $newArray
}

# Function to compare arrays and create a new array based on matching attributes
function Compare-Arrays {
    param (
        [array]$Array1, # entire EWIP pc list
        [array]$Array2  # list of pc names to search for
    )
    
    $newArray = @()

    foreach ($obj2 in $Array2) {
        $obj2NameNormalized = $obj2.'PC name'.Trim().Normalize()
        $matchedObj = $Array1 | Where-Object { $_.WORKSTATION_NAME.Trim().Normalize() -eq $obj2NameNormalized }
        if ($matchedObj) {
            $newArray += $matchedObj
        }
    }

    return $newArray
}

# Function to check for names not found
function Name-Not-Found {
    param (
        [array]$Array1, # list of pc names found (less than pc names to search for)
        [array]$Array2  # list of pc names to search for
    )

    $newArray = @()

    foreach ($obj2 in $Array2) {  
        $found_PC = $false
        $obj2NameNormalized = $obj2.'PC name'.Trim().Normalize()

        $matchedObj = $Array1 | Where-Object { $_.WORKSTATION_NAME.Trim().Normalize() -eq $obj2NameNormalized }
        if ($matchedObj) {
            $found_PC = $true
        }

        if($found_PC -eq $false){
            $newArray += $obj2
        }
    }

    return $newArray
}

# Function to display a welcome message and prompt the user for input
function Welcome-Message{
    Clear-Host

    Write-Host "Welcome to the Excel Parser. This program was written by Michael Morra."
    
    $invalid_input = $true

    while($invalid_input){
        $Global:number_entered = Read-Host -Prompt "Press 1 to process a large EWIP report. Press 2 to get information based on PC name"
    
        if($Global:number_entered -eq 1 -or $Global:number_entered -eq 2){
            $invalid_input=$false
        }
        else {
            Write-Host "Please enter a valid selection..."
        }
    }
    
    Clear-Host
    
    Write-Host "Processing Request..."
}

function Filter-Large-Report {

    param (
        [string]$excelPath, # Import excel file
        [string]$excelPathExport  # Export excel file
    )


$excelData = Import-Excel($excelPath)

# Convert LAST_HARDWARE_SCAN to DateTime objects
$excelData | ForEach-Object {
    if (![string]::IsNullOrWhiteSpace($_.LAST_HARDWARE_SCAN)) {
        $_.LAST_HARDWARE_SCAN = [DateTime]::ParseExact($_.LAST_HARDWARE_SCAN, 'M/d/yyyy H:mm', $null)
    }
}


$selectedData = $excelData | # Where-Object { ![string]::IsNullOrWhiteSpace($_.LAST_HARDWARE_SCAN) } |
    Group-Object -Property WORKSTATION_NAME | ForEach-Object {
        $_.Group | Sort-Object LAST_HARDWARE_SCAN -Descending | Select-Object WORKSTATION_NAME, LAST_HARDWARE_SCAN, LAST_LOGGED_USER_ID, PRIMARY_USER_ID, IP_ADDRESS, SUBNET -First 1
    }

$selectedData | Export-Csv -Path $excelPathExport -NoTypeInformation

}


# Main code

Welcome-Message


if($Global:number_entered -eq 1){

    Clear-Host

    # C:\Users\MORRAM\Documents\pcNamesFiltered.csv
    $excelPath = Read-Host -Prompt "Enter a filepath to the excel file generated by the EWIP report (large excel file with all PC information columns)"

    # C:\Users\MORRAM\Documents\all_PCs2.csv
    $excelPathExport = Read-Host -Prompt "Enter a file location to export the parsed file to"

    Clear-Host

    Write-Host "Processing..."

    Filter-Large-Report -excelPath $excelPath -excelPathExport $excelPathExport

}
else{
    #This section takes a list of PC names and compares it to the ewip report and returns the list of data associated with the PC names
    
    Clear-Host

    # C:\Users\MORRAM\Documents\pc_names.xlsx
    [string]$user_file = Read-Host -Prompt "Enter the filepath to the excel file that contains a list of PC names" 

    # C:\Users\MORRAM\Documents\all_PCs.csv
    [string]$EWIP_file= Read-Host -Prompt "Enter the filepath to the excel file that has all PC information (can be made from selecting 1 on the previous prompt)" 

    # C:\Users\MORRAM\Documents\pcs_found1.csv
    $path_matches = Read-Host -Prompt "Enter a file location to export the found PC information to"

    # C:\Users\MORRAM\Documents\pcs_not_found1.csv
    $path_no_matches = Read-Host -Prompt "Enter a file location to export the not-found PC to"

    Clear-Host

    Write-Host "Processing..."

    $list_pc_names = Import-Excel($user_file)

   #  $list_pc_names.'PC name'

   #  $list_pc_names.'PC name'

    $parsed_EWIP = Import-Excel($EWIP_file) #this is the parsed excel sheet from the if statement


    $list_pc_names = Convert-To-Upper -lower_case_PC_list $list_pc_names #Convert pc name list to uppercase

    
    # use for loop to compare each pc name with list, and when there's a match, store in array of obj and then after put in csv 

    
    $pc_full_info = Compare-Arrays -Array1 $parsed_EWIP -Array2 $list_pc_names    # $parsed_EWIP | Where-Object {$list_pc_names -contains $_.WORKSTATION_NAME}

    $pc_list_not_found = Name-Not-Found -Array1 $pc_full_info -Array2 $list_pc_names
    
    $pc_full_info | Export-Csv -Path $path_matches -NoTypeInformation

    $pc_list_not_found | Export-Csv -Path $path_no_matches -NoTypeInformation

    
    }
