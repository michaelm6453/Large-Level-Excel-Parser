<#
.SYNOPSIS
    EWIP parser script for parsing Excel files and filtering specific columns.
.DESCRIPTION
    This PowerShell script is designed to parse an Excel file (.xlsx) by filtering specific columns and displaying the results in a table format. 
.PARAMETER FilePath
    Specifies the path of the Excel file to be parsed.
.PARAMETER SheetName
    Specifies the name of the sheet within the Excel file to be parsed (optional).
.EXAMPLE
    $excelPath = "C:\Path\To\Your\File.xlsx"
    Import-Excel -FilePath $excelPath -SheetName "Sheet1"
.NOTES
    This script requires Excel to be installed on the system.
    The Import-Excel function converts the Excel file to CSV and imports the data.
    The script then filters the data by removing rows where any of the specified properties are null.
    Finally, it displays the filtered data in a table format and exports it to a CSV file.
#>

# FUNCTIONS 

function Import-Excel([string]$FilePath, [string]$SheetName = "")
{
    $csvFile = Join-Path $env:temp ("{0}.csv" -f (Get-Item -path $FilePath).BaseName)
    if (Test-Path -path $csvFile) { Remove-Item -path $csvFile }

    # convert Excel file to CSV file
    $xlCSVType = 6 
    $excelObject = New-Object -ComObject Excel.Application  
    $excelObject.Visible = $false 
    $workbookObject = $excelObject.Workbooks.Open($FilePath)
    SetActiveSheet $workbookObject $SheetName | Out-Null
    $workbookObject.SaveAs($csvFile,$xlCSVType) 
    $workbookObject.Saved = $true
    $workbookObject.Close()

     # cleanup 
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookObject) |
        Out-Null
    $excelObject.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelObject) |
        Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    # now import and return the data 
    Import-Csv -path $csvFile
}

function FindSheet([Object]$workbook, [string]$name)
{   # FindSheet function locates the sheet number based on the provided sheet name within a workbook.

    $sheetNumber = 0
    for ($i=1; $i -le $workbook.Sheets.Count; $i++) {
        if ($name -eq $workbook.Sheets.Item($i).Name) { $sheetNumber = $i; break }
    }
    return $sheetNumber
}

function SetActiveSheet([Object]$workbook, [string]$name)
{
    # SetActiveSheet function activates a specific sheet within a workbook based on the provided sheet name.

    if (!$name) { return }
    $sheetNumber = FindSheet $workbook $name
    if ($sheetNumber -gt 0) { $workbook.Worksheets.Item($sheetNumber).Activate() }
    return ($sheetNumber -gt 0)
}

function Compare-Arrays {
    # Function to compare and create a new array

    param (
        [array]$Array1, # entire EWIP pc list
        [array]$Array2  #list of pc names to search for
    )
    
    $newArray = @()

    foreach ($obj2 in $Array2) {
        $matchedObj = $Array1 | Where-Object { $_.WORKSTATION_NAME -eq $obj2.'PC name' }
        if ($matchedObj) {
            $newArray += $matchedObj
        }
    }

    return $newArray
}

function Name-Not-Found {
# Function to see 
    param (
        [array]$Array1, # list of pc names found
        [array]$Array2  # list of pc names to search for
    )

    $newArray = @()

    foreach ($obj2 in $Array2) {  
    
        $found_PC = $false


        $matchedObj = $Array1 | Where-Object { $_.WORKSTATION_NAME -eq $obj2.'PC name' }
        if ($matchedObj) {
            $found_PC = $true
        }

        if($found_PC -eq $false){
            $newArray += $obj2
        }
    }

    return $newArray
}

function Welcome-Message{
    Clear-Host

    # User enters a filepath to the .xlsx file from their computer 
    Write-Host "Welcome to the Excel Parser. This program was written by Michael Morra."
    
    [bool]$invalid_input = $true
    
    # Can easily add another prompt to ask for output path


    while($invalid_input){
    
    [int]$number_entered = Read-Host -Prompt "Press 1 to process a large EWIP report. Press 2 to get information based on PC name"
    
    
    if($number_entered -eq 1 -or $number_entered -eq 2){
        $invalid_input=$false
        }
    else {
    
        Write-Host "Please enter a valid selection..."
    
        }
    
    }
    
    Clear-Host
    
    Write-Host "Processing Request..."
}

function Update-File-Location {

    Clear-Host

    [string]$user_file = 'C:\Users\MORRAM\Documents\pc_names.xlsx' # Read-Host -Prompt "Enter the filepath to the excel file that contains a list of PC names" 

}

# Main code

Welcome-Message


# $excelPath = "C:\Users\MORRAM\Documents\testingFileEWIP.xlsx"



if($number_entered -eq 1){


$excelPath = "C:\Users\MORRAM\Documents\pcNamesFiltered.csv"

# $excelPath = "C:\Users\MORRAM\Documents\pcNamesInfo.csv"

$excelData = Import-Excel($excelPath)

# At this point we have imported the excel data/sheet

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

# Export the Data to CSV

$selectedData | Export-Csv -Path 'C:\Users\MORRAM\Documents\all_PCs.csv' -NoTypeInformation
}
else{
    #This section takes a list of PC names and compares it to the ewip report and returns the list of data associated with the PC names
    
    Clear-Host

    Write-Host "Processing..."

    [string]$user_file = 'C:\Users\MORRAM\Documents\pc_names.xlsx' # Read-Host -Prompt "Enter the filepath to the excel file that contains a list of PC names" 

    $list_pc_names = Import-Excel($user_file)

    $parsed_EWIP = Import-Excel('C:\Users\MORRAM\Documents\all_PCs.csv') #this is the parsed excel sheet from the if statement

    
    # use for loop to compare each pc name with list, and when there's a match, store in array of obj and then after put in csv 

    

    $pc_full_info = Compare-Arrays -Array1 $parsed_EWIP -Array2 $list_pc_names    # $parsed_EWIP | Where-Object {$list_pc_names -contains $_.WORKSTATION_NAME}

    $pc_list_not_found = Name-Not-Found -Array1 $pc_full_info -Array2 $list_pc_names
    
    $pc_full_info | Export-Csv -Path 'C:\Users\MORRAM\Documents\pcs_found.csv' -NoTypeInformation

    $pc_list_not_found | Export-Csv -Path 'C:\Users\MORRAM\Documents\pcs_not_found.csv' -NoTypeInformation

    
    }
