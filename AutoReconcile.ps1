#requires -version 3
<#
.SYNOPSIS
This script obtains and converts the password vault report, then runs the reconcilation script. Requires Microsoft Excel to be installed.
 
.DESCRIPTION
This is a companion script to Inventory_Reconcilation_SQLite.ps1.
This script obtains and converts the password vault report, then runs the reconcilation script.
Microsoft Excel must be installed on the system in order to successfully convert XLSX to CSV.
 
.PARAMETER workdir
The working directory for the script. Defaults to the temporary directory.
 
.OUTPUTS
Resulting reports are uploaded to SharePoint:
http://sharepoint.example.com/sites/Project42/Project%20Document%20Library/Reconciliation%20Reports/
 
.NOTES
Version:        0.9
Author:         Andrew Zbikowski <andrew@zibnet.us>
v0.9 Changes, 2015-10-27:
    *Development/Testing complete. Ready for the bugs that can only be discovered in production.
  
.EXAMPLE
AutoReconcile.ps1 -workdir C:\Windows\Temp
#>
 
#-------------[Parameters]-----------------------------------------------------
# Enable -Debug, -Verbose Paramaters. Write-Debug and Write-Verbose!
[CmdletBinding()]
Param (
    [Parameter()][string]$workdir = "$env:TEMP"
)
 
#-------------[Parameter Validation]-------------------------------------------
 
# Test that $workdir exists
if (-Not (Test-Path -Path $workdir -PathType Container)) {
    Throw "Work directory $workdir does not exist!"
}
 
#-------------[Initialisations]------------------------------------------------
# If the report script is renamed, this will have to be changed.
$reportScript = (Split-Path -parent $PSCommandPath) + '\vault_Inventory_Reconcilation_SQLite.ps1'
#-------------[Declarations]---------------------------------------------------
 
#-------------[Functions]------------------------------------------------------
 
<#
    Splits a full path to a file into Path and Filename.
    Returns a PSObject with the properties Path and Filename.
#>
Function Trim-Filename ($fullpath) {
    $position = $fullpath.LastIndexOf("\")
    $object = New-Object –TypeName PSObject
    $object | Add-Member –MemberType NoteProperty –Name Path –Value $fullpath.SubString(0,$position+1)
    $object | Add-Member –MemberType NoteProperty –Name Filename –Value $fullpath.SubString($position+1)
    $object | Add-Member –MemberType NoteProperty –Name Fullpath –Value $fullpath
    $object
}
 
 
<#
Download vault Report for the given date to working directory. 
Returns $false if the file does not exist, otherwise returns path to downloaded file.
#>
Function Get-vaultReport {
    Param (
        [Parameter(Mandatory=$true)]$date
    )
    # Convert the passed date object into a string of the needed format.
    $datestamp = Get-Date -Format s -Date $date
    $datestamp = $datestamp.Substring(0,10)
    Write-Verbose "Checking for vault Report from $datestamp."
    $datestamp = $datestamp.Replace("-","_")
    # URL to the SharePoint folder with the needed vault reports.
    $in_URL = 'http://sharepoint.example.com/sites/Project42/Project%20Document%20Library/Inventory%20Reports/'
    # Format of the report file name up to date stamp.
    $in_filename = 'Reporting%20-%20Enterprise%20Host%20Service_'
    # File extension of reports.
    $in_fileext = '.xlsx'
    # Full URL of the file the script will attempt to download.
    $in_fullURL = $in_URL + $in_filename + $datestamp + $in_fileext
    Write-Debug "Download URL: $in_fullURL"
    # Build full path to the downloaded file.
    $outfile = $in_filename + $datestamp + $in_fileext
    $outfile = "$workdir\$outfile"
    $outfile = $outfile.Replace('%20','_')
    Write-Debug "Download File: $outfile"
    # Create a .NET Webclient object to download the file.
    $webclient = New-Object System.Net.WebClient
    $webclient.UseDefaultCredentials = $true
    # if not using default credentials, $webclient.Credentials = Get-Credential
    Try { $webclient.DownloadFile($in_fullURL, $outfile) }
    Catch {
        $false
        Return # Exit function, return $false if the file does not download.
    }
    if (Test-Path -Path $outfile -PathType Leaf) {
        Write-Verbose "Found for vault Report from $datestamp."
        Trim-Filename $outfile
        Return # Exit function, return path to downloaded file.
    }
    else {
        $false
        Return # No file downloaded, return $false.
    }
}
 
<#
    Convert an Excel Spreadsheet to CSV file.
    Input: Full path to .xlsx file (The workbook)
    Input: Name of the worksheet to convert.
    Output: CSV file saved to same path as workbook.
#>
Function Convert-XLSX2CSV {
    Param (
        [Parameter(Mandatory=$true)][string]$xlsxPath
    )
    # Anything after the last \ is the file name.
    # At least dotSource things...
    if (-Not $xlsxPath.Contains("\")) {
        $xlsxPath = ".\" + $xlsxPath
    }
    # RegEx to extract path and filename.
    # $&: Entire match.
    # $1: Path to file.
    # $2: Filename without extension.
    # $3: File extension without dot.
    $pattern = '^(.*\\)(.*)\.\b(xlsx|xls)\b$'
    $outdir = $xlsxPath -replace $pattern, '$1'   
    $csvName = $xlsxPath -replace $pattern, '$2.csv'
    Write-Debug $xlsxPath
    Write-Debug $outdir
    Write-Debug $csvName
    # Setup the Excel Com Object
    $E = New-Object -ComObject Excel.Application
    $garbage = $E.Visible = $false
    $garbage = $E.DisplayAlerts = $false
    # Open the workbook
    $wb = $E.Workbooks.Open($xlsxPath)
    foreach ($ws in $wb.Worksheets)
    {
        $n = $excelFileName + "_" + $ws.Name
        $garbage = $ws.SaveAs($outdir + 'CSVTEMP' + $n + '_' + $csvName, 6)
    }
    $garbage = $E.Quit()
}

Function Cleanup-AutoReconcile {
    Param (
        [Parameter(Mandatory=$true)]$foundfile
    )
    Write-Verbose "Cleanup: Deleting $($foundfile.Fullpath)"
    Remove-Item -Path $foundfile.Fullpath
    $csv_wild = $foundfile.Path + "CSVTEMP*.csv"
    Write-Verbose "Cleanup: Deleting temporary CSV files from $($foundfile.Path)."
    Remove-Item -Path $csv_wild
}

Function Move-Reports ($reportDir) {
    # URL to SharePoint folder.
    $sharepoint = 'http://sharepoint.example.com/sites/Project42/Project%20Document%20Library/Reconciliation%20Reports/'
    # Create a .NET WebClient to upload files.
    $webclient = New-Object System.Net.WebClient
    $webclient.UseDefaultCredentials = $true
    # Create a WebServiceProxy to check in the files.
    $service = New-WebServiceProxy -UseDefaultCredential -uri 'http://sharepoint.example.com/_vti_bin/Lists.asmx?WSDL'
    Get-ChildItem -Path $reportDir -Filter "Report_*.csv" | ForEach-Object {
        Write-Verbose "Uploading $($_.Name)"
        $webclient.UploadFile($sharepoint + $_.Name, “PUT”, $_.FullName)
        $fileCheck = $sharepoint + $_.Name
        if ($service.CheckInFile($fileCheck,"PowerShell Upload","1")) {
            Write-Verbose "$($_.Name) Checked in to SharePoint."
            Remove-Item -Path $_.FullName
            Write-Verbose "Cleanup: $($_.FullName) was uploaded to SharePoint, local copy deleted."
        }
        else {
            Write-Error "Error: $($_.Name) could not be checked in to SharePoint."            
            Write-Error "Error: $($_.FullName) was not deleted."           
        }
    };
}
#-------------[Execution]------------------------------------------------------
# Today's date +10 days. For testing.
# $date = (Get-Date).AddDays(10)
$date = Get-Date
# Control the loop. Limit to previous 14 days.
$foundfile = $false
$count = 0
While (($foundfile -eq $false) -and ($count -lt 14)) {
    $foundfile = Get-vaultReport $date
    $date = $date.AddDays(-1)
    $count++
}
# If a vault report isn't found, throw error and exit.
if ($foundfile -eq $false) {
    Throw "Unable to find vault report in SharePoint."
}
Write-Verbose "Found vault Report: $($foundfile.Fullpath)"
Convert-XLSX2CSV $foundfile.Fullpath
# Build path to the CSV file created by Convert-XLSX2CSV.
$csvfile = 'CSVTEMP_Sheet1_' + $foundfile.Filename
$csvfile = $csvfile.Replace('.xlsx','.csv')
$csvfile = $foundfile.Path + $csvfile
if (Test-Path -Path $csvfile -PathType Leaf) {
    Write-Verbose "Found CSV: $csvfile"
}
else {
    Throw "File Not Found: $csvfile"
}
# Execute the report script.
if ($VerbosePreference) {
    Invoke-Expression "& `"$reportScript`" -vault $csvfile -workdir $workdir -Verbose"
} else {
    Invoke-Expression "& `"$reportScript`" -vault $csvfile -workdir $workdir"
}
# Move reports up to SharePoint
Move-Reports $workdir
# Cleanup after execution.
Cleanup-AutoReconcile $foundfile