#Requires -Version 3.0

<#
.SYNOPSIS
Compares CSV-Report Report to WebService inventory. Requires PSSQLite module. 

.DESCRIPTION
Takes a CSV-Report as input and compares to WebService inventory retrived
via the WebService web service. Requires the PSSQLite module to be availalbe. 

3rd party modules such as PSSQLite are typically installed to 
$env:USERPROFILE\Documents\WindowsPowerShell\Modules.

See https://msdn.microsoft.com/en-us/library/dd878350(v=vs.85).aspx for 
information on installing a Powershell module for all users. 

Author: Andrew Zbikowski (andy@zibnet.us)
Author Website: http://andy.zibnet.us

.PARAMETER CSV-Report
REQUIRED: Path and filename of CSV-Report CSV report. Example: C:\temp\CSV-Report_Report.csv 

.PARAMETER supportgroup
OPTIONAL: SupportGroup to lookup in WebService. 
DEFAULT: INT-SERVER  

.PARAMETER workdir
OPTIONAL: Directory where temporary files and final reports will be written. 
DEFAULT: $env:USERPROFILE\Documents

.PARAMETER nocleanup
OPTIONAL: Use this switch and the temporary SQLite database will not be deleted. 
DEFAULT: Temp file cleanup will be performed unless the -nocleanup switch is specified. 

.OUTPUTS
CSV Reports are saved to $env:USERPROFILE\Documents unless the workdir paramater is provided. 
The temporary SQLite database will be saved to $env:USERPROFILE\Documents if the Cleanup
parameter is set to $false. 

.NOTES
Current Version: 0.9
v0.1 Changes, 2015-08-03: 
    * First attempt. 
v0.2 Changes, 2015-08-03: 
    * First working attempt using Where-Object. This takes hours. 
v0.3 Changes, 2015-08-04: 
    * Replaced Where-Object with a stream filter, this is faster but still too slow. 
    * Writing the CSV file every 1000 records adds to slowness, but allows to verify that things are working as expected.   
v0.4 Changes: 
    * Investigating options for implementing a SQL like join of custom PowerShell objects. 
v0.5 Changes, 2015-08-05: 
    * Why reinvent the wheel? Transform data, insert in SQLite using PSSQLite module, and do a SQL JOIN! Now the report takes minutes instead of hours.   
v0.6 Changes, 2015-08-06: 
    * Rookie mistake...got bit by case sensitivity. Changing FQDNs to lowercase before inserting into SQLite database.  
v0.7 Changes, 2015-08-06: 
    * Work around SQLite's JOIN limitations. SQLite does not implement RIGHT OUTER JOIN or FULL OUTER JOIN. Adding an additional SQL query to get all desired data, then combining results of the queries.
v0.8 Changes, 2015-08-07: 
    Now a proper reusable script with parameters and help. 
V0.9 Changes, 2015-08-07: 
    * Changed the format of this change log. 
    * Updated Test-Path error checking to include -PathType. I accidently fed a directory into the CSV file parameter and discovered this bug.  
    * Split single large CSV report into three smaller reports. Large CSV report is still generated as well. 
    * Updated names of database columns for a more readable report. 
    * Still lots of redundant code that could be turned into functions. 
    * Added PowerShell v3.0 requirement as that is what this script was developed in. 

PSSQLite is availble at https://github.com/RamblingCookieMonster/PSSQLite

.EXAMPLE
Run script normally. Silent execution for running as a scheduled task. 
ScriptName.ps1 -CSV-Report Filename.csv

.EXAMPLE
Runing with the -verbose parameter is reccomended when running from an interactive shell.  
ScriptName.ps1 -CSV-Report Filename.csv -verbose

.EXAMPLE
If you want even more output, add -debug as well.
ScriptName.ps1 -CSV-Report Filename.csv -verbose -debug

.EXAMPLE 
Change WebService support group.
ScriptName.ps1 -CSV-Report Filename.csv -SupportGroup Client

.EXAMPLE
Change work directory. 
ScriptName.ps1 -CSV-Report Filename.csv -WorkDir C:\TEMP\Example

.EXAMPLE
Do not delete the SQLite database file after the report is generated. 
This could be handy if you want to perserve the SQLite database and come up with your own reports.  
ScriptName.ps1 -CSV-Report Filename.csv -nocleanup

#>

#------------------[Parameters]------------------------------------------------
### Deal with user input. 
# Enable -Debug, -Verbose Paramaters. Write-Debug and Write-Verbose! 

[CmdletBinding()]
 Param (
    [Parameter(Mandatory=$true)][string]$CSV-Report = $( Read-Host "Enter path and filename of CSV-Report CSV file"),
    [Parameter()][string]$supportgroup = "INT-SERVER",
    [Parameter()][string]$workdir = "$env:USERPROFILE\Documents",
    [Parameter()][switch]$nocleanup = $false
)

#------------------[Debug]-----------------------------------------------------
Write-Debug "Parameter CSV-Report: $CSV-Report"
Write-Debug "Parameter supportgroup: $supportgroup"
Write-Debug "Parameter workdir: $workdir"
Write-Debug "Parameter nocleanup: $nocleanup"

#------------------[Initialisations]-------------------------------------------
# If PSSQLite isn't available, just stop and throw a bright red error message. 
Import-Module PSSQLite -ErrorAction Stop

#------------------[Parameter Validation]--------------------------------------
# Verify that $workdir exists. 
if (!(Test-Path -PathType Container $workdir.trim())) {
    Throw "Working Directory not found: $workdir"
}

# Verify that $CSV-Report file exists.
if (!(Test-Path -PathType Leaf $CSV-Report.trim())) {
    Throw "CSV-Report Report File not found: $CSV-Report"
}

# Verify $supportgroup is a valid string (under 50 or so characters.)
if ( $supportgroup.length -gt 50 ) {
    Throw "Invalid WebService Support Group: $supportgroup"
} 

#------------------[Configuration]---------------------------------------------

# Setup WebService web service
$global:URI = "http://WebService.example.com/WebServiceServerInfo.svc"
Write-Debug "WebService URI: $URI"

#------------------[Declarations]----------------------------------------------
# Start automatic setup. Boring Stuff. 
# ISO 8601 Date Format. Accept no substuties!  
$global:iso8601 = Get-Date -Format s
# Colon (:) isn't a valid character in file names. 
$global:iso8601 = $global:iso8601.Replace(":","_")
# Just YYYY-MM-DD
#$datestamp = $iso8601.Substring(0,10)

# SQLite Database filename. 
$SQLiteDBFile = "$workdir\$iso8601.sqlite"
Write-Verbose "SQLite Database: $SQLiteDBFile"

# Tracking performance is fun and gives the script something to output while working. 
$global:TotalExecutionTime = 0

#------------------[Debug]-----------------------------------------------------
Write-Debug "Datestamp $iso8601"
Write-Debug "SQLite Database: $SQLiteDBFile"
Write-Debug "Total Execution time: $global:TotalExecutionTime"

#------------------[Functions]-------------------------------------------------
# There doesn't appear to be anything here. 

#------------------[Execution]-------------------------------------------------
Write-Verbose "Creating termporary SQLite database." 
$executionTime = Measure-Command {
    ### Step 1: Create SQLite Database and Tables
    #### CSV-Report Table:
    ### CSV-Report Fields availalbe:
    # Safe,Device type,Policy ID,Target system address,Target system user name,
    #  Group name,Last accessed date,Last accessed by,Last modified date,
    #  Last modified by,Change failure,Verification failure,Failure reason,
    #  database,description,servicename,userdn,Name,logondomain
    ### CSV-Report Fields we're interested in: 
    # Safe,PolicyID,TargetSystem_Address,TargetSystem_UserName,LastAccessedDate,
    #  LastAccessedBy,ChangeFailure,VerificationFailure,FailureReason
    ## Primary Key: Target System Address + Target system User Name.

    $sql = "CREATE TABLE CSV-Report (
        CSV-Report_Safe VARCHAR(255),
        CSV-Report_PolicyID VARCHAR(255),
        CSV-Report_TargetSystem_Address VARCHAR(255),
        CSV-Report_TargetSystem_Username VARCHAR(255),
        CSV-Report_LastAccessedDate VARCHAR(255),
        CSV-Report_ChangeFailure VARCHAR(255),
        CSV-Report_VerificationFailure VARCHAR(255),
        CSV-Report_FailureReason VARCHAR(255)
    );"
    Invoke-SqliteQuery -Query $sql -Database $SQLiteDBFile

    #### WebService Table: 
    # FQDN, BusinessSegmentDescription, Domain, OSDescription, OSType, Sitecode,
    #  SupportEnvDescription, SupportStageDescription
    ## Primary Key: FQDN
    $sql = "CREATE TABLE WebService (
        WebService_FQDN VARCHAR(255) PRIMARY KEY,
        WebService_BusinessSegmentDescription VARCHAR(255),
        WebService_Domain VARCHAR(255),
        WebService_OSDescription VARCHAR(255),
        WebService_OSType VARCHAR(255),
        WebService_SiteCode VARCHAR(255),
        WebService_SupportEnvDescription VARCHAR(255),
        WebService_SupportStageDescription VARCHAR(255)
    );"
    Invoke-SqliteQuery -Query $sql -Database $SQLiteDBFile
} | Select -ExpandProperty TotalSeconds

Write-Verbose "Creating termporary SQLite database took $executionTime seconds." 
$global:TotalExecutionTime = $global:TotalExecutionTime + $executionTime
Write-Debug "Total Execution time: $global:TotalExecutionTime"

### Step 2: Obtain and transform data, insert into SQLite using transactions. 
$WebService = New-WebServiceProxy $URI -class server -Namespace webservice
Write-Verbose "Retrieving data from WebService."

# Search WebService and put result object in $WebServiceResults
$executionTime = Measure-Command {
    # Setup the connection to WebService
    $objResults = $WebService.WebServiceInfoBySupportedBy($SupportGroup)
} | Select -ExpandProperty TotalSeconds
Write-Verbose "Retrieving data from WebService took $executionTime seconds." 
$global:TotalExecutionTime = $global:TotalExecutionTime + $executionTime
Write-Debug "Total Execution time: $global:TotalExecutionTime"

Write-Verbose "Transforming data from WebService."
$executionTime = Measure-Command {
    $dataTable = ForEach ($oResult in $objResults) {
        [pscustomobject]@{
            WebService_FQDN = $oResult.FQDN.ToLower()
            WebService_BusinessSegmentDescription = $oResult.BusinessSegmentDescription 
            WebService_Domain = $oResult.Domain
            WebService_OSDescription = $oResult.OSDescription
            WebService_OSType = $oResult.OSType
            WebService_SiteCode = $oResult.SiteCode
            WebService_SupportEnvDescription = $oResult.SupportEnvDescription
            WebService_SupportStageDescription = $oResult.SupportStageDescription
        }
    }
    ### Get rid of WebService Results
    $objResults = $null
    # Create a data table for CSV-Report! 
    $dataTable = $dataTable | Out-DataTable
} | Select -ExpandProperty TotalSeconds

Write-Verbose "Transforming WebService data took $executionTime seconds." 
$global:TotalExecutionTime = $global:TotalExecutionTime + $executionTime
Write-Debug "Total Execution time: $global:TotalExecutionTime"

# Add to database using a transaction
Write-Verbose "Adding WebService data to SQLite."
$executionTime = Measure-Command {
    Invoke-SQLiteBulkCopy -DataSource $SQLiteDBFile -DataTable $dataTable -Table WebService -Force
    $dataTable = $null
} | Select -ExpandProperty TotalSeconds

Write-Verbose "Adding WebService data into SQLite took $executionTime seconds." 
$global:TotalExecutionTime = $global:TotalExecutionTime + $executionTime
Write-Debug "Total Execution time: $global:TotalExecutionTime"

Write-Verbose "Importing CSV-Report CSV Report."
# Get CSV-Report results
$executionTime = Measure-Command {
    $objResults = Import-Csv $CSV-Report
} | Select -ExpandProperty TotalSeconds

Write-Verbose "Importing CSV-Report CSV Report took $executionTime seconds." 
$global:TotalExecutionTime = $global:TotalExecutionTime + $executionTime
Write-Debug "Total Execution time: $global:TotalExecutionTime"

Write-Verbose "Transforming data from CSV-Report."
# Load CSV-Report data into a new object that we can stuff into SQLite. 
$executionTime = Measure-Command {
    $dataTable = ForEach ($caResult in $objResults) {
        [pscustomobject]@{
            CSV-Report_Safe = $caResult.Safe
            CSV-Report_PolicyID = $caResult."Policy ID"
            CSV-Report_TargetSystem_Address = $caResult."Target system address".ToLower()
            CSV-Report_TargetSystem_Username = $caResult."Target system user name"
            CSV-Report_LastAccessedDate = $caResult."Last accessed date"
            CSV-Report_ChangeFailure = $caResult."Last accessed by"
            CSV-Report_VerificationFailure = $caResult."Change failure"
            CSV-Report_FailureReason = $caResult."Failure reason"
        }
    } 
    ### We can get rid of $CSV-ReportResults now. 
    $objResults = $null
    # Create a data table for CSV-Report! 
    $dataTable = $dataTable | Out-DataTable
} | Select -ExpandProperty TotalSeconds

Write-Verbose "Transforming data from CSV-Report took $executionTime seconds." 
$global:TotalExecutionTime = $global:TotalExecutionTime + $executionTime
Write-Debug "Total Execution time: $global:TotalExecutionTime"

Write-Verbose "Adding CSV-Report data to SQLite."
# Add to database using a transaction
$executionTime = Measure-Command {
    Invoke-SQLiteBulkCopy -DataSource $SQLiteDBFile -DataTable $dataTable -Table CSV-Report -Force
    $dataTable = $null
} | Select -ExpandProperty TotalSeconds

Write-Verbose "Adding CSV-Report data into SQLite took $executionTime seconds." 
$global:TotalExecutionTime = $global:TotalExecutionTime + $executionTime
Write-Debug "Total Execution time: $global:TotalExecutionTime"

### Step 3: Use SQL to get the desired reports. 
Write-Debug "Cross your fingers and clench your sphincters! HERE WE GO!!!"

### Report 1: CSV-Report and WebService Matches
$sql = "SELECT * FROM WebService
    INNER JOIN CSV-Report
    ON WebService.WebService_FQDN=CSV-Report.CSV-Report_TargetSystem_Address;"

Write-Verbose "Executing SQL Report 1: CSV-Report and WebService Matches."

$executionTime = Measure-Command {
    $report1 = Invoke-SqliteQuery -Query $sql -Database $SQLiteDBFile
} | Select -ExpandProperty TotalSeconds

Write-Verbose "SQL Report 1 completed in $executionTime seconds." 
$global:TotalExecutionTime = $global:TotalExecutionTime + $executionTime
Write-Debug "Total Execution time: $global:TotalExecutionTime"

### Report 2: Only WebService records with no CSV-Report Match. 

$sql = "SELECT * FROM WebService
    LEFT JOIN CSV-Report
    ON WebService.WebService_FQDN=CSV-Report.CSV-Report_TargetSystem_Address
    WHERE CSV-Report.CSV-Report_TargetSystem_Address is Null
    ORDER BY WebService.WebService_FQDN;"

Write-Verbose "Executing SQL Report 2: WebService records with no CSV-Report match."
$executionTime = Measure-Command {
    $report2 = Invoke-SqliteQuery -Query $sql -Database $SQLiteDBFile
} | Select -ExpandProperty TotalSeconds

Write-Verbose "SQL Report 2 completed in $executionTime seconds." 
$global:TotalExecutionTime = $global:TotalExecutionTime + $executionTime
Write-Debug "Total Execution time: $global:TotalExecutionTime"

### Report 3: Only CSV-Report records with no WebService Match. 

$sql = "SELECT * FROM CSV-Report
    LEFT JOIN WebService
    ON CSV-Report.CSV-Report_TargetSystem_Address=WebService.WebService_FQDN
    WHERE WebService.WebService_FQDN is Null
    ORDER BY CSV-Report.CSV-Report_TargetSystem_Address;"

Write-Verbose "Executing SQL Report 3: CSV-Report records with no WebService match."
$executionTime = Measure-Command {
    $report3 = Invoke-SqliteQuery -Query $sql -Database $SQLiteDBFile
} | Select -ExpandProperty TotalSeconds

Write-Verbose "SQL Report 3 completed in $executionTime seconds." 
$global:TotalExecutionTime = $global:TotalExecutionTime + $executionTime
Write-Debug "Total Execution time: $global:TotalExecutionTime"

### Report 4: Everything
$report4 = $report1 + $report2 + $report3

### Step 5: Write the final reports. 
# Write CSV Reports

$executionTime = Measure-Command {
    $outfile = $workdir + "\Report_" + $iso8601 +"_CSV-Report_AND_WebService_Matched.csv"
    Write-Verbose "Writing CSV Report 1: CSV-Report and WebService Matches."
    $report1 | Export-Csv $outfile -NoTypeInformation
    Write-Verbose "CSV Report 1 saved to $outfile."
    $outfile = $workdir + "\Report_" + $iso8601 + "_WebService_No_CSV-Report_Match.csv"

    Write-Verbose "Writing CSV Report 2: WebService records with no CSV-Report match."
    $report2 | Export-Csv $outfile -NoTypeInformation
    Write-Verbose "CSV Report 2 saved to $outfile."
    $outfile = $workdir + "\Report_" + $iso8601 +"_CSV-Report_No_WebService_Match.csv"
    
    Write-Verbose "Writing CSV Report 3: CSV-Report records with no WebService match."
    $report3 | Export-Csv $outfile -NoTypeInformation
    Write-Verbose "CSV Report 3 saved to $outfile."

    $outfile = $workdir + "\Report_" + $iso8601 + "_AllRecords_CSV-Report_WebService.csv"
    Write-Verbose "Writing CSV Report 4: All Records."
    $report4 | Export-Csv $outfile -NoTypeInformation
    Write-Verbose "CSV Report 4 saved to $outfile."
} | Select -ExpandProperty TotalSeconds

Write-Verbose "Writing CSV Reports completed in $executionTime seconds." 

$global:TotalExecutionTime = $global:TotalExecutionTime + $executionTime

Write-Debug "Total Execution time: $global:TotalExecutionTime"

### Step 6: Cleanup temp files, the end. 

# Delete SQLite database
If (!$nocleanup) {
    Remove-Item -Path $SQLiteDBFile
    Write-Verbose "Deleted $SQLiteDBFile"
    Write-Debug "Deleted $SQLiteDBFile"
} Else {
    Write-Verbose "SQLite Database located at $SQLiteDBFile"
    Write-Debug "SQLite Database located at $SQLiteDBFile"
}

Write-Verbose "Total Execution Time: $global:TotalExecutionTime seconds."
#------------------[End of Script]---------------------------------------------