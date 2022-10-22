#Requires -Version 7.1
#Requires -Modules SimplySQL,CognosModule

Param(
    [parameter(mandatory=$false)][switch]$SkipUpload,
    [parameter(mandatory=$false)][switch]$SkipDownload
)

<#

Clever Automation Script
Craig Millsap
Gentry Public Schools/CAMTech Computer Services LLC
 ___ _____ ___  ___   ___   ___    _  _  ___ _____         
/ __|_   _/ _ \| _ \ |   \ / _ \  | \| |/ _ \_   _|        
\__ \ | || (_) |  _/ | |) | (_) | | .` | (_) || |          
|___/ |_| \___/|_|   |___/ \___/  |_|\_|\___/ |_|

 ___ ___ ___ _____   _____ _  _ ___ ___   ___ ___ _    ___ 
| __|   \_ _|_   _| |_   _| || |_ _/ __| | __|_ _| |  | __|
| _|| |) | |  | |     | | | __ || |\__ \ | _| | || |__| _| 
|___|___/___| |_|     |_| |_||_|___|___/ |_| |___|____|___|
   
Please see https://github.com/AR-k12code/clever for more information.

During testing of this script I highly suggest you pause your sync @ https://schools.clever.com/sync/settings

This script requires you to have the CognosDefaults.ps1 file properly configured.
This script does not attempt to fix your student email addresses. That process is included in the eSchoolUpload
project and would need to be run prior to this script.

#>

try { Start-Transcript "$PSScriptRoot\clever-log.log" -Force } catch {
    Stop-Transcript; Start-Transcript "$PSScriptRoot\clever-log.log" -Force
}
$cleverhostkey = '76:0c:bb:e5:f7:df:97:c3:f2:77:0d:9a:2e:d7:92:18'

#Check for SimplySQL Module
try {
    Import-Module SimplySQL
} catch {
    write-host "Error: Failed to import SimplySQL module. You must install this module for this script to work." -ForegroundColor Red
    write-host "Info: Please run ""Install-Module SimplySQL -Scope AllUsers -Force"" from an administrator Powershell shell."
    exit(1)
}

#Check for Settings File
if (-Not(Test-Path $PSScriptRoot\settings.ps1)) {
    write-host "Error: Failed to find the settings.ps1 file. You can use the sample_settings.ps1 file as an example." -ForegroundColor Red
    exit(1)
} else {
    . $PSScriptRoot\settings.ps1
}

#Check that the Cognos Downloader has been properly installed.
try {
    Import-Module -Name CognosModule -ErrorAction STOP
} catch {
    Write-Error "Failed to find the CognosModule. https://github.com/AR-k12code/CognosModule Please follow the directions on the ARK12-Code Github for proper installation." -ForegroundColor Red
    exit(1)
}

#Required folders
if (-Not(Test-Path "$PSScriptRoot\downloads")) { New-Item -Path $PSScriptRoot\downloads -ItemType directory }
if (-Not(Test-Path "$PSScriptRoot\files")) { New-Item -Path $PSScriptRoot\files -ItemType directory }

#Current school year for pulling sections
if ([int](Get-Date -Format MM) -ge 7) {
    $schoolyear = [int](Get-Date -Format yyyy) + 1
} else {
    $schoolyear = [int](Get-Date -Format yyyy)
}

#process existing file on disk.
if (-Not($SkipDownload)) {

    #Establish Session Only. Report parameter is required but we can provide a fake one for authentication only.
    try {
        if (-Not($CognosConfigName)) {
            $CognosConfigName = "DefaultConfig"
        }
        Connect-ToCognos -ConfigName $CognosConfigName
    } catch {
        Write-Error "Failed to authenticate to Cognos."
        exit 1
    }

    $reports = @{
        'enrollments' = @{ 'parameters' = ''; 'reportname' = 'enrollments_markingperiod' }
        'schools' = @{ 'parameters' = ''; 'reportname' = 'schools' }
        'sections' = @{ 'parameters' = "p_year=$schoolyear"; 'reportname' = 'sections' }
        'students' = @{ 'parameters' = ''; 'reportname' = 'students_v2' }
        'teachers' = @{ 'parameters' = ''; 'reportname' = 'teachers' }
    }

    $results = $reports.Keys | ForEach-Object -Parallel {
        
        #report title
        $PSitem
        
        #pull in session to script block
        $CognosSession = $using:CognosSession
        $CognosDSN = $using:CognosDSN
        $CognosProfile = $using:CognosProfile
        $CognosUsername = $using:CognosUsername
        
        #Pull in properties for each hashtable key.
        $options = ($using:reports).$PSItem

        #Run Cognos Download using incoming options.
        try {
            Save-CognosReport -report "$($options.reportname)" -cognosfolder "_Shared Data File Reports\Clever Files" -savepath "$using:PSScriptRoot\downloads" -reportparams "$($options.parameters)" -FileName "$($PSItem).csv" -TeamContent
        } catch {
            Write-Output "$PSItem"
            throw
        }
        
    } -AsJob -ThrottleLimit 5 | Wait-Job #Please don't overload the Cognos Server.

    $results.ChildJobs | Where-Object { $PSItem.State -eq "Completed" } | Receive-Job

    #Output any failed jobs information.
    $failedJobs = $results.ChildJobs | Where-Object { $PSItem.State -ne "Completed" }
    $failedJobs | ForEach-Object {
        $PSItem | Receive-Job
    }

    if (($failedJobs | Measure-Object).count -ge 1) {
        Write-Host "Failed running", (($failedJobs | Measure-Object).count), "jobs." -ForegroundColor RED
        exit(2)
    }

}

#If full schedule then we need to build the sql tables and match enrollment to sections.
$database = "$PSScriptRoot\database.sqlite"

$sql_import = @"
drop table if exists enrollments_csv_import;
.mode csv
.separator ,
.import downloads\\enrollments.csv enrollments_csv_import
drop table if exists sections_csv_import;
.mode csv
.separator ,
.import downloads\\sections.csv sections_csv_import
"@

$sql_import | & $PSScriptRoot\bin\sqlite3.exe $database

Open-SQLiteConnection -DataSource $database

Invoke-SqlUpdate -Query '/* FINAL REVISION! */
    /* To ensure the table columns are correct I am dropping the entire table instead of truncating. */
    DROP TABLE IF EXISTS `enrollments_grouped`;
    CREATE TABLE `enrollments_grouped` (
      `School_id`	int(4),
      `Section_id` bigint(32),
      `Student_id` int(10),
      `Terms` TEXT,
      UNIQUE (`School_id`,`Section_id`,`Student_id`,`Terms`)
    );' | Out-Null

Start-SqlTransaction

Invoke-SqlUpdate -Query '/* I need the terms to be grouped together to query later. Make a copy of the enrollments table with all terms grouped together. */
    REPLACE INTO `enrollments_grouped`
      SELECT School_id,Section_id,Student_id,group_concat(Marking_period,'''') as Terms
      FROM (SELECT * FROM enrollments_csv_import ORDER BY Marking_period)
      GROUP BY Student_id,Section_id' | Out-Null

Complete-SqlTransaction

Invoke-SqlUpdate -Query '/* I need the terms to be grouped together to query later. Make a copy of the sections table with all terms grouped together. */
    DROP TABLE IF EXISTS `sections_grouped`;
    CREATE TABLE `sections_grouped` (
      `Terms` TEXT,
      `School_id` int(4),
      `Section_id` bigint(32),
      `Teacher_id` varchar(32),
      `Name` varchar(64),
      `Section_number` varchar(32),
      `Grade` varchar(15),
      `Course_name` varchar(64),
      `Course_number` varchar(32),
      `Course_description` varchar(64),
      `Period` varchar(10),
      `Subject` varchar(32),
      `Term_start` varchar(10),
      `Term_end` varchar(10),
      UNIQUE (`Section_id`,`Terms`)
    );' | Out-Null

    Start-SqlTransaction

Invoke-SqlUpdate -Query 'REPLACE INTO sections_grouped
      SELECT group_concat(Term_name,'''') as `Terms`,`School_id`, `Section_id`, `Teacher_id`, `Name`, `Section_number`, `Grade`, `Course_name`, `Course_number`, `Course_description`, `Period`, `Subject`, `Term_start`, `Term_end`
      FROM `sections_csv_import`
      GROUP BY `Section_id`
      ORDER BY `Terms`' | Out-Null

Complete-SqlTransaction

Invoke-SqlUpdate -Query 'DROP TABLE IF EXISTS `enrollments`;
    CREATE TABLE `enrollments` (
      `School_id` int(4),
      `Section_id` bigint(32),
      `Student_id` int(10),
      UNIQUE (`School_id`,`Section_id`,`Student_id`)
    );' | Out-Null

Start-SqlTransaction

Invoke-SqlUpdate -Query '/* This copies enrollments for students who are enrolled for the entire class terms. */
      INSERT INTO `enrollments`
      SELECT `enrollments_grouped`.`School_id`, `enrollments_grouped`.`Section_id`, `enrollments_grouped`.`Student_id`
      FROM `enrollments_grouped`
      INNER JOIN `sections_grouped` ON `enrollments_grouped`.`Section_id` = `sections_grouped`.`Section_id` AND `enrollments_grouped`.`Terms` = `sections_grouped`.`Terms`
      ORDER BY `enrollments_grouped`.`Student_id`,`sections_grouped`.`Period`;' | Out-Null


<#
    Get the Terms for each building.
#>

# We need to find all of the current terms and insert them into the enrollments table.
$dbTerms = Invoke-SqlQuery -Query "SELECT DISTINCT School_id,Term_name,Term_start,Term_end FROM sections_csv_import" | Select-Object -Property School_id,Term_name,Term_start,Term_end,
@{ Name = "Start"; Expression = { (Get-Date $PSitem.Term_start) } },
@{ Name = "End"; Expression = { (Get-Date $PSitem.Term_end) } } | Sort-Object -Property School_id,Start

#if there is a term 1 with a date in the future then we need to exclude any older data.
#This means new Master schedule has been defined. Maybe not for every campus but its time to move forward.
if ($dbTerms | Where-Object { $PSItem.Term_name -eq 1 -and $PSItem.Start -gt (Get-Date).AddDays(-1) }) {
    Write-Host "Info: New master schedule information available. Filtering to use new school year only."
    $dbTerms = $dbTerms | Where-Object { $PSItem.Start -gt (Get-Date).AddDays(-1) }
}

#we can order terms to find the current, future, then past by finding the ones with a future ending and then find the ones that have already ended.
$terms = @()
$terms += $dbTerms | Where-Object { $PSitem.End -gt (Get-Date) }
$terms += $dbTerms | Where-Object { $PSitem.End -lt (Get-Date) }

#digit terms only. (1,2,3,4,etc)
$digitTerms = $terms | Where-Object { $psitem.Term_name -match "^\d+$" } | Group-Object -Property School_id -AsHashTable

#prefixed terms. (A1,A2,A3,A4,R1,R2,R3,R4)
$prefixTerms = @{}
$terms | Where-Object { $psitem.Term_name -match "^\D\d+$" } | Select-Object *,@{ Name = "Prefix"; Expression = { ($PSitem.Term_name)[0]} } | Group-Object -Property School_id,Prefix | ForEach-Object { $prefixTerms.($PSitem.Name) = $PSitem.Group }

<#
    Deal with Enrollments
#>

$enrollmentsCurrentTermOnly = [System.Collections.ArrayList]@()

$digitTerms.Keys | ForEach-Object {
    $currentTerm = $digitTerms.($PSitem)[0]

    Write-Host "Info: Inserting term $($currentTerm.'Term_name') ($($currentTerm.Term_start) - $($currentTerm.Term_end)) enrollments for $($currentTerm.'School_id')"
    Invoke-SqlUpdate -Query "REPLACE INTO enrollments
        SELECT
            School_id,
            Section_id,
            Student_id
        FROM enrollments_csv_import
        WHERE School_id = $($currentTerm.School_id)
        AND Marking_period = ""$($currentTerm.Term_name)"""

    Invoke-SqlQuery -Query "SELECT
            School_id,
            Section_id,
            Student_id
        FROM enrollments_csv_import
        WHERE School_id = $($currentTerm.School_id)
        AND Marking_period = ""$($currentTerm.Term_name)""" | ForEach-Object {
            $enrollmentsCurrentTermOnly.Add($PSitem) | Out-Null
        }

}

$prefixTerms.Keys | ForEach-Object { 
    $currentTerm = $prefixTerms.($PSitem)[0]

    Write-Host "Info: Inserting term $($currentTerm.'Term_name') ($($currentTerm.Term_start) - $($currentTerm.Term_end)) enrollments for $($currentTerm.'School_id')"
    Invoke-SqlUpdate -Query "REPLACE INTO enrollments
        SELECT
            School_id,
            Section_id,
            Student_id
        FROM enrollments_csv_import
        WHERE School_id = $($currentTerm.School_id)
        AND Marking_period = ""$($currentTerm.Term_name)"""

    Invoke-SqlQuery -Query "SELECT
            School_id,
            Section_id,
            Student_id
        FROM enrollments_csv_import
        WHERE School_id = $($currentTerm.School_id)
        AND Marking_period = ""$($currentTerm.Term_name)""" | ForEach-Object {
            $enrollmentsCurrentTermOnly.Add($PSitem) | Out-Null
        }

}

Complete-SqlTransaction


Write-Host "Creating enrollments.csv file."

if ($current_term_only) {
    $enrollmentsCurrentTermOnly | Export-Csv -UseQuotes AsNeeded -NoTypeInformation -Path "$PSScriptRoot\files\enrollments.csv" -Force
} else {
    Invoke-SqlQuery -Query "SELECT School_id,Section_id,Student_id FROM enrollments" | Export-Csv -UseQuotes AsNeeded -NoTypeInformation -Path "$PSScriptRoot\files\enrollments.csv" -Force
}

<#
    Sections
#>
Write-Host "Creating sections.csv file."
Start-SqlTransaction
$sections = [System.Collections.ArrayList]@()

$digitTerms.Keys | ForEach-Object {

    $digitTermsAlreadyQueried = @()
    
    $digitTerms.$PSitem | ForEach-Object {

        Invoke-SqlQuery -Query (
            "SELECT * FROM sections_csv_import
            WHERE School_id = $($PSitem.School_id) AND Term_name = $($PSItem.Term_name)
            AND Section_id NOT IN (SELECT Section_id FROM sections_csv_import WHERE School_id = $($PSitem.School_id) AND Term_name IN (" + ($digitTermsAlreadyQueried -join ',') + "))"
        ) | ForEach-Object {
            $sections.Add($PSitem) | Out-Null
        }
            
        $digitTermsAlreadyQueried += $PSitem.Term_name
    }

}

$prefixTerms.Keys | ForEach-Object {
        
    $prefixTermsAlreadyQueried = @()

    $prefixTerms.$PSitem | ForEach-Object {

        Invoke-SqlQuery -Query (
            "SELECT * FROM sections_csv_import
            WHERE School_id = $($PSitem.School_id) AND Term_name =""$($PSItem.Term_name)""
            AND Section_id NOT IN (SELECT Section_id FROM sections_csv_import WHERE School_id = $($PSitem.School_id) AND Term_name IN (""" + ($prefixTermsAlreadyQueried -join '","') + """))"
        ) | ForEach-Object {
            $sections.Add($PSitem) | Out-Null
        }
            
        $prefixTermsAlreadyQueried += $PSitem.Term_name
    }

}

Complete-SqlTransaction

$sections | Export-CSV -UseQuotes AsNeeded -NoTypeInformation -Path "$PSScriptRoot\files\sections.csv" -Force

Copy-Item $PSScriptRoot\downloads\schools.csv $PSScriptRoot\files\schools.csv -Force
Copy-Item $PSScriptRoot\downloads\students.csv $PSScriptRoot\files\students.csv -Force
Copy-Item $PSScriptRoot\downloads\teachers.csv $PSScriptRoot\files\teachers.csv -Force

try {
    if ($SkipUpload) { exit 0 }
    Write-Host "Info: Uploading files to Clever..." -ForegroundColor YELLOW
    $exec = Start-Process -FilePath "$PSScriptRoot\bin\pscp.exe" -ArgumentList "-r -pw ""$cleverpassword"" -hostkey $cleverhostkey -batch $PSScriptRoot\files\ $($cleverusername)@sftp.clever.com:" -PassThru -Wait -NoNewWindow
    IF ($exec.ExitCode -ge 1) { Throw }
} catch {
    write-Host "ERROR: Failed to properly upload files to clever." -ForegroundColor RED
    Close-SqlConnection
    exit(3)
}

Close-SqlConnection

exit
