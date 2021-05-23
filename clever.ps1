#Requires -Version 7.1
<#

Clever Automation Script
Craig Millsap

During testing of this script I highly suggest you pause your sync @ https://schools.clever.com/sync/settings

This script assumes you have the CognosDefaults.ps1 file properly configured.
This script does not attempt to fix your student email addresses. That process is included in the eSchoolUpload project.

#>

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
} else {
    . $PSScriptRoot\settings.ps1
}

#Check that the Cognos Downloader has been properly installed.
if (-Not(Test-Path $PSScriptRoot\..\CognosDownload.ps1)) {
    Write-Host "Error: Failed to find the CognosDownload.ps1 at c:\scripts\CognosDownload.ps1. Please follow the directions on the ARK12-Code Github for proper installation." -ForegroundColor Red
    exit(1)
}

if ([int](Get-Date -Format MM) -ge 6) {
    $schoolyear = [int](Get-Date -Format yyyy) + 1
} else {
    $schoolyear = [int](Get-Date -Format yyyy)
}

$reports = @{
    'enrollments' = @{ 'parameters' = 'p_fake=fake'; 'reportname' = 'enrollments' } #You must make the API believe you have provided some prompts.
    'schools' = @{ 'parameters' = ''; 'reportname' = 'schools' }
    'sections' = @{ 'parameters' = "p_year=$schoolyear"; 'reportname' = 'sections' }
    'students' = @{ 'parameters' = ''; 'reportname' = 'students' }
    'teachers' = @{ 'parameters' = ''; 'reportname' = 'teachers' }
}

if ($full_schedule -eq $True) {
    $reports.'enrollments' = @{ 'parameters' = ''; 'reportname' = 'enrollments_allterms' }
}

#Establish Session Only. Report parameter is required but we can provide a fake one for authentication only.
. $PSScriptRoot\..\CognosDownload.ps1 -report FAKE -EstablishSessionOnly

$results = $reports.Keys | ForEach-Object -Parallel {
    #report title
    $PSitem
    
    #pull in session to script block
    $incomingsession = $using:session
    
    #Pull in properties for each hashtable key.
    $options = ($using:reports).$PSItem

    #Run Cognos Download using incoming options.
    & $using:PSScriptRoot\..\CognosDownload.ps1 -report "$($options.reportname)" -cognosfolder "_Shared Data File Reports\Clever Files" -SessionEstablished -savepath "$using:PSScriptRoot\downloads" -reportparams "$($options.parameters)" -FileName "$($PSItem).csv" -ShowReportDetails -TrimCSVWhiteSpace -TeamContent

    if ($LASTEXITCODE -ne 0) { throw }
    
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

#If full schedule then we need to build the sql tables and match enrollment to sections.
if ($full_schedule) {

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
      `Terms`	int(4),
      UNIQUE (`School_id`,`Section_id`,`Student_id`,`Terms`)
    );' | Out-Null

    Start-SqlTransaction

    Invoke-SqlUpdate -Query '/* I need the terms to be grouped together to query later. Make a copy of the enrollments table with all terms grouped together. */
    REPLACE INTO `enrollments_grouped`
      SELECT School_id,Section_id,Student_id,group_concat(Marking_period,'''') as Terms
      FROM enrollments_csv_import
      GROUP BY Student_id,Section_id
      ORDER BY Marking_period;' | Out-Null

    Complete-SqlTransaction

    Invoke-SqlUpdate -Query '/* I need the terms to be grouped together to query later. Make a copy of the sections table with all terms grouped together. */
    DROP TABLE IF EXISTS `sections_grouped`;
    CREATE TABLE `sections_grouped` (
      `Terms` int(4),
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

    Invoke-SqlUpdate -Query '/* This requires that teachers have an email address. This is on purpose for my district because of sections that are placeholders. */
      INSERT INTO `enrollments`
      SELECT `enrollments_grouped`.`School_id`, `enrollments_grouped`.`Section_id`, `enrollments_grouped`.`Student_id`
      FROM `enrollments_grouped`
      INNER JOIN `sections_grouped` ON `enrollments_grouped`.`Section_id` = `sections_grouped`.`Section_id` AND `enrollments_grouped`.`Terms` = `sections_grouped`.`Terms`
      ORDER BY `enrollments_grouped`.`Student_id`,`sections_grouped`.`Period`;' | Out-Null
    
    $term1 = Invoke-SqlQuery -Query "SELECT Term_name,Term_start,Term_end FROM sections_csv_import WHERE Term_name = 1 LIMIT 1"
    $term2 = Invoke-SqlQuery -Query "SELECT Term_name,Term_start,Term_end FROM sections_csv_import WHERE Term_name = 2 LIMIT 1"
    $term3 = Invoke-SqlQuery -Query "SELECT Term_name,Term_start,Term_end FROM sections_csv_import WHERE Term_name = 3 LIMIT 1"
    $term4 = Invoke-SqlQuery -Query "SELECT Term_name,Term_start,Term_end FROM sections_csv_import WHERE Term_name = 4 LIMIT 1"
  
    $today = Get-Date
    if (([Int]$(Get-Date -Format MM) -gt 7) -And ($today -le [datetime]$term1.Term_end)) {
        $currentTerm = 1
    } elseif (($today -gt [datetime]$term1.Term_end) -And ($today -le [datetime]$term2.Term_end)) {
        $currentTerm = 2
    } elseif (($today -gt [datetime]$term2.Term_end) -And ($today -le [datetime]$term3.Term_end)) {
        $currentTerm = 3
    } else {
        $currentTerm = 4 #why do math when its the only option left?
    }

    Invoke-SqlUpdate -Query 'INSERT OR REPLACE INTO `enrollments`
    SELECT School_id,Section_id,Student_id
    FROM `enrollments_csv_import`
    WHERE Marking_period = ',$currentTerm,'
    ORDER BY Student_id' | Out-Null

    Complete-SqlTransaction

    Invoke-SqlQuery -Query "SELECT School_id,Section_id,Student_id FROM enrollments" | ConvertTo-Csv -UseQuotes AsNeeded -NoTypeInformation | Out-File $PSScriptRoot\files\enrollments.csv -Force

} else {
    Copy-Item $PSScriptRoot\downloads\enrollments.csv $PSScriptRoot\files\enrollments.csv -Force
}

Copy-Item $PSScriptRoot\downloads\schools.csv $PSScriptRoot\files\schools.csv -Force
Copy-Item $PSScriptRoot\downloads\students.csv $PSScriptRoot\files\students.csv -Force
Copy-Item $PSScriptRoot\downloads\teachers.csv $PSScriptRoot\files\teachers.csv -Force
Copy-Item $PSScriptRoot\downloads\sections.csv $PSScriptRoot\files\sections.csv -Force

try {
    Write-Host "Info: Uploading files to Clever..." -ForegroundColor YELLOW
    $exec = Start-Process -FilePath "$PSScriptRoot\bin\pscp.exe" -ArgumentList "-r -pw ""$cleverpassword"" -hostkey $cleverhostkey -batch $PSScriptRoot\files\ $($cleverusername)@sftp.clever.com:" -PassThru -Wait -NoNewWindow
    IF ($exec.ExitCode -ge 1) { Throw }
} catch {
    write-Host "ERROR: Failed to properly upload files to clever." -ForegroundColor RED
    exit(1)
}

exit
