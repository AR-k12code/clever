# clever
**These scripts come without warranty of any kind. Use them at your own risk. I assume no liability for the accuracy, correctness, completeness, or usefulness of any information provided by this site nor for any sort of damages using these scripts may cause.**

This script will pull from a persistent folder in Cognos so you no longer have to manage the Cognos reports for yourself.

You will need the CognosDownloader from ````https://github.com/AR-k12code/CognosDownloader````. Please follow the installation process in the Readme.

You will need to copy the CognosDefaults.ps1 to c:\scripts\CognosDefaults.ps1 and configure for your username and school database.

## Requirements
Git ````https://git-scm.com/download/win````

Powershell 7 ````https://github.com/PowerShell/powershell/releases````

pscp and sqlite3 binary (Included in project, feel free to check the hashes.)

SimplySQL Powershell Module. (Installation instructions below.)

## Installation
Open Powershell 7 Administrative Shell
````
Install-Module SimplySQL -Scope AllUsers -Force
mkdir c:\scripts
cd \scripts
git clone https://github.com/AR-k12code/clever.git
cd clever
Copy-Item sample_settings.ps1 settings.ps1
````

## settings.ps1
You will need to configure your Clever username and password. You can get that information from ````https://schools.clever.com/sync/settings````.

## Run
Since this script is not trying to fix email addresses it does not need to have access to Active Directory. It can be run without saving the password in Task Scheduler. However, it must be run under the same local Windows user that encrypted the password for the CognosDownloader. I highly suggest this be a service account that will survive longer than you.

Open a Powershell 7 window (pwsh.exe)
````
cd \scripts\clever
.\clever.ps1
````

## Scheduling Task
This must be done using the account you used to save your encrypted Cognos password.

* Open Task Scheduler
* New Basic Task
* Name it, Daily (or more), Set Time. (Please choose an offset on minutes like 8:23am)
* Action: Start a Program.
    * Program/script = "pwsh.exe"
    * Add arguments = "-ExecutionPolicy bypass -File c:\scripts\clever\clever.ps1"


## Still needing
- [ ] Needs a lot more error control

### Profit