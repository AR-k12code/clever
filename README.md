# clever
**These scripts come without warranty of any kind. Use them at your own risk. I assume no liability for the accuracy, correctness, completeness, or usefulness of any information provided by this site nor for any sort of damages using these scripts may cause.**

This script will pull from a persistent folder in Cognos so you no longer have to manage the Cognos reports for yourself.

You will need the CognosModule from ````https://github.com/AR-k12code/CognosModule````. Please follow the installation process in the Readme.

**DO NOT INSTALL THESE SCRIPTS TO A DOMAIN CONTROLLER.**

Create a dedicated VM running Windows Server 2019 or Windows 10 Pro 1809+ for your automation scripts.

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
Since this script is not trying to fix email addresses it does not need to have access to Active Directory. It can be run without saving the password in Task Scheduler. However, it must be run under the same local Windows user that encrypted the password for the CognosModule. I highly suggest this be a service account that will survive longer than you.

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
    * Program/script = "c:\Program Files\PowerShell\7\pwsh.exe"
    * Add arguments = "-ExecutionPolicy bypass -File c:\scripts\clever\clever.ps1"
    * Start in = "c:\scripts\clever"
** Quotes are not needed for add arguments or start in with Server 2022; may not be needed for other versions as well!

## Troubleshooting
Review the clever-log.log file.

Delete the database.sqlite file and try again.

Verify your username and password are correct in the settings.ps1 file.

Report any other issues here on github.

### Profit
