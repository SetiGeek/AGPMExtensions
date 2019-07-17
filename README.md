# AGPMExtensions
Some Powershell scripts for the Advanced Group Policy Management tool

AGPM (Advance Group Policy Management Tool) is a versioning system for GPO, provided by Microsoft (https://docs.microsoft.com/en-us/microsoft-desktop-optimization-pack/agpm/)

## Warning
This script is done only by reverse engineering. I don't have Microsoft specifications for the GPOState.xml file (the central database for the archive).

## Available scripts
### Import-AGPMFromProduction
Script that import a GPO from the production into the AGPM archive.
GPO must already be under control with AGPM.

### Get-AGPMProductionStatus
Script that control version status into the archive, and compare it with the production environment.
A full report is done into a sheet that identify where are the differences: in the GPO name, in the computer or user policy, in the delegation, in the WMI filtering, or into the links list.
It doens't change anything, it just make an inventory.

## Author
@SetiGeek (https://github.com/SetiGeek)