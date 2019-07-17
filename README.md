# AGPMExtensions
Some Powershell scripts for the Advanced Group Policy Management tool

AGPM (Advance Group Policy Management Tool) is a versioning system for GPO.

## Warning
This script is done only by reverse engineering. I don't have Microsoft specifications for the GPOState.xml file (the central database for the archive).

## Available scripts
### Import-AGPMFromProduction
Script that import a GPO from the production into the AGPM archive.
GPO must already be under control with AGPM.

### Check-AGPMProductionStatus
