<#
.SYNOPSIS
Import GPO links changes into AGPM: import from production.
This one edit directly the AGPM gpostate.xml file

.EXAMPLE
Import-AGPMFromProduction "Domain GPO 1" "user@domain.local"


#>
# This one force common parameter support, like -Debug that set $debugPreference to Inquire (and require a confirmation at every Write-Debug), 
# or -Verbose that enable the Write-Verbose output
[CmdletBinding()]
param(
    # The GPO to import
    [Parameter(Mandatory=$true)]
    [String]$GpoName,
    # User that will be traced into AGPM
    [Parameter(Mandatory=$true)]
    [String]$UserName,
    # Location of AGPM data
    [Parameter(Mandatory=$false)]
    [String]$ArchivePath = "$($env:ProgramData)\Microsoft\AGPM",
    # Comment to record in AGPM history entry
    [Parameter(Mandatory=$false)]
    [String]$AGPMComment = "Automatic import"
)
#Location of the AGPM State file (the central database - XML document)
$GpoStateFile = "$ArchivePath\gpostate.xml"

# Probably useless, but I like this :)
Write-Verbose "Verbose mode activated"
if ($DebugPreference -ne "SilentlyContinue") { Write-Verbose "Debug mode activated" }

<# 
.SYNOPSIS 
Translate a string into a date time object, as it's store in AGPM DB. 
#>
function Get-Time($string) {
    # Date time format used by AGPM 
    $dateformat = "yyyy-MM-ddTHH:mm:ss.FFFFFFFZ"

    ## 
    try {
        [datetime]::ParseExact($string, $dateformat, $null)
    } catch {
        throw "Date time format is not correct. Format must be '$dateFormat'"
    }
}
<# 
.SYNOPSIS
Format a date time object into a string, as supported by AGPM DB.
#>
function Format-Time($date) {
    # Whatever type of $date is, it won't break. Worst could be having a bad information, but not a real issue.
    $dateformat = "yyyy-MM-ddTHH:mm:ss.FFFFFFFZ"
    "{0:$dateformat}" -f ($date)
}

<#
.SYNOPSIS
Get the XML object of the AGPM Database
#>
function Get-AGPMGpoStateXML($xmlFile) {
    # Load the state file
    Write-Verbose "[Get-AGPMGpoStateXML] Converting file $xmlFile into XML object"
    if (-not $(Test-Path -Path $xmlFile -PathType Leaf)) {
        throw "File GpoState.xml ($xmlFile) not found. Check that you run as an administrator and check the installation path of AGPM server."
    }
    [XML]$xml = Get-Content $xmlFile
    $xml
}

<# 
.SYNOPSIS
Find a GPO into the AGPM database and return the node object
#>
function Get-AGPMGpo($xmlObject, $GpoId) {
    if ($null -eq $xmlObject) { throw "XML object must not be null." }
    if ($null -eq $GpoId) { throw "Gpo ID must not be null." }
    # Namespaces
    $agpmNs = "http://schemas.microsoft.com/MDOP/2007/02/AGPM"
    $ns = @{agpm = $agpmNs }
    Write-Verbose "[Get-AGPMGpo] ID = $GpoId / XML = $xmlObject"

    $node = Select-Xml -XPath  "descendant::agpm:GPO[@agpm:id='$GpoId']" -Xml $xmlObject -Namespace $ns -ErrorAction Stop
    if ($null -eq $node) { throw "GPO is not found in the AGPM archive." }
    $node.Node
}

<#
.SYNOPSIS
Create an history element, that can be stored in AGPM database.
This element doesn't contains the "Links" list
#>
function New-AGPMImportProductionItem($xmlObject, $user, $ArchiveId, $comment) {
    if ($null -eq $xmlObject) { throw "XML object must not be null."}
    if ($null -eq $user) { throw "User SID must not be null."}
    if ($null -eq $ArchiveId) { throw "Archive ID must not be null."}
    if ($null -eq $comment) { throw "Comment must not be null."}

    # Namespace to use
    $agpmNs = "http://schemas.microsoft.com/MDOP/2007/02/AGPM"
    $prefix = "agpm"
    
    Write-Verbose "[New-AGPMImportProductionItem] New item in GPO state with user '$user', ArchiveId '$ArchiveId', and comment '$comment'"
    # Creation of a new element "Item" in history
    $element = $xmlObject.CreateElement("Item", $agpmNs)
    # Add attribute called 'state'
    $attrib = $xmlObject.CreateAttribute($prefix, "state", $agpmNs)
    $attrib.Value = "IMPORTED_FROM_LIVE"
    $element.Attributes.Append($attrib) | Out-Null # ignore the output
    
    # Add attribute called 'type'
    $attrib = $xmlObject.CreateAttribute($prefix, "type", $agpmNs) # "CHECKED_IN"))
    $attrib.Value = "CHECKED_IN"
    $element.Attributes.Append($attrib) | Out-Null # ignore the output
    
    # Add attribute called 'time'    
    $attrib = $xmlObject.CreateAttribute($prefix, "time", $agpmNs) #$(Format-Time(Get-Date))))
    $attrib.Value = Format-Time((Get-Date).ToUniversalTime()) # Time is stored in UTC
    $element.Attributes.Append($attrib) | Out-Null # ignore the output 

    # Add attribute called 'user' 
    $attrib = $xmlObject.CreateAttribute($prefix, "user", $agpmNs) #User "S-1-5-21-2748441835-270625253-4050110434-23191"
    $attrib.Value = "$user"
    $element.Attributes.Append($attrib) | Out-Null # ignore the output

    # Add attribute called 'archiveId' 
    $attrib = $xmlObject.CreateAttribute($prefix, "archiveId", $agpmNs) #ArchiveId
    $attrib.Value = "{$ArchiveId}"
    $element.Attributes.Append($attrib) | Out-Null # ignore the output

    # Add attribute called 'comment'     
    $attrib = $xmlObject.CreateAttribute($prefix, "comment", $agpmNs) #comment
    $attrib.Value = "$comment"
    $element.Attributes.Append($attrib) | Out-Null # ignore the output

    Write-Verbose "[New-AGPMImportProductionItem] Element created is $element"
    $element
}

<#
.SYNOPSIS
Archive the current version of the GPO and store it into the AGPM storage.
The AGPM database is not updated. 
This returns the ID of the archive folder
#>
function New-AGPMGpoArchive($GpoId, $ArchivePath) {
    # Backup the GPO (to get the ArchiveId)
    Write-Verbose "[New-AGPMGpoArchive] Backup of Gpo $GpoId in progress into $ArchivePath"
    $ArchivedGPO = Backup-GPO -Guid $GpoId -Path $ArchivePath -ErrorAction Stop
    if ($null -eq $ArchivedGPO) { throw "The GPO backup returned a null value."}
    Write-Verbose "[New-AGPMGpoArchive] Backup Gpo $GpoId created in $ArchivedGPO"
    #$ArchivedGPO
    $GPReportFile = "$ArchivePath\{$($ArchivedGPO.id)}\gpreport.xml"
    if (-not (Test-Path -Path $GPReportFile -PathType Leaf -ErrorAction Stop)) { throw "GPReport file not found ($GPReportFile)."}
    [XML]$GPReport = Get-Content "$ArchivePath\{$($ArchivedGPO.id)}\gpreport.xml" -ErrorAction Stop
    Write-Verbose "[New-AGPMGpoArchive] GPreport parsed into $GPReport"
    $ArchivedGPO.Id
}

<#
.SYNOPSIS
Search into the Active Directory for a specific GPO and return the ID.
#>
function Find-GPOID($GpoName) {
    Write-Verbose "[Find-GPOID] Looking for the GPO $GpoName"
    try {
        $Gpo = Get-GPO -Name $GpoName -ErrorAction Stop
        Write-Verbose "[Find-GPOID] GPO with name $GpoName is found: $GPO"
        "{$($Gpo.Id)}"
    }
    catch {
        throw "GPO $GpoName not found. Check its existence and your access rights."
    }
}

<#
.SYNOPSIS
Find an AD OrganizationalUnit from its canonical name, and returns only Distinguished Name
#>
function Convert-ADCanonicalNameOu($canonicalName) {
    # Cache the OU list
    $ouList = Get-ADOrganizationalUnit -Properties CanonicalName,DistinguishedName -Filter "*" -ErrorAction Stop
    # Return the distinguished Name
    $OuItem = ($ouList | Where-Object { $_.CanonicalName -eq $canonicalName })
    if ($null -eq $OuItem) { throw "$canonicalName OU not found."}
    if ($OuItem.count -lt 1) { throw "Found less than 1 OU with the name '$canonicalName'" }    
    if ($OuItem.count -ne 1) { throw "found more than 1 OU with the name '$canonicalName'" }
    $OuItem.DistinguishedName
}

<#
.SYNOPSIS
Find the user SID from the AD.
User name must be 'user@domain.tld' format
#>
function Get-ADSid($username) {    
    $array = $username.Split("@")
    if ($array.Count -ne 2) { throw "$username must be in 'user@domain' format" }
    (Get-ADUser -Identity $array[0] -Server $array[1] -ErrorAction Stop).SID
}

## Algorithm
try {
    $Activity = "Importing GPO $GpoName into AGPM"
    $Operations = @(
        @{ Name = "Looking for the GPO"; Percent = 5},
        @{ Name = "Loading the AGPM Archive database"; Percent = 10},
        @{ Name = "Loading the GPO $GpoName history"; Percent = 15},
        @{ Name = "Archiving the current GPO state"; Percent = 30},
        @{ Name = "Finding $UserName in Active Directory"; Percent = 75},
        @{ Name = "Registering the entry in the history"; Percent = 80},
        @{ Name = "Updating links list"; Percent = 85},
        @{ Name = "Saving the AGPM database"; Percent = 90}
    )

    # Step 1 - Get the GPO ID
    Write-Progress -Activity $Activity -PercentComplete $Operations[0].Percent -CurrentOperation $Operations[0].Name
    $GpoId = (Find-GPOID $GpoName).toUpper()
    Write-Verbose "GPO ID = $GpoId"

    # Step 2 - Open the XML database file
    Write-Progress -Activity $Activity -PercentComplete $Operations[1].Percent -CurrentOperation $Operations[1].Name
    $GpoStateXML = Get-AGPMGpoStateXML $GpoStateFile
    Write-Verbose "GPO XML file from $GpoStateFile = $GpoStateXML"
    
    # Step 3 - Find the GPO Node
    Write-Progress -Activity $Activity -PercentComplete $Operations[2].Percent -CurrentOperation $Operations[2].Name
    $GpoNode = Get-AGPMGpo $GpoStateXML $GpoId

    # Step 4 - Create a new Archive
    Write-Progress -Activity $Activity -PercentComplete $Operations[3].Percent -CurrentOperation $Operations[3].Name
    $ArchiveId = New-AGPMGpoArchive $GpoId $ArchivePath

    # Step 5 - Find the operating user in Active Directory
    Write-Progress -Activity $Activity -PercentComplete $Operations[4].Percent -CurrentOperation $Operations[4].Name
    $UserSid = Get-ADSid $UserName

    # Step 6 - Create History entry
    Write-Progress -Activity $Activity -PercentComplete $Operations[5].Percent -CurrentOperation $Operations[5].Name
    $GpoHistoryItem = New-AGPMImportProductionItem $GpoStateXML $UserSid $ArchiveId $AGPMComment
    
    # Step 7 - Get the GPO link list and add them as child of the Item
    Write-Progress -Activity $Activity -PercentComplete $Operations[6].Percent -CurrentOperation $Operations[6].Name
    # Namespace to use
    $agpmNs = "http://schemas.microsoft.com/MDOP/2007/02/AGPM"
    $prefix = "agpm"
        
    Write-Verbose "[] New Links"
    # Add the base "Links" element
    $LinksElement = $GpoStateXML.CreateElement("Links", $agpmNs)

    $gpReportFile = "$ArchivePath\{$ArchiveId}\gpreport.xml"
    [XML]$gpReportXml = Get-Content $gpReportFile
    $gpReportXml.GPO.LinksTo | ForEach-Object {
        $path = Convert-ADCanonicalNameOu $_.SOMPath
        $enabled = [byte][System.Convert]::ToBoolean($_.Enabled)
        $enforced = [byte][System.Convert]::ToBoolean($_.NoOverride)

        Write-Verbose "[] New Link: $path - Enabled: $enabled / Enforced: $enforced"
        $LinkElement = $GpoStateXML.CreateElement("Link", $agpmNs)
        
        # Path attribute
        $attrib = $GpoStateXML.CreateAttribute($prefix, "path", $agpmNs)
        $attrib.Value = "$path"
        $LinkElement.Attributes.Append($attrib) | Out-Null
        
        # Enabled attribute
        $attrib = $GpoStateXML.CreateAttribute($prefix, "enabled", $agpmNs)
        $attrib.Value = "$enabled"
        $LinkElement.Attributes.Append($attrib) | Out-Null
        
        # Enforced attribute
        $attrib = $GpoStateXML.CreateAttribute($prefix, "enforced", $agpmNs)
        $attrib.Value = "$enforced"
        $LinkElement.Attributes.Append($attrib) | Out-Null
        
        # Add the link into links list
        Write-Verbose "[] Adding the link to links element"
        $LinksElement.AppendChild($LinkElement) | Out-Null
    }
    Write-Verbose "[] Adding the links list into the AGPM new history item"
    $GpoHistoryItem.AppendChild($LinksElement) | Out-Null

    # Create new Item element in the history
    Write-Verbose "Adding new element into the XML tree"
    $GpoNode.History.PrependChild($GpoHistoryItem) | Out-Null
    # Set the state of the GPO
    Write-Verbose "Updating status of the current GPO in AGPM"
    $GpoNode.State.user = $GpoHistoryItem.user
    $GpoNode.State.time = $GpoHistoryItem.time
    $GpoNode.State.comment = $GpoHistoryItem.comment

    # Step 8 - Save the database file
    Write-Progress -Activity $Activity -PercentComplete $Operations[7].Percent -CurrentOperation $Operations[7].Name
    Write-Verbose "Saving the XML file"
    $GpoStateXML.Save($GpoStateFile)

    Write-Host "$GpoName successfuly imported."
}
catch {
    Write-Error "Import failed. $($_.Exception.ItemName) $($_.Exception.Message)"
    if ($null -ne $ArchiveId) {
        #drop the backup, because it's not used in AGPM anymore.
        $ArchiveItemPath = "$ArchivePath\{$ArchiveId}"
        if (Test-Path -Path $ArchiveItemPath -PathType Container) {
            Write-Verbose "[ERROR HANDLING] Dropping GPO Archive folder"
            Get-ChildItem -Path $ArchiveItemPath -Force -Recurse -ErrorAction Stop | Remove-Item -Force -Recurse -ErrorAction Stop
            Remove-Item -Path $ArchiveItemPath -Force -ErrorAction Stop 
        }
    }
}