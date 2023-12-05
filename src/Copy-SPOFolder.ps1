<#
.SYNOPSIS
    Copy a folder and its contents from a source folder to a target folder, including all available versions of each file. The target folder must exist. 
    Can either be run from the command line with a source and target, or will read source and target values from a SharePoint list.
    The source and target can be URLs pasted from the address bar of a browser open to a library or folder. Does not (yet) support links copied using the "Copy Link" button.
    Source repo: https://github.com/erobillard/M365/ 

.DESCRIPTION
    To provide source and target values, either:
        - Set up a list with 3 columns to specify the source, target and action. All rows where Action is Copy will be processed. Column names may be specified as parameters.
        - Command line parameters can indicate the source and target URLs.
    Constraints: 
        The target folder(s) must exist.
        The account executing the cmdlet requires read permissions (Member or Reader) for each source and read-write (Member or Contributor) permissions for each target. If reading from a list, read-write permissions on the list. 
        The cmdlet will not copy from or to the tenant's root SPO Site (e.g.: the top-level site found at https://contoso.sharepoint.com/).

.PARAMETER Verbose
    Optional, -Verbose displays detailed output if present

.PARAMETER ReadOnly
    Optional, -ReadOnly will run without executing any actual moves or row updates. Output includes a count of items that would be affected.

.PARAMETER SiteUrl 
    Deprecated, use ListUrl instead.

.PARAMETER ListUrl
    Either ListUrl or both Source and Target are required. ListUrl is the path to the list containing Source, Target, and Action columns. The site containing the list is used to connect for authentication.
    E.g.: -ListUrl "https://contoso.sharepoint.com/sites/SharePointTools/Copy%20a%20Folder"

.PARAMETER SourceColumnName
    Optional. The name of the column in the list referenced by ListUrl where source addresses are found. Default: SourceUrl
    E.g.: -SourceColumnName "Origin"

.PARAMETER TargetColumnName
    Optional. The name of the column in the list referenced by ListUrl where target addresses are found. Default: TargetUrl
    E.g.: -TargetColumnName "Destination"

.PARAMETER ActionColumnName
    Optional. The name of the column in the list referenced by ListUrl where the action/status ("Copy") is found. Default: Action
    E.g.: -ActionColumnName "Status"

.PARAMETER Source
    Required when not reading from a SPList, the URL of the library or folder to copy FROM. The site at this location is used for authentication. 
    E.g.: -Source "https://contoso.sharepoint.com/sites/SiteA/Shared%20Documents/General"

.PARAMETER Target
    Required when not reading from a SPlist, the URL of the library or folder to copy TO. 
    E.g.: -Source "https://contoso.sharepoint.com/sites/SiteB/Shared%20Documents/General"

.EXAMPLE
    Copy-SPOFolder -Verbose -ListUrl "https://contoso.sharepoint.com/sites/SharePointTools/Copy a Folder"
    Copy-SPOFolder -Verbose -ListUrl "https://contoso.sharepoint.com/sites/SharePointTools/Copy%20a%20Folder" -SourceColumnName "Source URL" -TargetColumnName "Target URL" -ActionColumnName "Action"
    Copy-SPOFolder -Verbose -Source "https://contoso.sharepoint.com/sites/SourceSite/Shared Documents" -Target "https://contoso.sharepoint.com/sites/TargetSite/Shared Documents" 
    Copy-SPOFolder -Source "https://contoso.sharepoint.com/sites/SourceSite/Shared%20Documents/General" -Target "https://contoso.sharepoint.com/sites/TargetSite/Shared%20Documents/General" 
    Copy-SPOFolder -Source https://contoso.sharepoint.com/sites/SourceSite/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FSourceSite%2FShared%20Documents%2FGeneral%2FTest%20Data -Target "https://contoso.sharepoint.com/sites/TargetSite/Shared Documents/General/Test Data" 

.NOTES
    To read source and target values from a SPList, your minimal path to awesome is: 
        1. Create and configure a SharePoint List: 
            a) Create a SharePoint List with the Blank List template named "Copy a Folder" 
            b) Open List Settings and rename the Title column to SourceUrl. 
            c) Create a Single line of text column named TargetUrl. Rename this column after creation if you prefer a different display name. 
            d) Create a Text column called Action. Set the column's default value to Copy
        2. If you chose different column names or labels in Step 1, locate the section "# Define variables" below and update the values as needed.
        3. Add a few rows with valid source and target values for testing. Note that the target library or subfolder must exist, create as needed.
        4. Install PnP PowerShell
        5. Execute the script, e.g.:  Copy-SPOFolder -verbose -siteUrl "https://contoso.sharepoint.com/sites/SharePointTools" 
    
.AUTHOR
    Eli Robillard, https://github.com/erobillard

.LASTEDIT
    2023-11-22 Eli Robillard    Version: 1.0.0.0    Refactored somewhat and converted to a cmdlet
    2023-11-28 Eli Robillard    Version: 1.0.0.1    See version notes.

.VERSION
    1.0.0.0 Initial Release
    1.0.0.1 Moved the URL processing functions to .\SPO-UrlMethods.ps1
            Added ListUrl parameter so we don't need to rely on an internal var for the List name. 
            Added parameters for SourceColumnName, TargetColumnName, and ActionColumnName - makes it easier to use non-default names.
            Added ReadOnly parameter to run with no changes.
            Added copyStatusValue var so a value other than "Copy" can be used for other languages.
            Display parameter settings at statup when -Verbose is present. This helps identify issues with column name casing.
    1.0.0.2 Fixed handling of ListUrl parameter, list name was not being extracted correctly, now looks for the element after the /Lists/ segment.
#>

param (
    [switch] $Verbose,
    [switch] $ReadOnly,
    [string] $Source,
    [string] $Target,
    [string] $SiteUrl,
    [string] $ListUrl,
    [string] $SourceColumnName,
    [string] $TargetColumnName,
    [string] $ActionColumnName
)

# Using the PnP Powershell library for connection and copy commands
Import-Module PnP.PowerShell

# Required to decode escaped characters in the URL
Add-Type -AssemblyName System.Web

# Include the file with the URL parsing and translation functions
. .\SPO-UrlMethods.ps1

if ($Verbose) { Write-Host "Copy-SPOFolder Begin" }

# Define variables
$listName = "Copy a Folder"
$listSiteUrl = ""
$copyStatusValue = "Copy"
$newStatusValue = "Copy-Complete"
$listExecution = $true
$itemCount=0

# Read parameters into vars
if ($PSBoundParameters.ContainsKey('ListUrl') -and $ListUrl -ne $true) { 
    $listSiteUrl = Get-SPWebPath($ListUrl)
    $listName = Get-ListName($ListUrl)
    if ($Verbose) { 
        Write-Host "ListUrl was provided: " $ListUrl 
        Write-Host "List site: " $listSiteUrl 
        Write-Host "List name: " $listName
    }
}
elseif ($PSBoundParameters.ContainsKey('SiteUrl') -and $SiteUrl -ne $true) { 
    if ($Verbose) { Write-Host "SiteUrl was provided: " $SiteUrl }
    $listSiteUrl = $SiteUrl
}
else { 
    if($PSBoundParameters.ContainsKey('Source')) {
        $listExecution = $false
        $listSiteUrl = Get-SPWebPath($Source)
        $sourceUrl = $Source
        $targetUrl = $Target
        if ($Verbose) { Write-Host "Site: " $listSiteUrl " || Source: " $sourceUrl " || Target: " $targetUrl }
    }
    else {
        Write-Host "Error: At least one of these as parameters must be used: -ListUrl [value] -SiteUrl [value] OR -Source [value]" -ForegroundColor Red
        return
    }
}
# Set name of Source column
if($PSBoundParameters.ContainsKey('SourceColumnName')) {
    $sourceUrlColumnName = $SourceColumnName
}
else { $sourceUrlColumnName = "Title" }
#Set name of Target column
if($PSBoundParameters.ContainsKey('TargetColumnName')) {
    $targetUrlColumnName = $TargetColumnName
}
else { $targetUrlColumnName = "TargetUrl" }
# Set name of Action or Status column
if($PSBoundParameters.ContainsKey('ActionColumnName')) {
    $statusColumnName = $ActionColumnName
}
else { $statusColumnName = "Action" }

if ($ReadOnly) { Write-Host "ReadOnly specified, no files will be copied." }
if ($Verbose) { 
    Write-Host "Verbose mode" $Verbose
    Write-Host "Source column name:" $sourceUrlColumnName
    Write-Host "Target column name:" $targetUrlColumnName
    Write-Host "Action column name:" $statusColumnName
}

# Connect to the site that contains the list with the library mappings
$connectionSiteUrl = Connect-PnPOnline -Url $listSiteUrl -Interactive

if ($listExecution) {
    # Get all items from the list where status is "Copy"
    $listItems = Get-PnPListItem -List $listName -Query "<View><Query><Where><Eq><FieldRef Name='$statusColumnName'/><Value Type='Text'>$copyStatusValue</Value></Eq></Where></Query></View>"
}
else {
    # Create a one-row array using the command-line parameters. 
    if ($Verbose) { Write-Host "Creating the array:" }
    $listItems = @{ $sourceUrlColumnName = $sourceUrl; $targetUrlColumnName = $targetUrl; $statusColumnName = $copyStatusValue }
    if ($Verbose) { Write-Host $listItems }
}

# Loop through each item
foreach ($item in $listItems) {
        # Process the source URL
        $sourceUrl = $item[$sourceUrlColumnName]
        if ($sourceUrl -eq "" -or ($null -eq $sourceUrl)) {
            Write-Host "Error: Source URL is an empty string. To resolve: Check that a value was provided. If reading from a list, confirm that the column name is actually" $sourceUrlColumnName -ForegroundColor Red
            return
        }
        if ($Verbose) { Write-Host "Source: " $sourceUrl }
        # Convert the source URL to a natural SPO path ("https://[tenant].sharepoint.com/sites[/sitename][/library][/folder (optional)]")
        $sourceUrl = Get-SPOFolderNaturalUrl($sourceUrl) 

        # Process the target URL
        $targetUrl = $item[$targetUrlColumnName]
        if ($targetUrl -eq "" -or ($null -eq $targetUrl)) {
            Write-Host "Error: Target URL is an empty string. To resolve: Check that a value was provided. If reading from a list, confirm that the column name is actually" $targetUrlColumnName -ForegroundColor Red
            return
        }
        if ($Verbose) { Write-Host "Target: " $targetUrl }
        # Convert the target URL to a natural SPO path 
        $targetUrl = Get-SPOFolderNaturalUrl($targetUrl) 

        Write-Host "Copying from: " $sourceUrl "to:" $targetUrl 

        # Use the System.Uri class to parse the URL
        $uri = New-Object System.Uri($sourceUrl)

        # Extract the site path and library name
        $sitePath = Get-SPWebPath($uri)
        $libraryPath = Get-LibraryPath($uri)

        $connection = Connect-PnPOnline -Url $sitePath -Interactive

        #Set the vars we'll use in the actual operations
        $sourceIndex = $sourceUrl.IndexOf("/sites")
        $sourceRelative = $SourceUrl.Substring($sourceIndex)
        $targetIndex = $TargetUrl.IndexOf("/sites")
        $targetRelative = $TargetUrl.Substring($targetIndex)
        $lastSlashIndex = $sourceRelative.LastIndexOf("/")
        $sourceName = $sourceRelative.Substring($lastSlashIndex + 1)

        if ($Verbose) { Write-Host "   Library name: " $libraryPath " || Folder name: " $sourceName " || ServerRelativeUrl: " $sourceRelative " || Target Path: " $targetRelative }

        $allItems = Get-PnPListItem -List "$libraryPath" -FolderServerRelativeUrl "$sourceRelative" -Connection $connection

        foreach ($dirItem in $allItems) {
            $sourceRelative = $sourceRelative -replace "[~#%&*:?|]", ""
            if ($dirItem.FieldValues.FileDirRef -eq $sourceRelative) {
                Write-Host "Copying file: $($dirItem.FieldValues.FileLeafRef)" -ForegroundColor green
                if (-not ($ReadOnly)) {
                    try {
                        Copy-PnPFile -SourceUrl "$($dirItem.FieldValues.FileRef)" -TargetUrl "$targetRelative" -Force
                    }
                    catch {
                        if ($_.Exception.Message -like "*The system cannot find the file specified*") {
                            Write-Host "Error: The target folder could not be found. Please check the target address and try again." -ForegroundColor Red
                        } else {
                            Write-Host "An unexpected error occurred: $($_.Exception.Message)" -ForegroundColor Red
                        }
                    }
                }
            $itemCount++
            }
        }

        if ($listExecution -and -not ($ReadOnly)) {
            # Update the status on the row to newStatusValue (e.g. Copy-Complete)
            $connectionSiteUrl = Connect-PnPOnline -Url $listSiteUrl -Interactive
            Set-PnPListItem -List $listName -Identity $item.Id -Values @{ $statusColumnName = $newStatusValue } -Connection $connectionSiteUrl
        }
}

Write-Host "Copy-SPOFolder complete:" $itemCount.ToString() "items processed."
