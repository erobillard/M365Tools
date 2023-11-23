<#
.SYNOPSIS
    Copy a folder and its contents from a source folder to a target folder, including all available versions of each file. The target folder must exist. 
    Can either be run from the command line with a source and target, or will read source and target values from a SharePoint list.
    The source and target can be URLs pasted from the address bar of a browser open to a library or folder. Does not (yet) support links copied using the "Copy Link" button.

.DESCRIPTION
    To provide source and target values, either:
        - Set up a list with 3 columns that specify the source, target and action. All rows with an Action of Copy will be processed. The list's name and column names are specified as vars below.
        - Command line parameters can indicate the source and target URLs.
    Constraints: 
        The target folder must exist.
        The account executing the cmdlet requires read permissions for the source and write permissions for the target. 
        The cmdlet will not copy from or to the tenant's root SPO Site (e.g.: the top-level site found at https://contoso.sharepoint.com/)

.PARAMETER verbose
    Optional, -verbose displays detailed output if present

.PARAMETER siteUrl
    Required when reading from a list, this is the path of the site to connect with for authentication.

.PARAMETER source
    Required when not reading from a SPList, the URL of the library or folder to copy FROM. The site at this location is used for authentication. 

.PARAMETER target
    Required when not reading from a SPlist, the URL of the library or folder to copy TO. 

.EXAMPLE
    Copy-SPOFolder -Verbose -siteUrl "https://contoso.sharepoint.com/sites/SharePointTools" 
    Copy-SPOFolder -Verbose -source "https://contoso.sharepoint.com/sites/SourceSite/Shared Documents" -Target "https://contoso.sharepoint.com/sites/TargetSite/Shared Documents" 
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
    2023-11-22 Refactored somewhat and converted to a cmdlet

.VERSION
    1.0.0.0 Initial Release
#>

param (
    [switch] $Verbose,
    [string] $SiteUrl, 
    [string] $Source,
    [string] $Target
)

Import-Module PnP.PowerShell

# Get-SPWebPath: Extracts the path to the SharePoint Site Collection from a complete SPO URL, e.g. to use as a target for authentication.
# Exmaple: https://contoso.sharepoint.com/sites/MySPOSite/Shared%20Documents/General 
# Returns: https://contoso.sharepoint.com/sites/MySPOSite
Function Get-SPWebPath ($url) {
    $parts = $url -split "/"
    $index = $parts.IndexOf("sites")
    $leftPortion = [string]::Join("/", $parts[0..($index+1)])
    return $leftPortion
}

# Get-LibraryPath: Extracts the name of the library from a complete path.
# Example: https://contoso.sharepoint.com/sites/MySPOSite/Shared%20Documents/General
# Returns: Shared%20Documents
Function Get-LibraryPath ($url) {
    $parts = $url -split "/"
    $index = $parts.IndexOf("sites")
    # Get the second segment after the "sites" part (the library name)
    $secondSegment = $parts[$index + 2]
    return $secondSegment
}

# Get-RealUri: Translates a browser URL for a SharePoint library or folder into the natural URL for the same. 
# Example: https://contoso.sharepoint.com/sites/MySPOSite/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FMySPOSite%2FShared%20Documents%2FGeneral&viewid=51319bc8%2Dc5cd%2D5a53%2D97fc%2Dae17afd3c80b
# Returns: https://contoso.sharepoint.com/sites/MySPOSite/Shared%20Documents/General 
Function Get-RealUri ($url) {
    # Decode the URL
    $decodedUrl = [System.Web.HttpUtility]::UrlDecode($url)

    # Define the regex pattern to match the "id" parameter
    $pattern = "(?<=\?|&)id=([^&]+)"

    # Use the -match operator to find matches
    if ($decodedUrl -match $pattern) {
        # The $matches automatic variable contains the matches
        # The value of the "id" parameter is in $matches[1]

        # Extract the domain from the URL
        $domain = $url -split "/sites" | Select-Object -First 1

        # Combine the domain with the id parameter to get the actual URL
        $actualUrl = $domain + $matches[1]

        return $actualUrl
    } else {
        return "" 
    }
}

if ($Verbose) { Write-Host "Copy-SPOFolder Begin" }

# Define variables
$listName = "Copy a Folder"
$statusColumnName = "Action"
$newStatusValue = "Copy-Complete"
$sourceUrlColumnName = "Title"
$targetUrlColumnName = "TargetUrl"
$listExecution = $true
$iteration=0
$itemCount=0

# Check for parameters 
if ($PSBoundParameters.ContainsKey('SiteUrl') -and $SiteUrl -ne $true) { 
    if ($Verbose) { Write-Host "SiteUrl was provided: " $SiteUrl }
}
else { 
    if($PSBoundParameters.ContainsKey('Source')) {
        $listExecution = $false
        $SiteUrl = Get-SPWebPath($Source)
        $sourceUrl = $Source
        $targetUrl = $Target
        if ($Verbose) { Write-Host "Site Url: " $SiteUrl " || Source URL: " $sourceUrl " || Target URL: " $targetUrl }
    }
    else {
        Write-Host "Error: Must include either -SiteUrl [value] or -Source [value] as parameters." -ForegroundColor Red
        return
    }
}

# Connect to the site that contains the list with the library mappings
$connectionSiteUrl = Connect-PnPOnline -Url $SiteUrl -Interactive

if ($listExecution) {
    # Get all items from the list where status is "Copy"
    $listItems = Get-PnPListItem -List $listName -Query "<View><Query><Where><Eq><FieldRef Name='$statusColumnName'/><Value Type='Text'>Copy</Value></Eq></Where></Query></View>"
}
else {
    # Create a one-row array using the command-line parameters. 
    if ($Verbose) { Write-Host "Creating the array:" }
    $listItems = @{ $sourceUrlColumnName = $sourceUrl; $targetUrlColumnName = $targetUrl; $statusColumnName = "Copy" }
    if ($Verbose) { Write-Host $listItems }
}

# Loop through each item
foreach ($item in $listItems) {
    
        # Convert the source URL to a natural SPO path ("https://[tenant].sharepoint.com/sites[/sitename][/library][/folder (optional)]")
        $sourceUrl = $item[$sourceUrlColumnName]
        if ($sourceUrl -eq "" -or $sourceUrl -eq $null) {
            Write-Host "Error: Source URL is an empty string. To resolve: Check that a value was provided. If reading from a list, confirm that the column name is actually" $sourceUrlColumnName -ForegroundColor Red
            return
        }
        if ($Verbose) { Write-Host "Source Url Before: " $sourceUrl }
        # Look for a string indicating that the URL was pasted from the browser, .aspx will work in most cases.
        if ($sourceUrl -match ".aspx") {
            $sourceUrl = Get-RealUri($sourceUrl) 
        }
        $sourceUrl = $sourceUrl.Replace("%20", " ")
        $SourceUrl = $sourceUrl.Replace("%2F", "/")
        if ($Verbose) { Write-Host "Source Url After: " $sourceUrl }

        # Convert the target URL
        $targetUrl = $item[$targetUrlColumnName]
        if ($targeteUrl -eq "" -or $targetUrl -eq $null) {
            Write-Host "Error: Target URL is an empty string. To resolve: Check that a value was provided. If reading from a list, confirm that the column name is actually" $targetUrlColumnName -ForegroundColor Red
            return
        }
        # Look for a string indicating that the URL was pasted from the browser
        if ($targetUrl -match ".aspx") {
            $targetUrl = Get-RealUri($targetUrl) 
        }
        $targetUrl = $targetUrl.Replace("%20", " ")
        $targetUrl = $targetUrl.Replace("%2F", "/")
        if ($Verbose) { Write-Host "Target Url: " $targetUrl }

        Write-Host "Copying from: " $sourceUrl " to: " $targetUrl 

        # Use the System.Uri class to parse the URL
        $uri = New-Object System.Uri($SourceUrl)

        # Extract the site path and library name
        $sitePath = Get-SPWebPath($uri)
        $libraryPath = Get-LibraryPath($uri)

        if ($iteration -eq 0) {
            # Connect to the source site (where files and folders will be copied from). This needs to match the relative Url (sourceRelative) or an error will be thrown during Get-PnPListItem
            $connection = Connect-PnPOnline -Url $sitePath -Interactive
            $iteration++
        }

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
            $itemCount++
            }
        }

        if ($listExecution) {
            # Update the status on the row to newStatusValue (e.g. Copy-Complete)
            $connectionSiteUrl = Connect-PnPOnline -Url $siteUrl -Interactive
            $set = Set-PnPListItem -List $listName -Identity $item.Id -Values @{ $statusColumnName = $newStatusValue } -Connection $connectionSiteUrl
        }
}

Write-Host "Copy-SPOFolder Complete:" $itemCount.ToString() "items processed."
