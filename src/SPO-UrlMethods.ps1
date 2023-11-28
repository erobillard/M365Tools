﻿<#
.SYNOPSIS
    Methods to parse and translate SharePoint Online URLs 

    Decode-CopyLinkUrl -Url "[address]"
        Translates a folder URL generated by the Copy Link button to a natural URL

    Get-SPOFolderNaturalUrl -Url "[address]" 
        Translates a URL from the browser address bar or a Copy Link button for a SharePoint library or folder into the natural URL for the same. 
        Example 1: 
            Get-SPOFolderNaturalUrl -Url "https://contoso.sharepoint.com/sites/MySPOSite/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FMySPOSite%2FShared%20Documents%2FGeneral&viewid=51319bc8%2Dc5cd%2D5a53%2D97fc%2Dae17afd3c80b"
        Example 2: 
            Get-SPOFolderNaturalUrl -Url "https://contoso.sharepoint.com/:f:/r/sites/MySPOSite/Shared%20Documents/General?csf=1&web=1&e=BR2B7F"
        Both return: "https://contoso.sharepoint.com/sites/MySPOSite/Shared Documents/General"

    Get-SPWebPath -Url "[address]"
        Takes a natural SPO Url and provides the Url of that site. Handy to get a path for authentication. 

    Get-LibraryPath -Url "[address]" 
        Takes a natural SPO Url of a library or folder, and returns the segment with the library path.
        Example: Get-LibraryPath -Url "https://contoso.sharepoint.com/sites/MySPOSite/Shared Documents/General"
        Returns: "Shared Documents"

.EXAMPLE

.NOTES
    Right now the only error checking on the input is to confirm ".sharepoint.com" is in the Url. 
    This script only contains string converstions, no connection to SPO is actually required. 
#>

# Get-SPWebPath: Extracts the path to the SharePoint Site Collection from a complete SPO URL, e.g. to use as a target for authentication.
# Example: https://contoso.sharepoint.com/sites/MySPOSite/Shared%20Documents/General 
# Returns: https://contoso.sharepoint.com/sites/MySPOSite
Function Get-SPWebPath ($Url) {
    $parts = $Url -split "/"
    $index = $parts.IndexOf("sites")
    $leftPortion = [string]::Join("/", $parts[0..($index+1)])
    return $leftPortion
}

# Get-LibraryPath: Extracts the name of the library from a complete path.
# Example: https://contoso.sharepoint.com/sites/MySPOSite/Shared%20Documents/General
# Returns: Shared%20Documents
Function Get-LibraryPath ($Url) {
    $parts = $Url -split "/"
    $index = $parts.IndexOf("sites")
    # Get the second segment after the "sites" part (the library name)
    $secondSegment = $parts[$index + 2]
    return $secondSegment
}

Function Get-SPOFolderNaturalUrl ($Url) {
    Add-Type -AssemblyName System.Web

    # Translate special characters from %99 format to the actual character (spaces, slashes, etc.)
    $decodedUrl = [System.Web.HttpUtility]::UrlDecode($Url)

    if ($decodedUrl -match ".sharepoint.com") {
    }
    else {
        $error = "Error: The -Url parameter must be a SharePoint Online URL. Input provided:" + $Url
        Write-Host $error -ForegroundColor Red
        return $error 
    }

    # Define the regex pattern to match the "id=" parameter 
    $pattern = "(?<=\?|&)id=([^&]+)"

    # Use the -match operator to find matches in a URL copied from the address bar
    if ($decodedUrl -match $pattern) {
        # The $matches automatic variable contains the matches
        # The value of the "id" parameter is in $matches[1]

        # Extract the domain from the URL
        $domain = $decodedUrl -split "/sites" | Select-Object -First 1

        # Combine the domain with the id parameter to get the actual URL
        $actualUrl = $domain + $matches[1]

        return $actualUrl
    }
    # Check for a folder shared from a Copy Link button to people with existing access
    elseif ($decodedUrl -match ":f") { 
        return Decode-CopyLinkUrl($decodedUrl)
    }
    else {
        $error = "Error: Please use URLs for sharing with people who have existing access to the library or folder. This URL may expire: " + $decodedUrl 
        Write-Host $error -ForegroundColor Red
        return $error
    }
}

# Translates a folder URL generated by the Copy Link button to a natural URL
Function Decode-CopyLinkUrl ($Url) {
    $decodedUrl = [System.Web.HttpUtility]::UrlDecode($Url)
    if ($decodedUrl -match ":f:/r") { 
    
        # Create a Uri object from the URL
        $uri = New-Object System.Uri($decodedUrl)

        # Extract the path and unescape it
        $path = [System.Uri]::UnescapeDataString($uri.AbsolutePath)

        # Remove the leading "/:f" and any characters up to "/sites" from the path
        $path = $path -replace "^/:f.*?/sites", "/sites"

        # Construct the new URL
        $newUrl = $uri.Scheme + "://" + $uri.Host + $path

        # Return the new URL
        return $newUrl
    }
    else {
        $error = "Error: Please use URLs for sharing with people who have existing access to the library or folder. This URL may expire:" 
        Write-Host $error $decodedUrl -ForegroundColor Red
        return $error
    }
}

