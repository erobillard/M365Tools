---
Module Name: erobillard.M365Tools
applicable: SharePoint Online
title: Copy-SPFolder
---
  
# Copy-SPFolder

## SYNOPSIS
Copy a folder and its contents from a source folder to a target folder, including all available versions of each file. The target folder must exist. 
Source and target URLs can either be specified in the command line with parameters, or read from a SharePoint list.
Source and target URLs may be pasted from the address bar of a browser open to a library or folder. Does not (yet) support links copied using the "Copy Link" button.


## SYNTAX

```powershell
Copy-SPFolder [-Verbose] [-SiteUrl <Url>] [-Source <Url>] [-Target <Url>]
```

## DESCRIPTION
To provide source and target values, either:
    - Set up a list with 3 columns that specify the source, target and action. All rows with an Action of Copy will be processed. The list's name and column names are specified as inline vars.
    - Command line parameters can indicate the source and target URLs.
Constraints: 
    The PnP Powershell module is a prerequisite: [Microsoft PnP Powershell](https://aka.ms/m365pnp) 
    The target folder must exist.
    The account executing the cmdlet requires a minimum of read permissions at the source and read-write permissions at the target. 
    Will not copy from or to the tenant's root SPO Site (e.g.: the top-level site found at https://[tenant].sharepoint.com/)

## EXAMPLES

### EXAMPLE 1
```powershell
Copy-SPOFolder -Verbose -siteUrl "https://contoso.sharepoint.com/sites/SharePointTools" 
```

Using the list located in the SharePointTools site, where the action column is "Copy", copy the folders listed in the source column into the folders listed in the target column. Provide verbose output. 

### EXAMPLE 2
```powershell
Copy-SPOFolder -Verbose -source "https://contoso.sharepoint.com/sites/SourceSite/Shared Documents" -Target "https://contoso.sharepoint.com/sites/TargetSite/Shared Documents" 
```

Copy the folder specified in the -source parameter into the folder specified in the -target parameter. Provide verbose output. 

### EXAMPLE 3
```powershell
Copy-SPOFolder -Source "https://contoso.sharepoint.com/sites/SourceSite/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FSourceSite%2FShared%20Documents%2FGeneral" -Target "[https://contoso.sharepoint.com/sites/TargetSite/Shared Documents/General/Test Data](https://contoso.sharepoint.com/sites/TargetSite/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FSiteB%2FShared%20Documents%2FGeneral&viewid=9427acf1%2De1ae%2D4ecc%2Db456%2D4a1bead7726b)" 
```

Similar to Example 2, this time the source and target values are pasted from the browser rather than "natural" SharePoint Online URLs.

## PARAMETERS

### -Verbose
Provide verbose output.

```yaml
Type: Boolean
Accepted values: Either present or not.

Required: False
Default value: None
```

### -SiteUrl
The location of the site where the "Copy Folders" list is created. The name of the list is specified as a var in the script's source. 
Note: This parameter should be superceded by the Url of the list in a future update, e.g.: ListUrl.

```yaml
Type: String
Parameter Sets: (Read from list)
Accepted values: The URL of a valid SharePoint Online site. 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Source
The URL of the folder to copy from. 

```yaml
Type: String
Parameter Sets: (Command-line)
Accepted values: The URL of a valid SharePoint Online library or folder. 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Target
The URL of the folder to copy to. 

```yaml
Type: String
Parameter Sets: (Command-line)
Accepted values: The URL of a valid SharePoint Online library or folder. 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

## RELATED LINKS

[Microsoft PnP Powershell](https://aka.ms/m365pnp)
