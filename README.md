# M365Tools
This repo is for any useful tools and scripts for Microsoft 365 that I build or enhance from other projects. 

## PowerShell Cmdlets: 
[Copy-SPOFolder](https://github.com/erobillard/M365Tools/blob/main/docs/Copy-SPOFolder.md)

### Why it exists: 
The PnP Powershell library has a Copy-PnPFile cmdlet with important advantages over copying or moving with the SharePoint UI: 
   - Copies include the version history
   - Copies include the correct Created, CreatedBy, LastModified, and LastModifiedBy metadata wherever possible.

However there was nothing (when I wrote this) to copy an entire library or folder. I was working with a few clients on migrations, and once we did a lift-and-shift into the tenant, we needed a way to further copy/move folders around. Since no one loves building delimited files, I went with a SharePoint list with a UX similar to most migration tools, i.e., with Source and Target columns, and an Action column to indicate the status of each row. 

### Minimal path to awesome: 
   - Open Powershell and run Copy-SPOFolder with source and target parameters. 

### Alternative path to awesome:
1. Create and configure a SharePoint List: 
   - Create a SharePoint List with the Blank List template named "Copy a Folder" 
   - Open List Settings and rename the Title column to SourceUrl.
   - Create a Single line of text column named TargetUrl. Rename this column after creation if you prefer a different display name. 
   - Create a Text column called Action. Set the column's default value to Copy
2. If you chose different column names or labels in Step 1, locate the section "# Define variables" below and update the values as needed.
3. Add a few rows with valid source and target values for testing. Note that the target library or subfolder must exist, create as needed.
4. Install PnP PowerShell
5. Execute the script, e.g.:  Copy-SPOFolder -verbose -siteUrl "https://contoso.sharepoint.com/sites/SharePointTools" 

Step 1:
![Screenshot of the List Settings page.](./docs/Guide-List-Settings-Copy-a-Folder.png)

Step 3: 
![Adding a new entry to the Copy a Folder list.](./docs/Guide-Copy-Tool-New-Item.png)
![Viewing the Copy a Folder list with one row ready to go.](./docs/Guide-List-Copy-a-Folder-Ready-to-go.png)

Step 5:
![Command line example.](https://github.com/erobillard/M365Tools/blob/main/docs/Guide-PSCommandLine-Execute.png)

### Suggestions
There are a few ideas already considered for the backlog: 
   - Change the -SiteUrl parameter to -ListUrl and get rid of the hardcoded list name.
   - Either use this as a starting point of a Move-SPOList cmdlet, or add a -DeleteSource parameter to accomplish the same.
   - Support the SPO uri format provided by the Copy Link buttons.
   - Provide the option to set the Copy List column names as parameters.
   - Submit a command-line only version to the PnP Powershell project.
   - Document the steps to deploy the script as an Azure Function App, and provide a means of kicking off the Copy from the Copy List. This would require extensive changes to the authentication sections in order to use credentials securely stored with the Azure Function.

All are welcome to join the project, pull requests to implement these and other are ideas more than welcome.
