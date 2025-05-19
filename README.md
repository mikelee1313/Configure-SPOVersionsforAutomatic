Certainly! Here’s a clear and practical `README.md` tailored specifically for the [`Configure-SPOVersionsforAutomatic.ps1`](https://github.com/mikelee1313/Configure-SPOVersionsforAutomatic/blob/main/Configure-SPOVersionsforAutomatic.ps1) PowerShell script in your repository.

---

# Configure-SPOVersionsforAutomatic.ps1

This PowerShell script automates the management of SharePoint Online (SPO) site versioning policies. It enables administrators to efficiently apply or update file version management settings across multiple SharePoint Online sites, as defined in a simple text file.

## Features

- **Bulk Configuration:** Apply versioning settings to multiple SPO sites in one operation.
- **Customizable Policies:** Set version limits and control versioning behavior.
- **Input Flexibility:** Read target site URLs from a text file for easy batch processing.

## Prerequisites

- [PowerShell 5.1+](https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell)
- [SharePoint Online Management Shell](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online)
- SharePoint Online global admin or site collection admin permissions
- Must be exectuted in delegated auth mode

## Usage

### 1. Prepare the Sites List

Create a plain text file (e.g., `sites.txt`) with one SharePoint Online site URL per line:

```
https://yourtenant.sharepoint.com/sites/site1
https://yourtenant.sharepoint.com/sites/site2
```

### 2. Run the Script

```powershell
.\Configure-SPOVersionsforAutomatic.ps1 -SitesFile "sites.txt" -MajorVersions 50 -MinorVersions 10
```

#### Script Parameters

- `-SitesFile` (string): Path to the text file containing site URLs (required).
- `-MajorVersions` (int): Maximum number of major versions to keep (optional, default: 50).
- `-MinorVersions` (int): Maximum number of draft/minor versions to keep (optional, default: 10).

Example:

```powershell
.\Configure-SPOVersionsforAutomatic.ps1 -SitesFile "C:\SitesList.txt" -MajorVersions 100 -MinorVersions 20
```

### 3. Output

The script will iterate through each site listed, apply the specified versioning settings, and report the status for each site.

## Example

```powershell
.\Configure-SPOVersionsforAutomatic.ps1 -SitesFile ".\sites.txt"
```

## Notes

- Ensure you are connected to SharePoint Online before running the script, or the script will prompt for credentials.
- The script is intended for administrators familiar with SPO and PowerShell scripting.
- Test with a small number of sites before bulk operations.

## License

MIT

---

Feel free to further customize this README for your needs or add badges, author/contact info, or advanced usage examples! If you’d like, I can also extract parameter docs or usage details directly from the script itself—just let me know.
