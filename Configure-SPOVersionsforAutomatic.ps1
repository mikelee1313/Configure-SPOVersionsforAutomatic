<#
.SYNOPSIS
    PowerShell script to manage SharePoint Online site version policies across multiple sites.

.DESCRIPTION
    This script provides functionality to manage SharePoint Online site version policies and file version management across multiple sites defined in a text file. 
    
    It includes capabilities to:

    - Get current version policies
    - Enable auto-expiration version trimming
    - Check version policy status and storage usage
    - Create batch delete jobs for version cleanup
    - Monitor batch deletion job status

    The script implements throttling handling to manage SharePoint Online request limits and provides detailed logging.

.PARAMETER tenantId
    The Microsoft 365 tenant ID.

.PARAMETER clientId
    The application (client) ID for authentication.

.PARAMETER url
    The SharePoint Online admin center URL.

.EXAMPLE
    .\Configure-SPOVersionsforAutomatic.ps1

.NOTES
    Authors: Mike Lee /Luis DuSolier
    Date: 5/19/25
    
    File Name      : Configure-SPOVersionsforAutomatic.ps1
    Prerequisites  : 
    - PnP.PowerShell module installed (Tested with 3.1.0)
    - Text file with site URLs at C:\temp\M365CPI13246019-Sites.txt
    - Proper permissions to connect to SPO and modify sites
    
    The script uses interactive authentication. Make sure you have appropriate permissions
    to perform operations on the specified SharePoint sites.

.Disclaimer: The sample scripts are provided AS IS without warranty of any kind. 
    Microsoft further disclaims all implied warranties including, without limitation, 
    any implied warranties of merchantability or of fitness for a particular purpose. 
    The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
    In no event shall Microsoft, its authors, or anyone else involved in the creation, 
    production, or delivery of the scripts be liable for any damages whatsoever 
    (including, without limitation, damages for loss of business profits, business interruption, 
    loss of business information, or other pecuniary loss) arising out of the use of or inability 
    to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

.INPUTS
    Site URLs from a text file located at C:\temp\M365CPI13246019-Sites.txt.

.OUTPUTS
    - Console output showing operation status
    - Detailed log file in %TEMP% directory named 'configure_versions_SPO[date]_logfile.log'

.FUNCTIONALITY
    SharePoint Online, Version Management, Site Management, PnP PowerShell
#>

# This is the logging function
Function Write-LogEntry {
    param(
        [string] $LogName,
        [string] $LogEntryText,
        [string] $LogLevel = "INFO"  # Default log level is INFO
    )
    if ($LogName -ne $null) {
        # Skip DEBUG level messages if Debug is set to False
        if ($LogLevel -eq "DEBUG" -and $Debug -eq $False) {
            return
        }
        
        # log the date and time in the text file along with the data passed
        "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToShortTimeString()) : [$LogLevel] $LogEntryText" | Out-File -FilePath $LogName -append;
    }
}

############################################
################configuration###############

#tenant Properties
$tenantId = '9cfc42cb-51da-4055-87e9-b20a170b6ba3'
$clientId = '1e892341-f9cd-4c54-82d6-0fc3287954cf'
$url = "https://m365cpi13246019-admin.sharepoint.com"

# Read sites from file
Write-LogEntry -LogName $log -LogEntryText "Reading site list from: C:\temp\M365CPI13246019-Sites.txt" -LogLevel "INFO"
$sites = Get-Content -Path "C:\temp\M365CPI13246019-Sites.txt"
Write-LogEntry -LogName $log -LogEntryText "Found $($sites.Count) sites to process" -LogLevel "INFO"

# Initialize logging
$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$log = "$env:TEMP\" + 'configure_versions_SPO' + $date + '_' + "logfile.log"
$Debug = $true

#################section####################
############################################

# Log script start
Write-LogEntry -LogName $log -LogEntryText "Script execution started. Connecting to tenant admin site: $url" -LogLevel "INFO"

# Connect to the SharePoint Online admin site
$connnecton = Connect-PnPOnline -Url $url -ClientId $clientId -Tenant $tenantId -interactive -returnConnection
Write-LogEntry -LogName $log -LogEntryText "Successfully connected to admin site" -LogLevel "INFO"



# Function to handle throttling
function Execute-WithThrottlingHandling {
    param (
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,
        
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [int]$MaxRetries = 5,
        [int]$InitialRetrySeconds = 30
    )
    
    $retryCount = 0
    $success = $false
    
    Write-Host "Executing operation on site: $SiteUrl" -ForegroundColor Cyan
    Write-LogEntry -LogName $log -LogEntryText "Executing operation on site: $SiteUrl" -LogLevel "INFO"
    
    while (-not $success -and $retryCount -lt $MaxRetries) {
        try {
            # Execute the command and capture output
            $output = & $ScriptBlock
            $success = $true
            
            # Display the output to console if it's not empty
            if ($output) {
                Write-Host "Output from site $SiteUrl :" -ForegroundColor Green
                $output | Format-Table -AutoSize
            }
            
            Write-Host "Successfully executed command for site: $SiteUrl" -ForegroundColor Green
            Write-LogEntry -LogName $log -LogEntryText "Successfully executed command for site: $SiteUrl" -LogLevel "INFO"
        }
        catch {
            if ($_.Exception.Response.StatusCode -eq 429 -or $_.Exception.Response.StatusCode -eq 503) {
                $retryAfter = $_.Exception.Response.Headers["Retry-After"]
                if (-not $retryAfter) {
                    $retryAfter = $InitialRetrySeconds * [math]::Pow(2, $retryCount)
                }
                
                $retryCount++
                $warningMsg = "Throttling detected for site $SiteUrl. Waiting for $retryAfter seconds before retry $retryCount of $MaxRetries..."
                Write-Warning $warningMsg
                Write-LogEntry -LogName $log -LogEntryText $warningMsg -LogLevel "WARNING"
                Start-Sleep -Seconds $retryAfter
            }
            else {
                $errorMsg = "Error processing site $SiteUrl : $_"
                Write-Error $errorMsg
                Write-Host $_.Exception.ToString() -ForegroundColor Red
                Write-LogEntry -LogName $log -LogEntryText $errorMsg -LogLevel "ERROR"
                throw $_
            }
        }
    }
    
    if (-not $success) {
        $errorMsg = "Failed to execute command for $SiteUrl after $MaxRetries retries."
        Write-Error $errorMsg
        Write-LogEntry -LogName $log -LogEntryText $errorMsg -LogLevel "ERROR"
    }
}

# Function to process each site with a specific operation
function Invoke-SiteBatch {
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$SiteUrls,
        
        [Parameter(Mandatory = $true)]
        [scriptblock]$Operation,
        
        [Parameter(Mandatory = $true)]
        [string]$ClientId,
        
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        
        [Parameter(Mandatory = $true)]
        [object]$Connection,
        
        [string]$OperationDescription = "operation"
    )
    
    Write-Host "Starting batch processing for operation: $OperationDescription on $($SiteUrls.Count) sites" -ForegroundColor Yellow
    Write-LogEntry -LogName $log -LogEntryText "Starting batch processing for operation: $OperationDescription on $($SiteUrls.Count) sites" -LogLevel "INFO"
    
    foreach ($siteUrl in $SiteUrls) {
        Write-Host "Processing site: $siteUrl" -ForegroundColor Cyan
        Write-LogEntry -LogName $log -LogEntryText "Processing site: $siteUrl for $OperationDescription" -LogLevel "INFO"
        
        try {
            # Connect to the site using delegate access
            Write-Host "Connecting to site: $siteUrl" -ForegroundColor Cyan
            Write-LogEntry -LogName $log -LogEntryText "Connecting to site: $siteUrl" -LogLevel "DEBUG"
            Connect-PnPOnline -Url $siteUrl -ClientId $ClientId -Tenant $TenantId -Interactive -Connection $Connection
            
            # Apply site operation with throttling handling
            Execute-WithThrottlingHandling -SiteUrl $siteUrl -ScriptBlock $Operation
        }
        catch {
            $errorMsg = "Failed to connect to site $siteUrl. Error: $_"
            Write-Error $errorMsg
            Write-Host $_.Exception.ToString() -ForegroundColor Red
            Write-LogEntry -LogName $log -LogEntryText $errorMsg -LogLevel "ERROR"
        }
    }
    
    Write-Host "Processing completed for all sites" -ForegroundColor Green
    Write-LogEntry -LogName $log -LogEntryText "Completed batch processing for operation: $OperationDescription" -LogLevel "INFO"
}

# Create operation script blocks
$getVersionPolicyOperation = {
    $policy = Get-PnPSiteVersionPolicy
    # Return policy object for display
    Write-Host "  - Site version policy retrieved successfully" -ForegroundColor Green
    Write-LogEntry -LogName $log -LogEntryText "Site version policy retrieved: EnableAutoExpirationVersionTrim = $($policy.EnableAutoExpirationVersionTrim)" -LogLevel "INFO"
    return $policy | fl # Format list for better readability
}

$setVersionPolicyOperation = {
    $result = Set-PnPSiteVersionPolicy -EnableAutoExpirationVersionTrim $true
    Write-Host "  - Site version policy set successfully" -ForegroundColor Green
    Write-LogEntry -LogName $log -LogEntryText "Site version policy set to EnableAutoExpirationVersionTrim = True" -LogLevel "INFO"
    return $result | fl # Format list for better readability
}

$getVersionPolicyStatusOperation = {
    $status = Get-PnPSiteVersionPolicyStatus
    # Return status object for display
    Write-Host "  - Site version policy status retrieved successfully" -ForegroundColor Green
    Write-LogEntry -LogName $log -LogEntryText "Site version policy status: StorageUsageMB = $($status.StorageUsageMB), VersionStorageUsageMB = $($status.VersionStorageUsageMB)" -LogLevel "INFO"
    return $status | fl
}

$createBatchDeleteJobOperation = {
    $job = New-PnPSiteFileVersionBatchDeleteJob -Automatic -Force
    # Return job object for display
    Write-Host "  - Site file version batch delete job created successfully" -ForegroundColor Green
    Write-LogEntry -LogName $log -LogEntryText "Batch delete job created with ID: $($job.Id)" -LogLevel "INFO"
    return $job | fl # Format list for better readability
}

$getBatchDeleteJobStatusOperation = {
    $jobStatus = Get-PnPSiteFileVersionBatchDeleteJobStatus
    # Return job status object for display
    Write-Host "  - Site file version batch delete job status retrieved successfully" -ForegroundColor Green
    Write-LogEntry -LogName $log -LogEntryText "Batch delete job status: State = $($jobStatus.State), ProgressPercentage = $($jobStatus.ProgressPercentage)%" -LogLevel "INFO"
    return $jobStatus | fl # Format list for better readability
}

# Display menu and get user selection
function Show-OperationMenu {
    Clear-Host
    Write-Host "==== SharePoint Site Version Policy Operations ====" -ForegroundColor Cyan
    Write-Host "1: Get current version policy for all sites"
    Write-Host "2: Set auto-expiration version trim to enabled for all sites"
    Write-Host "3: Get version policy status for all sites"
    Write-Host "4: Create batch delete job for all sites"
    Write-Host "5: Get batch delete job status for all sites"
    Write-Host "Q: Quit"
    Write-Host "=================================================" -ForegroundColor Cyan
    
    $selection = Read-Host "Please select an operation (1-5, or Q to quit)"
    Write-LogEntry -LogName $log -LogEntryText "User selected menu option: $selection" -LogLevel "INFO"
    return $selection
}

# Main execution loop
function Start-OperationsMenu {
    $continue = $true
    Write-LogEntry -LogName $log -LogEntryText "Starting operations menu" -LogLevel "INFO"
    
    while ($continue) {
        $choice = Show-OperationMenu
        
        switch ($choice) {
            "1" {
                Write-Host "Running: Get current version policy" -ForegroundColor Yellow
                Write-LogEntry -LogName $log -LogEntryText "Starting operation: Get current version policy" -LogLevel "INFO"
                Invoke-SiteBatch -SiteUrls $sites -Operation $getVersionPolicyOperation -ClientId $clientId -TenantId $tenantId -Connection $connnecton -OperationDescription "get version policy"
                Read-Host "Press Enter to return to menu"
            }
            "2" {
                Write-Host "Running: Set auto-expiration version trim" -ForegroundColor Yellow
                Write-LogEntry -LogName $log -LogEntryText "Starting operation: Set auto-expiration version trim" -LogLevel "INFO"
                Invoke-SiteBatch -SiteUrls $sites -Operation $setVersionPolicyOperation -ClientId $clientId -TenantId $tenantId -Connection $connnecton -OperationDescription "set version policy"
                Read-Host "Press Enter to return to menu"
            }
            "3" {
                Write-Host "Running: Get version policy status" -ForegroundColor Yellow
                Write-LogEntry -LogName $log -LogEntryText "Starting operation: Get version policy status" -LogLevel "INFO"
                Invoke-SiteBatch -SiteUrls $sites -Operation $getVersionPolicyStatusOperation -ClientId $clientId -TenantId $tenantId -Connection $connnecton -OperationDescription "get version policy status"
                Read-Host "Press Enter to return to menu"
            }
            "4" {
                Write-Host "Running: Create batch delete job" -ForegroundColor Yellow
                Write-LogEntry -LogName $log -LogEntryText "Starting operation: Create batch delete job" -LogLevel "INFO"
                Invoke-SiteBatch -SiteUrls $sites -Operation $createBatchDeleteJobOperation -ClientId $clientId -TenantId $tenantId -Connection $connnecton -OperationDescription "create batch delete job"
                Read-Host "Press Enter to return to menu"
            }
            "5" {
                Write-Host "Running: Get batch delete job status" -ForegroundColor Yellow
                Write-LogEntry -LogName $log -LogEntryText "Starting operation: Get batch delete job status" -LogLevel "INFO"
                Invoke-SiteBatch -SiteUrls $sites -Operation $getBatchDeleteJobStatusOperation -ClientId $clientId -TenantId $tenantId -Connection $connnecton -OperationDescription "get batch delete job status"
                Read-Host "Press Enter to return to menu"
            }
            "Q" {
                $continue = $false
                Write-Host "Exiting script..." -ForegroundColor Yellow
                Write-LogEntry -LogName $log -LogEntryText "User exited script" -LogLevel "INFO"
            }
            "q" {
                $continue = $false
                Write-Host "Exiting script..." -ForegroundColor Yellow
                Write-LogEntry -LogName $log -LogEntryText "User exited script" -LogLevel "INFO"
            }
            default {
                Write-Host "Invalid selection. Please try again." -ForegroundColor Red
                Write-LogEntry -LogName $log -LogEntryText "Invalid menu selection: $choice" -LogLevel "WARNING"
                Start-Sleep -Seconds 2
            }
        }
    }
}

# Start the interactive menu
Write-LogEntry -LogName $log -LogEntryText "Displaying operations menu" -LogLevel "INFO"
Start-OperationsMenu

# Log script completion
Write-LogEntry -LogName $log -LogEntryText "Script execution completed. Log file: $log" -LogLevel "INFO"
