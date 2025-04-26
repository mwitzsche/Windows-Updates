<#
.SYNOPSIS
    Manages Windows Updates with various filtering and installation options.

.DESCRIPTION
    This script provides comprehensive management of Windows Updates using the PSWindowsUpdate module.
    It allows for searching, downloading, installing, and hiding updates with various filtering options
    including update categories, severity levels, and KB article IDs. The script includes detailed
    logging and error handling.

.PARAMETER Action
    Action to perform (Search, Download, Install, Hide, Unhide, GetHistory).

.PARAMETER Categories
    Categories of updates to include (Security, Critical, Important, Optional, Feature, Driver).

.PARAMETER KBArticleIDs
    Specific KB article IDs to target.

.PARAMETER MaxUpdates
    Maximum number of updates to process.

.PARAMETER AutoReboot
    Whether to automatically reboot after installing updates if required.

.PARAMETER IgnoreReboot
    Whether to ignore reboot requirements after installing updates.

.PARAMETER ScheduleReboot
    Schedule a reboot at a specific time after installing updates.

.PARAMETER RebootTime
    Time to schedule reboot (used with ScheduleReboot parameter).

.PARAMETER ComputerName
    Remote computer names to manage updates on.

.PARAMETER Credential
    Credentials to use for remote computers.

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    .\Manage-WindowsUpdates.ps1 -Action Search -Categories Security,Critical

.EXAMPLE
    .\Manage-WindowsUpdates.ps1 -Action Install -Categories Security,Critical -AutoReboot $true

.EXAMPLE
    .\Manage-WindowsUpdates.ps1 -Action Install -KBArticleIDs KB5001567,KB5003173 -IgnoreReboot $true

.NOTES
    Author: Michael Witzsche
    Date: April 26, 2025
    Version: 1.0.0
    Requires: PSWindowsUpdate module
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidateSet("Search", "Download", "Install", "Hide", "Unhide", "GetHistory")]
    [string]$Action,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Security", "Critical", "Important", "Optional", "Feature", "Driver")]
    [string[]]$Categories,

    [Parameter(Mandatory = $false)]
    [string[]]$KBArticleIDs,

    [Parameter(Mandatory = $false)]
    [int]$MaxUpdates = 1000,

    [Parameter(Mandatory = $false)]
    [bool]$AutoReboot = $false,

    [Parameter(Mandatory = $false)]
    [bool]$IgnoreReboot = $false,

    [Parameter(Mandatory = $false)]
    [bool]$ScheduleReboot = $false,

    [Parameter(Mandatory = $false)]
    [datetime]$RebootTime,

    [Parameter(Mandatory = $false)]
    [string[]]$ComputerName,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "$env:USERPROFILE\Documents\WindowsUpdates\WindowsUpdates_$(Get-Date -Format 'yyyyMMdd').log"
)

# Function to write log messages
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Create log directory if it doesn't exist
    $logDir = Split-Path -Path $LogPath -Parent
    if (-not (Test-Path -Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }
    
    Add-Content -Path $LogPath -Value $logMessage
    
    switch ($Level) {
        "INFO" { Write-Host $logMessage -ForegroundColor Green }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "ERROR" { Write-Host $logMessage -ForegroundColor Red }
    }
}

# Function to check if PSWindowsUpdate module is installed and install if not
function Ensure-PSWindowsUpdateModule {
    try {
        Write-Log -Message "Checking if PSWindowsUpdate module is installed..." -Level "INFO"
        
        if (-not (Get-Module -Name PSWindowsUpdate -ListAvailable)) {
            Write-Log -Message "PSWindowsUpdate module not found. Installing..." -Level "INFO"
            
            # Check if running as administrator
            $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
            
            if (-not $isAdmin) {
                throw "Administrator privileges required to install PSWindowsUpdate module. Please run PowerShell as Administrator."
            }
            
            try {
                Install-Module -Name PSWindowsUpdate -Force -Scope AllUsers -Confirm:$false
                Write-Log -Message "PSWindowsUpdate module installed successfully." -Level "INFO"
            }
            catch {
                throw "Failed to install PSWindowsUpdate module: $_"
            }
        }
        else {
            Write-Log -Message "PSWindowsUpdate module is already installed." -Level "INFO"
        }
        
        # Import the module
        Import-Module -Name PSWindowsUpdate
        Write-Log -Message "PSWindowsUpdate module imported successfully." -Level "INFO"
        
        return $true
    }
    catch {
        Write-Log -Message "Failed to ensure PSWindowsUpdate module: $_" -Level "ERROR"
        return $false
    }
}

# Function to build parameters for Windows Update commands
function Build-WindowsUpdateParams {
    $params = @{}
    
    # Add common parameters
    if ($ComputerName) {
        $params.Add("ComputerName", $ComputerName)
        
        if ($Credential) {
            $params.Add("Credential", $Credential)
        }
    }
    
    # Add KB article IDs if specified
    if ($KBArticleIDs) {
        $params.Add("KBArticleID", $KBArticleIDs)
    }
    
    # Add categories if specified
    if ($Categories) {
        $categoryFilter = @()
        
        foreach ($category in $Categories) {
            switch ($category) {
                "Security" { $categoryFilter += "Security Updates" }
                "Critical" { $categoryFilter += "Critical Updates" }
                "Important" { $categoryFilter += "Important" }
                "Optional" { $categoryFilter += "Optional" }
                "Feature" { $categoryFilter += "Feature Packs" }
                "Driver" { $categoryFilter += "Drivers" }
            }
        }
        
        if ($categoryFilter.Count -gt 0) {
            $params.Add("Category", $categoryFilter)
        }
    }
    
    # Add installation parameters if needed
    if ($Action -eq "Install") {
        $params.Add("AcceptAll", $true)
        
        if ($AutoReboot) {
            $params.Add("AutoReboot", $true)
        }
        elseif ($IgnoreReboot) {
            $params.Add("IgnoreReboot", $true)
        }
        
        if ($ScheduleReboot -and $RebootTime) {
            $params.Add("ScheduleReboot", $RebootTime)
        }
    }
    
    return $params
}

# Function to search for Windows updates
function Search-WindowsUpdates {
    param (
        [hashtable]$Params
    )
    
    try {
        Write-Log -Message "Searching for Windows updates..." -Level "INFO"
        
        if ($Params.ContainsKey("ComputerName")) {
            Write-Log -Message "Searching for updates on remote computers: $($Params.ComputerName -join ', ')" -Level "INFO"
        }
        
        $filterInfo = ""
        
        if ($Params.ContainsKey("KBArticleID")) {
            $filterInfo += "KB Articles: $($Params.KBArticleID -join ', ') "
        }
        
        if ($Params.ContainsKey("Category")) {
            $filterInfo += "Categories: $($Params.Category -join ', ')"
        }
        
        if ($filterInfo) {
            Write-Log -Message "Applying filters: $filterInfo" -Level "INFO"
        }
        
        $updates = Get-WindowsUpdate @Params -MaxUpdates $MaxUpdates
        
        if ($updates.Count -eq 0) {
            Write-Log -Message "No updates found matching the specified criteria." -Level "INFO"
        }
        else {
            Write-Log -Message "Found $($updates.Count) updates:" -Level "INFO"
            
            foreach ($update in $updates) {
                Write-Log -Message "- $($update.Title) [KB$($update.KB)] - $($update.Size)" -Level "INFO"
            }
        }
        
        return $updates
    }
    catch {
        Write-Log -Message "Failed to search for Windows updates: $_" -Level "ERROR"
        return $null
    }
}

# Function to download Windows updates
function Download-WindowsUpdates {
    param (
        [hashtable]$Params
    )
    
    try {
        Write-Log -Message "Downloading Windows updates..." -Level "INFO"
        
        if ($Params.ContainsKey("ComputerName")) {
            Write-Log -Message "Downloading updates on remote computers: $($Params.ComputerName -join ', ')" -Level "INFO"
        }
        
        $filterInfo = ""
        
        if ($Params.ContainsKey("KBArticleID")) {
            $filterInfo += "KB Articles: $($Params.KBArticleID -join ', ') "
        }
        
        if ($Params.ContainsKey("Category")) {
            $filterInfo += "Categories: $($Params.Category -join ', ')"
        }
        
        if ($filterInfo) {
            Write-Log -Message "Applying filters: $filterInfo" -Level "INFO"
        }
        
        $updates = Get-WindowsUpdate @Params -Download -MaxUpdates $MaxUpdates
        
        if ($updates.Count -eq 0) {
            Write-Log -Message "No updates downloaded." -Level "INFO"
        }
        else {
            Write-Log -Message "Downloaded $($updates.Count) updates:" -Level "INFO"
            
            foreach ($update in $updates) {
                Write-Log -Message "- $($update.Title) [KB$($update.KB)] - $($update.Size)" -Level "INFO"
            }
        }
        
        return $updates
    }
    catch {
        Write-Log -Message "Failed to download Windows updates: $_" -Level "ERROR"
        return $null
    }
}

# Function to install Windows updates
function Install-WindowsUpdates {
    param (
        [hashtable]$Params
    )
    
    try {
        Write-Log -Message "Installing Windows updates..." -Level "INFO"
        
        if ($Params.ContainsKey("ComputerName")) {
            Write-Log -Message "Installing updates on remote computers: $($Params.ComputerName -join ', ')" -Level "INFO"
        }
        
        $filterInfo = ""
        
        if ($Params.ContainsKey("KBArticleID")) {
            $filterInfo += "KB Articles: $($Params.KBArticleID -join ', ') "
        }
        
        if ($Params.ContainsKey("Category")) {
            $filterInfo += "Categories: $($Params.Category -join ', ')"
        }
        
        if ($filterInfo) {
            Write-Log -Message "Applying filters: $filterInfo" -Level "INFO"
        }
        
        $rebootInfo = ""
        
        if ($Params.ContainsKey("AutoReboot")) {
            $rebootInfo = "with automatic reboot if required"
        }
        elseif ($Params.ContainsKey("IgnoreReboot")) {
            $rebootInfo = "ignoring reboot requirements"
        }
        elseif ($Params.ContainsKey("ScheduleReboot")) {
            $rebootInfo = "with scheduled reboot at $($Params.ScheduleReboot)"
        }
        
        if ($rebootInfo) {
            Write-Log -Message "Installing updates $rebootInfo" -Level "INFO"
        }
        
        $updates = Install-WindowsUpdate @Params -MaxUpdates $MaxUpdates
        
        if ($updates.Count -eq 0) {
            Write-Log -Message "No updates installed." -Level "INFO"
        }
        else {
            Write-Log -Message "Installed $($updates.Count) updates:" -Level "INFO"
            
            foreach ($update in $updates) {
                $status = if ($update.Result -eq "Installed") { "Successfully installed" } else { "Failed to install" }
                Write-Log -Message "- $status: $($update.Title) [KB$($update.KB)]" -Level "INFO"
            }
            
            # Check if reboot is required
            $rebootRequired = Get-WURebootStatus -Silent
            
            if ($rebootRequired) {
                Write-Log -Message "System reboot is required to complete the update process." -Level "WARNING"
                
                if ($Params.ContainsKey("AutoReboot")) {
                    Write-Log -Message "System will reboot automatically." -Level "WARNING"
                }
                elseif ($Params.ContainsKey("ScheduleReboot")) {
                    Write-Log -Message "System will reboot at scheduled time: $($Params.ScheduleReboot)" -Level "WARNING"
                }
                elseif ($Params.ContainsKey("IgnoreReboot")) {
                    Write-Log -Message "Reboot requirement ignored as specified." -Level "WARNING"
                }
                else {
                    Write-Log -Message "Please reboot the system manually to complete the update process." -Level "WARNING"
                }
            }
            else {
                Write-Log -Message "No reboot is required." -Level "INFO"
            }
        }
        
        return $updates
    }
    catch {
        Write-Log -Message "Failed to install Windows updates: $_" -Level "ERROR"
        return $null
    }
}

# Function to hide Windows updates
function Hide-WindowsUpdates {
    param (
        [hashtable]$Params
    )
    
    try {
        Write-Log -Message "Hiding Windows updates..." -Level "INFO"
        
        if (-not $Params.ContainsKey("KBArticleID")) {
            throw "KB article IDs must be specified when hiding updates."
        }
        
        Write-Log -Message "Hiding updates with KB articles: $($Params.KBArticleID -join ', ')" -Level "INFO"
        
        $results = foreach ($kb in $Params.KBArticleID) {
            Hide-WindowsUpdate -KBArticleID $kb -Hide $true -Confirm:$false
        }
        
        Write-Log -Message "Successfully processed hide requests for updates." -Level "INFO"
        
        return $results
    }
    catch {
        Write-Log -Message "Failed to hide Windows updates: $_" -Level "ERROR"
        return $null
    }
}

# Function to unhide Windows updates
function Unhide-WindowsUpdates {
    param (
        [hashtable]$Params
    )
    
    try {
        Write-Log -Message "Unhiding Windows updates..." -Level "INFO"
        
        if (-not $Params.ContainsKey("KBArticleID")) {
            throw "KB article IDs must be specified when unhiding updates."
        }
        
        Write-Log -Message "Unhiding updates with KB articles: $($Params.KBArticleID -join ', ')" -Level "INFO"
        
        $results = foreach ($kb in $Params.KBArticleID) {
            Hide-WindowsUpdate -KBArticleID $kb -Hide:$false -Confirm:$false
        }
        
        Write-Log -Message "Successfully processed unhide requests for updates." -Level "INFO"
        
        return $results
    }
    catch {
        Write-Log -Message "Failed to unhide Windows updates: $_" -Level "ERROR"
        return $null
    }
}

# Function to get Windows update history
function Get-WindowsUpdateHistory {
    param (
        [hashtable]$Params
    )
    
    try {
        Write-Log -Message "Getting Windows update history..." -Level "INFO"
        
        if ($Params.ContainsKey("ComputerName")) {
            Write-Log -Message "Getting update history from remote computers: $($Params.ComputerName -join ', ')" -Level "INFO"
        }
        
        $history = Get-WUHistory
        
        if ($history.Count -eq 0) {
            Write-Log -Message "No update history found." -Level "INFO"
        }
        else {
            Write-Log -Message "Found $($history.Count) update history entries:" -Level "INFO"
            
            foreach ($entry in $history) {
                $status = $entry.Status
                $date = $entry.Date
                $title = $entry.Title
                
                Write-Log -Message "- [$date] $status: $title" -Level "INFO"
            }
        }
        
        return $history
    }
    catch {
        Write-Log -Message "Failed to get Windows update history: $_" -Level "ERROR"
        return $null
    }
}

# Main script execution
try {
    Write-Log -Message "Starting Windows Update management process." -Level "INFO"
    Write-Log -Message "Action: $Action" -Level "INFO"
    
    # Ensure PSWindowsUpdate module is installed
    $moduleReady = Ensure-PSWindowsUpdateModule
    
    if (-not $moduleReady) {
        Write-Log -Message "Failed to ensure PSWindowsUpdate module is ready. Exiting..." -Level "ERROR"
        exit 1
    }
    
    # Build parameters for Windows Update commands
    $params = Build-WindowsUpdateParams
    
    # Perform the requested action
    switch ($Action) {
        "Search" {
            $result = Search-WindowsUpdates -Params $params
        }
        "Download" {
            $result = Download-WindowsUpdates -Params $params
        }
        "Install" {
            $result = Install-WindowsUpdates -Params $params
        }
        "Hide" {
            $result = Hide-WindowsUpdates -Params $params
        }
        "Unhide" {
            $result = Unhide-WindowsUpdates -Params $params
        }
        "GetHistory" {
            $result = Get-WindowsUpdateHistory -Params $params
        }
    }
    
    Write-Log -Message "Windows Update management process completed." -Level "INFO"
    
    return $result
}
catch {
    Write-Log -Message "An error occurred during Windows Update management process: $_" -Level "ERROR"
    exit 1
}
