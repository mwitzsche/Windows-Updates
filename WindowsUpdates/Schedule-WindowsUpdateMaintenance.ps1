<#
.SYNOPSIS
    Schedules and manages Windows Update maintenance windows.

.DESCRIPTION
    This script creates and manages scheduled maintenance windows for Windows Updates.
    It allows for defining maintenance schedules, configuring update behavior during
    maintenance windows, and generating reports on update compliance.

.PARAMETER Action
    Action to perform (Create, Remove, List, Start, Report).

.PARAMETER Name
    Name of the maintenance window.

.PARAMETER StartTime
    Start time for the maintenance window.

.PARAMETER Duration
    Duration of the maintenance window in minutes.

.PARAMETER DaysOfWeek
    Days of the week when the maintenance window should run.

.PARAMETER UpdateCategories
    Categories of updates to install during the maintenance window.

.PARAMETER AllowReboot
    Whether to allow reboots during the maintenance window.

.PARAMETER ComputerName
    Remote computer names to manage maintenance windows on.

.PARAMETER Credential
    Credentials to use for remote computers.

.PARAMETER LogPath
    Path where logs will be stored.

.EXAMPLE
    .\Schedule-WindowsUpdateMaintenance.ps1 -Action Create -Name "Weekly Security Updates" -StartTime "22:00" -Duration 120 -DaysOfWeek Sunday -UpdateCategories Security,Critical -AllowReboot $true

.EXAMPLE
    .\Schedule-WindowsUpdateMaintenance.ps1 -Action Start -Name "Weekly Security Updates"

.EXAMPLE
    .\Schedule-WindowsUpdateMaintenance.ps1 -Action Report -Name "Weekly Security Updates"

.NOTES
    Author: Michael Witzsche
    Date: April 26, 2025
    Version: 1.0.0
    Requires: PSWindowsUpdate module
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidateSet("Create", "Remove", "List", "Start", "Report")]
    [string]$Action,

    [Parameter(Mandatory = $false)]
    [string]$Name,

    [Parameter(Mandatory = $false)]
    [string]$StartTime,

    [Parameter(Mandatory = $false)]
    [int]$Duration = 120,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")]
    [string[]]$DaysOfWeek,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Security", "Critical", "Important", "Optional", "Feature", "Driver")]
    [string[]]$UpdateCategories = @("Security", "Critical"),

    [Parameter(Mandatory = $false)]
    [bool]$AllowReboot = $false,

    [Parameter(Mandatory = $false)]
    [string[]]$ComputerName,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "$env:USERPROFILE\Documents\WindowsUpdates\MaintenanceWindows_$(Get-Date -Format 'yyyyMMdd').log"
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

# Function to validate maintenance window parameters
function Validate-MaintenanceWindowParameters {
    try {
        Write-Log -Message "Validating maintenance window parameters..." -Level "INFO"
        
        if ([string]::IsNullOrEmpty($Name)) {
            throw "Maintenance window name is required."
        }
        
        if ($Action -eq "Create") {
            if ([string]::IsNullOrEmpty($StartTime)) {
                throw "Start time is required for creating a maintenance window."
            }
            
            if ($Duration -le 0) {
                throw "Duration must be greater than 0 minutes."
            }
            
            if (-not $DaysOfWeek -or $DaysOfWeek.Count -eq 0) {
                throw "At least one day of the week must be specified."
            }
            
            # Validate start time format
            try {
                $timeFormat = [datetime]::ParseExact($StartTime, "HH:mm", [System.Globalization.CultureInfo]::InvariantCulture)
            }
            catch {
                throw "Start time must be in 24-hour format (HH:mm)."
            }
        }
        
        return $true
    }
    catch {
        Write-Log -Message "Parameter validation failed: $_" -Level "ERROR"
        return $false
    }
}

# Function to get maintenance window configuration path
function Get-MaintenanceWindowConfigPath {
    param (
        [Parameter(Mandatory = $false)]
        [string]$Computer = $env:COMPUTERNAME
    )
    
    if ($Computer -eq $env:COMPUTERNAME) {
        return "$env:ProgramData\WindowsUpdateMaintenance"
    }
    else {
        # For remote computers, we'll use a temporary path and then copy the configuration
        return "$env:TEMP\WindowsUpdateMaintenance_$Computer"
    }
}

# Function to create a maintenance window
function Create-MaintenanceWindow {
    try {
        Write-Log -Message "Creating maintenance window '$Name'..." -Level "INFO"
        
        # Get configuration path
        $configPath = Get-MaintenanceWindowConfigPath
        
        # Create configuration directory if it doesn't exist
        if (-not (Test-Path -Path $configPath)) {
            New-Item -Path $configPath -ItemType Directory -Force | Out-Null
            Write-Log -Message "Created configuration directory: $configPath" -Level "INFO"
        }
        
        # Check if maintenance window already exists
        $maintenanceWindowFile = Join-Path -Path $configPath -ChildPath "$Name.xml"
        
        if (Test-Path -Path $maintenanceWindowFile) {
            Write-Log -Message "Maintenance window '$Name' already exists. Updating configuration..." -Level "WARNING"
        }
        
        # Create maintenance window configuration
        $maintenanceWindow = @{
            Name = $Name
            StartTime = $StartTime
            Duration = $Duration
            DaysOfWeek = $DaysOfWeek
            UpdateCategories = $UpdateCategories
            AllowReboot = $AllowReboot
            CreatedOn = Get-Date
            LastModified = Get-Date
            LastRun = $null
            NextRun = $null
        }
        
        # Calculate next run time
        $now = Get-Date
        $nextRun = $null
        
        foreach ($day in $DaysOfWeek) {
            $dayOfWeek = [System.DayOfWeek]::$day
            $daysUntil = ($dayOfWeek - $now.DayOfWeek + 7) % 7
            
            if ($daysUntil -eq 0) {
                # Today is the scheduled day, check if the start time has passed
                $startDateTime = [datetime]::ParseExact($StartTime, "HH:mm", [System.Globalization.CultureInfo]::InvariantCulture)
                $scheduledTime = $now.Date.AddHours($startDateTime.Hour).AddMinutes($startDateTime.Minute)
                
                if ($scheduledTime -gt $now) {
                    # Today's scheduled time is in the future
                    $candidateNextRun = $scheduledTime
                }
                else {
                    # Today's scheduled time has passed, use next week
                    $candidateNextRun = $scheduledTime.AddDays(7)
                }
            }
            else {
                # Scheduled day is in the future
                $startDateTime = [datetime]::ParseExact($StartTime, "HH:mm", [System.Globalization.CultureInfo]::InvariantCulture)
                $candidateNextRun = $now.Date.AddDays($daysUntil).AddHours($startDateTime.Hour).AddMinutes($startDateTime.Minute)
            }
            
            if ($null -eq $nextRun -or $candidateNextRun -lt $nextRun) {
                $nextRun = $candidateNextRun
            }
        }
        
        $maintenanceWindow.NextRun = $nextRun
        
        # Save maintenance window configuration
        $maintenanceWindow | Export-Clixml -Path $maintenanceWindowFile
        
        Write-Log -Message "Maintenance window '$Name' created successfully." -Level "INFO"
        Write-Log -Message "Next run scheduled for: $($nextRun)" -Level "INFO"
        
        # Create scheduled task
        $taskName = "WindowsUpdateMaintenance_$Name"
        $taskDescription = "Windows Update Maintenance Window: $Name"
        $scriptPath = $MyInvocation.MyCommand.Path
        
        # Create trigger for each day of the week
        $triggers = @()
        
        foreach ($day in $DaysOfWeek) {
            $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $day -At $StartTime
            $triggers += $trigger
        }
        
        # Create action to run this script with Start action
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"$scriptPath`" -Action Start -Name `"$Name`""
        
        # Register the scheduled task
        $taskExists = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        
        if ($taskExists) {
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
            Write-Log -Message "Removed existing scheduled task: $taskName" -Level "INFO"
        }
        
        Register-ScheduledTask -TaskName $taskName -Description $taskDescription -Trigger $triggers -Action $action -RunLevel Highest
        Write-Log -Message "Created scheduled task: $taskName" -Level "INFO"
        
        return $maintenanceWindow
    }
    catch {
        Write-Log -Message "Failed to create maintenance window: $_" -Level "ERROR"
        return $null
    }
}

# Function to remove a maintenance window
function Remove-MaintenanceWindow {
    try {
        Write-Log -Message "Removing maintenance window '$Name'..." -Level "INFO"
        
        # Get configuration path
        $configPath = Get-MaintenanceWindowConfigPath
        
        # Check if maintenance window exists
        $maintenanceWindowFile = Join-Path -Path $configPath -ChildPath "$Name.xml"
        
        if (-not (Test-Path -Path $maintenanceWindowFile)) {
            throw "Maintenance window '$Name' does not exist."
        }
        
        # Remove maintenance window configuration
        Remove-Item -Path $maintenanceWindowFile -Force
        Write-Log -Message "Removed maintenance window configuration file." -Level "INFO"
        
        # Remove scheduled task
        $taskName = "WindowsUpdateMaintenance_$Name"
        $taskExists = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        
        if ($taskExists) {
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
            Write-Log -Message "Removed scheduled task: $taskName" -Level "INFO"
        }
        
        Write-Log -Message "Maintenance window '$Name' removed successfully." -Level "INFO"
        
        return $true
    }
    catch {
        Write-Log -Message "Failed to remove maintenance window: $_" -Level "ERROR"
        return $false
    }
}

# Function to list maintenance windows
function List-MaintenanceWindows {
    try {
        Write-Log -Message "Listing maintenance windows..." -Level "INFO"
        
        # Get configuration path
        $configPath = Get-MaintenanceWindowConfigPath
        
        # Check if configuration directory exists
        if (-not (Test-Path -Path $configPath)) {
            Write-Log -Message "No maintenance windows found." -Level "INFO"
            return @()
        }
        
        # Get all maintenance window configuration files
        $maintenanceWindowFiles = Get-ChildItem -Path $configPath -Filter "*.xml"
        
        if ($maintenanceWindowFiles.Count -eq 0) {
            Write-Log -Message "No maintenance windows found." -Level "INFO"
            return @()
        }
        
        # Load maintenance window configurations
        $maintenanceWindows = @()
        
        foreach ($file in $maintenanceWindowFiles) {
            $maintenanceWindow = Import-Clixml -Path $file.FullName
            $maintenanceWindows += $maintenanceWindow
        }
        
        Write-Log -Message "Found $($maintenanceWindows.Count) maintenance windows:" -Level "INFO"
        
        foreach ($window in $maintenanceWindows) {
            $nextRun = if ($window.NextRun) { $window.NextRun.ToString("yyyy-MM-dd HH:mm") } else { "Not scheduled" }
            $lastRun = if ($window.LastRun) { $window.LastRun.ToString("yyyy-MM-dd HH:mm") } else { "Never" }
            
            Write-Log -Message "- $($window.Name)" -Level "INFO"
            Write-Log -Message "  Schedule: $($window.StartTime) on $($window.DaysOfWeek -join ', ') for $($window.Duration) minutes" -Level "INFO"
            Write-Log -Message "  Categories: $($window.UpdateCategories -join ', ')" -Level "INFO"
            Write-Log -Message "  Allow Reboot: $($window.AllowReboot)" -Level "INFO"
            Write-Log -Message "  Last Run: $lastRun" -Level "INFO"
            Write-Log -Message "  Next Run: $nextRun" -Level "INFO"
        }
        
        return $maintenanceWindows
    }
    catch {
        Write-Log -Message "Failed to list maintenance windows: $_" -Level "ERROR"
        return $null
    }
}

# Function to start a maintenance window
function Start-MaintenanceWindow {
    try {
        Write-Log -Message "Starting maintenance window '$Name'..." -Level "INFO"
        
        # Get configuration path
        $configPath = Get-MaintenanceWindowConfigPath
        
        # Check if maintenance window exists
        $maintenanceWindowFile = Join-Path -Path $configPath -ChildPath "$Name.xml"
        
        if (-not (Test-Path -Path $maintenanceWindowFile)) {
            throw "Maintenance window '$Name' does not exist."
        }
        
        # Load maintenance window configuration
        $maintenanceWindow = Import-Clixml -Path $maintenanceWindowFile
        
        Write-Log -Message "Loaded maintenance window configuration:" -Level "INFO"
        Write-Log -Message "- Name: $($maintenanceWindow.Name)" -Level "INFO"
        Write-Log -Message "- Schedule: $($maintenanceWindow.StartTime) on $($maintenanceWindow.DaysOfWeek -join ', ') for $($maintenanceWindow.Duration) minutes" -Level "INFO"
        Write-Log -Message "- Categories: $($maintenanceWindow.UpdateCategories -join ', ')" -Level "INFO"
        Write-Log -Message "- Allow Reboot: $($maintenanceWindow.AllowReboot)" -Level "INFO"
        
        # Ensure PSWindowsUpdate module is installed
        $moduleReady = Ensure-PSWindowsUpdateModule
        
        if (-not $moduleReady) {
            throw "Failed to ensure PSWindowsUpdate module is ready."
        }
        
        # Build parameters for Windows Update installation
        $params = @{
            AcceptAll = $true
        }
        
        # Add categories
        $categoryFilter = @()
        
        foreach ($category in $maintenanceWindow.UpdateCategories) {
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
        
        # Add reboot parameters
        if ($maintenanceWindow.AllowReboot) {
            $params.Add("AutoReboot", $true)
            Write-Log -Message "Automatic reboot is enabled." -Level "INFO"
        }
        else {
            $params.Add("IgnoreReboot", $true)
            Write-Log -Message "Automatic reboot is disabled." -Level "INFO"
        }
        
        # Calculate end time
        $startTime = Get-Date
        $endTime = $startTime.AddMinutes($maintenanceWindow.Duration)
        
        Write-Log -Message "Maintenance window started at $startTime" -Level "INFO"
        Write-Log -Message "Maintenance window will end at $endTime" -Level "INFO"
        
        # Install updates
        Write-Log -Message "Installing Windows updates..." -Level "INFO"
        $updates = Install-WindowsUpdate @params
        
        # Update maintenance window configuration
        $maintenanceWindow.LastRun = $startTime
        
        # Calculate next run time
        $now = Get-Date
        $nextRun = $null
        
        foreach ($day in $maintenanceWindow.DaysOfWeek) {
            $dayOfWeek = [System.DayOfWeek]::$day
            $daysUntil = ($dayOfWeek - $now.DayOfWeek + 7) % 7
            
            if ($daysUntil -eq 0) {
                # Today is the scheduled day, use next week
                $startDateTime = [datetime]::ParseExact($maintenanceWindow.StartTime, "HH:mm", [System.Globalization.CultureInfo]::InvariantCulture)
                $scheduledTime = $now.Date.AddHours($startDateTime.Hour).AddMinutes($startDateTime.Minute).AddDays(7)
                $candidateNextRun = $scheduledTime
            }
            else {
                # Scheduled day is in the future
                $startDateTime = [datetime]::ParseExact($maintenanceWindow.StartTime, "HH:mm", [System.Globalization.CultureInfo]::InvariantCulture)
                $candidateNextRun = $now.Date.AddDays($daysUntil).AddHours($startDateTime.Hour).AddMinutes($startDateTime.Minute)
            }
            
            if ($null -eq $nextRun -or $candidateNextRun -lt $nextRun) {
                $nextRun = $candidateNextRun
            }
        }
        
        $maintenanceWindow.NextRun = $nextRun
        
        # Save updated maintenance window configuration
        $maintenanceWindow | Export-Clixml -Path $maintenanceWindowFile
        
        Write-Log -Message "Maintenance window completed." -Level "INFO"
        Write-Log -Message "Next run scheduled for: $($nextRun)" -Level "INFO"
        
        return $updates
    }
    catch {
        Write-Log -Message "Failed to start maintenance window: $_" -Level "ERROR"
        return $null
    }
}

# Function to generate a maintenance window report
function Generate-MaintenanceWindowReport {
    try {
        Write-Log -Message "Generating report for maintenance window '$Name'..." -Level "INFO"
        
        # Get configuration path
        $configPath = Get-MaintenanceWindowConfigPath
        
        # Check if maintenance window exists
        $maintenanceWindowFile = Join-Path -Path $configPath -ChildPath "$Name.xml"
        
        if (-not (Test-Path -Path $maintenanceWindowFile)) {
            throw "Maintenance window '$Name' does not exist."
        }
        
        # Load maintenance window configuration
        $maintenanceWindow = Import-Clixml -Path $maintenanceWindowFile
        
        # Get update history
        $history = Get-WUHistory
        
        # Filter history to include only updates installed during the maintenance window
        $maintenanceWindowHistory = @()
        
        if ($maintenanceWindow.LastRun) {
            $lastRunDate = $maintenanceWindow.LastRun.Date
            
            foreach ($entry in $history) {
                if ($entry.Date -ge $lastRunDate) {
                    $maintenanceWindowHistory += $entry
                }
            }
        }
        
        # Generate report
        $report = @{
            MaintenanceWindow = $maintenanceWindow
            UpdateHistory = $maintenanceWindowHistory
            GeneratedAt = Get-Date
            ComputerName = $env:COMPUTERNAME
        }
        
        # Save report
        $reportDir = "$env:ProgramData\WindowsUpdateMaintenance\Reports"
        
        if (-not (Test-Path -Path $reportDir)) {
            New-Item -Path $reportDir -ItemType Directory -Force | Out-Null
        }
        
        $reportFile = Join-Path -Path $reportDir -ChildPath "$($Name)_Report_$(Get-Date -Format 'yyyyMMdd').xml"
        $report | Export-Clixml -Path $reportFile
        
        Write-Log -Message "Report generated successfully and saved to: $reportFile" -Level "INFO"
        
        # Display report summary
        Write-Log -Message "Report Summary:" -Level "INFO"
        Write-Log -Message "- Maintenance Window: $($maintenanceWindow.Name)" -Level "INFO"
        Write-Log -Message "- Last Run: $($maintenanceWindow.LastRun)" -Level "INFO"
        Write-Log -Message "- Next Run: $($maintenanceWindow.NextRun)" -Level "INFO"
        Write-Log -Message "- Updates Installed: $($maintenanceWindowHistory.Count)" -Level "INFO"
        
        foreach ($entry in $maintenanceWindowHistory) {
            Write-Log -Message "  - $($entry.Title) [Status: $($entry.Status)]" -Level "INFO"
        }
        
        return $report
    }
    catch {
        Write-Log -Message "Failed to generate maintenance window report: $_" -Level "ERROR"
        return $null
    }
}

# Main script execution
try {
    Write-Log -Message "Starting Windows Update maintenance window management process." -Level "INFO"
    Write-Log -Message "Action: $Action" -Level "INFO"
    
    # Validate parameters
    $parametersValid = Validate-MaintenanceWindowParameters
    
    if (-not $parametersValid) {
        Write-Log -Message "Parameter validation failed. Exiting..." -Level "ERROR"
        exit 1
    }
    
    # Perform the requested action
    switch ($Action) {
        "Create" {
            $result = Create-MaintenanceWindow
        }
        "Remove" {
            $result = Remove-MaintenanceWindow
        }
        "List" {
            $result = List-MaintenanceWindows
        }
        "Start" {
            $result = Start-MaintenanceWindow
        }
        "Report" {
            $result = Generate-MaintenanceWindowReport
        }
    }
    
    Write-Log -Message "Windows Update maintenance window management process completed." -Level "INFO"
    
    return $result
}
catch {
    Write-Log -Message "An error occurred during Windows Update maintenance window management process: $_" -Level "ERROR"
    exit 1
}
