# Windows Update PowerShell Scripts

A comprehensive collection of PowerShell scripts for managing Windows Updates in on-premises environments.

## Overview

This repository contains a set of PowerShell scripts designed to help administrators manage Windows updates in on-premises environments. The scripts provide robust functionality for automating common update management tasks with detailed logging, error handling, and comprehensive parameter options.

## Table of Contents

- [Requirements](#requirements)
- [Installation](#installation)
- [Windows Update Management Scripts](#windows-update-management-scripts)
- [Usage Examples](#usage-examples)
- [Contributing](#contributing)
- [License](#license)

## Requirements

- Windows 10/11 or Windows Server 2019+
- PowerShell 5.1 or PowerShell 7.x
- Administrator privileges for some operations
- Internet connectivity for downloading updates

### Module Dependencies

- **PSWindowsUpdate** - Required for Windows Update scripts

## Installation

1. Clone or download this repository to your local machine
2. Ensure you have the required modules installed:

```powershell
# Install PSWindowsUpdate module
Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser
```

3. Run the scripts with appropriate parameters as needed

## Windows Update Management Scripts

### Manage-WindowsUpdates.ps1

A comprehensive script for managing Windows Updates with various filtering and installation options.

#### Features

- Search, download, install, hide, and unhide Windows updates
- Filter updates by category (Security, Critical, Important, etc.)
- Target specific KB article IDs
- Control reboot behavior (auto-reboot, ignore reboot, schedule reboot)
- Manage updates on remote computers
- Detailed logging and error handling

#### Parameters

| Parameter | Description |
|-----------|-------------|
| Action | Action to perform (Search, Download, Install, Hide, Unhide, GetHistory) |
| Categories | Categories of updates to include |
| KBArticleIDs | Specific KB article IDs to target |
| MaxUpdates | Maximum number of updates to process |
| AutoReboot | Whether to automatically reboot after installing updates |
| IgnoreReboot | Whether to ignore reboot requirements |
| ScheduleReboot | Schedule a reboot at a specific time |
| RebootTime | Time to schedule reboot |
| ComputerName | Remote computer names to manage updates on |
| Credential | Credentials to use for remote computers |
| LogPath | Path where logs will be stored |

#### Examples

```powershell
# Search for security and critical updates
.\Manage-WindowsUpdates.ps1 -Action Search -Categories Security,Critical

# Install security and critical updates with automatic reboot
.\Manage-WindowsUpdates.ps1 -Action Install -Categories Security,Critical -AutoReboot $true

# Install specific updates by KB article ID and ignore reboot
.\Manage-WindowsUpdates.ps1 -Action Install -KBArticleIDs KB5001567,KB5003173 -IgnoreReboot $true
```

### Schedule-WindowsUpdateMaintenance.ps1

Creates and manages scheduled maintenance windows for Windows Updates.

#### Features

- Define maintenance schedules for update installation
- Configure update behavior during maintenance windows
- Generate reports on update compliance
- Schedule recurring maintenance windows
- Detailed logging and reporting

#### Parameters

| Parameter | Description |
|-----------|-------------|
| Action | Action to perform (Create, Remove, List, Start, Report) |
| Name | Name of the maintenance window |
| StartTime | Start time for the maintenance window |
| Duration | Duration of the maintenance window in minutes |
| DaysOfWeek | Days of the week when the maintenance window should run |
| UpdateCategories | Categories of updates to install during the maintenance window |
| AllowReboot | Whether to allow reboots during the maintenance window |
| ComputerName | Remote computer names to manage maintenance windows on |
| Credential | Credentials to use for remote computers |
| LogPath | Path where logs will be stored |

#### Examples

```powershell
# Create a weekly maintenance window for security updates
.\Schedule-WindowsUpdateMaintenance.ps1 -Action Create -Name "Weekly Security Updates" -StartTime "22:00" -Duration 120 -DaysOfWeek Sunday -UpdateCategories Security,Critical -AllowReboot $true

# Start a maintenance window manually
.\Schedule-WindowsUpdateMaintenance.ps1 -Action Start -Name "Weekly Security Updates"

# Generate a report for a maintenance window
.\Schedule-WindowsUpdateMaintenance.ps1 -Action Report -Name "Weekly Security Updates"
```

## Usage Examples

### Basic Windows Update Management

```powershell
# Check for available updates
.\Manage-WindowsUpdates.ps1 -Action Search

# Install all available updates without rebooting
.\Manage-WindowsUpdates.ps1 -Action Install -IgnoreReboot $true

# Get update history
.\Manage-WindowsUpdates.ps1 -Action GetHistory
```

### Scheduled Update Maintenance

```powershell
# Create a maintenance window for updates
.\Schedule-WindowsUpdateMaintenance.ps1 -Action Create -Name "Monthly Updates" -StartTime "03:00" -Duration 180 -DaysOfWeek Saturday -UpdateCategories Security,Critical,Important -AllowReboot $true

# List all configured maintenance windows
.\Schedule-WindowsUpdateMaintenance.ps1 -Action List
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Author

Michael Witzsche
