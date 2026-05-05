# AD Group User Listing Script

A PowerShell script for Active Directory administrators to interactively browse groups and view their members.

## Description

This script provides an easy-to-use interactive interface for viewing Active Directory groups and their members. It displays all AD groups in a numbered list, allows you to select a group by number, and then shows all members of that group.

## Features

- Lists all Active Directory groups with numbered selection
- Interactive group selection
- Displays all members with their object type (user, computer, group, etc.)
- Shows total member count
- Color-coded output for better readability
- Input validation
- Handles empty groups gracefully
- Loop functionality: Press ENTER to run again or ESC to exit

## Prerequisites

- Windows PowerShell 5.1 or PowerShell 7+
- Active Directory PowerShell module (RSAT-AD-PowerShell)
- Appropriate Active Directory permissions to read group information

## Installation

### Installing AD PowerShell Module

**Windows 10/11:**
```powershell
# Run as Administrator
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
```

**Windows Server:**
```powershell
# Run as Administrator
Install-WindowsFeature RSAT-AD-PowerShell
```

### Getting the Script

**Option 1: Clone this repository**
```bash
git clone https://github.com/yourusername/ad-group-listing.git
cd ad-group-listing
```

**Option 2: Download ZIP**
- Click the green "Code" button above
- Select "Download ZIP"
- Extract the files to your desired location

**Option 3: Download script directly**
```powershell
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/yourusername/ad-group-listing/main/Get-ADGroupMembers.ps1" -OutFile "Get-ADGroupMembers.ps1"
```

## Usage

1. Open PowerShell with appropriate AD permissions
2. Navigate to the script directory
3. Run the script:
```powershell
.\Get-ADGroupMembers.ps1
```

4. Enter the number of the group you want to view
5. View the list of members
6. Press ENTER to search another group or ESC to exit

**Example:**