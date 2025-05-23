# Holy Hand Grenade - MECM & AD Device Manager

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue?logo=powershell&logoColor=white)
![Windows](https://img.shields.io/badge/Windows-10%2F11-blue?logo=windows&logoColor=white)
![Lines of Code](https://img.shields.io/badge/Lines%20of%20Code-762-brightgreen)
![File Size](https://img.shields.io/badge/File%20Size-35KB-orange)

## 💣 General Description

**Holy Hand Grenade** is intended for IT administrators to manage computer objects across both Microsoft Endpoint Configuration Manager (MECM/SCCM) and Active Directory on corporate domain. It provides a simple interface for common operations, eliminating the need to switch between multiple management consoles and reducing the time-sink of bulk device operations.

The name is an ode to a similiar script called "Killer Bunny" (a [Monty Python reference](https://www.youtube.com/watch?v=-IOMNUayJjI)).

> *"First shalt thou take out the Holy Pin. Then shalt thou count to three, no more, no less. Three shall be the number thou shalt count, and the number of the counting shall be three. Four shalt thou not count, neither count thou two, excepting that thou then proceed to three."* 🐰💣

<hr>
<div align="center">
<img src="https://github.com/user-attachments/assets/1a6d2219-bd9d-4e6e-a05c-4e5b8d58c8a7">
</div>

## ✨ Features

### 🔐 **Automatic Privilege Elevation**
- **Forced UAC Elevation**: Automatically detects if running without administrator privileges and re-launches with elevation
- **Hidden Console**: Runs with a hidden PowerShell console for a cleaner user experience

### 🔍 **Intelligent Dependency Management**
- **RSAT Auto-Installation**: Automatically detects and offers to install Remote Server Administration Tools (Active Directory module)
- **MECM Console Detection**: Checks for Configuration Manager Console availability and gracefully handles missing components
- **PowerShell Version Validation**: Requires PowerShell 5.1+ with clear version feedback
- **Active Directory Connectivity**: Tests AD Web Services connection and provides meaningful error messages

### 📊 **Flexible Input Methods**
- **Manual Entry**: Type computer names directly in a multi-line text box
- **File Import**: Import computer lists from .txt files with one computer name per line
- **Line-Separated Processing**: Handles large batches of computers efficiently
- **Input Validation**: Real-time validation with visual feedback

### 🎯 **Comprehensive Action Set**
The tool supports 8 distinct operations for complete device lifecycle management:

| Action | Description |
|--------|-------------|
| **Move and Enable** | 📦➡️✅ Moves computer to target OU and enables the account |
| **Move and Disable** | 📦➡️❌ Moves computer to target OU and disables the account |
| **Move Only** | 📦➡️ Moves computer to target OU without changing enabled state |
| **Enable Only** | ✅ Enables computer account in current location |
| **Delete - AD Only** | 🗑️ Removes computer from Active Directory only |
| **Delete - MECM Only** | 🗑️ Removes computer from MECM/Configuration Manager only |
| **Delete - MECM & AD** | 🗑️🗑️ Removes computer from both systems simultaneously |
| **Analyze** | 🔍 Reports computer status in both AD and MECM without making changes |

## 📋 Prerequisites

- **Windows 10/11** with PowerShell 5.1 or later
- **Administrator privileges** (A-account, automatic elevation handled)
- **RSAT Tools** (automatic installation offered)
- **Configuration Manager Console** (optional, for MECM operations)
- **Active Directory Access**
- **Network connectivity** to domain controllers and MECM infrastructure

## 🚀 Quick Start

1. **Download** `HolyHandGrenade.ps1`
2. **Right-click** → Run with PowerShell (UAC prompt will appear)
3. **Allow dependency installation** if prompted (RSAT)
4. **Enter target OU** and computer names
5. **Select desired action** from dropdown
6. **Click Run** to execute operation

## ⚙️ Configuration

The script includes several configurable parameters at the top:

```powershell
# Predefined OUs for dropdown
$predefinedOUs = @(
    "OU=General Use,OU=Workstations,DC=corp,DC=domain,DC=com",
    "OU=Disposed,OU=Workstations,DC=corp,DC=domain,DC=com"
)

# MECM Configuration
$CollectionId = 'COL0001'
$SiteCode = "SITE"
$ProviderMachineName = "server.provider.com"
```
