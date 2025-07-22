README.md
MarkDown

# Distribution List Email Activity Checker

A PowerShell script to check the last email activity (sent/received) for Exchange Online Distribution Lists and Microsoft 365 Groups.

## 📋 Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Limitations](#limitations)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

## 🔍 Overview

This PowerShell script helps administrators track email activity for Distribution Lists and Microsoft 365 Groups in Exchange Online. It provides detailed information about the last emails sent to or from a distribution list, helping determine if the list is actively used.

## ✨ Features

- **Real-time Email Tracking**: Check email activity for the last 10 days
- **Historical Search**: Extended search up to 90 days using Exchange Online Historical Search
- **Distribution List Details**: Get comprehensive information about the DL/Group
- **Member Information**: View distribution list membership
- **Failed Delivery Tracking**: Identify failed or pending email deliveries
- **Modern Authentication**: Uses Exchange Online PowerShell V3 with modern authentication
- **Detailed Reporting**: Comprehensive output with timestamps, senders, subjects, and status

## 📋 Prerequisites

### Required Software
- **PowerShell 5.1** or **PowerShell 7+**
- **Exchange Online Management Module**

### Required Permissions
- **Exchange Online Administrator** or **Global Administrator** role
- **Message Trace** permissions
- **Distribution Group** read permissions

### Licensing Requirements
- **Basic functionality**: Exchange Online Plan 1 or higher
- **Historical Search (90 days)**: Exchange Online Plan 2, Office 365 E3/E5, or Microsoft 365 Business Premium

## 🚀 Installation

### 1. Install Required PowerShell Module

```powershell
# Install Exchange Online Management module
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force

# Import the module
Import-Module ExchangeOnlineManagement
2. Clone or Download the Script
Bash

# Clone the repository
git clone https://github.com/yourusername/dl-email-activity-checker.git

# Or download the script directly
📖 Usage
Basic Usage
Run the script:

PowerShell

.\Check-DLEmailActivity.ps1
When prompted, enter:

Your target distribution list email address
Authentication credentials (if not already connected)
Follow the interactive prompts to:

Perform extended historical search
View distribution list members
Download historical data
Advanced Usage
Check Specific Distribution List
PowerShell

# Modify the script to target a specific DL
$DistributionListEmail = "your-dl@company.com"
Automated Execution
PowerShell

# For automated runs, pre-connect to Exchange Online
Connect-ExchangeOnline -UserPrincipalName admin@company.com

# Run the script
.\Check-DLEmailActivity.ps1
⚠️ Limitations
Exchange Online Message Trace Limitations
Feature	Time Limit	Notes
Get-MessageTrace	10 days	Real-time search, immediate results
Start-HistoricalSearch	90 days	Requires Plan 2+ licensing, results via CSV download
Message retention	90 days	Maximum searchable history
Known Issues
10-Day Limitation: Recent activity searches limited to 10 days
Historical Search Requirements: Extended searches require higher licensing tiers
Large Distribution Lists: Member enumeration may be slow for very large lists
🔧 Troubleshooting
Common Errors
"Invalid StartDate value. The StartDate can't be older than 10 days"
Cause: Attempting to use Get-MessageTrace beyond 10-day limit
Solution: Use the Historical Search option for older data
"Access Denied" or "Insufficient Permissions"
Cause: User lacks required Exchange permissions
Solution: Ensure you have Exchange Administrator role or equivalent
"Distribution list not found"
Cause: Email address doesn't exist or insufficient permissions
Solution: Verify the email address and check permissions
Performance Optimization
PowerShell

# For better performance with large datasets
$ReceivedMessages = Get-MessageTrace -RecipientAddress $DLEmail -StartDate $StartDate -EndDate $EndDate -PageSize 1000
📊 Sample Output
=== Distribution List Email Activity Report ===
Target: sales-team@company.com
Started: 01/15/2024 10:30:00

Distribution List Details:
Name: Sales Team Distribution List
Email: sales-team@company.com
Total Members: 15
Created: 01/01/2024 08:00:00
Last Modified: 01/10/2024 14:30:00
Accepts External Email: True

--- Checking Email Activity ---
✓ LATEST EMAIL RECEIVED:
  Date: 01/15/2024 09:45:00
  From: customer@client.com
  Subject: Q1 Sales Inquiry
  Status: Delivered
  Total emails received in last 10 days: 47

✗ No emails sent from this DL in the last 10 days
  (Note: Most distribution lists only receive and forward emails)
🤝 Contributing
Fork the repository
Create a feature branch (git checkout -b feature/amazing-feature)
Commit your changes (git commit -m 'Add amazing feature')
Push to the branch (git push origin feature/amazing-feature)
Open a Pull Request
Development Guidelines
Follow PowerShell best practices
Include error handling for all external calls
Add comments for complex logic
Test with different distribution list types
📝 License
This project is licensed under the MIT License - see the LICENSE file for details.

📞 Support
Issues: GitHub Issues
Discussions: GitHub Discussions
🏷️ Version History
v1.2.0 - Added historical search functionality
v1.1.0 - Added member information display
v1.0.0 - Initial release with basic email tracking
🔗 Related Resources
Exchange Online PowerShell V3
Message Trace in Exchange Online
Exchange Online Permissions
⭐ Star this repository if you find it helpful!
