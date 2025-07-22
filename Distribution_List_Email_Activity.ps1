```powershell
<#
.SYNOPSIS
    Checks email activity for Distribution Lists and Microsoft 365 Groups in Exchange Online.

.DESCRIPTION
    This script provides comprehensive email activity reporting for Exchange Online Distribution Lists,
    including recent activity (10 days) and historical search options (up to 90 days).
    
    Features:
    - Recent email activity tracking
    - Distribution list/group information
    - Member enumeration
    - Historical search capabilities
    - Failed delivery tracking

.PARAMETER DistributionListEmail
    The email address of the distribution list to check

.PARAMETER DaysToCheck
    Number of days to check for recent activity (maximum 10 for real-time search)

.EXAMPLE
    .\Check-DLEmailActivity.ps1
    Runs the script interactively, prompting for distribution list email

.EXAMPLE
    .\Check-DLEmailActivity.ps1 -DistributionListEmail "team@company.com" -DaysToCheck 7
    Checks specific distribution list for the last 7 days

.NOTES
    Author: Your Name
    Version: 1.2.0
    Requires: Exchange Online Management PowerShell module
    Permissions: Exchange Administrator or Global Administrator
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$DistributionListEmail,
    
    [Parameter(Mandatory=$false)]
    [ValidateRange(1,10)]
    [int]$DaysToCheck = 9
)

# Import required modules
Import-Module ExchangeOnlineManagement

function Connect-ToExchangeOnlineIfNeeded {
    <#
    .SYNOPSIS
        Connects to Exchange Online if not already connected
    #>
    try {
        # Check if already connected
        $ConnectionInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
        
        if (-not $ConnectionInfo) {
            Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
            Connect-ExchangeOnline -ShowProgress $false
            Write-Host "Successfully connected to Exchange Online" -ForegroundColor Green
        }
        else {
            Write-Host "Already connected to Exchange Online" -ForegroundColor Green
        }
        return $true
    }
    catch {
        Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
        return $false
    }
}

function Get-DLEmailActivity {
    <#
    .SYNOPSIS
        Gets email activity for a distribution list
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$DistributionListEmail,
        
        [Parameter(Mandatory=$false)]
        [int]$DaysBack = 9
    )
    
    Write-Host "`n=== Checking Distribution List Email Activity ===" -ForegroundColor Cyan
    Write-Host "DL Email: $DistributionListEmail" -ForegroundColor White
    Write-Host "Checking last $DaysBack days..." -ForegroundColor White
    
    # Get distribution list information
    try {
        $DL = Get-DistributionGroup -Identity $DistributionListEmail -ErrorAction SilentlyContinue
        
        if ($DL) {
            Write-Host "`nDistribution List Details:" -ForegroundColor Green
            Write-Host "Name: $($DL.DisplayName)" -ForegroundColor White
            Write-Host "Email: $($DL.PrimarySmtpAddress)" -ForegroundColor White
            Write-Host "Total Members: $($DL.MemberCount)" -ForegroundColor White
            Write-Host "Created: $($DL.WhenCreated)" -ForegroundColor White
            Write-Host "Last Modified: $($DL.WhenChanged)" -ForegroundColor White
            Write-Host "Accepts External Email: $(-not $DL.RequireSenderAuthenticationEnabled)" -ForegroundColor White
        }
        else {
            # Check if it's a Microsoft 365 Group
            $group = Get-UnifiedGroup -Identity $DistributionListEmail -ErrorAction SilentlyContinue
            if ($group) {
                Write-Host "`nMicrosoft 365 Group Details:" -ForegroundColor Green
                Write-Host "Name: $($group.DisplayName)" -ForegroundColor White
                Write-Host "Email: $($group.PrimarySmtpAddress)" -ForegroundColor White
                Write-Host "Created: $($group.WhenCreated)" -ForegroundColor White
                Write-Host "Last Modified: $($group.WhenChanged)" -ForegroundColor White
                $DL = $group
            }
            else {
                Write-Host "No distribution list or group found with email: $DistributionListEmail" -ForegroundColor Red
                return $false
            }
        }
    }
    catch {
        Write-Error "Error retrieving distribution list information: $($_.Exception.Message)"
        return $false
    }
    
    # Check email activity within 10-day limit
    Write-Host "`n--- Checking Email Activity (Last $DaysBack days) ---" -ForegroundColor Yellow
    
    $EndDate = Get-Date
    $StartDate = $EndDate.AddDays(-$DaysBack)
    
    try {
        # Search for emails RECEIVED by the DL
        Write-Host "`nSearching for emails RECEIVED by the DL..." -ForegroundColor Cyan
        
        $ReceivedMessages = Get-MessageTrace -RecipientAddress $DistributionListEmail -StartDate $StartDate -EndDate $EndDate -PageSize 5000 |
                           Sort-Object Received -Descending
        
        if ($ReceivedMessages) {
            $LatestReceived = $ReceivedMessages | Select-Object -First 1
            Write-Host "✓ LATEST EMAIL RECEIVED:" -ForegroundColor Green
            Write-Host "  Date: $($LatestReceived.Received)" -ForegroundColor White
            Write-Host "  From: $($LatestReceived.SenderAddress)" -ForegroundColor White
            Write-Host "  Subject: $($LatestReceived.Subject)" -ForegroundColor White
            Write-Host "  Status: $($LatestReceived.Status)" -ForegroundColor White
            Write-Host "  Total emails received in last $DaysBack days: $($ReceivedMessages.Count)" -ForegroundColor White
            
            # Show recent activity summary
            if ($ReceivedMessages.Count -gt 1) {
                Write-Host "`n  Recent received emails:" -ForegroundColor Gray
                $ReceivedMessages | Select-Object -First 5 | ForEach-Object {
                    Write-Host "    $($_.Received) - From: $($_.SenderAddress)" -ForegroundColor Gray
                }
            }
        }
        else {
            Write-Host "✗ No emails received in the last $DaysBack days" -ForegroundColor Yellow
        }
        
        # Search for emails SENT from the DL
        Write-Host "`nSearching for emails SENT from the DL..." -ForegroundColor Cyan
        
        $SentMessages = Get-MessageTrace -SenderAddress $DistributionListEmail -StartDate $StartDate -EndDate $EndDate -PageSize 5000 |
                       Sort-Object Received -Descending
        
        if ($SentMessages) {
            $LatestSent = $SentMessages | Select-Object -First 1
            Write-Host "✓ LATEST EMAIL SENT:" -ForegroundColor Green
            Write-Host "  Date: $($LatestSent.Received)" -ForegroundColor White
            Write-Host "  To: $($LatestSent.RecipientAddress)" -ForegroundColor White
            Write-Host "  Subject: $($LatestSent.Subject)" -ForegroundColor White
            Write-Host "  Status: $($LatestSent.Status)" -ForegroundColor White
            Write-Host "  Total emails sent in last $DaysBack days: $($SentMessages.Count)" -ForegroundColor White
        }
        else {
            Write-Host "✗ No emails sent from this DL in the last $DaysBack days" -ForegroundColor Yellow
            Write-Host "  (Note: Most distribution lists only receive and forward emails)" -ForegroundColor Gray
        }
        
        # Check for failed deliveries
        $FailedMessages = $ReceivedMessages | Where-Object { $_.Status -ne "Delivered" }
        
        if ($FailedMessages) {
            Write-Host "`n⚠ Failed/Pending Deliveries:" -ForegroundColor Red
            $FailedMessages | Select-Object -First 5 | ForEach-Object {
                Write-Host "  $($_.Received) - From: $($_.SenderAddress) - Status: $($_.Status)" -ForegroundColor Red
            }
        }
        
        return $true
        
    }
    catch {
        Write-Error "Error checking email activity: $($_.Exception.Message)"
        return $false
    }
}

function Start-HistoricalEmailSearch {
    <#
    .SYNOPSIS
        Starts a historical email search for extended date ranges
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$DistributionListEmail
    )
    
    Write-Host "`n--- Starting Historical Email Search ---" -ForegroundColor Cyan
    
    try {
        # Calculate date range (11-90 days ago)
        $EndDate = (Get-Date).AddDays(-10)
        $StartDate = (Get-Date).AddDays(-90)
        
        Write-Host "Starting historical search for emails RECEIVED by $DistributionListEmail..." -ForegroundColor Yellow
        Write-Host "Date range: $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))" -ForegroundColor Gray
        
        $SearchName = "DL-Activity-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
        
        $HistoricalSearch = Start-HistoricalSearch -ReportTitle $SearchName -StartDate $StartDate -EndDate $EndDate -ReportType MessageTrace -RecipientAddress $DistributionListEmail
        
        if ($HistoricalSearch) {
            Write-Host "✓ Historical search started successfully!" -ForegroundColor Green
            Write-Host "Search ID: $($HistoricalSearch.JobId)" -ForegroundColor White
            Write-Host "Status: $($HistoricalSearch.Status)" -ForegroundColor White
            
            # Monitor search progress
            Write-Host "`nMonitoring search progress..." -ForegroundColor Yellow
            $MaxWaitTime = 300  # 5 minutes maximum wait
            $WaitTime = 0
            
            do {
                Start-Sleep -Seconds 10
                $WaitTime += 10
                $SearchStatus = Get-HistoricalSearch -JobId $HistoricalSearch.JobId
                Write-Host "Status: $($SearchStatus.Status) (waited $WaitTime seconds)" -ForegroundColor Gray
                
                if ($WaitTime -ge $MaxWaitTime) {
                    Write-Host "Search is taking longer than expected. You can check status later with:" -ForegroundColor Yellow
                    Write-Host "Get-HistoricalSearch -JobId $($HistoricalSearch.JobId)" -ForegroundColor White
                    break
                }
            } while ($SearchStatus.Status -eq "InProgress")
            
            if ($SearchStatus.Status -eq "Done") {
                Write-Host "✓ Historical search completed!" -ForegroundColor Green
                
                # Get download link
                $Results = Get-HistoricalSearch -JobId $HistoricalSearch.JobId
                
                if ($Results -and $Results.FileUrl) {
                    Write-Host "✓ Results are available for download:" -ForegroundColor Green
                    Write-Host "Download URL: $($Results.FileUrl)" -ForegroundColor White
                    Write-Host "`nThe CSV file contains all email activity for the specified period." -ForegroundColor Yellow
                }
                else {
                    Write-Host "No historical email activity found for this distribution list." -ForegroundColor Yellow
                }
            }
            elseif ($WaitTime -lt $MaxWaitTime) {
                Write-Host "Historical search status: $($SearchStatus.Status)" -ForegroundColor Yellow
            }
        }
        
    }
    catch {
        Write-Warning "Historical search failed: $($_.Exception.Message)"
        Write-Host "Note: Historical search requires Exchange Online Plan 2 or Office 365 E3/E5 licensing." -ForegroundColor Yellow
    }
}

function Get-DLMembers {
    <#
    .SYNOPSIS
        Gets distribution list member information
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$DistributionListEmail
    )
    
    Write-Host "`n--- Distribution List Member Information ---" -ForegroundColor Cyan
    
    try {
        $Members = Get-DistributionGroupMember -Identity $DistributionListEmail -ErrorAction Stop
        
        Write-Host "Total Members: $($Members.Count)" -ForegroundColor White
        
        if ($Members.Count -le 20) {
            Write-Host "`nMembers:" -ForegroundColor Gray
            $Members | ForEach-Object {
                Write-Host "  $($_.DisplayName) ($($_.PrimarySmtpAddress)) - $($_.RecipientType)" -ForegroundColor White
            }
        }
        else {
            Write-Host "Member list is large ($($Members.Count) members). First 20 members:" -ForegroundColor Gray
            $Members | Select-Object -First 20 | ForEach-Object {
                Write-Host "  $($_.DisplayName) ($($_.PrimarySmtpAddress)) - $($_.RecipientType)" -ForegroundColor White
            }
            Write-Host "  ... and $($Members.Count - 20) more members" -ForegroundColor Gray
        }
        
    }
    catch {
        Write-Warning "Could not retrieve member list: $($_.Exception.Message)"
    }
}

# Main execution
function Main {
    Write-Host "=== Distribution List Email Activity Checker ===" -ForegroundColor Magenta
    Write-Host "Started: $(Get-Date)" -ForegroundColor Gray
    
    # Connect to Exchange Online
    if (-not (Connect-ToExchangeOnlineIfNeeded)) {
        Write-Error "Cannot proceed without Exchange Online connection."
        return
    }
    
    # Get distribution list email if not provided
    if (-not $DistributionListEmail) {
        $DistributionListEmail = Read-Host "Enter the distribution list email address"
        
        if ([string]::IsNullOrWhiteSpace($DistributionListEmail)) {
            Write-Error "Distribution list email address is required."
            return
        }
    }
    
    Write-Host "Target: $DistributionListEmail" -ForegroundColor White
    
    # Check recent email activity
    $ActivityFound = Get-DLEmailActivity -DistributionListEmail $DistributionListEmail -DaysBack $DaysToCheck
    
    if ($ActivityFound) {
        # Offer historical search
        Write-Host "`n--- Extended Search Options ---" -ForegroundColor Magenta
        $DoHistoricalSearch = Read-Host "Do you want to perform a historical search (10-90 days ago)? (Y/N)"
        
        if ($DoHistoricalSearch -eq "Y" -or $DoHistoricalSearch -eq "y") {
            Start-HistoricalEmailSearch -DistributionListEmail $DistributionListEmail
        }
        
        # Offer to show members
        $ShowMembers = Read-Host "`nDo you want to see the distribution list members? (Y/N)"
        
        if ($ShowMembers -eq "Y" -or $ShowMembers -eq "y") {
            Get-DLMembers -DistributionListEmail $DistributionListEmail
        }
    }
    
    Write-Host "`n=== SUMMARY ===" -ForegroundColor Magenta
    Write-Host "✓ Email activity check completed for: $DistributionListEmail" -ForegroundColor Green
    Write-Host "✓ Checked recent activity (last $DaysToCheck days)" -ForegroundColor Green
    
    if ($ActivityFound) {
        Write-Host "✓ Distribution list information retrieved" -ForegroundColor Green
    }
    
    Write-Host "`nNote: For searches beyond 10 days, use the Historical Search option." -ForegroundColor Yellow
    Write-Host "Script completed at: $(Get-Date)" -ForegroundColor Gray
}

# Execute main function
Main
