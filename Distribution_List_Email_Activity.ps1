# Corrected script for Distribution List Email Activity
# Handles Exchange Online's 10-day limitation for Get-MessageTrace

Import-Module ExchangeOnlineManagement

# Connect to Exchange Online if not already connected
if (-not (Get-ConnectionInformation)) {
    Connect-ExchangeOnline
}

function Get-DLEmailActivityCorrected {
    param(
        [Parameter(Mandatory=$true)]
        [string]$DistributionListEmail
    )
    
    Write-Host "`n=== Checking Distribution List Email Activity ===" -ForegroundColor Cyan
    Write-Host "DL Email: $DistributionListEmail" -ForegroundColor White
    
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
                $DL = $group  # Use for further processing
            }
            else {
                Write-Host "No distribution list or group found with email: $DistributionListEmail" -ForegroundColor Red
                return
            }
        }
    }
    catch {
        Write-Error "Error retrieving distribution list information: $($_.Exception.Message)"
        return
    }
    
    # Check last 10 days (Get-MessageTrace limitation)
    Write-Host "`n--- Checking Email Activity (Last 10 days) ---" -ForegroundColor Yellow
    
    $EndDate = Get-Date
    $StartDate = $EndDate.AddDays(-9)  # Use 9 days to stay within limit
    
    try {
        # Search for emails RECEIVED by the DL (last 10 days)
        Write-Host "`nSearching for emails RECEIVED by the DL (last 10 days)..." -ForegroundColor Cyan
        
        $ReceivedMessages = Get-MessageTrace -RecipientAddress $DistributionListEmail -StartDate $StartDate -EndDate $EndDate -PageSize 5000 |
                           Sort-Object Received -Descending
        
        if ($ReceivedMessages) {
            $LatestReceived = $ReceivedMessages | Select-Object -First 1
            Write-Host "✓ LATEST EMAIL RECEIVED (Last 10 days):" -ForegroundColor Green
            Write-Host "  Date: $($LatestReceived.Received)" -ForegroundColor White
            Write-Host "  From: $($LatestReceived.SenderAddress)" -ForegroundColor White
            Write-Host "  Subject: $($LatestReceived.Subject)" -ForegroundColor White
            Write-Host "  Status: $($LatestReceived.Status)" -ForegroundColor White
            Write-Host "  Total emails received in last 10 days: $($ReceivedMessages.Count)" -ForegroundColor White
            
            # Show recent activity
            if ($ReceivedMessages.Count -gt 1) {
                Write-Host "`n  Recent received emails:" -ForegroundColor Gray
                $ReceivedMessages | Select-Object -First 5 | ForEach-Object {
                    Write-Host "    $($_.Received) - From: $($_.SenderAddress)" -ForegroundColor Gray
                }
            }
        }
        else {
            Write-Host "✗ No emails received in the last 10 days" -ForegroundColor Yellow
        }
        
        # Search for emails SENT from the DL (last 10 days)
        Write-Host "`nSearching for emails SENT from the DL (last 10 days)..." -ForegroundColor Cyan
        
        $SentMessages = Get-MessageTrace -SenderAddress $DistributionListEmail -StartDate $StartDate -EndDate $EndDate -PageSize 5000 |
                       Sort-Object Received -Descending
        
        if ($SentMessages) {
            $LatestSent = $SentMessages | Select-Object -First 1
            Write-Host "✓ LATEST EMAIL SENT (Last 10 days):" -ForegroundColor Green
            Write-Host "  Date: $($LatestSent.Received)" -ForegroundColor White
            Write-Host "  To: $($LatestSent.RecipientAddress)" -ForegroundColor White
            Write-Host "  Subject: $($LatestSent.Subject)" -ForegroundColor White
            Write-Host "  Status: $($LatestSent.Status)" -ForegroundColor White
            Write-Host "  Total emails sent in last 10 days: $($SentMessages.Count)" -ForegroundColor White
        }
        else {
            Write-Host "✗ No emails sent from this DL in the last 10 days" -ForegroundColor Yellow
        }
        
    }
    catch {
        Write-Error "Error checking recent email activity: $($_.Exception.Message)"
    }
    
    # Offer historical search for longer periods
    Write-Host "`n--- Historical Search Option (Beyond 10 days) ---" -ForegroundColor Magenta
    Write-Host "For searches beyond 10 days, we need to use Historical Search." -ForegroundColor White
    
    $DoHistoricalSearch = Read-Host "Do you want to start a historical search (10-90 days ago)? (Y/N)"
    
    if ($DoHistoricalSearch -eq "Y" -or $DoHistoricalSearch -eq "y") {
        Start-HistoricalEmailSearch -DistributionListEmail $DistributionListEmail
    }
}

function Start-HistoricalEmailSearch {
    param(
        [Parameter(Mandatory=$true)]
        [string]$DistributionListEmail
    )
    
    Write-Host "`n--- Starting Historical Email Search ---" -ForegroundColor Cyan
    
    try {
        # Calculate date range (11-90 days ago to avoid overlap with recent search)
        $EndDate = (Get-Date).AddDays(-10)
        $StartDate = (Get-Date).AddDays(-90)
        
        # Start historical search for received emails
        Write-Host "Starting historical search for emails RECEIVED by $DistributionListEmail..." -ForegroundColor Yellow
        Write-Host "Date range: $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))" -ForegroundColor Gray
        
        $SearchName = "DL-Received-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
        
        $HistoricalSearch = Start-HistoricalSearch -ReportTitle $SearchName -StartDate $StartDate -EndDate $EndDate -ReportType MessageTrace -RecipientAddress $DistributionListEmail
        
        if ($HistoricalSearch) {
            Write-Host "✓ Historical search started successfully!" -ForegroundColor Green
            Write-Host "Search ID: $($HistoricalSearch.JobId)" -ForegroundColor White
            Write-Host "Status: $($HistoricalSearch.Status)" -ForegroundColor White
            
            # Wait for completion and check status
            Write-Host "`nWaiting for search to complete..." -ForegroundColor Yellow
            do {
                Start-Sleep -Seconds 10
                $SearchStatus = Get-HistoricalSearch -JobId $HistoricalSearch.JobId
                Write-Host "Current status: $($SearchStatus.Status)" -ForegroundColor Gray
            } while ($SearchStatus.Status -eq "InProgress")
            
            if ($SearchStatus.Status -eq "Done") {
                Write-Host "✓ Historical search completed!" -ForegroundColor Green
                
                # Get the results
                $Results = Get-HistoricalSearch -JobId $HistoricalSearch.JobId -ReportLocation
                
                if ($Results -and $Results.FileUrl) {
                    Write-Host "✓ Results are available for download:" -ForegroundColor Green
                    Write-Host "Download URL: $($Results.FileUrl)" -ForegroundColor White
                    Write-Host "`nYou can download the CSV file to see all email activity in the 90-day period." -ForegroundColor Yellow
                }
                else {
                    Write-Host "No historical email activity found for this distribution list." -ForegroundColor Yellow
                }
            }
            else {
                Write-Host "Historical search failed or was not completed. Status: $($SearchStatus.Status)" -ForegroundColor Red
            }
        }
        
    }
    catch {
        Write-Warning "Historical search failed: $($_.Exception.Message)"
        Write-Host "Note: Historical search requires Exchange Online Plan 2 or Office 365 E3/E5 licensing." -ForegroundColor Yellow
    }
}

function Get-DLMemberActivity {
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
            Write-Host "Member list is large ($($Members.Count) members). First 10 members:" -ForegroundColor Gray
            $Members | Select-Object -First 10 | ForEach-Object {
                Write-Host "  $($_.DisplayName) ($($_.PrimarySmtpAddress)) - $($_.RecipientType)" -ForegroundColor White
            }
        }
        
    }
    catch {
        Write-Warning "Could not retrieve member list: $($_.Exception.Message)"
    }
}

# Main execution
$DistributionListEmail = "kelly@hostway.com"

Write-Host "=== Distribution List Email Activity Report (Corrected) ===" -ForegroundColor Magenta
Write-Host "Target: $DistributionListEmail" -ForegroundColor White
Write-Host "Started: $(Get-Date)" -ForegroundColor Gray

# Run the corrected analysis
Get-DLEmailActivityCorrected -DistributionListEmail $DistributionListEmail

# Show members
$ShowMembers = Read-Host "`nDo you want to see the distribution list members? (Y/N)"
if ($ShowMembers -eq "Y" -or $ShowMembers -eq "y") {
    Get-DLMemberActivity -DistributionListEmail $DistributionListEmail
}

Write-Host "`n=== SUMMARY ===" -ForegroundColor Magenta
Write-Host "✓ Checked recent email activity (last 10 days)" -ForegroundColor Green
Write-Host "✓ Distribution list details retrieved" -ForegroundColor Green
Write-Host "Note: For complete historical data beyond 10 days, use the Historical Search option." -ForegroundColor Yellow

Write-Host "`nScript completed at: $(Get-Date)" -ForegroundColor Gray
