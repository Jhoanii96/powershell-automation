<#
.SYNOPSIS
    Retrieves all Power BI report subscriptions across the tenant using Admin APIs.
    
.DESCRIPTION
    This script connects to Power BI Service, enumerates all reports via the Admin API,
    and retrieves subscriptions for each report. Results are exported to CSV with
    enriched metadata including report name and workspace ID.
    
.NOTES
    Prerequisites:
    - Power BI Management Module: Install-Module -Name MicrosoftPowerBIMgmt
    - Power BI Administrator permissions required for /admin/ endpoints
    - Ensure you have appropriate licensing for admin operations
#>

function Format-UserToString {
    param (
        [Parameter(Mandatory=$true)]
        [object]$UserObject
    )
    
    if ($null -eq $UserObject) {
        return "N/A"
    }
    
    $properties = @()
    
    # Define the properties we want to extract in order
    $propertyNames = @('emailAddress', 'displayName', 'identifier', 'graphId', 'principalType', 'userType')
    
    foreach ($prop in $propertyNames) {
        if ($UserObject.PSObject.Properties.Name -contains $prop -and $UserObject.$prop) {
            $properties += "$prop=$($UserObject.$prop)"
        }
    }
    
    return $properties -join " | "
}
function Format-UsersArrayToString {
    param (
        [Parameter(Mandatory=$true)]
        [array]$UsersArray
    )
    
    if ($null -eq $UsersArray -or $UsersArray.Count -eq 0) {
        return "N/A"
    }
    
    $formattedUsers = @()
    
    foreach ($user in $UsersArray) {
        $formattedUsers += (Format-UserToString -UserObject $user)
    }
    
    return $formattedUsers -join " ; "
}

# Initialize collection for all subscription data
$allSubscriptions = [System.Collections.ArrayList]::new()

# Initialize counters for progress tracking
$processedCount = 0
$successCount = 0
$errorCount = 0

# =============================================================================
# SECTION 3: RETRIEVE ALL REPORTS (ADMIN SCOPE)
# =============================================================================

Write-Host "`n--- Retrieving all reports from tenant (Admin API) ---" -ForegroundColor Cyan


# Invoke Admin API to get all reports across the tenant
$reportsResponse = Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/admin/reports" -Method Get | ConvertFrom-Json

$totalReports = $reportsResponse.value.Count
Write-Host "Found $totalReports reports to process" -ForegroundColor Green


# =============================================================================
# SECTION 4: ITERATE THROUGH REPORTS & RETRIEVE SUBSCRIPTIONS
# =============================================================================

Write-Host "`n--- Processing individual report subscriptions ---" -ForegroundColor Cyan

$baseAdminUrl = "https://api.powerbi.com/v1.0/myorg/admin"
$subscriptionsEndpointTemplate = "$baseAdminUrl/reports/{0}/subscriptions"

foreach ($report in $reportsResponse.value) {
    $processedCount++
    $reportId = $report.id
    $reportName = $report.name
    $workspaceId = $report.workspaceId
    
    # Display progress
    Write-Host "Retrieving Subscriptions - Report $($processedCount) of $($totalReports) : $($reportName)"

    $subscriptionsUrl = $subscriptionsEndpointTemplate -f $reportId
    
    try {
        # Build the subscriptions URL for this specific report
        #$subscriptionsUrl = "https://api.powerbi.com/v1.0/myorg/admin/reports/$($reportId)/subscriptions"
        # Call the Subscriptions API for this report
        $rawResponse = Invoke-PowerBIRestMethod -Url $subscriptionsUrl -Method Get -ErrorAction Stop

        if ([string]::IsNullOrWhiteSpace($rawResponse)) {
            throw "Empty response received - possible rate limiting or API error"
        }
        
        $subscriptionsResponse = $rawResponse | ConvertFrom-Json
        
        # Process each subscription found for this report
        if ($subscriptionsResponse.value -and $subscriptionsResponse.value.Count -gt 0) {
            foreach ($subscription in $subscriptionsResponse.value) {
                $subscriptionRow = [PSCustomObject]@{
                    # Original subscription properties
                    Id = $subscription.id
                    Title = $subscription.title
                    ArtifactId = $subscription.artifactId
                    ArtifactDisplayName = $subscription.artifactDisplayName
                    SubArtifactDisplayName = $subscription.subArtifactDisplayName
                    ArtifactType = $subscription.artifactType
                    IsEnabled = $subscription.isEnabled
                    Frequency = $subscription.frequency
                    StartDate = $subscription.startDate
                    EndDate = $subscription.endDate
                    LinkToContent = $subscription.linkToContent
                    PreviewImage = $subscription.previewImage
                    AttachmentFormat = $subscription.attachmentFormat
                
                    # Formatted owner field
                    owner = Format-UserToString -UserObject $subscription.owner
                
                    # Formatted users field (array of users)
                    users = Format-UsersArrayToString -UsersArray $subscription.users
                
                    # Enriched report context
                    ReportId = $reportId
                    ReportName = $reportName
                    WorkspaceId = $workspaceId
                    RetrievedAt = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                }
            
                # Add to master collection
                [void]$allSubscriptions.Add($subscriptionRow)
            }
            $successCount++
            Write-Host "[$processedCount/$totalReports] Found $($subscriptionsResponse.value.Count) subscription(s) for: $reportName" -ForegroundColor Green
        } else {
            Write-Host "[$processedCount/$totalReports] No subscriptions for: $reportName" -ForegroundColor DarkGray
        }
        
    } catch {
        $errorCount++ 
        # CRITICAL: Check for Rate Limit (429) error - STOP EXECUTION IMMEDIATELY
        if ($_.Exception.Message -like "*429*" -or 
            $_.Exception.Response.StatusCode -eq 429 -or
            $_.FullyQualifiedErrorId -like "*429*") {
            
            Write-Host "`n" + ("=" * 60) -ForegroundColor Red
            Write-Host "CRITICAL: RATE LIMIT (429) DETECTED" -ForegroundColor Red
            Write-Host ("=" * 60) -ForegroundColor Red
            Write-Host "Stopping execution to prevent further throttling." -ForegroundColor Yellow
            Write-Host "Partial results collected up to report #$processedCount will be saved." -ForegroundColor Yellow
            
            $rateLimitHit = $true
            break  # Exit foreach loop immediately
        }
        
        # Distinguish between permission errors and other failures
        if ($_.Exception.Message -like "*401*" -or $_.Exception.Message -like "*403*") {
            Write-Warning "[$processedCount/$totalReports] Permission denied for Report '$reportName' (ID: $reportId) in Workspace: $workspaceId"
        } else {
            Write-Warning "[$processedCount/$totalReports] Failed to retrieve subscriptions for '$reportName' (ID: $reportId). Error: $($_.Exception.Message)"
        }
    }
}

# Clear progress bar
Write-Progress -Activity "Retrieving Subscriptions" -Completed

# =============================================================================
# SECTION 5: DATA EXPORT & SUMMARY
# =============================================================================

Write-Host "`n--- Export Results ---" -ForegroundColor Cyan

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupPath = "C:\ps1\workspaces\PowerBI_Report_Subscriptions_$timestamp.csv"

if ($allSubscriptions.Count -gt 0) {
    try {
        # Export to primary location
        #$reportsResponse.value | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8 -Force
        $allSubscriptions | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8 -Force
        
        # Also create timestamped backup
        #$allSubscriptions | Export-Csv -Path $backupPath -NoTypeInformation -Encoding UTF8 -Force
        
        Write-Host "SUCCESS: Exported $($allSubscriptions.Count) subscriptions to:" -ForegroundColor Green
        Write-Host "  Primary: $outputPath" -ForegroundColor White
        Write-Host "  Backup:  $backupPath" -ForegroundColor White
        
        # Display sample of retrieved data
        Write-Host "`nSample of retrieved subscriptions (first 5):" -ForegroundColor Cyan
        $allSubscriptions | Select-Object -First 5 | 
            Format-Table ReportName, WorkspaceId, id, title, frequency -AutoSize
        
    } catch {
        Write-Error "Failed to export CSV: $_"
    }
} else {
    Write-Host "No subscriptions found across all $totalReports reports." -ForegroundColor Yellow
}

# =============================================================================
# SECTION 6: EXECUTION SUMMARY
# =============================================================================

Write-Host "`n=== Execution Summary ===" -ForegroundColor Cyan
Write-Host "Total Reports Processed:    $totalReports" -ForegroundColor White
Write-Host "Reports with Subscriptions: $successCount" -ForegroundColor Green
Write-Host "Errors Encountered:         $errorCount" -ForegroundColor $(if($errorCount -gt 0){"Red"}else{"Green"})
Write-Host "Total Subscriptions Found:  $($allSubscriptions.Count)" -ForegroundColor $(if($allSubscriptions.Count -gt 0){"Green"}else{"Yellow"})

# Optional: Disconnect session (uncomment if needed)
# Disconnect-PowerBIServiceAccount
# Write-Host "`nDisconnected from Power BI Service" -ForegroundColor Gray



