
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputPath = "C:\ps1\workspaces\PowerBI_Report_Subscriptions_$timestamp.csv"

# Initialize collection for all subscription data
$allSubscriptions = [System.Collections.ArrayList]::new()

# Initialize counters for progress tracking
$processedCount = 0
$successCount = 0
$errorCount = 0

Write-Host "`n--- Retrieving all reports from tenant (Admin API) ---" -ForegroundColor Cyan

$reportsResponse = Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/admin/reports" -Method Get | ConvertFrom-Json

$totalReports = $reportsResponse.value.Count
Write-Host "Found $($totalReports) reports to process" -ForegroundColor Green

foreach ($report in $reportsResponse.value) {
    $processedCount++
    $reportId = $report.id
    $reportName = $report.name
    $workspaceId = $report.workspaceId
    
    # Display progress
    Write-Host "Retrieving Subscriptions - Report $($processedCount) of $($totalReports) : $($reportName)"

    try {
        # Build the subscriptions URL for this specific report
        $subscriptionsUrl = "https://api.powerbi.com/v1.0/myorg/admin/reports/$($reportId)/subscriptions"
        # Call the Subscriptions API for this report
        $subscriptionsResponse = Invoke-PowerBIRestMethod -Url $subscriptionsUrl -Method Get -ErrorAction Stop | ConvertFrom-Json

        # Process each subscription found for this report
        if ($subscriptionsResponse.value -and $subscriptionsResponse.value.Count -gt 0) {
            foreach ($subscription in $subscriptionsResponse.value) {
                # Format owner as pipe-separated string inline
                $ownerString = if ($subscription.owner) { 
                    ($subscription.owner.PSObject.Properties | Where-Object { $_.Value } | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join " | " 
                } else { "N/A" }

                # Format users array as semicolon-separated pipe-separated strings inline
                $usersString = if ($subscription.users) { 
                    ($subscription.users | ForEach-Object { 
                        ($_.PSObject.Properties | Where-Object { $_.Value } | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join " | " 
                    }) -join " ; " 
                } else { "N/A" }

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
                    Owner = $ownerString
                
                    # Formatted users field (array of users)
                    Users = $usersString
                
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
        # Distinguish between permission errors and other failures
        if ($_.Exception.Message -like "*401*" -or $_.Exception.Message -like "*403*") {
            Write-Warning "[$processedCount/$totalReports] Permission denied for Report '$reportName' (ID: $reportId) in Workspace: $workspaceId"
        } else {
            Write-Warning "[$processedCount/$totalReports] Failed to retrieve subscriptions for '$reportName' (ID: $reportId). Error: $($_.Exception.Message)"
        }
    }
}

Write-Host "`n--- Export Results ---" -ForegroundColor Cyan

if ($allSubscriptions.Count -gt 0) {
    try {
        # Export to primary location
        $allSubscriptions | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8 -Force
        
        Write-Host "SUCCESS: Exported $($allSubscriptions.Count) subscriptions to:" -ForegroundColor Green
        Write-Host "  $outputPath" -ForegroundColor White
    } catch {
        Write-Error "Failed to export CSV: $_"
    }
} else {
    Write-Host "No subscriptions found across all $totalReports reports." -ForegroundColor Yellow
}
