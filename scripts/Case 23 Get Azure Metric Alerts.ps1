<#
.SYNOPSIS
    Exports Azure Metric Alerts to a CSV file.
.DESCRIPTION
    This script retrieves all metric alerts from a specified Azure subscription and resource group,
    then exports them to a CSV file with timestamp in the filename.
.PARAMETER Route
    The base directory path where the CSV file will be saved.
.PARAMETER FileName
    The base name for the output CSV file (without extension).
.PARAMETER SubscriptionId
    The Azure subscription ID to query for metric alerts.
.PARAMETER ResourceGroup
    The Azure resource group name to query for metric alerts.
.PARAMETER OutputEncoding
    The encoding for the output CSV file (default: UTF8).
.EXAMPLE
    .\Export-AzureMetricAlerts.ps1 -Route "C:\ps1\ADF" -FileName "AzureMetricAlerts" -SubscriptionId "518ad9ab-1a06-4ded-a39b-8651f6b0b319" -ResourceGroup "ABXMaintenanceDev"
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$Route = "C:\ps1\ADF",
    
    [Parameter(Mandatory = $true)]
    [string]$FileName = "AzureMetricAlerts",
    
    [Parameter(Mandatory = $true)]
    [string]$SubscriptionId = "518ad9ab-1a06-4ded-a39b-8651f6b0b319",
    
    [Parameter(Mandatory = $true)]
    [string]$ResourceGroup = "ABXMaintenanceDev"
)

# Initialize variables
$ErrorActionPreference = "Stop"

Disconnect-AzAccount
Connect-AzAccount

# Create timestamp for filename
$timestamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
$fileNameWithTime = "${FileName}_${timestamp}.csv"
$fullPath = Join-Path -Path $Route -ChildPath $fileNameWithTime

# Ensure the output directory exists
if (-not (Test-Path -Path $Route)) {
    New-Item -ItemType Directory -Path $Route | Out-Null
}

# Get Azure access token
Write-Host "Acquiring Azure access token..."
$token = (Get-AzAccessToken -ResourceUrl "https://management.azure.com").Token

$headers = @{
    Authorization = "Bearer $token"
}

# Build the API URI
$uri = "https://management.azure.com/subscriptions/$($SubscriptionId)/resourceGroups/$($ResourceGroup)/providers/Microsoft.Insights/metricAlerts?api-version=2018-03-01"

# Get metric alerts
Write-Host "Retrieving metric alerts from subscription $SubscriptionId, resource group $ResourceGroup..."
$response = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers

# Process alerts
Write-Host "Processing $($response.value.Count) metric alerts..."
$alerts = foreach ($alert in $response.value) {
    # Process criteria
    $criteriaStrings = @()
    foreach ($criteria in $alert.properties.criteria.allOf) {
        $criteriaStrings += "Whenever $($criteria.metricName) is $($criteria.operator) $($criteria.threshold)"
    }
    $criterias = $criteriaStrings -join " | "

    # Process action groups
    $actionGroupStrings = @()
    foreach ($actionGroup in $alert.properties.actions.actionGroupId) {
        try {
            $actionUri = "https://management.azure.com$($actionGroup)?api-version=2018-03-01"
            $actionResponse = Invoke-RestMethod -Uri $actionUri -Method Get -Headers $headers
            
            $emails = if ($actionResponse.properties.emailReceivers) {
                $actionResponse.properties.emailReceivers.emailAddress -join ", "
            } else {
                "No email receivers configured"
            }
            
            $actionGroupStrings += "Action Group: $($actionResponse.name) - User Group: $emails"
        } catch {
            Write-Warning "Failed to retrieve action group details for $($actionGroup): $_"
            $actionGroupStrings += "Action Group: Error retrieving details"
        }
    }
    $emailAddress = $actionGroupStrings -join " | "

    # Create output object
    [PSCustomObject]@{
        Name         = $alert.name
        Type         = $alert.type
        Location     = $alert.location
        Description  = $alert.properties.description
        Severity     = $alert.properties.severity
        Enabled      = $alert.properties.enabled
        Criteria     = $criterias
        ActionEmails = $emailAddress
        Tags         = if ($alert.tags) { $alert.tags -join '; ' } else { $null }
    }
}

# Export to CSV
Write-Host "Exporting data to $fullPath..."
$alerts | Export-Csv -Path $fullPath -NoTypeInformation -Encoding UTF8
Write-Host "CSV file created successfully at: $fullPath"



