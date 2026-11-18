Connect-AzAccount

$csvData = Import-Csv -Path 'C:\Users\gopro\Downloads\StartUp.csv' -Delimiter ','

# Filter records with Failed status
$failedPipelines = $csvData | Where-Object { $_.Status -eq 'Failed' }

# Define parameters
$subscriptionId = "1236c889-0d22-41e7-b25f-4a254679be49"
$dataFactory = "GLBADFOPSBIPROD"
$resourceGroup = "CACC-RG-OPSBI-PROD"

$token = (Get-AzAccessToken -ResourceUrl "https://management.azure.com").Token
$headers = @{
    Authorization = "Bearer $token"
    "Content-Type" = "application/json"
}

$body = @{
    lastUpdatedAfter = "2025-01-01"
    lastUpdatedBefore = "2025-03-13"
} | ConvertTo-Json -Depth 100 -Compress


# Iterate through failed pipeline runs
$pipelinesResultsObject = @()

function Process-PipelineRun {
    param (
        [object]$Run,
        [string]$TriggerName
    )

    try {
        # Get activity runs for current pipeline run
        $uri = "https://management.azure.com/subscriptions/$($subscriptionId)/resourceGroups/$($resourceGroup)/providers/Microsoft.DataFactory/factories/$($dataFactory)/pipelineruns/$($Run.RunId)/queryActivityruns?api-version=2018-06-01"
        $response = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body
        $failedResponse = $response.value | Where-Object { $_.status -eq 'Failed' }

        foreach ($item in $failedResponse) {
            # Add result to collection
            $pipelinesResultsObject += [PSCustomObject]@{
                RunId             = $Run.RunId
                TriggerName       = $TriggerName
                PipelineName      = $Run.PipelineName
                StartTime         = $Run.Start
                EndTime           = $Run.End
                Duration_min      = $Run.Duration_min
                ActivityName      = $item.activityName
                ActivityRunStart  = $item.activityRunStart
                ActivityRunEnd    = $item.activityRunEnd
                Status            = $item.status
                Message           = $item.error.message
            }

            # Recursive processing for ExecutePipeline activities
            if ($item.activityType -eq 'ExecutePipeline' -and $item.output.pipelineRunId) {
                # Create synthetic run object for child pipeline
                $childRun = [PSCustomObject]@{
                    RunId        = $Run.RunId
                    PipelineName = $item.output.pipelineName
                    RunStart     = $item.activityRunStart
                    RunEnd       = $item.activityRunEnd
                }

                # Recursive call with child pipeline details
                Process-PipelineRun -Run $childRun -TriggerName $TriggerName
            }
        }
    }
    catch {
        Write-Warning "Error processing RunId $($Run.RunId): $_"
    }
}

# Process initial failed pipeline runs
foreach ($run in $failedPipelines) {
    Process-PipelineRun -Run $run -TriggerName $run.PredecessorsName
}


# Create output directory if it doesn't exist
$outputDir = "C:\ps1\ADF"
if (-not (Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# Generate timestamped filename
$currentDateTime = Get-Date -Format "yyyyMMdd_HHmmss"
$csvFilePath = Join-Path -Path $outputDir -ChildPath "PipelinesFailedMessage_DataFactory_$currentDateTime.csv"

# Export results to CSV
$pipelinesResultsObject | Select-Object * | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8

Write-Host "Results exported to: $csvFilePath"

#Get-AzDataFactoryV2PipelineRun -DataFactoryName GLBADFOPSBIPROD -PipelineRunId 5b0ae153-5e6b-4e64-9d59-407df177ca0d -ResourceGroupName CACC-RG-OPSBI-PROD

#$results = Get-AzDataFactoryV2ActivityRun -DataFactoryName GLBADFOPSBIPROD -PipelineRunId 98ce6abf-5749-4573-8dc0-6af8a6be4c36 -ResourceGroupName CACC-RG-OPSBI-PROD -RunStartedAfter 2025-01-01 -RunStartedBefore 2025-03-13 



#$results | ConvertTo-Json -Depth 100





<#
$uri = "https://management.azure.com/subscriptions/$($subscriptionId)/resourceGroups/CACC-RG-OPSBI-PROD/providers/Microsoft.DataFactory/factories/GLBADFOPSBIPROD/pipelineruns/$($runId)/queryActivityruns?api-version=2018-06-01"
$response = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body
$failedResponse = $response.value | Where-Object { $_.status -eq 'Failed' }
$failedResponse[3].executionDetails.integrationRuntime
#>


