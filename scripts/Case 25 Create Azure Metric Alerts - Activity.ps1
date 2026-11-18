
[CmdletBinding()]
param (
    [string]$SubscriptionId = "518ad9ab-1a06-4ded-a39b-8651f6b0b319",
    [string]$ResourceGroup = "ABXMaintenanceDev",
    [string]$DataFactoryName = "USAEU2VELMANDFDEV",

    [Parameter(Mandatory=$true)]
    [string]$CsvPath = "C:\Users\User\Desktop\Workspace\ANALYTICS\PowerShell\Finales\AlertsForActivity.csv"
)

# Ensure Azure connection
try {
    # Disconnect first to ensure clean state (silently continue if not connected)
    Disconnect-AzAccount -ErrorAction SilentlyContinue
    
    # Connect to Azure (will prompt if not already authenticated)
    Connect-AzAccount -ErrorAction Stop
    
    # Get access token
    $token = (Get-AzAccessToken -ResourceUrl "https://management.azure.com").Token
}
catch {
    Write-Error "Failed to authenticate with Azure: $_"
    exit 1
}

# Prepare headers
$headers = @{
    Authorization = "Bearer $token"
    "Content-Type" = "application/json"
}

$csvData = Import-Csv -Path $CsvPath

foreach ($row in $csvData) {
    $AlertRuleName = $row.AlertRuleName
    $Description = $row.Description
    $Severity = $row.Severity
    $Enable = $row.Enable
    $EvaluationFrequency = $row.EvaluationFrequency
    $WindowSize = $row.WindowSize
    $Threshold = $row.Threshold
    $ActivityType = $row.ActivityType -split ','
    $ActivityName = $row.ActivityName -split ','
    $PipelineName = $row.PipelineName -split ','
    $FailureTypes = $row.FailureTypes -split ','
    $Operator = $row.Operator
    $TimeAggregation = $row.TimeAggregation
    $ActionGroupsName = $row.ActionGroupsName

    # Construct URIs and scope
    $uri = "https://management.azure.com/subscriptions/$($SubscriptionId)/resourceGroups/$($ResourceGroup)/providers/Microsoft.Insights/metricAlerts/$($AlertRuleName)?api-version=2018-03-01"
    $scope = "/subscriptions/$($SubscriptionId)/resourceGroups/$($ResourceGroup)/providers/Microsoft.DataFactory/factories/$($DataFactoryName)"
    $actionGroupId = "/subscriptions/$($SubscriptionId)/resourceGroups/$($ResourceGroup)/providers/microsoft.insights/actionGroups/$($ActionGroupsName)"

    # Build the request body
    $body = @{
        location = "global"
        tags = @{}
        properties = @{
            description = $Description 
            severity = $Severity 
            enabled = $Enable 
            scopes = @($scope) 
            evaluationFrequency = $EvaluationFrequency 
            windowSize = $WindowSize
            criteria = @{
                "odata.type" = "Microsoft.Azure.Monitor.SingleResourceMultipleMetricCriteria"
                allOf = @(
                    @{
                        criterionType = "StaticThresholdCriterion"
                        name = "ActivityFailedRuns"
                        metricNamespace = "Microsoft.DataFactory/factories"
                        metricName = "ActivityFailedRuns"
                        dimensions = @(
                            @{
                                name = "ActivityType"
                                operator = "Include"
                                values = $ActivityType
                            }
                            @{
                                name = "Name"
                                operator = "Include"
                                values = $ActivityName
                            }
                            @{
                                name = "PipelineName"
                                operator = "Include"
                                values = $PipelineName
                            }
                            @{
                                name = "FailureType"
                                operator = "Include"
                                values = $FailureTypes
                            }
                        )
                        operator = $Operator
                        threshold = $Threshold
                        timeAggregation = $TimeAggregation
                    }
                )
            }
            actions = @(
                @{
                    actionGroupId = $actionGroupId
                    webHookProperties = @{}
                }
            )
        }
    } | ConvertTo-Json -Depth 10

    try {
        # Execute the request
        Invoke-RestMethod -Uri $uri -Method Put -Headers $headers -Body $body -ContentType "application/json" -ErrorAction Stop
        Write-Host "Metric alert '$AlertRuleName' created/updated successfully!" -ForegroundColor Green
    }
    catch {
        $errorDetails = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
        if ($errorDetails) {
            Write-Error "Failed to create/update alert. Error: $($errorDetails.error.message)"
        }
        else {
            Write-Error "Failed to create/update alert. Error: $_"
        }
    }

}




































Disconnect-AzAccount
Connect-AzAccount

$token = (Get-AzAccessToken -ResourceUrl "https://management.azure.com").Token

$subscriptionId = "518ad9ab-1a06-4ded-a39b-8651f6b0b319"
$resourceGroup = "ABXMaintenanceDev"
$dataFactoryName = "USAEU2VELMANDFDEV"
$actionGroupsName = "Monitoring"
$alertRuleName = "DemoRuleName01"

$headers = @{
    Authorization = "Bearer $token"
    "Content-Type" = "application/json"

}

$uri = "https://management.azure.com/subscriptions/$($subscriptionId)/resourceGroups/$($resourceGroup)/providers/Microsoft.Insights/metricAlerts/$($alertRuleName)?api-version=2018-03-01"
$scope = "/subscriptions/$($subscriptionId)/resourceGroups/$($resourceGroup)/providers/Microsoft.DataFactory/factories/$($dataFactoryName)"
$actionGroupId = "/subscriptions/$($subscriptionId)/resourceGroups/$($resourceGroup)/providers/microsoft.insights/actionGroups/$($actionGroupsName)"

$body = @{
    location = "global"
    tags = @{}
    properties = @{
        description = "Test Rule"
        severity = 2
        enabled = $true
        scopes = @($scope)
        evaluationFrequency = "PT1M"
        windowSize = "PT1M"
        criteria = @{
            "odata.type" = "Microsoft.Azure.Monitor.SingleResourceMultipleMetricCriteria"
            allOf = @(
                @{
                    criterionType = "StaticThresholdCriterion"
                    name = "PipelineFailedRuns"
                    metricName = "PipelineFailedRuns"
                    dimensions = @(
                        @{
                            name = "FailureType"
                            operator = "Include"
                            values = @("UserError", "SystemError", "BadGateway")
                        }
                        @{
                            name = "Name"
                            operator = "Include"
                            values = @("Monitoring")
                        }
                    )
                    operator = "GreaterThan"
                    threshold = 0
                    timeAggregation = "Total"
                }
            )
        }
        actions = @(
            @{
                actionGroupId = $actionGroupId
                webHookProperties = @{}
            }
        )
    }
} | ConvertTo-Json -Depth 10

# Execute the request
$response = Invoke-RestMethod -Uri $uri -Method Put -Headers $headers -Body $body -ContentType "application/json"

# Check response
if ($response.StatusCode -in 200, 201) {
    Write-Host "Metric alert created successfully!"
} else {
    Write-Host "Failed to create alert. Status code: $($response.StatusCode)"
    $response.Content
}