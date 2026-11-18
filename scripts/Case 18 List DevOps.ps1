<#
.SYNOPSIS
Automates repository operations including branch creation, code modification, and pull request creation.
#>

#region Initialization
Write-Output "`n=== Initializing Script Parameters ==="

$org = "abxclient"
$project = "Veladero Vision Link"
$repositoryId = "adfrepo"
$token = "************************************"
$headers = @{ Authorization = "Basic $([Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$token")))" }

# Generate unique branch name
$new_guid = (New-Guid).Guid.Split("-")[0]
$current_date = Get-Date -Format "yyyyMMdd"
$targetBranch = "API_$($current_date)_$($new_guid)"
$baseBranch = "main"

# Path replacements
$oldPath = "/DatabricksRepo/"
$newPath = "/DatabricksRepoUC/"

Write-Output "  Target branch: $targetBranch"
Write-Output "  Path replacement: '$oldPath' -> '$newPath'"
Start-Sleep -Seconds 2
#endregion

#region Azure DevOps Connection
function Get-BranchCommit($branchName) {
    $url = "https://dev.azure.com/$org/$project/_apis/git/repositories/$repositoryId/refs?filter=heads/$branchName&api-version=5.1"
    return (Invoke-RestMethod -Uri $url -Method Get -Headers $headers).value[0].objectId
}

Write-Output "`n=== Connecting to Azure DevOps ==="
$parentCommitId = Get-BranchCommit $baseBranch
Write-Output "  Base branch commit ID: $parentCommitId"
#endregion

#region Branch Creation
Write-Output "`n=== Creating Target Branch ==="

$body = ConvertTo-Json @(@{
    name = "refs/heads/$targetBranch"
    newObjectId = $parentCommitId
    oldObjectId = "0" * 40
})

$url = "https://dev.azure.com/$org/$project/_apis/git/repositories/$repositoryId/refs?api-version=5.1"
$response = Invoke-RestMethod -Uri $url -Headers $headers -Method Post -Body $body -ContentType "application/json"

Write-Output "  Successfully created branch: $targetBranch"
Start-Sleep -Seconds 2
#endregion

#region Azure Data Factory Processing
Write-Output "`n=== Connecting to Azure ==="
Connect-AzAccount
$resourceGroup = "ABXMaintenanceDev"
$dataFactoryName = "USAEU2VELMANDFDEV"

Write-Output "`n=== Processing Data Factory Pipelines ==="
$pipelines = Get-AzDataFactoryV2Pipeline -DataFactoryName $dataFactoryName -ResourceGroupName $resourceGroup
$changesObject = [System.Collections.ArrayList]@()

foreach ($pipeline in $pipelines) {
    Write-Output "  Processing pipeline: $($pipeline.Name)"
    
    $url_items = "https://dev.azure.com/$org/$project/_apis/git/repositories/$repositoryId/items?versionDescriptor.version=$targetBranch&path=/pipeline/$($pipeline.Name).json&api-version=5.1"
    $response_items = Invoke-RestMethod -Uri $url_items -Headers $headers -Method Get -ContentType "application/json; charset=utf-8"

    $modified = $false
    $activities = $response_items.properties.activities

    function Process-Activities {
        param([array]$activities)
        foreach ($activity in $activities) {
            #if($activity.type -eq "SetVariable") {
            #    $activityValue = $activity.typeProperties.value
            #    if ($activityValue -is [string]) {
            #        # Solo modificar si es un string
            #        Write-Output "  [MODIFIED] Activity: $($activity.name) (SetVariable)"
            #        $activity.typeProperties.value = $activityValue -replace $oldPath, $newPath
            #        $script:modified = $true
            #    }
            #    elseif ($activityValue.type -eq "Expression") {
            #        # Si notebookPath es una expresión, modificar su valor interno
            #        Write-Output "  [MODIFIED] Activity: $($activity.name) (SetVariable) - Expression Path"
            #        $activity.typeProperties.value.value = $activityValue.value -replace $oldPath, $newPath
            #        $script:modified = $true
            #    }
            #}
            # If the activity is a DatabricksNotebook, update its notebookPath.
            if ($activity.type -eq "DatabricksNotebook" -and $activity.typeProperties -and $activity.typeProperties.notebookPath) {
                $notebookPath = $activity.typeProperties.notebookPath
                
                # Verificar si notebookPath es una expresión
                if ($notebookPath -is [string]) {
                    # Solo modificar si es un string
                    Write-Output "  [MODIFIED] Activity: $($activity.name) (DatabricksNotebook)"
                    $activity.typeProperties.notebookPath = $notebookPath -replace $oldPath, $newPath
                    $script:modified = $true
                }
                elseif ($notebookPath.type -eq "Expression") {
                    # Si notebookPath es una expresión, modificar su valor interno
                    Write-Output "  [MODIFIED] Activity: $($activity.name) (DatabricksNotebook) - Expression Path"
                    $activity.typeProperties.value.value = $notebookPath.value -replace $oldPath, $newPath
                    $script:modified = $true
                }
            }

            # Recursive processing for nested activities
            if ($activity.activities) { Process-Activities $activity.activities }
            if ($activity.type -eq "IfCondition") {
                if ($activity.typeProperties.ifTrueActivities) { Process-Activities $activity.typeProperties.ifTrueActivities }
                if ($activity.typeProperties.ifFalseActivities) { Process-Activities $activity.typeProperties.ifFalseActivities }
            }
            if ($activity.type -eq "Switch") {
                $activity.typeProperties.cases | ForEach-Object { Process-Activities $_.activities }
                if ($activity.typeProperties.defaultActivities) { Process-Activities $activity.typeProperties.defaultActivities }
            }
            if ($activity.type -eq "ForEach" -and $activity.typeProperties.activities) {
                Process-Activities $activity.typeProperties.activities
            }
        }
    }

    Process-Activities $activities

    if ($modified) {
        Write-Output "    Pipeline modified: $($pipeline.Name)"
        $jsonPayload = $response_items | ConvertTo-Json -Depth 100
        $jsonPayload = [regex]::Replace($jsonPayload, '\\u([0-9a-fA-F]{4})', { [char][Convert]::ToInt16($args[0].Groups[1].Value, 16) })
        
        $changesObject.Add(@{
            changeType = "edit"
            item = @{ path = "/pipeline/$($pipeline.Name).json" }
            newContent = @{
                content = $jsonPayload
                contentType = "rawtext"
            }
        }) | Out-Null
    }
}
#endregion

if ($changesObject.Count -gt 0) {
    #region Commit Changes
    Write-Output "`n=== Committing Changes ==="
    $newCommitId = Get-BranchCommit $targetBranch
    $commit = @{
        comment = "Updating active tasks API, but saving in a new branch."
        changes = $changesObject
    }
    $body = @{
        refUpdates = @(
            @{
                name        = "refs/heads/$($targetBranch)"
                oldObjectId = $newCommitId
            }
        )
        commits = @($commit)
    } | ConvertTo-Json -Depth 10

    $uri = "https://dev.azure.com/$org/$project/_apis/git/repositories/$repositoryId/pushes?api-version=6.0"
    $response = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body -ContentType "application/json"

    Write-Output "  Successfully pushed changes to $targetBranch"
    Start-Sleep -Seconds 2
    #endregion

    #region Pull Request Creation
    Write-Output "`n=== Creating Pull Request ==="
    try {
        $prBody = @{
            sourceRefName = "refs/heads/$targetBranch"
            targetRefName = "refs/heads/main"
            title = "Update Activity Via API"
            description = "Automated path updates for Databricks UC migration"
            completionOptions = @{ deleteSourceBranch = $true }
        } | ConvertTo-Json

        $prUrl = "https://dev.azure.com/$org/$project/_apis/git/repositories/$repositoryId/pullrequests?api-version=5.0"
        $pr = Invoke-RestMethod -Uri $prUrl -Method Post -Headers $headers -Body $prBody -ContentType "application/json"
    
        Write-Output "  PR Created: $($pr.url)"
        Write-Output "  Setting auto-complete parameters..."
    
        $updateBody = @{
            autoCompleteSetBy = @{ id = $pr.createdBy.id }
            completionOptions = @{ 
                deleteSourceBranch = $true
                mergeStrategy = "squash"
            }
        } | ConvertTo-Json

        Invoke-RestMethod -Uri "$($pr.url)?api-version=5.0" -Method Patch -Headers $headers -Body $updateBody -ContentType "application/json"
        Start-Sleep -Seconds 2
    
        Write-Output "`n=== Operation Completed Successfully ==="
        Write-Output "  PR Link: $($pr.url)"
    }
    catch {
        Write-Output "  ERROR: Failed to create pull request"
    }
    #endregion
} else {
    Write-Host "No changes detected"
}

