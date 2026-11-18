

Connect-AzAccount
Connect-AzAccount -DeviceCode 


$token = (Get-AzAccessToken -ResourceUrl "https://management.azure.com").Token

$subscriptionId = "518ad9ab-1a06-4ded-a39b-8651f6b0b319"
$resourceGroup = "ABXMaintenanceProd"
$dataFactoryName = "USAEU2VELMAN-PRODUC-ADF"
$apiVersion = "2018-06-01"

$headers = @{
    Authorization = "Bearer $token"
    "Content-Type" = "application/json"
}

# Get all pipelines
$pipelines = Get-AzDataFactoryV2Pipeline -DataFactoryName $dataFactoryName -ResourceGroupName $resourceGroup

$count_pipelines = 0
$count_activities = 0
foreach ($pipeline in $pipelines) {
    Write-Output "`n================ Processing Pipeline: $($pipeline.Name) ================"
    $pipeline_uri = "https://management.azure.com/subscriptions/$($subscriptionId)/resourceGroups/$($resourceGroup)/providers/Microsoft.DataFactory/factories/$($dataFactoryName)/pipelines/$($pipeline.Name)?api-version=$($apiVersion)"
    $pipeline_response = Invoke-RestMethod -Uri $pipeline_uri -Method Get -Headers $headers
    $script:count_pipelines = $script:count_pipelines + 1
    $activities = $pipeline_response.properties.activities
    
    function Process-Activities {
        param([array]$activities)
        foreach ($activity in $activities) {
            if($activity.type -eq "SetVariable") {
                Write-Output "  Activity: $($activity.name) (SetVariable)"
                Write-Output "  Activity: $($activity.name) $($activity.typeProperties.notebookPath) (SetVariable)"
                #$activity.typeProperties.value = $activity.typeProperties.value -replace $oldPath, $newPath 
                #$script:modified = $true
                $script:count_activities = $script:count_activities + 1
            }
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
                elseif ($notebookPath -is [hashtable] -and $notebookPath["type"] -eq "Expression") {
                    # Si notebookPath es una expresión, modificar su valor interno
                    Write-Output "  [MODIFIED] Activity: $($activity.name) (DatabricksNotebook) - Expression Path"
                    $activity.typeProperties.notebookPath["value"] = $notebookPath["value"] -replace $oldPath, $newPath
                    $script:modified = $true
                }
            }

            # Process any direct child activities.
            if ($activity.activities) {
                Process-Activities $activity.activities
            }

            # Process IfCondition activities (look inside ifTrueActivities and ifFalseActivities).
            if ($activity.type -eq "IfCondition" -and $activity.typeProperties) {
                if ($activity.typeProperties.ifTrueActivities) {
                    Process-Activities $activity.typeProperties.ifTrueActivities
                }
                if ($activity.typeProperties.ifFalseActivities) {
                    Process-Activities $activity.typeProperties.ifFalseActivities
                }
            }

            # Process Switch activities (look inside each case and defaultActivities).
            if ($activity.type -eq "Switch" -and $activity.typeProperties) {
                if ($activity.typeProperties.cases) {
                    foreach ($case in $activity.typeProperties.cases) {
                        if ($case.activities) {
                            Process-Activities $case.activities
                        }
                    }
                }
                if ($activity.typeProperties.defaultActivities) {
                    Process-Activities $activity.typeProperties.defaultActivities
                }
            }

            # Process ForEach activities (activities are nested in typeProperties.activities).
            if ($activity.type -eq "ForEach" -and $activity.typeProperties) {
                if ($activity.typeProperties.activities) {
                    Process-Activities $activity.typeProperties.activities
                }
            }
        }
    }
    
    Process-Activities $activities
}

Write-Output "Total Pipelines: $($count_pipelines)"
Write-Output "Total Activities: $($count_activities)"
