
Connect-AzAccount

# Define your resource group and data factories
$resourceGroup1 = "ABXMaintenanceDev"
$dataFactory1 = "USAEU2VELMANDFDEV"
$resourceGroup2 = "ABXMaintenanceProd"
$dataFactory2 = "USAEU2VELMAN-PROD-ADF"

# Retrieve pipelines from both data factories
$pipelines1 = Get-AzDataFactoryV2Pipeline -DataFactoryName $dataFactory1 -ResourceGroupName $resourceGroup1
$pipelines2 = Get-AzDataFactoryV2Pipeline -DataFactoryName $dataFactory2 -ResourceGroupName $resourceGroup2

# Find common pipeline names
$commonPipelines = Compare-Object -ReferenceObject $pipelines1.Name -DifferenceObject $pipelines2.Name -IncludeEqual |
    Where-Object { $_.SideIndicator -eq '==' } |
    Select-Object -ExpandProperty InputObject

if (-not $commonPipelines) {
    Write-Output "No common pipelines found between the data factories"
    return
}

# Compare activity names in common pipelines
foreach ($pipelineName in $commonPipelines) {
    # Get activities from both versions of the pipeline
    $activities1 = ($pipelines1 | Where-Object { $_.Name -eq $pipelineName }).Activities.Name
    $activities2 = ($pipelines2 | Where-Object { $_.Name -eq $pipelineName }).Activities.Name

    # Compare activity names
    $comparison = Compare-Object -ReferenceObject $activities1 -DifferenceObject $activities2 -IncludeEqual

    # Group results
    $matched = $comparison | Where-Object { $_.SideIndicator -eq '==' } | Select-Object -ExpandProperty InputObject
    $onlyInDF1 = $comparison | Where-Object { $_.SideIndicator -eq '<=' } | Select-Object -ExpandProperty InputObject
    $onlyInDF2 = $comparison | Where-Object { $_.SideIndicator -eq '=>' } | Select-Object -ExpandProperty InputObject

    # Display matches
    #if ($matched) {
    #    Write-Output "Matching Activities:"
    #    $matched | ForEach-Object { Write-Output "- $_" }
    #    Write-Output ""
    #}

    # Display unique to DF1
    if ($onlyInDF1) {
        Write-Output "`n==============================================="
        Write-Output "Pipeline: $pipelineName"
        Write-Output "===============================================`n"
        Write-Output "Activities only in $($dataFactory1):"
        $onlyInDF1 | ForEach-Object { Write-Output "- $_" }
        Write-Output ""
    }

    # Display unique to DF2
    if ($onlyInDF2) {
        Write-Output "`n==============================================="
        Write-Output "Pipeline: $pipelineName"
        Write-Output "===============================================`n"
        Write-Output "Activities only in $($dataFactory2):"
        $onlyInDF2 | ForEach-Object { Write-Output "- $_" }
        Write-Output ""
    }
}

$df1Names = $pipelines1.Name
$df2Names = $pipelines2.Name

# Compare pipeline names from both data factories
$compareResult = Compare-Object -ReferenceObject $df1Names -DifferenceObject $df2Names -IncludeEqual

$unmatched = $compareResult | Where-Object { $_.SideIndicator -ne '==' } | ForEach-Object {
    [PSCustomObject]@{
        Name = $_.InputObject
        DataFactory = if ($_.SideIndicator -eq '<=') { $dataFactory1 } else { $dataFactory2 }
    }
}

# Output non-matching pipelines
Write-Output "`n=============================================================================================="
Write-Output "`nNon-Matching Pipelines:"
$unmatched | ForEach-Object { Write-Output "- $($_.Name) (Found in: $($_.DataFactory))" }






