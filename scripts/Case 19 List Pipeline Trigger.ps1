

Connect-AzAccount
# Retrieve pipelines and triggers
$resourceGroup = "ABXMaintenanceDev"
$dataFactory = "USAEU2VELMANDFDEV"
$pipelines = Get-AzDataFactoryV2Pipeline -DataFactoryName $dataFactory -ResourceGroupName $resourceGroup
$triggers = Get-AzDataFactoryV2Trigger -DataFactoryName $dataFactory -ResourceGroupName $resourceGroup
$pipelines[0].Parameters
# Iterate through each pipeline and find linked triggers
$count_pipelines = 0
foreach ($pipeline in $pipelines) {
    $script:count_pipelines = $script:count_pipelines + 1
    $pipelineName = $pipeline.Name
    $linkedTriggers = $triggers | Where-Object {
        # Check if the trigger's "Pipelines" property references this pipeline
        $_.Properties.Pipelines.PipelineReference.ReferenceName -contains $pipelineName
    }

    # Output pipeline name and its associated triggers
    Write-Output "Pipeline: $pipelineName"
    if ($linkedTriggers) {
        $linkedTriggers | ForEach-Object {
            Write-Output "  -> Trigger: $($_.Name) (State: $($_.RuntimeState))"
        }
    } else {
        Write-Output "  -> No triggers linked to this pipeline."
    }
    Write-Output "------------------------"
}
Write-Output "Total Pipelines: $($count_pipelines)"

























# Login to Azure
Connect-AzAccount
# Define your resource group and data factories
$resourceGroup1 = "ABXMaintenanceDev"
$dataFactory1 = "USAEU2VELMANDFDEV"
$resourceGroup2 = "ABXMaintenanceProd"
$dataFactory2 = "USAEU2VELMAN-PROD-ADF"

# Retrieve pipelines from both data factories
$triggers1 = Get-AzDataFactoryV2Trigger -DataFactoryName $dataFactory1 -ResourceGroupName $resourceGroup1
$triggers2 = Get-AzDataFactoryV2Trigger -DataFactoryName $dataFactory2 -ResourceGroupName $resourceGroup2
#$triggers1[1].Properties.Pipelines.Parameters
function Get-RecurrenceDetails {
    param (
        [Parameter(Mandatory=$true)]
        [Microsoft.Azure.Commands.DataFactoryV2.Models.PSTrigger]$Trigger
    )

    $recurrence = $Trigger.Properties.Recurrence
    $pipelineParams = $Trigger.Properties.Pipelines | ForEach-Object { $_.Parameters } | Where-Object { $_ }
    if (-not $recurrence) {
        return [PSCustomObject]@{
            Frequency    = $null
            Interval     = $null
            StartTime    = $null
            EndTime      = $null
            TimeZone     = $null
            Hours        = $null
            Minutes      = $null
            RuntimeState = $null
            Parameters   = $null
        }
    }

    [PSCustomObject]@{
        Frequency    = $recurrence.Frequency
        Interval     = $recurrence.Interval
        StartTime    = $recurrence.StartTime
        EndTime      = $recurrence.EndTime
        TimeZone     = $recurrence.TimeZone
        Hours        = if ($recurrence.Schedule.Hours) { $recurrence.Schedule.Hours -join ', ' } else { $null }
        Minutes      = if ($recurrence.Schedule.Minutes) { $recurrence.Schedule.Minutes -join ', ' } else { $null }
        RuntimeState = $Trigger.RuntimeState
        Parameters   = if ($pipelineParams) { $pipelineParams | ConvertTo-Json -Depth 10 -Compress } else { $null }
    }
}

$results = @()
$properties = @('Frequency', 'Interval', 'StartTime', 'EndTime', 'TimeZone', 'Hours', 'Minutes', 'RuntimeState', 'Parameters')

# Compare each trigger from Data Factory 1
foreach ($trigger1 in $triggers1) {
    $name = $trigger1.Name
    $trigger2 = $triggers2 | Where-Object { $_.Name -ceq $name }

    if ($trigger2) {
        $recurrence1 = Get-RecurrenceDetails $trigger1
        $recurrence2 = Get-RecurrenceDetails $trigger2

        $match = $true
        foreach ($prop in $properties) {
            if ($recurrence1.$prop -ne $recurrence2.$prop) {
                $match = $false
                break
            }
        }

        if ($match) {
            $results += [PSCustomObject]@{
                TriggerName = $name
                Status      = 'Match'
                Details     = $recurrence1
            }
        } else {
            $results += [PSCustomObject]@{
                TriggerName = $name
                Status      = 'Mismatch'
                Details1    = $recurrence1
                Details2    = $recurrence2
            }
        }
    } else {
        $results += [PSCustomObject]@{
            TriggerName = $name
            Status      = 'OnlyInDataFactory1'
        }
    }
}

# Check for triggers only in Data Factory 2
foreach ($trigger2 in $triggers2) {
    $name = $trigger2.Name
    if ($name -cnotin $triggers1.Name) {
        $results += [PSCustomObject]@{
            TriggerName = $name
            Status      = 'OnlyInDataFactory2'
        }
    }
}

# Display results
foreach ($result in $results) {
    switch ($result.Status) {
        #'Match' {
        #    Write-Host "======================================================================="
        #    Write-Host "Trigger '$($result.TriggerName)' matches. Recurrence Details:"
        #    Write-Host "======================================================================="
        #    $result.Details | Format-List
        #}
        'Mismatch' {
            Write-Host "======================================================================="
            Write-Host "Trigger '$($result.TriggerName)' has mismatched recurrence:"
            Write-Host "======================================================================="
            Write-Host "--- Data Factory 1 ($dataFactory1) ---"
            $result.Details1 | Format-List
            Write-Host "--- Data Factory 2 ($dataFactory2) ---"
            $result.Details2 | Format-List
        }
        'OnlyInDataFactory1' {
            Write-Host "Trigger '$($result.TriggerName)' exists only in Data Factory 1, Status: '$($result.Details1.RuntimeState)' ($dataFactory1)."
        }
        'OnlyInDataFactory2' {
            Write-Host "Trigger '$($result.TriggerName)' exists only in Data Factory 2, Status: '$($result.Details1.RuntimeState)' ($dataFactory2)."
        }
    }
    #Write-Host "`n"
}

<#
# Generate output file
$outputDir = "C:\ps1\ADF"
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}
$currentDateTime = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFilePath = Join-Path -Path $outputDir -ChildPath "Triggers_$currentDateTime.txt"
# Write results to the output file
$results | ForEach-Object {
    $output = "Trigger Name: $($_.TriggerName)`n"
    $output += "Status: $($_.Status)`n"
    switch ($_.Status) {
        'Mismatch' {
            $output += "--- Data Factory 1 ($dataFactory1) ---`n"
            $output += ($_.Details1 | Format-List | Out-String)
            $output += "--- Data Factory 2 ($dataFactory2) ---`n"
            $output += ($_.Details2 | Format-List | Out-String)
        }
        'OnlyInDataFactory1' {
            $output += "Exists only in Data Factory 1 ($dataFactory1).`n"
            $output += "Runtime State: $($_.RuntimeState)`n"
        }
        'OnlyInDataFactory2' {
            $output += "Exists only in Data Factory 2 ($dataFactory2).`n"
            $output += "Runtime State: $($_.RuntimeState)`n"
        }
        'Match' {
            $output += "Details:`n"
            $output += ($_.Details | Format-List | Out-String)
        }
    }
    $output += "`n"
    Add-Content -Path $outputFilePath -Value $output
}

Write-Host "Results have been saved to: $outputFilePath"

#>




$triggers2 = Get-AzDataFactoryV2Trigger -DataFactoryName $dataFactory2 -ResourceGroupName $resourceGroup2
$dataFactory2Results = @()
foreach ($trigger2 in $triggers2) {
    $recurrence = $Trigger2.Properties.Recurrence
    $pipelineReference = $Trigger2.Properties.Pipelines | ForEach-Object { $_.PipelineReference } | Where-Object { $_ }
    $pipelineParams = $Trigger2.Properties.Pipelines | ForEach-Object { $_.Parameters } | Where-Object { $_ }
    $dataFactory2Results += [PSCustomObject]@{
        ResourceGroupName  = $trigger2.ResourceGroupName
        DataFactoryName  = $trigger2.DataFactoryName
        TriggerName  = $trigger2.Name
        Frequency   = $recurrence.Frequency
        Interval    = $recurrence.Interval
        StartTime   = $recurrence.StartTime
        EndTime     = $recurrence.EndTime
        TimeZone    = $recurrence.TimeZone
        Hours       = if ($recurrence.Schedule.Hours) { $recurrence.Schedule.Hours -join ', ' } else { $null }
        Minutes     = if ($recurrence.Schedule.Minutes) { $recurrence.Schedule.Minutes -join ', ' } else { $null }
        RuntimeState = $trigger2.RuntimeState
        PipelineReference  = if ($pipelineReference) { $pipelineReference | ConvertTo-Json -Depth 10 -Compress } else { $null }
        Parameters  = if ($pipelineParams) { $pipelineParams | ConvertTo-Json -Depth 10 -Compress } else { $null }
    }
}

$outputDir = "C:\ps1\ADF"
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}
$currentDateTime = Get-Date -Format "yyyyMMdd_HHmmss"
$csvFilePath = Join-Path -Path $outputDir -ChildPath "Triggers_DataFactory2_$currentDateTime.csv"
$dataFactory2Results | Export-Csv -Path $csvFilePath -NoTypeInformation










$pipelines2 = Get-AzDataFactoryV2Pipeline -DataFactoryName $dataFactory2 -ResourceGroupName $resourceGroup2

$dataFactory2Results = @()
foreach ($pipeline2 in $pipelines2) {
    $dataFactory2Results += [PSCustomObject]@{
        ResourceGroupName  = $pipeline2.ResourceGroupName
        DataFactoryName  = $pipeline2.DataFactoryName
        PipelineName  = $pipeline2.Name
        Parameters  = if ($pipeline2.Parameters) { $pipeline2.Parameters | ConvertTo-Json -Depth 10 -Compress } else { $null }
    }
}

$outputDir = "C:\ps1\ADF"
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}
$currentDateTime = Get-Date -Format "yyyyMMdd_HHmmss"
$csvFilePath = Join-Path -Path $outputDir -ChildPath "Pipelines_Parameters_DataFactory2_$currentDateTime.csv"
$dataFactory2Results | Export-Csv -Path $csvFilePath -NoTypeInformation

