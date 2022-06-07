﻿function New-WordList {
    [CmdletBinding()]
    param(
        [ScriptBlock] $ListItems,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [int] $BehaviourOption = 0,
        [alias('ListType')][Xceed.Document.NET.ListItemType] $Type = [Xceed.Document.NET.ListItemType]::Bulleted,
        [bool] $Supress = $true
    )

    if ($ListItems) {
        [Array] $Parameters = Invoke-Command -ScriptBlock $ListItems
        if ($Parameters.Count -gt 0) {
            $List = $null
            foreach ($Item in $Parameters) {
                if ($null -eq $List) {
                    $List = $WordDocument.AddList($Item.Text, $Item.Level, $Type, $Item.StartNumber, $Item.TrackChanges, $Item.ContinueNumbering)
                    $Paragraph = $List.Items[$List.Items.Count - 1]
                } else {
                    $List = $WordDocument.AddListItem($List, $Item.Text, $Item.Level, $Type, $Item.StartNumber, $Item.TrackChanges, $Item.ContinueNumbering)
                    $Paragraph = $List.Items[$List.Items.Count - 1]
                }
            }
            Add-WordListItem -WordDocument $WordDocument -List $List -Supress $true
            if (-not $Supress) {
                $List
            }
        } else {
            Write-Warning 'New-WordList - Empty list provided. Skipping.'
        }
    }
}