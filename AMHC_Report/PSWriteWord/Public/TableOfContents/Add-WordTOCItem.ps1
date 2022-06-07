﻿function Add-WordTocItem {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container] $WordDocument,
        [alias('Level')] [ValidateRange(0, 8)] [int] $ListLevel,
        [alias('ListType')][Xceed.Document.NET.ListItemType] $ListItemType = [Xceed.Document.NET.ListItemType]::Bulleted,
        [alias('Value', 'ListValue')]$Text,
        [alias ("HT")] [Xceed.Document.NET.HeadingType] $HeadingType = [Xceed.Document.NET.HeadingType]::Heading1,
        [nullable[int]] $StartNumber,
        [bool]$TrackChanges = $false,
        [bool]$ContinueNumbering = $true,
        [bool]$Supress = $false
    )
    $List = New-WordListItemInternal -WordDocument $WordDocument -List $null -Text $Text -ListItemType $ListItemType -ContinueNumbering $ContinueNumbering -ListLevel $ListLevel -StartNumber $StartNumber -TrackChanges $TrackChanges
    $List = Add-WordListItem -WordDocument $WordDocument -List $List
    $Paragraph = Convert-ListToHeadings -WordDocument $WordDocument -List $List -HeadingType $HeadingType
    if ($Supress) { return } else { return $Paragraph }
}

