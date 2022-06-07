﻿function New-WordBlockList {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline, Mandatory = $true)][Xceed.Document.NET.Container]$WordDocument,
        [bool] $TocEnable,
        [string] $TocText,
        [int] $TocListLevel,
        [Xceed.Document.NET.ListItemType] $TocListItemType,
        [Xceed.Document.NET.HeadingType] $TocHeadingType,
        [int] $EmptyParagraphsBefore,
        [int] $EmptyParagraphsAfter,
        [string] $Text,
        [string] $TextListEmpty,

        [Object] $ListData,
        [Xceed.Document.NET.ListItemType] $ListType
    )
    if ($TocEnable) {
        $TOC = $WordDocument | Add-WordTocItem -Text $TocText -ListLevel $TocListLevel -ListItemType $TocListItemType -HeadingType $TocHeadingType
    }
    New-WordBlockParagraph -EmptyParagraphs $EmptyParagraphsBefore -WordDocument $WordDocument
    $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text $Text
    if ((Get-ObjectCount $ListData) -gt 0) {
        $List = Add-WordList -WordDocument $WordDocument -ListType $ListType -Paragraph $Paragraph -ListData $ListData #-Verbose
    } else {
        $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text $TextListEmpty
    }
    New-WordBlockParagraph -EmptyParagraphs $EmptyParagraphsAfter -WordDocument $WordDocument

}