function New-WordBlock {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline, Mandatory = $true)][Xceed.Document.NET.Container]$WordDocument,
        ### TOC GLOBAL
        [nullable[bool]] $TocGlobalDefinition,
        [string] $TocGlobalTitle,
        [int] $TocGlobalRightTabPos,
        [Xceed.Document.NET.TableOfContentsSwitches[]] $TocGlobalSwitches,

        ### TOC
        [nullable[bool]] $TocEnable,
        [string] $TocText,
        [int] $TocListLevel,
        [nullable[Xceed.Document.NET.ListItemType]] $TocListItemType,
        [nullable[Xceed.Document.NET.HeadingType]] $TocHeadingType,

        ### Paragraphs/PageBreaks
        [int] $EmptyParagraphsBefore,
        [int] $EmptyParagraphsAfter,
        [int] $PageBreaksBefore,
        [int] $PageBreaksAfter,

        ### Text Data
        [string] $Text,
        [string] $TextNoData,
        [nullable[Xceed.Document.NET.Alignment][]] $TextAlignment = [Xceed.Document.NET.Alignment]::Both,

        ### Table Data
        [Object] $TableData,
        [nullable[Xceed.Document.NET.TableDesign]] $TableDesign = [Xceed.Document.NET.TableDesign]::None,
        [nullable[int]] $TableMaximumColumns = 5,
        [nullable[bool]] $TableTitleMerge,
        [string] $TableTitleText,
        [nullable[Xceed.Document.NET.Alignment]] $TableTitleAlignment = 'center',
        [nullable[System.Drawing.KnownColor]] $TableTitleColor = 'Black',
        [switch] $TableTranspose,
        [float[]] $TableColumnWidths,

        ### List Data
        [Object] $ListData,
        [nullable[Xceed.Document.NET.ListItemType]] $ListType,
        [string] $ListTextEmpty,

        ### List Builder
        [string[]] $ListBuilderContent,
        [Xceed.Document.NET.ListItemType[]] $ListBuilderType,
        [int[]] $ListBuilderLevel,

        ### String Based Data - for functions that return String type data
        [Object] $TextBasedData,
        [nullable[Xceed.Document.NET.Alignment][]] $TextBasedDataAlignment = [Xceed.Document.NET.Alignment]::Both,

        ### Chart Data
        [nullable[bool]] $ChartEnable,
        [string] $ChartTitle,
        $ChartKeys,
        $ChartValues,
        [Xceed.Document.NET.ChartLegendPosition] $ChartLegendPosition = [Xceed.Document.NET.ChartLegendPosition]::Bottom,
        [bool] $ChartLegendOverlay
    )
    ### PAGE BREAKS BEFORE
    $WordDocument | New-WordBlockPageBreak -PageBreaks $PageBreaksBefore

    ### TOC GLLOBAL PROCESSING
    if ($TocGlobalDefinition) {
        Add-WordToc -WordDocument $WordDocument -Title $TocGlobalTitle -Switches $TocGlobalSwitches -RightTabPos $TocGlobalRightTabPos -Supress $True
    }

    ### TOC PROCESSING
    if ($TocEnable) {
        $TOC = $WordDocument | Add-WordTocItem -Text $TocText -ListLevel $TocListLevel -ListItemType $TocListItemType -HeadingType $TocHeadingType
    }

    ### EMPTY PARAGRAPHS BEFORE
    $WordDocument | New-WordBlockParagraph -EmptyParagraphs $EmptyParagraphsBefore

    ### TEXT PROCESSING
    if ($Text) {
        if ($TableData -or $ListData -or ($ChartEnable -and ($ChartKeys.Count -gt 0) -or ($ChartValues.Count -gt 0) ) -or $ListBuilderContent -or (-not $TextNoData)) {
            $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text $Text -Alignment $TextAlignment
        } else {
            if ($TextNoData) {
                $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text $TextNoData -Alignment $TextAlignment
            }
        }
    }
    ### TABLE PROCESSING
    if ($TableData -and $TableDesign) {

        if ($TableTitleMerge) {
            $OverwriteTitle = $TableTitleText
        }

        #if ($TableMaximumColumns -eq $null) { $TableMaximumColumns = 5 }
        if ($TableColumnWidths) {
            Add-WordTable -WordDocument $WordDocument -Paragraph $Paragraph -DataTable $TableData -AutoFit Window -Design $TableDesign -DoNotAddTitle:$TableTitleMerge -MaximumColumns $TableMaximumColumns -Transpose:$TableTranspose -ColumnWidth $TableColumnWidths -OverwriteTitle $OverwriteTitle -Supress $True
        } else {
            Add-WordTable -WordDocument $WordDocument -Paragraph $Paragraph -DataTable $TableData -AutoFit Window -Design $TableDesign -DoNotAddTitle:$TableTitleMerge -MaximumColumns $TableMaximumColumns -Transpose:$TableTranspose -OverwriteTitle $OverwriteTitle -Supress $True
        }
        #if ($TableTitleMerge) {
        #    $Table = Set-WordTableRowMergeCells -Table $Table -RowNr 0 -MergeAll  # -ColumnNrStart 0 -ColumnNrEnd 1
        #    if ($TableTitleText -ne $null) {
        #        $TableParagraph = Get-WordTableRow -Table $Table -RowNr 0 -ColumnNr 0
        #        $TableParagraph = Set-WordText -Paragraph $TableParagraph -Text $TableTitleText -Alignment $TableTitleAlignment -Color $TableTitleColor
        #    }
        #}
    }
    ### LIST PROCESSING
    if ($ListData) {
        if ((Get-ObjectCount $ListData) -gt 0) {
            Write-Verbose 'New-WordBlock - Adding ListData'
            $List = Add-WordList -WordDocument $WordDocument -ListType $ListType -Paragraph $Paragraph -ListData $ListData #-Verbose
        } else {
            Write-Verbose 'New-WordBlock - Adding ListData - Empty List'
            $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text $ListTextEmpty
        }
    }

    ### LIST BUILDER PROCESSING
    <#
    if ($ListBuilderContent) {
        $ListDomainInformation = $null
        for ($a = 0; $a -lt $ListBuilderContent.Count; $a++) {
            $ListDomainInformation = $ListDomainInformation | New-WordListItem -WordDocument $WordDocument -ListLevel $ListBuilderLevel[$a] -ListItemType $ListBuilderType[$a] -ListValue $ListBuilderContent[$a]
        }
        $Paragraph = Add-WordListItem -WordDocument $WordDocument -Paragraph $Paragraph -List $ListDomainInformation #-Supress $true
    }
    #>

    if ($ListBuilderContent) {
        $Paragraph = New-WordList -WordDocument $WordDocument -Type $ListBuilderType[0] {
            #$ListDomainInformation = $null
            for ($a = 0; $a -lt $ListBuilderContent.Count; $a++) {
                #$ListDomainInformation = $ListDomainInformation | New-WordListItem -WordDocument $WordDocument -ListLevel $ListBuilderLevel[$a] -ListItemType $ListBuilderType[$a] -ListValue $ListBuilderContent[$a]

                New-WordListItem -ListLevel $ListBuilderLevel[$a] -ListValue $ListBuilderContent[$a]
            }
            # $Paragraph = Add-WordListItem -WordDocument $WordDocument -Paragraph $Paragraph -List $ListDomainInformation #-Supress $true
        } -Supress $False
    }

    ### SIMPLE TEXT PROCESSING - if source is bunch of text this is the way to go
    if ($TextBasedData) {
        $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text $TextBasedData -Alignment $TextBasedDataAlignment
    }

    ### CHART PROCESSING
    if ($ChartEnable) {
        $WordDocument | New-WordBlockParagraph -EmptyParagraphs 1
        if (($ChartKeys.Count -eq 0) -or ($ChartValues.Count -eq 0)) {
            # If chart had no values or keys it would create an empty chart and prevent saving of document in Word
            # Handling this case with TextNoData above
        } else {
            Add-WordPieChart -WordDocument $WordDocument -ChartName $ChartTitle -Names $ChartKeys -Values $ChartValues -ChartLegendPosition $ChartLegendPosition -ChartLegendOverlay $ChartLegendOverlay
        }
    }
    ### EMPTY PARAGRAPHS AFTER
    $WordDocument | New-WordBlockParagraph -EmptyParagraphs $EmptyParagraphsAfter

    ### PAGE BREAKS AFTER
    $WordDocument | New-WordBlockPageBreak -PageBreaks $PageBreaksAfter
}