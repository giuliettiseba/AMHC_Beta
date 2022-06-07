<#

Pay version of Xceed 1.5 support this:

In Section, setting the page orientation individually for the different sections will now be supported.
In Section, the following properties can now be set to configure each section of the Document: Headers, Footers, DifferentFirstPage, MarginTop, MarginBottom, MarginLeft, MarginRight, MarginHeader, MarginFooter, MirrorMargins, PageWidth, PageHeight, PageBorders, PageLayout.
In Section, the SectionBreakType property will now correctly get/set the Xml and therefore contain the desired value.

Free version (currently at 1.1 of Xceed) doesn't yet. Therefore orientation, page margins etc can only be applied globally.
#>


function Set-WordPageSize {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container]$WordDocument,
        [nullable[single]] $PageWidth,
        [nullable[single]] $PageHeight
    )
    if ($PageWidth -ne $null) {$WordDocument.PageWidth = $PageWidth }
    if ($PageHeight -ne $null) {$WordDocument.PageHeight = $PageHeight }
}
