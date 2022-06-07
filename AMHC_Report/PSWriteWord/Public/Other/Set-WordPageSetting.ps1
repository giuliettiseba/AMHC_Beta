﻿function Set-WordPageSettings {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container]$WordDocument,
        [nullable[single]] $MarginLeft,
        [nullable[single]] $MarginRight,
        [nullable[single]] $MarginTop,
        [nullable[single]] $MarginBottom,
        [nullable[single]] $PageWidth,
        [nullable[single]] $PageHeight,
        [alias ("PageLayout")][nullable[Xceed.Document.NET.Orientation]] $Orientation
    )

    Set-WordMargins -WordDocument $WordDocument -MarginLeft $MarginLeft -MarginRight $MarginRight -MarginTop $MarginTop -MarginBottom $Mar
    Set-WordPageSize -WordDocument $WordDocument -PageWidth $PageWidth -PageHeight $PageHeight
    Set-WordOrientation -WordDocument $WordDocument -Orientation $Orientation
}